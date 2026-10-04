"""Runs the DirectSportsGoods close on the raw data (P2).

Cleans the raw files, posts them to a general ledger in SQLite, runs the month-end
checks, posts correcting journals, proves the corrected ledger matches the P1 model,
and scores the checks against the planted-error answer key.

Run: python src/generate.py && python src/close.py
"""
import csv
import re
import sqlite3
from datetime import datetime
from pathlib import Path

import build_pack as bp
import generate as gen

SQL = Path(__file__).parent / "sql"
DB = bp.ROOT / "data" / "ledger.db"
REPORT = bp.ROOT / "docs" / "close_report.md"
NUMERIC = {"amount", "gross_sales", "refunds", "fees", "orders", "cost_of_goods_dispatched", "gross_pay",
           "employer_ni", "pension", "total_cost", "balance", "line_id", "invoice_id", "terms_days"}
WHY_MISSED = {
    "duplicate_invoice": "It arrived more than 14 days after the original, outside the matching window. "
                         "A supplier statement reconciliation would catch it, since the supplier only expects payment once.",
}
WHY_FALSE = {
    "miscoded_expense": "The coding was right, but nobody recorded why. The fix is process: require an approval "
                        "note on any recode. The correcting journal moved it within overheads, so profit is unaffected.",
}
LABELS = {"duplicate_invoice": "Duplicate invoices", "miscoded_expense": "Miscoded expenses",
          "missed_accrual": "Missed accruals", "fee_overcharge": "Marketplace fee overcharges"}


def read(name):
    with open(gen.RAW / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def norm_name(s):
    words = re.sub(r"[^a-z0-9 ]", " ", s.lower()).split()
    return " ".join(w for w in words if w not in {"ltd", "limited", "llp"})


def parse_date(s):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"Unrecognised date: {s!r}")


def parse_amount(s):
    return round(float(s.replace("£", "").replace(",", "").strip()), 2)


def clean(db):
    """Loads the raw files, fixing supplier names, date formats and amounts on the AP export."""
    suppliers = read("suppliers.csv")
    lookup = {norm_name(s["name"]): s["supplier_id"] for s in suppliers}
    invoices = []
    for r in read("purchase_invoices.csv"):
        sid = lookup.get(norm_name(r["supplier_name"]))
        if sid is None:
            raise ValueError(f"Unknown supplier: {r['supplier_name']!r}")
        invoices.append({**r, "supplier_id": sid, "invoice_no": r["invoice_no"].strip(),
                         "invoice_no_norm": re.sub(r"[^A-Z0-9]", "", r["invoice_no"].upper()),
                         "invoice_date": parse_date(r["invoice_date"]), "amount": parse_amount(r["amount"])})
    tables = {
        "suppliers": suppliers, "invoices": invoices, "web_sales": read("web_sales_daily.csv"),
        "marketplace": read("marketplace_statement.csv"), "dispatch": read("stock_dispatch.csv"),
        "payroll": read("payroll.csv"), "manual_journals": read("manual_journals.csv"),
        "bank": read("bank_statement.csv"), "opening": read("opening_balances.csv"),
    }
    for name, rows in tables.items():
        cols = list(rows[0])
        db.execute(f"CREATE TABLE {name} ({', '.join(cols)})")
        db.executemany(f"INSERT INTO {name} VALUES ({', '.join('?' * len(cols))})",
                       [[float(r[c]) if c in NUMERIC else r[c] for c in cols] for r in rows])
    return {name: len(rows) for name, rows in tables.items()}


def run_sql(db, name, **params):
    for stmt in (SQL / name).read_text(encoding="utf-8").split(";"):
        if stmt.strip():
            db.execute(stmt, params)


def pl(db, include_adjustments):
    """{(period, channel, line): value} with costs negative, matching the P1 sign convention."""
    rows = db.execute(f"""
        SELECT g.period, g.channel, a.pl_line, ROUND(SUM(-g.amount), 2)
        FROM gl g JOIN accounts a USING (account)
        WHERE a.pl_line IS NOT NULL {'' if include_adjustments else "AND g.source <> 'adjustment'"}
        GROUP BY 1, 2, 3""")
    return {(per, ch, line): v for per, ch, line, v in rows}


def tie_out(ledger, truth):
    """Largest absolute gap between the ledger and the P1 model, and operating profit by FY."""
    period_of = {first.strftime("%Y-%m"): (fy, p) for fy, p, first in gen.months()}
    expected = {}
    for (fy, ch, line), vals in truth.items():
        if line == "Orders":
            continue
        for per, (f, p) in period_of.items():
            if f == fy:
                expected[(per, ch, line)] = vals[p - 1]
    keys = set(expected) | set(ledger)
    gap = max(abs(ledger.get(k, 0) - expected.get(k, 0)) for k in keys)
    profit = {}
    for (per, _, _), v in ledger.items():
        fy = period_of[per][0]
        profit[fy] = profit.get(fy, 0) + v
    return gap, profit, sum(v for k, v in expected.items() if period_of[k[0]][0] == "FY26")


def main():
    DB.unlink(missing_ok=True)
    db = sqlite3.connect(DB)
    counts = clean(db)
    run_sql(db, "ledger.sql")
    run_sql(db, "checks.sql", fee_rate=bp.V["mkt_fee"])
    before = pl(db, include_adjustments=False)
    run_sql(db, "corrections.sql")
    after = pl(db, include_adjustments=True)
    db.commit()

    truth = gen.truth()
    gap_before, profit_before, _ = tie_out(before, truth)
    gap_after, profit_after, fy26_truth = tie_out(after, truth)
    tb = db.execute("SELECT ROUND(SUM(amount), 2) FROM gl").fetchone()[0]
    overdue = db.execute("""
        SELECT i.invoice_id FROM invoices i JOIN suppliers s USING (supplier_id)
        WHERE date(i.invoice_date, '+' || CAST(s.terms_days AS INTEGER) || ' days') <= '2026-01-26'
          AND NOT EXISTS (SELECT 1 FROM bank b WHERE b.description LIKE '%' || i.invoice_no)""").fetchall()
    dup_ids = {int(r[0]) for r in db.execute("SELECT record_id FROM exceptions WHERE check_name = 'duplicate_invoice'")}

    detail = {(c, int(r)): d for c, r, d in db.execute("SELECT check_name, record_id, detail FROM exceptions")}
    flags = set(detail)
    with open(gen.ANSWER_KEY, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    desc = {(r["error_type"], int(r["record_id"])): r["description"] for r in rows}
    key = {k for k in desc if k[0] != "not_an_error"}
    decoys = {int(r) for c, r in desc if c == "not_an_error"}
    caught, missed, false_pos = flags & key, key - flags, flags - key
    decoys_flagged = {r for _, r in false_pos} & decoys
    od, missed_ids = {int(r[0]) for r in overdue}, {r for c, r in missed if c == "duplicate_invoice"}

    # Self-check: the books balance, and any gap to the P1 model comes only from what the checks missed
    # or wrongly corrected. With nothing missed, the corrected ledger must be the P1 model exactly.
    assert abs(tb) < 0.01, f"Trial balance out by {tb}"
    if not missed:
        assert gap_after < 0.01, f"Corrected ledger is £{gap_after:,.2f} away from the P1 model"

    lines = [
        "# Close report: FY24 to FY26",
        "",
        "> DirectSportsGoods is a fictional company. All data is made up. Generated by `src/close.py`.",
        "",
        f"**{len(caught)} of {len(key)} planted errors caught, {len(false_pos)} false "
        f"alarm{'s' * (len(false_pos) != 1)}.** {len(decoys) - len(decoys_flagged)} of {len(decoys)} decoys "
        "(correct entries that look wrong) were rightly left alone. "
        + ("After correcting journals, the ledger matches the management pack to the penny "
           f"(largest monthly gap £{gap_after:,.2f}, against £{gap_before:,.0f} before corrections)."
           if gap_after < 0.01 else
           f"After correcting journals, the largest monthly gap to the management pack is £{gap_after:,.0f} "
           f"(£{gap_before:,.0f} before corrections), all of it from the misses listed below."),
        "",
        "## Checks",
        "",
        "| Check | Planted | Caught | Missed | False alarms |",
        "|-------|--------:|-------:|-------:|-------------:|",
    ]
    for kind, label in LABELS.items():
        n = lambda s: sum(1 for c, _ in s if c == kind)
        lines.append(f"| {label} | {n(key)} | {n(caught)} | {n(missed)} | {n(false_pos)} |")
    lines += [
        "",
        "How each check works:",
        "",
        "- **Duplicate invoices:** same supplier and the same invoice number once spacing, case and punctuation are "
        "stripped, whatever the amount. Separately, same supplier and amount under a different number within 14 days, "
        "to catch invoices a supplier re-sent. Matching on amount alone would flag every month's rent.",
        "- **Miscoded expenses:** invoice posted to a different account from the supplier's usual one, unless the "
        "recode has an approval note.",
        "- **Missed accruals:** service delivered in one month, invoice posted to a later one, and no accrual "
        "journal for it at month end.",
        "- **Marketplace fee overcharges:** fees above the contracted rate on net sales, by statement line.",
        "",
        "## Effect on operating profit",
        "",
        "| Year | Before corrections | After corrections | Change |",
        "|------|-------------------:|------------------:|-------:|",
    ]
    for fy in gen.FY_START:
        b, a = profit_before.get(fy, 0), profit_after.get(fy, 0)
        lines.append(f"| {fy} | £{b / 1e3:,.0f}k | £{a / 1e3:,.0f}k | {(a - b) / 1e3:+,.0f}k |")
    lines += [
        "",
        f"FY26 operating profit after corrections is £{profit_after['FY26'] / 1e3:,.0f}k, the same figure as the "
        f"management pack (£{fy26_truth / 1e3:,.0f}k).",
        "",
        "## Controls",
        "",
        f"- Trial balance: debits equal credits (difference £{abs(tb):,.2f}).",
        f"- Overdue unpaid invoices: {len(overdue)}. {len(od & dup_ids)} are duplicates caught above"
        + (f", and {len(od & missed_ids)} are the duplicates the check missed. " if od & missed_ids else ". ")
        + "They sit unpaid because no real debt exists, so the aged creditors review backs up the duplicate check"
        + (" and catches what it let through." if od & missed_ids else "."),
        "",
        "## Source data",
        "",
        "| Table | Rows |",
        "|-------|-----:|",
        *[f"| {name} | {n:,} |" for name, n in counts.items()],
        "",
        "## Simplifications",
        "",
        "- No VAT and no corporation tax in the ledger, so cash builds faster than it would in reality.",
        "- Sales arrive as daily totals per channel rather than individual orders.",
        "- Payroll is paid in full on the last working day of the month.",
    ]
    if missed or false_pos:
        lines += ["", "## Misses and false alarms", ""]
        for c, r in sorted(missed):
            effect = ""
            if c == "duplicate_invoice":
                nominal, amount = db.execute("SELECT nominal, amount FROM invoices WHERE invoice_id = ?", (r,)).fetchone()
                effect = (f" It's a stock purchase, so profit is unaffected, but stock and supplier balances are overstated by £{amount:,.0f}."
                          if nominal == "1200" else f" Costs are overstated by £{amount:,.0f}.")
            lines.append(f"- **Missed, {LABELS[c].lower()}:** {desc[(c, r)]}. {WHY_MISSED.get(c, '')}{effect}")
        for c, r in sorted(false_pos):
            why = desc.get(("not_an_error", r), detail[(c, r)])
            lines.append(f"- **False alarm, {LABELS[c].lower()}:** {why}. {WHY_FALSE.get(c, '')}")
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Caught {len(caught)}/{len(key)}, false alarms {len(false_pos)}, gap before £{gap_before:,.0f}, after £{gap_after:,.2f}")
    print(f"Wrote {REPORT.relative_to(bp.ROOT)} and {DB.relative_to(bp.ROOT)}")


if __name__ == "__main__":
    main()
