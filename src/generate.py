"""Generates three years of raw transaction data for DirectSportsGoods (FY24 to FY26).

DirectSportsGoods is fictional and so is all of this data. Monthly totals tie to the
P1 model in build_pack.py, so once the planted errors are corrected the ledger should
reproduce the pack's actuals to the penny. Every planted error is logged in
data/answer_key.csv.

Run: python src/generate.py
"""
import calendar
import csv
import random
from datetime import date, timedelta

import build_pack as bp

RAW = bp.ROOT / "data" / "raw"
ANSWER_KEY = bp.ROOT / "data" / "answer_key.csv"
FY_START = {"FY24": 2023, "FY25": 2024, "FY26": 2025}  # each FY starts 1 Feb of this year
START, END = date(2023, 2, 1), date(2026, 1, 31)
LINES = ["Orders", "Gross sales", "Returns", "Cost of goods sold", "Delivery",
         "Marketplace fees", "Marketing", "Payroll", "Other overheads", "Depreciation"]

# id, name, default nominal, payment terms (days), invoice number format
SUPPLIERS = [
    ("S01", "Apex Athletic Supply Ltd", "1200", 60, "APX-{:05d}"),
    ("S02", "Northfield Sportswear Ltd", "1200", 60, "NS{:06d}"),
    ("S03", "Peak Footwear Distribution Ltd", "1200", 60, "PFD/{:05d}"),
    ("S04", "Velocity Equipment Co Ltd", "1200", 60, "VE-{:05d}"),
    ("S05", "Stride Apparel Ltd", "1200", 45, "SA{:05d}"),
    ("S06", "Summit Outdoor Goods Ltd", "1200", 45, "SOG-{:05d}"),
    ("S10", "ParcelCo Ltd", "6000", 14, "PC-{:06d}"),
    ("S20", "Searchly Ads Ltd", "6200", 30, "SRCH-{:05d}"),
    ("S21", "Socialreach Media Ltd", "6200", 30, "SRM{:05d}"),
    ("S22", "Linkpartner Affiliates Ltd", "6200", 30, "LPA-{:05d}"),
    ("S30", "Midshire Logistics Park Ltd", "7100", 30, "MLP{:04d}"),
    ("S31", "Cloudstack Software Ltd", "7110", 30, "CS-{:05d}"),
    ("S32", "Kestrel Grid Energy Ltd", "7120", 30, "KGE{:06d}"),
    ("S33", "Ashcombe & Rowe LLP", "7130", 30, "AR-{:04d}"),
    ("S34", "Harbourline Insurance Ltd", "7140", 30, "HI{:05d}"),
    ("S40", "Fulcrum Automation Ltd", "1500", 30, "FA-{:05d}"),
    ("S41", "Brightdesk IT Ltd", "1500", 30, "BIT{:05d}"),
]
SUP = {s[0]: s for s in SUPPLIERS}
STOCK = ["S01", "S02", "S03", "S04", "S05", "S06"]
MARKETING = [("S20", 0.6), ("S21", 0.3), ("S22", 0.1)]
OVERHEADS = [("S30", 0.40), ("S31", 0.22), ("S32", 0.12), ("S33", 0.16), ("S34", 0.10)]
CAPEX = ["S40", "S41"]
OPENING = {"bank": 2_500_000.00, "fixed_assets": 1_500_000.00}


def months():
    """(fy, period 1-12, first day of month) for all 36 months."""
    for fy, y in FY_START.items():
        for p in range(12):
            m = (1 + p) % 12 + 1
            yield fy, p + 1, date(y + (1 if m == 1 else 0), m, 1)


def month_days(d):
    return [date(d.year, d.month, i) for i in range(1, calendar.monthrange(d.year, d.month)[1] + 1)]


def black_friday(y):
    nov30 = date(y, 11, 30)
    return nov30 - timedelta((nov30.weekday() - 4) % 7)


def day_weights(days):
    w = []
    for d in days:
        x = [1.1, 0.95, 0.95, 0.95, 1.0, 1.0, 1.15][d.weekday()]
        x *= {black_friday(d.year): 3.0, black_friday(d.year) + timedelta(3): 2.0}.get(d, 1)
        x *= {(12, 25): 0.3, (12, 26): 1.8}.get((d.month, d.day), 1)
        w.append(x)
    return [x / sum(w) for x in w]


def split(total, weights):
    """Splits a total into pence-rounded parts that add back exactly."""
    parts = [round(total * w / sum(weights), 2) for w in weights]
    parts[-1] = round(total - sum(parts[:-1]), 2)
    return parts


def truth():
    """Correct monthly figures: {(fy, channel, line): [12 values]}, costs negative as in P1."""
    _, data = bp.build_data()
    t = {}
    for scn, ch, line, per, _, val in data:
        fy = {"Actual": "FY26", "Prior year": "FY25"}.get(scn)
        if fy and per >= 1 and line in LINES:
            t.setdefault((fy, ch, line), [0.0] * 12)[per - 1] = val
    rng = random.Random(24)
    for (fy, ch, line), vals in list(t.items()):
        if fy == "FY25":
            t[("FY24", ch, line)] = vals[:] if line == "Depreciation" else \
                [round(v / (1 + bp.V["growth"]) * rng.gauss(1, 0.02), 2) for v in vals]
    # Fees are contractual, so FY24 fees follow FY24 sales at the contracted rate.
    gross, ret = t[("FY24", "Marketplace", "Gross sales")], t[("FY24", "Marketplace", "Returns")]
    t[("FY24", "Marketplace", "Marketplace fees")] = [-round(bp.V["mkt_fee"] * (g + r), 2) for g, r in zip(gross, ret)]
    return t


def workday(d):
    return d + timedelta(max(0, 7 - d.weekday()) if d.weekday() >= 5 else 0)


def generate():
    rng = random.Random(2027)
    T = truth()
    get = lambda fy, ch, line, p: abs(T[(fy, ch, line)][p - 1])

    web, mkt, dispatch, payroll, journals, invoices = [], [], [], [], [], []
    seq = {s[0]: rng.randint(1000, 9000) for s in SUPPLIERS}

    def invoice(sid, inv_date, amount, period, channel="Central", svc=None, desc=""):
        seq[sid] += 1
        svc_from, svc_to = svc or (inv_date.replace(day=1), inv_date)
        invoices.append({
            "invoice_id": len(invoices) + 1, "supplier_id": sid, "supplier_name": SUP[sid][1],
            "invoice_no": SUP[sid][4].format(seq[sid]), "invoice_date": inv_date,
            "period_posted": period, "nominal": SUP[sid][2], "channel": channel,
            "amount": round(amount, 2), "service_from": svc_from, "service_to": svc_to,
            "description": desc,
        })

    # Stock cover: purchases = COGS + change in stock, with stock at Frasers' stock days on the next 12 months' COGS.
    cogs = [get(fy, "Website", "Cost of goods sold", p) + get(fy, "Marketplace", "Cost of goods sold", p) for fy, p, _ in months()]
    cogs += [c * (1 + bp.V["growth"]) for c in cogs[-12:]]
    stock = [bp.V["inv_days"] / 365 * sum(cogs[k:k + 12]) for k in range(37)]
    OPENING["inventory"] = round(stock[0], 2)

    settlement, s_start = 1, START
    for k, (fy, p, first) in enumerate(months()):
        days = month_days(first)
        w = day_weights(days)
        period = first.strftime("%Y-%m")
        for ch, rows in (("Website", web), ("Marketplace", mkt)):
            g = split(get(fy, ch, "Gross sales", p), w)
            r = split(get(fy, ch, "Returns", p), w)
            o = split(get(fy, ch, "Orders", p), w)
            f = split(get(fy, ch, "Marketplace fees", p), w) if ch == "Marketplace" else None
            for i, d in enumerate(days):
                row = {"date": d, "orders": round(o[i]), "gross_sales": g[i], "refunds": r[i]}
                if f:
                    if (d - s_start).days >= 14:
                        settlement, s_start = settlement + 1, d
                    row = {"line_id": len(mkt) + 1, "settlement_id": f"S-{settlement:04d}", **row, "fees": f[i]}
                rows.append(row)
            dispatch.append({"period": period, "channel": ch, "cost_of_goods_dispatched": get(fy, ch, "Cost of goods sold", p)})

            # Carrier bills each channel's account twice a month. The second invoice arrives in the
            # next month but belongs to this one.
            dl = split(get(fy, ch, "Delivery", p), w)
            for lo, hi in ((0, 15), (15, len(days))):
                svc = (days[lo], days[hi - 1])
                inv_date = days[hi - 1] + timedelta(3)
                invoice("S10", inv_date, sum(dl[lo:hi]), period, ch, svc, f"Parcel delivery, {ch.lower()} account")

        mk = split(get(fy, "Website", "Marketing", p), [s for _, s in MARKETING])
        for (sid, _), amt in zip(MARKETING, mk):
            invoice(sid, days[-1], amt, period, "Website", desc="Advertising")
        oh = split(get(fy, "Central", "Other overheads", p), [s for _, s in OVERHEADS])
        for (sid, _), amt in zip(OVERHEADS, oh):
            invoice(sid, days[0], amt, period, desc="Monthly charge")
        invoice(CAPEX[k % 2], days[rng.randint(5, 20)], bp.V["capex"], period, desc="Warehouse and IT equipment")

        purchases = cogs[k] + stock[k + 1] - stock[k]
        n = rng.randint(20, 30)
        for amt, d in zip(split(purchases, [rng.random() + 0.2 for _ in range(n)]), sorted(rng.choices(days, k=n))):
            invoice(rng.choice(STOCK), d, amt, period, desc="Stock purchase")

        pay = split(get(fy, "Central", "Payroll", p), [0.45, 0.20, 0.35])
        for dept, total in zip(["Warehouse", "Customer service", "Head office"], pay):
            ni, pen = round(total * 0.09, 2), round(total * 0.03, 2)
            payroll.append({"period": period, "department": dept, "gross_pay": round(total - ni - pen, 2),
                            "employer_ni": ni, "pension": pen, "total_cost": total})
        journals.append({"period": period, "journal": "Depreciation", "debit": "7500", "credit": "1510",
                         "amount": get(fy, "Central", "Depreciation", p), "channel": "Central", "ref": ""})

    # ---------- planted errors ----------
    key = []

    def plant(kind, record_type, record_id, period, amount, desc):
        key.append({"error_id": len(key) + 1, "error_type": kind, "record_type": record_type,
                    "record_id": record_id, "period": period, "amount": round(amount, 2), "description": desc})

    originals = list(invoices)
    candidates = [i for i in originals if i["supplier_id"] in STOCK + [s for s, _ in OVERHEADS]]
    for n, inv in enumerate(rng.sample(candidates, 16)):
        dup = {**inv, "invoice_id": len(invoices) + 1}
        sid = inv["supplier_id"]
        if n < 10:  # keyed twice, number formatted differently
            dup["invoice_date"] += timedelta(rng.randint(0, 6))
            dup["invoice_no"] = rng.choice([inv["invoice_no"], inv["invoice_no"].replace("-", ""), " " + inv["invoice_no"].lower()])
            how = "entered twice"
        elif n < 13:  # keyed twice, with a typo in the amount
            dup["invoice_date"] += timedelta(rng.randint(0, 6))
            dup["amount"] = round(inv["amount"] + rng.choice([-1, 1]) * rng.uniform(5, 90), 2)
            how = f"entered twice, the second time as £{dup['amount']:,.2f}"
        else:  # supplier re-sent it under a new number
            seq[sid] += 1
            dup["invoice_no"] = SUP[sid][4].format(seq[sid])
            dup["invoice_date"] += timedelta(rng.randint(3, 30))
            how = f"re-sent as {dup['invoice_no']} and entered again"
        dup["invoice_date"] = min(dup["invoice_date"], END)
        dup["period_posted"] = dup["invoice_date"].strftime("%Y-%m")
        dup["service_from"], dup["service_to"] = dup["invoice_date"].replace(day=1), dup["invoice_date"]
        invoices.append(dup)
        plant("duplicate_invoice", "invoice", dup["invoice_id"], dup["period_posted"], dup["amount"],
              f"Invoice {inv['invoice_no']} from {inv['supplier_name']} {how}")

    wrong = {"6200": "7110", "7110": "6200", "7130": "6200", "1500": "7150"}
    for inv in rng.sample([i for i in originals if i["nominal"] in wrong], 10):
        old = inv["nominal"]
        inv["nominal"] = wrong[old]
        plant("miscoded_expense", "invoice", inv["invoice_id"], inv["period_posted"], inv["amount"],
              f"{inv['supplier_name']} coded to {wrong[old]} instead of {old}")

    # Decoys: correct entries a careless check would flag. Logged as not_an_error.
    for inv in rng.sample([i for i in originals if i["supplier_id"] == "S32"], 2):
        inv["nominal"], inv["description"] = "7150", "Approved recode: emergency generator repair"
        plant("not_an_error", "invoice", inv["invoice_id"], inv["period_posted"], inv["amount"],
              "Energy supplier invoice for a generator repair, recoded to repairs with an approval note")
    inv = rng.choice([i for i in originals if i["supplier_id"] == "S34"])
    inv["nominal"], inv["description"] = "7130", "Broker advisory fee"
    plant("not_an_error", "invoice", inv["invoice_id"], inv["period_posted"], inv["amount"],
          "Insurer invoice for broker advice, correctly coded to professional fees but with no approval note")

    late = [i for i in originals if i["supplier_id"] == "S10" and i["invoice_date"].month != i["service_to"].month and i["invoice_date"] <= END]
    picks = rng.sample(late, 11)
    for inv in picks[:8]:
        inv["period_posted"] = inv["invoice_date"].strftime("%Y-%m")
        plant("missed_accrual", "invoice", inv["invoice_id"], inv["service_to"].strftime("%Y-%m"), inv["amount"],
              f"Delivery for {inv['service_from']:%d %b} to {inv['service_to']:%d %b %Y} not accrued, posted in {inv['period_posted']}")
    for inv in picks[8:]:
        inv["period_posted"] = inv["invoice_date"].strftime("%Y-%m")
        svc = inv["service_to"].strftime("%Y-%m")
        for per, name, dr, cr in ((svc, "Delivery accrual", "6000", "2100"), (inv["period_posted"], "Delivery accrual reversal", "2100", "6000")):
            journals.append({"period": per, "journal": name, "debit": dr, "credit": cr, "amount": inv["amount"],
                             "channel": inv["channel"], "ref": inv["invoice_no"]})
        plant("not_an_error", "invoice", inv["invoice_id"], svc, inv["amount"],
              f"Delivery invoice posted in {inv['period_posted']}, but accrued properly in {svc}")

    for row in rng.sample(mkt[:-40], 5):
        extra = round(rng.uniform(300, 2500), 2)
        row["fees"] = round(row["fees"] + extra, 2)
        plant("fee_overcharge", "marketplace_line", row["line_id"], row["date"].strftime("%Y-%m"), extra,
              f"Marketplace charged £{extra:,.2f} above the contracted rate on {row['date']:%d %b %Y} ({row['settlement_id']})")

    # ---------- bank statement ----------
    bank = []
    for row in web:
        bank.append((row["date"] + timedelta(1), f"CARDPAY SETTLEMENT {row['date']:%Y%m%d}", round(row["gross_sales"] - row["refunds"], 2)))
    payouts = {}
    for row in mkt:
        s = payouts.setdefault(row["settlement_id"], [row["date"], 0.0])
        s[0] = row["date"]
        s[1] += row["gross_sales"] - row["refunds"] - row["fees"]
    for sid, (last, amt) in payouts.items():
        bank.append((last + timedelta(3), f"MKTPLACE PAYOUT {sid}", round(amt, 2)))
    for inv in originals:
        bank.append((workday(inv["invoice_date"] + timedelta(SUP[inv["supplier_id"]][3])),
                     f"BACS {inv['supplier_name'].upper()} {inv['invoice_no']}", -inv["amount"]))
    for per in sorted({r["period"] for r in payroll}):
        y, m = map(int, per.split("-"))
        last = date(y, m, calendar.monthrange(y, m)[1])
        bank.append((last - timedelta(max(0, last.weekday() - 4)), f"PAYROLL {per}",
                     -round(sum(r["total_cost"] for r in payroll if r["period"] == per), 2)))
    bank = sorted(b for b in bank if b[0] <= END)
    bal, statement = OPENING["bank"], []
    for d, desc, amt in bank:
        bal = round(bal + amt, 2)
        statement.append({"date": d, "description": desc, "amount": amt, "balance": bal})

    # ---------- write, with the mess a real AP export has ----------
    def messy(inv):
        name = inv["supplier_name"]
        name = rng.choice([name, name.upper(), name.replace(" Ltd", " Limited"), name + " ", name.replace(" Ltd", "")])
        d = inv["invoice_date"]
        return {**inv, "supplier_name": name,
                "invoice_date": d.isoformat() if rng.random() < 0.7 else d.strftime("%d/%m/%Y"),
                "amount": f"£{inv['amount']:,.2f}" if rng.random() < 0.3 else f"{inv['amount']:.2f}"}

    RAW.mkdir(parents=True, exist_ok=True)
    files = {
        "web_sales_daily.csv": web,
        "marketplace_statement.csv": mkt,
        "stock_dispatch.csv": dispatch,
        "purchase_invoices.csv": [{k: v for k, v in messy(i).items() if k != "supplier_id"} for i in invoices],
        "payroll.csv": payroll,
        "manual_journals.csv": journals,
        "bank_statement.csv": statement,
        "suppliers.csv": [{"supplier_id": s[0], "name": s[1], "default_nominal": s[2], "terms_days": s[3]} for s in SUPPLIERS],
        "opening_balances.csv": [{"account": a, "amount": v} for a, v in
                                 (("1000", OPENING["bank"]), ("1200", OPENING["inventory"]), ("1500", OPENING["fixed_assets"]))],
    }
    for name, rows in files.items():
        write(RAW / name, rows)
    write(ANSWER_KEY, key)
    return T, files, key


def write(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)


if __name__ == "__main__":
    T, files, key = generate()
    for name, rows in files.items():
        print(f"  {name:28s} {len(rows):>6,} rows")
    print(f"  answer_key.csv               {len(key):>6,} rows ({sum(k['error_type'] != 'not_an_error' for k in key)} errors, the rest decoys)")
