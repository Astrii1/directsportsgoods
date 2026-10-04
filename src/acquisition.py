"""Builds the DirectSportsGoods acquisition case (P4): two targets, one rejected, one bought.

DirectSportsGoods, Fernbrook Racquets and Halvergate Hockey are all fictional, and so is
every figure. Each target gets a data room of raw files. The due diligence below finds
the problems from those files, the same way the P2 close found errors in the ledger.

Run: python src/acquisition.py   (after src/forecast.py's inputs exist, i.e. build_pack works)
"""
import csv
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as col

import build_pack as bp
import forecast as fc
from build_pack import GBP, INPUT, PCT, head_row, header, style_row
from generate import split

ROOT = bp.ROOT
ROOM = ROOT / "data" / "dataroom"
OUT = ROOT / "pack" / "DirectSportsGoods_acquisition.xlsx"
PAPER = ROOT / "docs" / "board_paper.md"
LTM = [date(2025 + (1 + i) // 12, (1 + i) % 12 + 1, 1) for i in range(12)]  # Feb-25 to Jan-26
PRIOR = [d.replace(year=d.year - 1) for d in LTM]

# Our policies and market inputs. Cited where they are facts; the rest are company choices.
POLICY = {
    "hurdle": 0.12,          # board hurdle rate for synergies (company policy)
    "target_rate": 0.15,     # hurdle plus 3 points for a small founder-run business (company policy)
    "terminal_growth": 0.02,
    "synergy_share": 0.30,   # most of the synergy value we will pay away
    "p_renew": 0.5,          # chance the brand agreement survives the change of control
    "tax": bp.V["tax"],
    "bank_rate": 0.0375,     # Bank of England Bank Rate, 17 September 2026
    "margin": 0.030,         # bank's term sheet margin (fictional)
    "loan": 2_500_000,
    "loan_months": 60,
    "min_cash": 2_000_000,   # board's minimum cash policy (company policy)
    "deal_costs": 250_000,   # legal, tax and financial due diligence (fictional quotes)
    "md_market_salary": 120_000,  # recruiter's estimate for a replacement MD (fictional)
}
COMPLETION = 3  # forecast month index: completion on 1 May 2026
EARNOUT_MONTH = 12  # Feb-27, after the brand agreement's renewal date (31 Dec 2026)
HOLDSPORT = {"price": 122.9e6, "revenue": 168.8e6, "op_profit": 26.3e6}  # Frasers AR 2026 p.179-180
XXL = {"price": 68.6e6, "revenue": 474.5e6}  # Frasers AR 2026 p.179-180, loss-making


# ---------- data rooms ----------

def write(path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)


def read(target, name):
    with open(ROOM / target / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def monthly(annual, weights):
    return split(annual, weights)


def fernbrook_room(rng):
    season = [7, 7, 8, 10, 11, 11, 10, 8, 7, 7, 8, 6]  # Feb..Jan: padel and tennis peak in summer
    rev = 10_400_000
    cogs = rev * 0.58
    brand_a_purchases = cogs * 0.35
    roles = [("Managing Director (founder)", 1, 30_000), ("Operations manager", 1, 52_000),
             ("Warehouse operative", 14, 25_400), ("Customer service advisor", 3, 26_500), ("Buyer / marketer", 4, 38_000)]
    staff = sum(n * pay for _, n, pay in roles) * 1.12
    overheads = [("Ecommerce platform and software", 95_000), ("Audit and accountancy", 28_000), ("Insurance", 22_000),
                 ("Payment processing", 208_000), ("Packaging", 151_000), ("Utilities", 64_000), ("Other", None)]
    lines = {"Revenue": rev, "Cost of sales": -cogs, "Supplier rebates": brand_a_purchases * 0.05,
             "Other income": 72_000, "Staff costs": -staff, "Warehouse rent": -180_000,
             "Delivery": -150_000 * 5.40, "Marketing": -rev * 0.11}
    other_total = sum(lines.values()) - 950_000  # overheads that leave reported EBITDA at £950k
    overheads[-1] = ("Other", other_total - sum(v for _, v in overheads[:-1]))
    lines["Other overheads"] = -other_total
    accounts = []
    for months, scale, label in ((PRIOR, 1 / 1.22, "Prior year"), (LTM, 1, "Last 12 months")):
        split_lines = {}
        for k, v in lines.items():
            if k == "Other income":
                v = 11_000 * 12 / 12 if scale != 1 else 12_000
                vals = monthly(v, [1] * 12)
                if scale == 1:
                    vals[7] += 60_000  # Sep-25 consultancy fee
                split_lines[k] = vals
            elif k in ("Staff costs", "Warehouse rent", "Other overheads"):
                split_lines[k] = monthly(v * (scale if k == "Other overheads" else 1), [1] * 12)
            else:
                split_lines[k] = monthly(v * scale, season)
        for i, d in enumerate(months):
            row = {"month": d.strftime("%Y-%m"), "period": label, **{k: split_lines[k][i] for k in lines}}
            row["EBITDA"] = round(sum(split_lines[k][i] for k in lines), 2)
            accounts.append(row)
    other_income = [{"date": d.strftime("%Y-%m-28"), "description": "Affiliate commission", "amount": round((11_000 if d in PRIOR else 12_000) / 12, 2)}
                    for d in PRIOR + LTM]
    other_income.append({"date": "2025-09-19", "description": "Consultancy fee: padel club fit-out advice", "amount": 60_000})

    lines_stock = ["Padel rackets", "Tennis rackets", "Badminton rackets", "Squash rackets", "Strings", "Shoes", "Bags", "Balls and shuttles", "Apparel"]
    old = split(360_000, [6, 2, 1, 1, 0.5, 1.5, 0.5, 0.2, 0.8])     # over 12 months, mostly last season's padel rackets
    mid = split(170_000, [4, 2, 1, 1, 0.5, 1, 0.5, 0.5, 1])
    fresh = split(1_950_000 - 530_000, [5, 4, 2, 1, 1, 2, 1, 1.5, 1.5])
    stock = []
    for i, name in enumerate(lines_stock):
        f1, f2 = split(fresh[i], [0.65, 0.35])
        stock.append({"product_line": name, "age_0_3m": f1, "age_3_6m": f2, "age_6_12m": mid[i], "age_over_12m": old[i],
                      "provision_held": 0})

    cohorts, opened, active, repeat = [], [5200, 7400, 9800, 11600, 15400, 17300, 18300], [900, 1500, 2300, 3200, 5600, 8100, 10400], [0.44, 0.43, 0.42, 0.40, 0.38, 0.31, 0.24]
    for y, o, a, r in zip(range(2019, 2026), opened, active, repeat):
        cohorts.append({"cohort_year": y, "accounts_opened": o, "active_last_12m": a, "repeat_within_12m": r})
    activity = [{"period": "Prior year", "active_customers": 30_400, "orders": 129_000, "revenue": round(rev / 1.22, 2)},
                {"period": "Last 12 months", "active_customers": sum(active), "orders": 150_000, "revenue": rev}]

    suppliers = [("Brand A", 45, 0.35), ("Brand B", 45, 0.14), ("Brand C", 60, 0.10), ("Brand D", 45, 0.09), ("Brand E", 30, 0.08),
                 ("Strings wholesaler", 45, 0.07), ("Shoe distributor", 45, 0.07), ("Apparel supplier", 30, 0.05),
                 ("Packaging supplier", 30, 0.03), ("Balls wholesaler", 45, 0.02)]
    payables = cogs * 70 / 365  # the founder stretched suppliers to about 70 days
    creditors = []
    for (name, terms, share), bal in zip(suppliers, split(payables, [s for *_, s in suppliers])):
        within, late = round(bal * terms / 70, 2), None
        late = round(bal - within, 2)
        creditors.append({"supplier": name, "terms_days": terms, "within_terms": within, "overdue": late, "total": round(bal, 2)})

    contracts = [
        {"counterparty": "Brand A", "type": "Distribution agreement", "share_of_sales": 0.35, "expiry": "2026-12-31",
         "change_of_control": "Brand may terminate on change of control",
         "notes": "5% volume rebate, personal to the current owner, ends on a change of control. Brand A launched its own UK web shop in 2025."},
        {"counterparty": "Unit 4 landlord", "type": "Warehouse lease", "share_of_sales": "", "expiry": "2031-01-31",
         "change_of_control": "None", "notes": "Rent £180,000 a year. Tenant break on 31 January 2027 with a £60,000 break fee."},
        {"counterparty": "German fulfilment centre", "type": "Third-party warehousing", "share_of_sales": "", "expiry": "Rolling",
         "change_of_control": "None", "notes": "Holds stock in Germany since 2022 to speed up EU delivery."},
        {"counterparty": "Ecommerce platform", "type": "Software licence", "share_of_sales": "", "expiry": "Rolling", "change_of_control": "None",
         "notes": "Monthly subscription."},
    ]
    eu, quarters = [], [(y, q) for y in (2022, 2023, 2024, 2025) for q in (1, 2, 3, 4)][1:]
    for (y, q), amt in zip(quarters, split(975_000, [1 + 0.15 * i for i in range(len(quarters))])):
        eu.append({"quarter": f"{y}-Q{q}", "sales_dispatched_from_germany_gross": amt, "german_vat_accounted": 0})

    payroll = [{"role": r, "fte": n, "salary_each": pay, "annual_cost_with_oncosts": round(n * pay * 1.12, 2)} for r, n, pay in roles]
    files = {"management_accounts.csv": accounts, "other_income_ledger.csv": other_income, "payroll.csv": payroll,
             "overheads.csv": [{"category": c, "annual": round(v, 2)} for c, v in overheads],
             "stock_ageing.csv": stock, "customer_cohorts.csv": cohorts, "customer_activity.csv": activity,
             "creditor_ageing.csv": creditors, "contracts.csv": contracts, "eu_fulfilment_sales.csv": eu}
    for name, rows in files.items():
        write(ROOM / "fernbrook" / name, rows)


def halvergate_room(rng):
    season = [9, 6, 3, 2, 2, 3, 12, 13, 12, 11, 10, 17]  # Feb..Jan: hockey runs September to March
    framework, retail = 3_772_000, 5_428_000
    roles = [("Managing Director (founder)", 1, 85_000), ("Schools account manager", 2, 42_000), ("Warehouse operative", 12, 25_400),
             ("Customer service advisor", 3, 26_500), ("Buyer / marketer", 3, 39_000), ("Finance assistant", 1, 31_000)]
    staff = sum(n * pay for _, n, pay in roles) * 1.12
    lines = {"Revenue": framework + retail, "Cost of sales": -(framework * 0.70 + retail * 0.50), "Supplier rebates": 240_000,
             "Other income": 0, "Staff costs": -staff, "Warehouse rent": -210_000, "Delivery": -(72_000 * 5.60 + 150_000),
             "Marketing": -retail * 0.06}
    other_total = sum(lines.values()) - 1_000_000
    lines["Other overheads"] = -other_total
    accounts = []
    for months, scale, label in ((PRIOR, 1 / 1.04, "Prior year"), (LTM, 1, "Last 12 months")):
        split_lines = {k: monthly(v * (1 if k in ("Staff costs", "Warehouse rent") else scale),
                                  [1] * 12 if k in ("Staff costs", "Warehouse rent", "Other overheads", "Other income") else season)
                       for k, v in lines.items()}
        for i, d in enumerate(months):
            row = {"month": d.strftime("%Y-%m"), "period": label, **{k: split_lines[k][i] for k in lines}}
            row["EBITDA"] = round(sum(split_lines[k][i] for k in lines), 2)
            accounts.append(row)
    segments = [{"segment": "County schools and clubs framework", "revenue": framework, "gross_margin": 0.30,
                 "direct_costs": 150_000 + 2 * 42_000 * 1.12},
                {"segment": "Online retail", "revenue": retail, "gross_margin": 0.50, "direct_costs": ""}]
    contracts = [
        {"counterparty": "County schools sports consortium", "type": "Framework supply agreement", "share_of_sales": 0.41,
         "expiry": "2026-11-30", "change_of_control": "None", "notes": "Goes to open tender in November 2026. Incumbent has no preference."},
        {"counterparty": "Halvergate Properties Ltd (owned by the founder's family)", "type": "Warehouse lease", "share_of_sales": "",
         "expiry": "2036-01-31", "change_of_control": "None",
         "notes": "Rent £210,000 a year, no break clause. Agent's letter puts market rent at £105,000."},
        {"counterparty": "Three stick and ball suppliers", "type": "Rebate letters", "share_of_sales": "", "expiry": "None stated",
         "change_of_control": "Not addressed", "notes": "Rebates of £240,000 a year agreed personally with the founder. No written contract."},
    ]
    payroll = [{"role": r, "fte": n, "salary_each": pay, "annual_cost_with_oncosts": round(n * pay * 1.12, 2)} for r, n, pay in roles]
    files = {"management_accounts.csv": accounts, "segments.csv": segments, "contracts.csv": contracts, "payroll.csv": payroll}
    for name, rows in files.items():
        write(ROOM / "halvergate" / name, rows)


# ---------- due diligence: everything below reads the data room files ----------

def num(x):
    return float(x) if x not in ("", None) else 0.0


def ltm(rows, line):
    return sum(num(r[line]) for r in rows if r["period"] == "Last 12 months")


def dd_fernbrook():
    acc = read("fernbrook", "management_accounts.csv")
    reported = ltm(acc, "EBITDA")
    revenue = ltm(acc, "Revenue")
    cogs = -ltm(acc, "Cost of sales")
    findings = []

    founder = next(r for r in read("fernbrook", "payroll.csv") if "founder" in r["role"].lower())
    md_adj = -(POLICY["md_market_salary"] - num(founder["salary_each"])) * 1.12
    findings.append(("Founder underpaid", "payroll.csv", f"Founder MD paid £{num(founder['salary_each']):,.0f}; a replacement costs about £{POLICY['md_market_salary']:,.0f}", "EBITDA", md_adj))

    oi = read("fernbrook", "other_income_ledger.csv")
    prior_desc = {r["description"] for r in oi if r["date"] < "2025-02"}
    one_off = sum(num(r["amount"]) for r in oi if r["date"] >= "2025-02" and r["description"] not in prior_desc)
    findings.append(("One-off income in EBITDA", "other_income_ledger.csv", "A £60k consultancy fee appears once, with nothing like it the year before", "EBITDA", -one_off))

    contracts = {c["counterparty"]: c for c in read("fernbrook", "contracts.csv")}
    rebate = ltm(acc, "Supplier rebates") if "ends on a change of control" in contracts["Brand A"]["notes"] else 0
    findings.append(("Rebate ends on sale", "contracts.csv", "Brand A's 5% rebate is personal to the owner and stops when the business changes hands", "EBITDA", -rebate))
    normalised = reported + sum(f[4] for f in findings)

    stock = read("fernbrook", "stock_ageing.csv")
    old, mid = sum(num(r["age_over_12m"]) for r in stock), sum(num(r["age_6_12m"]) for r in stock)
    book = sum(num(r[k]) for r in stock for k in ("age_0_3m", "age_3_6m", "age_6_12m", "age_over_12m"))
    provision = 0.6 * old + 0.2 * mid - sum(num(r["provision_held"]) for r in stock)
    findings.append(("Old stock at full value", "stock_ageing.csv", f"£{old / 1e3:,.0f}k over 12 months old and £{mid / 1e3:,.0f}k at 6-12 months, with no provision. Our policy writes them down 60% and 20%", "Completion accounts", -provision))

    cohorts = read("fernbrook", "customer_cohorts.csv")
    act = {r["period"]: r for r in read("fernbrook", "customer_activity.csv")}
    accounts_total = sum(num(r["accounts_opened"]) for r in cohorts)
    active = num(act["Last 12 months"]["active_customers"])
    active_growth = active / num(act["Prior year"]["active_customers"]) - 1
    rep = {int(r["cohort_year"]): num(r["repeat_within_12m"]) for r in cohorts}
    findings.append(("Customer numbers overstated", "customer_cohorts.csv", f"{accounts_total / 1e3:,.0f}k is every account ever opened. {active / 1e3:,.0f}k bought in the last year (up {active_growth:.0%}), and repeat rates fell from {rep[2023]:.0%} to {rep[2025]:.0%} in the padel cohorts", "Growth", 0))

    findings.append(("One brand is 35% of sales", "contracts.csv", "Brand A can end the agreement on a change of control, it expires 31 Dec 2026, and Brand A now sells direct", "Earn-out", 0))

    cred = read("fernbrook", "creditor_ageing.csv")
    payables = sum(num(r["total"]) for r in cred)
    terms = sum(num(r["total"]) * num(r["terms_days"]) for r in cred) / payables
    days = payables / cogs * 365
    catch_up = payables - cogs * terms / 365
    findings.append(("Suppliers stretched before sale", "creditor_ageing.csv", f"Payables at {days:.0f} days against contractual terms of {terms:.0f} days", "Completion accounts", -catch_up))

    eu = read("fernbrook", "eu_fulfilment_sales.csv")
    eu_sales = sum(num(r["sales_dispatched_from_germany_gross"]) for r in eu)
    vat = eu_sales * 0.19 / 1.19 * 1.10  # German VAT at 19%, plus about 10% for interest and penalties
    findings.append(("German VAT never registered", "eu_fulfilment_sales.csv + contracts.csv", f"£{eu_sales / 1e3:,.0f}k of sales shipped from a German warehouse since 2022 with no German VAT accounted", "Escrow", -vat))

    normal_nwc = (book - provision) + 100_000 - cogs * terms / 365
    return {"reported": reported, "normalised": normalised, "revenue": revenue, "findings": findings, "provision": provision,
            "catch_up": catch_up, "vat": vat, "active_growth": active_growth, "normal_nwc": normal_nwc, "brand_share": 0.35,
            "cogs": cogs, "asking": 6_500_000}


def dd_halvergate():
    acc = read("halvergate", "management_accounts.csv")
    reported, revenue = ltm(acc, "EBITDA"), ltm(acc, "Revenue")
    seg = {r["segment"]: r for r in read("halvergate", "segments.csv")}
    fw = seg["County schools and clubs framework"]
    fw_contribution = num(fw["revenue"]) * num(fw["gross_margin"]) - num(fw["direct_costs"])
    cut = num(fw["revenue"]) * 0.05  # expected price cut to win the re-tender
    rebates = ltm(acc, "Supplier rebates")
    lease = next(c for c in read("halvergate", "contracts.csv") if c["type"] == "Warehouse lease")
    months = [r for r in acc if r["period"] == "Last 12 months"]
    cum, trough = 0, 0
    for r in months[2:] + months[:2]:  # running EBITDA from April, when the season ends
        cum += num(r["EBITDA"])
        trough = min(trough, cum)
    season_share = sum(num(r["Revenue"]) for r in months if r["month"][5:] in ("09", "10", "11", "12", "01", "02", "03")) / revenue
    findings = [
        ("Rebates agreed personally", "contracts.csv", "£240k a year of supplier rebates rest on the founder's relationships, with nothing in writing", "EBITDA", -rebates),
        ("41% of sales go to tender", "contracts.csv + segments.csv", f"The schools framework (£{num(fw['revenue']) / 1e6:.1f}m of sales, £{fw_contribution / 1e3:,.0f}k contribution) is re-tendered in November 2026", "Scenario", -fw_contribution),
        ("Above-market lease from the founder", "contracts.csv", "Warehouse rented from the founder's family at £210k against £105k market rent, 10 years left, no break. We could never close it", "Synergies", -105_000 * 10),
        ("Seasonal cash drain", "management_accounts.csv", f"{season_share:.0%} of sales fall between September and March. Spring and summer months lose money", "Funding", trough),
    ]
    normalised = reported - rebates
    return {"reported": reported, "normalised": normalised, "revenue": revenue, "findings": findings, "fw_contribution": fw_contribution,
            "cut": cut, "fw_revenue": num(fw["revenue"]), "asking": 5_500_000, "floor": 4_500_000, "season_share": season_share}


# ---------- valuation ----------

def dcf(rev0, growth, margin, r, g=POLICY["terminal_growth"], da=40_000, capex=50_000, nwc=0.10):
    """Enterprise value from five years of free cash flow plus a terminal value."""
    rev, prev, pv, rows = rev0, rev0, 0, []
    for t in range(5):
        rev = prev * (1 + growth[t])
        ebitda = rev * margin[t]
        fcf = ebitda - POLICY["tax"] * (ebitda - da) - capex - nwc * (rev - prev)
        pv += fcf / (1 + r) ** (t + 1)
        rows.append((rev, ebitda, fcf))
        prev = rev
    tv = rows[-1][2] * (1 + g) / (r - g) / (1 + r) ** 5
    return pv + tv, rows


def cases(F, H):
    m = F["normalised"] / F["revenue"]
    lost = F["brand_share"] * 0.6  # lose Brand A; about 40% of its customers switch to other brands we stock
    fern = {
        "Seller's case": (F["revenue"], [0.20, 0.15, 0.12, 0.10, 0.08], [F["reported"] / F["revenue"]] * 5, POLICY["target_rate"]),
        "Fernbrook, brand renewed": (F["revenue"], [0.06, 0.05, 0.04, 0.03, 0.03], [m] * 5, POLICY["target_rate"]),
        "Fernbrook, brand lost": (F["revenue"], [0.06, -lost, 0.03, 0.03, 0.03], [m, m - 0.01, m - 0.01, m - 0.01, m - 0.01], POLICY["target_rate"]),
    }
    hm = H["normalised"] / H["revenue"]
    keep = (H["normalised"] - H["cut"]) / (H["revenue"] - H["cut"])
    lost_m = (H["normalised"] - H["fw_contribution"]) / (H["revenue"] - H["fw_revenue"])
    halv = {
        "Halvergate, framework kept at 5% lower prices": (H["revenue"], [-0.05 * H["fw_revenue"] / H["revenue"], 0.02, 0.02, 0.02, 0.02], [keep] * 5, POLICY["target_rate"]),
        "Halvergate, framework lost": (H["revenue"], [0.0, -H["fw_revenue"] / H["revenue"], 0.02, 0.02, 0.02], [hm, lost_m, lost_m, lost_m, lost_m], POLICY["target_rate"]),
    }
    return {**fern, **halv}


def synergies(F):
    """Annual synergies with Fernbrook, years 1-5, before tax. Year 1 is the first 12 months after completion."""
    orders = 150_000
    run = {
        "Delivery on our carrier contract (£4.45 an order, from P3)": orders * (5.40 - fc.OPTIONS["quote"][0]),
        "Close their warehouse (rent)": 180_000,
        "Pick and pack in our warehouse": 14 * 2080 * 12.21 * 1.12 - orders / 5.6 * 12.71 * 1.12,
        "Duplicate software, audit and insurance": 70_000 + 28_000 + 22_000,
    }
    phase = [0.5, 1, 1, 1, 1]
    one_off = {"Warehouse move": 150_000, "Platform migration": 120_000, "Lease break fee": 60_000, "Redundancy": 40_000}
    attrition = F["revenue"] * 0.02 * 0.25  # 2% of sales lost in year 1 at 25% contribution
    years = [sum(run.values()) * p - (attrition if t == 0 else 0) - (sum(one_off.values()) if t == 0 else 0) for t, p in enumerate(phase)]
    r = POLICY["hurdle"]
    after_tax = [y * (1 - POLICY["tax"]) for y in years]
    npv = sum(v / (1 + r) ** (t + 1) for t, v in enumerate(after_tax)) + after_tax[-1] / r / (1 + r) ** 5
    return {"run": run, "one_off": one_off, "attrition": attrition, "years": years, "npv": npv}


def halvergate_synergy_npv(H):
    run = 120_000 * (5.60 - fc.OPTIONS["quote"][0]) * 0.6 + 60_000  # carrier on retail parcels, some overheads; warehouse locked in
    r = POLICY["hurdle"]
    yearly = [run * 0.5] + [run] * 4
    return sum(v * 0.75 / (1 + r) ** (t + 1) for t, v in enumerate(yearly)) + run * 0.75 / r / (1 + r) ** 5


def round_to(x, step=250_000, down=True):
    return (x // step) * step if down else round(x / step) * step


def offer(F, val, syn):
    renewed, lost = val["Fernbrook, brand renewed"][0], val["Fernbrook, brand lost"][0]
    upfront = round_to(lost + POLICY["synergy_share"] * syn["npv"])
    earnout = round_to(renewed - lost, down=False)
    expected = POLICY["p_renew"] * renewed + (1 - POLICY["p_renew"]) * lost
    return {"upfront": upfront, "earnout": earnout, "escrow": -(-F["vat"] // 50_000) * 50_000,  # VAT exposure rounded up to the next £50k
            "catch_up": F["catch_up"], "expected_value": expected, "walk_away": expected + POLICY["synergy_share"] * syn["npv"],
            "ceiling": renewed + syn["npv"]}


# ---------- funding: combined cash on top of the P3 forecast ----------

def combined_cash(F, syn, deal, base_runs):
    """Month-end group cash for each P3 scenario, with the brand renewed or lost."""
    season = [7, 7, 8, 10, 11, 11, 10, 8, 7, 7, 8, 6]
    w = [season[c % 12] / sum(season) for c in range(fc.N)]
    rate = (POLICY["bank_rate"] + POLICY["margin"]) / 12
    out = {}
    for scn, (P, _) in base_runs.items():
        for brand in ("renewed", "lost"):
            cash, debt, series = 0.0, 0.0, []
            ebitda_y = F["normalised"]
            for c in range(fc.N):
                flow = 0.0
                if c == COMPLETION:
                    debt = POLICY["loan"]
                    flow += POLICY["loan"] - (deal["upfront"] - deal["catch_up"]) - POLICY["deal_costs"]
                if c > COMPLETION:
                    k = c - COMPLETION  # months since completion
                    yr = (k - 1) // 12
                    e = ebitda_y * w[c] * 12 / 12
                    if brand == "lost" and c >= 11:  # agreement ends 31 Dec 2026
                        e *= 1 - F["brand_share"] * 0.6
                    s = syn["years"][min(yr, 4)] / 12 if yr < 5 else 0
                    flow += (e + s) * (1 - POLICY["tax"]) - 50_000 / 12
                    if k <= 2:
                        flow -= deal["catch_up"] / 2  # pay suppliers back to normal terms
                    repay = POLICY["loan"] / POLICY["loan_months"]
                    flow -= debt * rate + repay
                    debt -= repay
                    if c == EARNOUT_MONTH and brand == "renewed":
                        flow -= deal["earnout"]
                cash += flow
                series.append(P["cash"][c] + cash)
            out[(scn, brand)] = series
    return out


# ---------- workbook ----------

def build_workbook(F, H, val, syn, deal, cash, hv):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cover"
    header(ws, "DirectSportsGoods: acquisition case")
    ws.column_dimensions["A"].width = 100
    for i, t in enumerate([
        "Two founder-owned racket and hockey retailers, screened and taken through due diligence. Both are fictional.",
        "",
        "Summary: the two targets side by side",
        "Fernbrook DD and Halvergate DD: what due diligence found in each data room, and what it costs",
        "Valuation: discounted cash flow for each case, live formulas",
        "Bridge: from Fernbrook's asking price to our offer",
        "Synergies: what Fernbrook is worth to us on top of its own cash flows",
        "Funding: group cash month by month after the deal, for each P3 scenario",
        "",
        "Figures are in £k unless stated. Data room files are in data/dataroom/.",
    ], start=4):
        ws.cell(i, 1, t)

    ws = wb.create_sheet("Summary")
    header(ws, "Targets side by side (£k)")
    head_row(ws, 5, ["", "Fernbrook Racquets", "Halvergate Hockey"])
    rows = [
        ("Niche", "Racket sports, heavy in padel", "Hockey"),
        ("Revenue, last 12 months", F["revenue"], H["revenue"]),
        ("Reported EBITDA", F["reported"], H["reported"]),
        ("Asking price", F["asking"], H["asking"]),
        ("Asking price / reported EBITDA", F["asking"] / F["reported"], H["asking"] / H["reported"]),
        ("EBITDA after due diligence", F["normalised"], H["normalised"]),
        ("Asking price / EBITDA after due diligence", F["asking"] / F["normalised"], H["asking"] / H["normalised"]),
        ("Walk-away price", deal["walk_away"], hv["walk_away"]),
        ("Seller's lowest acceptable price", "", H["floor"]),
        ("Decision", f"Buy: £{deal['upfront'] / 1e6:.2f}m upfront plus up to £{deal['earnout'] / 1e6:.2f}m earn-out", "Walk away"),
    ]
    for i, (label, a, b) in enumerate(rows):
        r = 6 + i
        ws.cell(r, 1, label)
        for j, v in enumerate((a, b)):
            c = ws.cell(r, 2 + j, v)
            c.number_format = "0.0x" if "/" in label else GBP
            c.alignment = Alignment(horizontal="right")
    ws.cell(18, 1, f"Benchmark: Frasers paid {HOLDSPORT['price'] / HOLDSPORT['op_profit']:.1f}x operating profit for Holdsport and "
                   f"{XXL['price'] / XXL['revenue']:.2f}x revenue for loss-making XXL in FY26 (Frasers AR 2026, p.179-180).").font = bp.F_NOTE
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 22

    for name, D in (("Fernbrook DD", F), ("Halvergate DD", H)):
        ws = wb.create_sheet(name)
        header(ws, f"{name.replace(' DD', '')}: due diligence findings (£k)", "Every finding comes from a file in the data room.")
        head_row(ws, 5, ["Finding", "Source file", "What the file shows", "Affects", "£k"])
        for i, (finding, src, what, hits, amt) in enumerate(D["findings"]):
            r = 6 + i
            for j, v in enumerate((finding, src, what, hits, amt if amt else "")):
                c = ws.cell(r, 1 + j, v)
                c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(r, 5).number_format = GBP
            ws.row_dimensions[r].height = 15 * (len(what) // 60 + 1)
        r = 7 + len(D["findings"])
        ws.cell(r, 1, "Reported EBITDA").font = Font(bold=True)
        ws.cell(r, 5, D["reported"]).number_format = GBP
        ws.cell(r + 1, 1, "EBITDA after due diligence").font = Font(bold=True)
        ws.cell(r + 1, 5, D["normalised"]).number_format = GBP
        for c, w in zip("ABCDE", (30, 26, 60, 18, 10)):
            ws.column_dimensions[c].width = w

    ws = wb.create_sheet("Valuation")
    header(ws, "Discounted cash flow by case (£k)", "Blue cells are inputs. Free cash flow = EBITDA - tax on (EBITDA - depreciation) - capex - working capital on growth.")
    head_row(ws, 5, ["", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Terminal", "Value"])
    r = 6
    for name, (rev0, growth, margin, rate) in val["inputs"].items():
        ws.cell(r, 1, name).font = Font(bold=True, color=bp.NAVY)
        ws.cell(r + 1, 1, "Starting revenue")
        ws.cell(r + 1, 2, rev0).number_format = GBP
        ws.cell(r + 1, 3, "Discount rate")
        ws.cell(r + 1, 4, rate).number_format = "0.0%"
        ws.cell(r + 1, 5, "Terminal growth")
        ws.cell(r + 1, 6, POLICY["terminal_growth"]).number_format = "0.0%"
        for c in (2, 4, 6):
            ws.cell(r + 1, c).font = Font(color=INPUT)
        labels = ["Revenue growth", "EBITDA margin", "Revenue", "EBITDA", "Free cash flow", "Discount factor", "Present value"]
        R = {lab: r + 2 + i for i, lab in enumerate(labels)}
        for lab in labels:
            ws.cell(R[lab], 1, lab)
        for t in range(5):
            L, P = col(2 + t), col(1 + t)
            ws.cell(R["Revenue growth"], 2 + t, growth[t]).number_format = PCT
            ws.cell(R["EBITDA margin"], 2 + t, margin[t]).number_format = PCT
            ws.cell(R["Revenue growth"], 2 + t).font = Font(color=INPUT)
            ws.cell(R["EBITDA margin"], 2 + t).font = Font(color=INPUT)
            prev = f"$B${r + 1}" if t == 0 else f"{P}{R['Revenue']}"
            ws.cell(R["Revenue"], 2 + t, f"={prev}*(1+{L}{R['Revenue growth']})").number_format = GBP
            ws.cell(R["EBITDA"], 2 + t, f"={L}{R['Revenue']}*{L}{R['EBITDA margin']}").number_format = GBP
            ws.cell(R["Free cash flow"], 2 + t, f"={L}{R['EBITDA']}-{POLICY['tax']}*({L}{R['EBITDA']}-40000)-50000-0.1*({L}{R['Revenue']}-{prev})").number_format = GBP
            ws.cell(R["Discount factor"], 2 + t, f"=1/(1+$D${r + 1})^{t + 1}").number_format = "0.000"
            ws.cell(R["Present value"], 2 + t, f"={L}{R['Free cash flow']}*{L}{R['Discount factor']}").number_format = GBP
        ws.cell(R["Free cash flow"], 7, f"=F{R['Free cash flow']}*(1+$F${r + 1})/($D${r + 1}-$F${r + 1})").number_format = GBP
        ws.cell(R["Discount factor"], 7, f"=F{R['Discount factor']}").number_format = "0.000"
        ws.cell(R["Present value"], 7, f"=G{R['Free cash flow']}*G{R['Discount factor']}").number_format = GBP
        ws.cell(R["Present value"], 8, f"=SUM(B{R['Present value']}:G{R['Present value']})").number_format = GBP
        ws.cell(R["Present value"], 8).font = Font(bold=True)
        r = R["Present value"] + 2
    ws.column_dimensions["A"].width = 44
    for c in range(2, 9):
        ws.column_dimensions[col(c)].width = 11

    ws = wb.create_sheet("Bridge")
    header(ws, "Fernbrook: from asking price to offer (£k)")
    head_row(ws, 5, ["Step", "£k"])
    for i, (label, v) in enumerate(deal["bridge"]):
        ws.cell(6 + i, 1, label)
        ws.cell(6 + i, 2, v).number_format = GBP
        if label.startswith(("Asking", "Walk-away", "Upfront", "Paid at completion")):
            ws.cell(6 + i, 1).font = ws.cell(6 + i, 2).font = Font(bold=True)
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 12

    ws = wb.create_sheet("Synergies")
    header(ws, "Synergies with Fernbrook (£k, before tax)")
    head_row(ws, 5, ["", "Run rate"])
    r = 6
    for k, v in syn["run"].items():
        ws.cell(r, 1, k)
        ws.cell(r, 2, v).number_format = GBP
        r += 1
    ws.cell(r, 1, "Total run rate").font = Font(bold=True)
    ws.cell(r, 2, sum(syn["run"].values())).number_format = GBP
    r += 2
    head_row(ws, r, ["One-off costs, year 1", "£k"])
    for k, v in syn["one_off"].items():
        r += 1
        ws.cell(r, 1, k)
        ws.cell(r, 2, -v).number_format = GBP
    r += 1
    ws.cell(r, 1, "Customers lost in the switch, year 1 (2% of sales)")
    ws.cell(r, 2, -syn["attrition"]).number_format = GBP
    r += 2
    head_row(ws, r, ["Net synergies by year", "£k"])
    for t, v in enumerate(syn["years"]):
        r += 1
        ws.cell(r, 1, f"Year {t + 1}")
        ws.cell(r, 2, v).number_format = GBP
    r += 2
    ws.cell(r, 1, f"Net present value after tax at {POLICY['hurdle']:.0%}, with year 5 held flat after").font = Font(bold=True)
    ws.cell(r, 2, syn["npv"]).number_format = GBP
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 12

    ws = wb.create_sheet("Funding")
    header(ws, "Group cash after the deal (£k)", f"P3 forecast plus the deal: £{POLICY['loan'] / 1e6:.1f}m loan at Bank Rate + {POLICY['margin']:.1%}, "
                                               f"repaid over {POLICY['loan_months'] // 12} years. Board minimum: £{POLICY['min_cash'] / 1e6:.1f}m.")
    head_row(ws, 5, [""] + [""] * fc.N + ["Lowest"])
    for c, dt in enumerate(fc.FM):
        ws.cell(5, 2 + c, dt).number_format = "mmm-yy"
    for i, ((scn, brand), series) in enumerate(cash.items()):
        r = 6 + i
        ws.cell(r, 1, f"{scn}, brand {brand}")
        for c, v in enumerate(series):
            ws.cell(r, 2 + c, round(v, 2))
        ws.cell(r, 2 + fc.N, f"=MIN(B{r}:{col(1 + fc.N)}{r})")
        style_row(ws, r, 2 + fc.N, GBP)
    ws.column_dimensions["A"].width = 26
    for c in range(2, 3 + fc.N):
        ws.column_dimensions[col(c)].width = 9.5
    wb.save(OUT)


def main():
    rng = random.Random(4)
    fernbrook_room(rng)
    halvergate_room(rng)
    (ROOM / "README.md").write_text("# Data rooms\n\nFernbrook Racquets and Halvergate Hockey are fictional companies. "
                                    "Every file and figure here is made up for the DirectSportsGoods acquisition case (P4).\n", encoding="utf-8")
    F, H = dd_fernbrook(), dd_halvergate()
    inputs = cases(F, H)
    val = {name: dcf(rev0, g, m, r) for name, (rev0, g, m, r) in inputs.items()}
    val["inputs"] = inputs
    syn = synergies(F)
    deal = offer(F, val, syn)
    hv_keep, hv_lost = val["Halvergate, framework kept at 5% lower prices"][0], val["Halvergate, framework lost"][0]
    hv = {"expected": 0.5 * hv_keep + 0.5 * hv_lost, "syn": halvergate_synergy_npv(H)}
    hv["walk_away"] = max(0, hv["expected"]) + POLICY["synergy_share"] * hv["syn"]
    hv["best"] = hv_keep + POLICY["synergy_share"] * hv["syn"]
    seller = val["Seller's case"][0]
    deal["bridge"] = [
        ("Asking price", F["asking"]),
        ("Seller's own forecast, valued at our 15% rate", seller),
        ("Less: diligence adjustments to earnings and growth", val["Fernbrook, brand renewed"][0] - seller),
        ("Less: brand agreement at risk (50% chance of losing it)", deal["expected_value"] - val["Fernbrook, brand renewed"][0]),
        (f"Plus: {POLICY['synergy_share']:.0%} of our synergies", POLICY["synergy_share"] * syn["npv"]),
        ("Walk-away price", deal["walk_away"]),
        ("Upfront price offered (brand-lost value plus synergy share, rounded down)", deal["upfront"]),
        ("Earn-out if Brand A renews for 3 years (gap between renewed and lost values)", deal["earnout"]),
        ("Less at completion: suppliers stretched to 70 days, back to terms", -deal["catch_up"]),
        ("Paid at completion", deal["upfront"] - deal["catch_up"]),
        ("Held in escrow from that payment for German VAT, 24 months", deal["escrow"]),
    ]
    base_runs = fc_runs()
    cash = combined_cash(F, syn, deal, base_runs)

    # Self-checks: due diligence finds what the data rooms hide, the rejected target stays rejected,
    # and the deal keeps cash above the board's minimum in every scenario where the brand renews.
    assert abs(F["normalised"] - 683_600) < 2_000, F["normalised"]
    assert abs(F["provision"] - 250_000) < 1, F["provision"]
    assert 380_000 < F["catch_up"] < 450_000, F["catch_up"]
    assert hv["best"] < H["floor"], "Halvergate should be out of reach even if the framework is kept"
    assert deal["upfront"] + deal["earnout"] <= deal["ceiling"]
    assert all(min(s) > POLICY["min_cash"] for (scn, b), s in cash.items() if b == "renewed")

    build_workbook(F, H, val, syn, deal, cash, hv)
    PAPER.write_text(paper(F, H, val, syn, deal, cash, hv), encoding="utf-8")
    print(f"Wrote {OUT.relative_to(ROOT)}, {PAPER.relative_to(ROOT)} and data/dataroom/")
    print(f"  Fernbrook: reported {F['reported'] / 1e3:,.0f}k, normalised {F['normalised'] / 1e3:,.0f}k, seller DCF {seller / 1e6:.2f}m, "
          f"renewed {val['Fernbrook, brand renewed'][0] / 1e6:.2f}m, lost {val['Fernbrook, brand lost'][0] / 1e6:.2f}m, syn NPV {syn['npv'] / 1e6:.2f}m")
    print(f"  Offer: upfront {deal['upfront'] / 1e6:.2f}m, earn-out {deal['earnout'] / 1e6:.2f}m, walk-away {deal['walk_away'] / 1e6:.2f}m, "
          f"ceiling {deal['ceiling'] / 1e6:.2f}m, catch-up {deal['catch_up'] / 1e3:,.0f}k, VAT {F['vat'] / 1e3:,.0f}k, escrow {deal['escrow'] / 1e3:,.0f}k")
    print(f"  Halvergate: reported {H['reported'] / 1e3:,.0f}k, normalised {H['normalised'] / 1e3:,.0f}k, keep {hv_keep / 1e6:.2f}m, "
          f"lost {hv_lost / 1e6:.2f}m, walk-away {hv['walk_away'] / 1e6:.2f}m, best {hv['best'] / 1e6:.2f}m vs floor {H['floor'] / 1e6:.1f}m")
    for k, s in cash.items():
        print(f"  {k}: lowest cash {min(s) / 1e6:.2f}m, Jul-27 {s[-1] / 1e6:.2f}m")


def fc_runs():
    A, fixed = fc.calibrate()
    X = {k: v[0] for k, v in fixed.items()}
    runs = {}
    for j, s in enumerate(fc.SCN):
        d = {key: vals[j] for key, _, *vals in fc.DRIVERS}
        runs[s] = (fc.forecast(A, X, d), None)
    return runs


def paper(F, H, val, syn, deal, cash, hv):
    k = lambda v: f"£{abs(v) / 1e3:,.0f}k"
    m = lambda v: f"£{v / 1e6:.2f}m"
    renewed, lost = val["Fernbrook, brand renewed"][0], val["Fernbrook, brand lost"][0]
    paid_now = deal["upfront"] - deal["catch_up"]
    npv_renewed = renewed + syn["npv"] - deal["upfront"] - deal["earnout"] - POLICY["deal_costs"]
    npv_lost = lost + syn["npv"] - deal["upfront"] - POLICY["deal_costs"]
    low = {key: min(s) for key, s in cash.items()}
    tight = min(low, key=low.get)
    holdsport = HOLDSPORT["price"] / HOLDSPORT["op_profit"]
    fd = "\n".join(f"| {f} | `{src}` | {what} | {('-' if a < 0 else '') + k(a) if a else 'see below'} |" for f, src, what, _, a in F["findings"])
    hd = "\n".join(f"- **{f}.** {what}." for f, src, what, _, a in H["findings"])
    cash_rows = "\n".join(f"| {scn} | {m(low[(scn, 'renewed')])} | {m(low[(scn, 'lost')])} |" for scn in fc.SCN)
    syn_rows = "\n".join(f"| {name} | {k(v)} |" for name, v in syn["run"].items())
    bridge = "\n".join(f"| {label} | {v / 1e6:+.2f} |" if label.startswith(("Less", "Plus")) else f"| **{label}** | **{v / 1e6:.2f}** |"
                       for label, v in deal["bridge"][:6])
    return f"""# Board paper: first acquisition

> DirectSportsGoods, Fernbrook Racquets and Halvergate Hockey are fictional companies, and every figure here is made up. The data room files are in [`data/dataroom/`](../data/dataroom/) and the model is in the [acquisition workbook](../pack/DirectSportsGoods_acquisition.xlsx).

**Decision requested:** approve an offer for Fernbrook Racquets of **{m(deal['upfront'])} upfront plus an earn-out of up to {m(deal['earnout'])}**, funded by a {m(POLICY['loan'])} bank loan and cash. Withdraw from Halvergate Hockey.

## Summary

- We took two founder-owned specialists through due diligence. Both looked good at first glance. Diligence cut Fernbrook's earnings by {1 - F['normalised'] / F['reported']:.0%} and found a problem in Halvergate that no price fixes.
- **Fernbrook** is asking {m(F['asking'])}: {F['asking'] / F['reported']:.1f}x the EBITDA it reports, and {F['asking'] / F['normalised']:.1f}x the EBITDA it really makes. Our walk-away price is {m(deal['walk_away'])}.
- **Halvergate** is worth {m(hv['best'])} to us even if it keeps its biggest contract. The founder won't go below {m(H['floor'])}.
- Even if Fernbrook loses its biggest brand, the deal is worth an estimated {m(npv_lost)} more to us than it costs, including synergies. If the brand renews and we pay the full earn-out, the figure is {m(npv_renewed)}.

## The two targets

| | Fernbrook Racquets | Halvergate Hockey |
|---|---:|---:|
| Revenue, last 12 months | {m(F['revenue'])} | {m(H['revenue'])} |
| Reported EBITDA | {k(F['reported'])} | {k(H['reported'])} |
| EBITDA after due diligence | {k(F['normalised'])} | {k(H['normalised'])} |
| Asking price | {m(F['asking'])} | {m(H['asking'])} |
| Our walk-away price | {m(deal['walk_away'])} | {m(hv['walk_away'])} |
| Decision | Offer | Walk away |

For scale: Frasers Group paid {holdsport:.1f}x operating profit for Holdsport in FY26 (Frasers Annual Report 2026, p.179-180). At that multiple, Fernbrook's post-diligence EBITDA is worth {m(holdsport * F['normalised'])}, close to our upfront offer.

## Why we walk away from Halvergate

Halvergate looked the better business: higher margin, a lower asking multiple ({H['asking'] / H['reported']:.1f}x), and an established schools market. Diligence found:

{hd}

If the schools framework is lost, Halvergate makes a loss. If it is kept, it will be at lower prices. Weighting the two outcomes equally, it is worth {m(hv['walk_away'])} to us including synergies, and the warehouse lease rules out the biggest synergy. No deal structure bridges the gap to {m(H['floor'])}.

## Fernbrook: what due diligence found

| Finding | Data room file | What it shows | Effect |
|---|---|---|---:|
{fd}

EBITDA falls from {k(F['reported'])} to **{k(F['normalised'])}**. The customer cohorts cut our growth assumption from the seller's 20% to 6%. Brand A is handled through the earn-out below.

## Valuation

| Step | £m |
|---|---:|
{bridge}

Values are discounted cash flows over five years plus a terminal value, at {POLICY['target_rate']:.0%} for Fernbrook on its own (our {POLICY['hurdle']:.0%} hurdle plus 3 points for a small founder-run business). Synergies are valued at {POLICY['hurdle']:.0%}. We give away {POLICY['synergy_share']:.0%} of the synergies at most, because we carry the risk of delivering them.

## The offer

- **{m(deal['upfront'])} upfront**, on a cash-free, debt-free basis with normal working capital of about {m(F['normal_nwc'])}. Suppliers have been stretched to 70 days, so bringing them back to terms takes {k(deal['catch_up'])} off the price: **{m(paid_now)} paid at completion**.
- **Earn-out of up to {m(deal['earnout'])}**, paid in February 2027 only if Brand A signs a new three-year agreement on the same terms. That is roughly the gap between Fernbrook's value with and without the brand, so the founder carries the risk they know most about.
- **{k(deal['escrow'])} held in escrow for 24 months** against German VAT (estimated {k(F['vat'])} with interest and penalties). We register for German VAT from day one.
- **Stock valued on our provisioning policy** in the completion accounts, {k(F['provision'])} below book value.
- **Founder stays six months** on a paid handover, to introduce us to Brand A and the club coaches.

## Synergies

| Run-rate synergy | £k a year |
|---|---:|
{syn_rows}
| **Total** | **{k(sum(syn['run'].values()))}** |

One-off costs of {k(sum(syn['one_off'].values()))} in year one (warehouse move, platform migration, lease break, redundancy), plus about {k(syn['attrition'])} of contribution from customers lost in the switch. Half the run rate arrives in year one. Net present value after tax: **{m(syn['npv'])}**. The warehouse move happens after the summer peak, in September.

## Funding and cash

{m(POLICY['loan'])} five-year term loan at Bank Rate ({POLICY['bank_rate']:.2%} at the time of writing) plus {POLICY['margin']:.1%}, with the rest from cash. Lowest month-end group cash over the 18-month forecast, against a board minimum of {m(POLICY['min_cash'])}:

| P3 scenario | Brand A renews (earn-out paid) | Brand A lost |
|---|---:|---:|
{cash_rows}

Cash stays above the minimum in every case, but in the **{tight[0].lower()} case with the earn-out paid, headroom is only {k(low[tight] - POLICY['min_cash'])}**. Before signing, we should either agree to pay the earn-out in two halves (February and August 2027) or put a £1m revolving credit facility in place.

## Risks

- **Brand A walks away.** Protected by the earn-out. Fernbrook is then worth {m(lost)} on its own, and {m(lost + syn['npv'])} to us with synergies, against {m(deal['upfront'])} paid upfront.
- **Padel cools faster than expected.** Repeat rates in the 2025 cohort are already down to 24%. Our 6% growth assumption sits well below the seller's 20%, but a sharper fall would hurt.
- **Integration slips into peak season.** A delayed warehouse move would push the rent and labour synergies into the following year.
- **Size.** The combined group passes the £54m turnover line for a medium-sized company. If it also passes a second size test two years running, large-company reporting and audit rules apply. Finance will plan for that from FY28.

## Next steps

1. Board approval of the offer and the funding.
2. Send a revised offer letter to Fernbrook's founder setting out the diligence findings behind the price.
3. Agree the loan terms and either the earn-out instalments or the revolving facility.
4. Target completion on 1 May 2026.
"""


if __name__ == "__main__":
    main()
