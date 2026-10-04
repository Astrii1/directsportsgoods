# Board paper: first acquisition

> DirectSportsGoods, Fernbrook Racquets and Halvergate Hockey are fictional companies, and every figure here is made up. The data room files are in [`data/dataroom/`](../data/dataroom/) and the model is in the [acquisition workbook](../pack/DirectSportsGoods_acquisition.xlsx).

**Decision requested:** approve an offer for Fernbrook Racquets of **£3.25m upfront plus an earn-out of up to £1.25m**, funded by a £2.50m bank loan and cash. Withdraw from Halvergate Hockey.

## Summary

- We took two founder-owned specialists through due diligence. Both looked good at first glance. Diligence cut Fernbrook's earnings by 28% and found a problem in Halvergate that no price fixes.
- **Fernbrook** is asking £6.50m: 6.8x the EBITDA it reports, and 9.5x the EBITDA it really makes. Our walk-away price is £3.96m.
- **Halvergate** is worth £3.16m to us even if it keeps its biggest contract. The founder won't go below £4.50m.
- Even if Fernbrook loses its biggest brand, the deal is worth an estimated £1.58m more to us than it costs, including synergies. If the brand renews and we pay the full earn-out, the figure is £1.50m.

## The two targets

| | Fernbrook Racquets | Halvergate Hockey |
|---|---:|---:|
| Revenue, last 12 months | £10.40m | £9.20m |
| Reported EBITDA | £950k | £1,000k |
| EBITDA after due diligence | £684k | £760k |
| Asking price | £6.50m | £5.50m |
| Our walk-away price | £3.96m | £1.60m |
| Decision | Offer | Walk away |

For scale: Frasers Group paid 4.7x operating profit for Holdsport in FY26 (Frasers Annual Report 2026, p.179-180). At that multiple, Fernbrook's post-diligence EBITDA is worth £3.19m, close to our upfront offer.

## Why we walk away from Halvergate

Halvergate looked the better business: higher margin, a lower asking multiple (5.5x), and an established schools market. Diligence found:

- **Rebates agreed personally.** £240k a year of supplier rebates rest on the founder's relationships, with nothing in writing.
- **41% of sales go to tender.** The schools framework (£3.8m of sales, £888k contribution) is re-tendered in November 2026.
- **Above-market lease from the founder.** Warehouse rented from the founder's family at £210k against £105k market rent, 10 years left, no break. We could never close it.
- **Seasonal cash drain.** 78% of sales fall between September and March. Spring and summer months lose money.

If the schools framework is lost, Halvergate makes a loss. If it is kept, it will be at lower prices. Weighting the two outcomes equally, it is worth £1.60m to us including synergies, and the warehouse lease rules out the biggest synergy. No deal structure bridges the gap to £4.50m.

## Fernbrook: what due diligence found

| Finding | Data room file | What it shows | Effect |
|---|---|---|---:|
| Founder underpaid | `payroll.csv` | Founder MD paid £30,000; a replacement costs about £120,000 | -£101k |
| One-off income in EBITDA | `other_income_ledger.csv` | A £60k consultancy fee appears once, with nothing like it the year before | -£60k |
| Rebate ends on sale | `contracts.csv` | Brand A's 5% rebate is personal to the owner and stops when the business changes hands | -£106k |
| Old stock at full value | `stock_ageing.csv` | £360k over 12 months old and £170k at 6-12 months, with no provision. Our policy writes them down 60% and 20% | -£250k |
| Customer numbers overstated | `customer_cohorts.csv` | 85k is every account ever opened. 32k bought in the last year (up 5%), and repeat rates fell from 38% to 24% in the padel cohorts | see below |
| One brand is 35% of sales | `contracts.csv` | Brand A can end the agreement on a change of control, it expires 31 Dec 2026, and Brand A now sells direct | see below |
| Suppliers stretched before sale | `creditor_ageing.csv` | Payables at 70 days against contractual terms of 44 days | -£428k |
| German VAT never registered | `eu_fulfilment_sales.csv + contracts.csv` | £975k of sales shipped from a German warehouse since 2022 with no German VAT accounted | -£171k |

EBITDA falls from £950k to **£684k**. The customer cohorts cut our growth assumption from the seller's 20% to 6%. Brand A is handled through the earn-out below.

## Valuation

| Step | £m |
|---|---:|
| **Asking price** | **6.50** |
| **Seller's own forecast, valued at our 15% rate** | **7.22** |
| Less: diligence adjustments to earnings and growth | -3.41 |
| Less: brand agreement at risk (50% chance of losing it) | -0.59 |
| Plus: 30% of our synergies | +0.73 |
| **Walk-away price** | **3.96** |

Values are discounted cash flows over five years plus a terminal value, at 15% for Fernbrook on its own (our 12% hurdle plus 3 points for a small founder-run business). Synergies are valued at 12%. We give away 30% of the synergies at most, because we carry the risk of delivering them.

## The offer

- **£3.25m upfront**, on a cash-free, debt-free basis with normal working capital of about £1.07m. Suppliers have been stretched to 70 days, so bringing them back to terms takes £428k off the price: **£2.82m paid at completion**.
- **Earn-out of up to £1.25m**, paid in February 2027 only if Brand A signs a new three-year agreement on the same terms. That is roughly the gap between Fernbrook's value with and without the brand, so the founder carries the risk they know most about.
- **£200k held in escrow for 24 months** against German VAT (estimated £171k with interest and penalties). We register for German VAT from day one.
- **Stock valued on our provisioning policy** in the completion accounts, £250k below book value.
- **Founder stays six months** on a paid handover, to introduce us to Brand A and the club coaches.

## Synergies

| Run-rate synergy | £k a year |
|---|---:|
| Delivery on our carrier contract (£4.45 an order, from P3) | £143k |
| Close their warehouse (rent) | £180k |
| Pick and pack in our warehouse | £17k |
| Duplicate software, audit and insurance | £120k |
| **Total** | **£459k** |

One-off costs of £370k in year one (warehouse move, platform migration, lease break, redundancy), plus about £52k of contribution from customers lost in the switch. Half the run rate arrives in year one. Net present value after tax: **£2.43m**. The warehouse move happens after the summer peak, in September.

## Funding and cash

£2.50m five-year term loan at Bank Rate (3.75% at the time of writing) plus 3.0%, with the rest from cash. Lowest month-end group cash over the 18-month forecast, against a board minimum of £2.00m:

| P3 scenario | Brand A renews (earn-out paid) | Brand A lost |
|---|---:|---:|
| Base | £5.72m | £5.72m |
| Downside | £4.15m | £5.06m |
| Upside | £6.10m | £6.10m |

Cash stays at least £2.15m above the minimum in every case, including the downside case with the earn-out paid. The deal doesn't need extra facilities, but we should keep the loan's covenants loose enough to allow a second acquisition.

## Risks

- **Brand A walks away.** Protected by the earn-out. Fernbrook is then worth £2.64m on its own, and £5.08m to us with synergies, against £3.25m paid upfront.
- **Padel cools faster than expected.** Repeat rates in the 2025 cohort are already down to 24%. Our 6% growth assumption sits well below the seller's 20%, but a sharper fall would hurt.
- **Integration slips into peak season.** A delayed warehouse move would push the rent and labour synergies into the following year.
- **Size.** The combined group passes the £54m turnover line for a medium-sized company. If it also passes a second size test two years running, large-company reporting and audit rules apply. Finance will plan for that from FY28.

## Next steps

1. Board approval of the offer and the funding.
2. Send a revised offer letter to Fernbrook's founder setting out the diligence findings behind the price.
3. Agree the loan terms and either the earn-out instalments or the revolving facility.
4. Target completion on 1 May 2026.
