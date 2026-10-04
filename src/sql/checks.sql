-- Month-end checks. Each one writes its findings to exceptions.
-- :fee_rate is the contracted marketplace fee rate, passed in from the pipeline.

CREATE TABLE exceptions (check_name TEXT, record_type TEXT, record_id INTEGER, period TEXT, amount REAL, detail TEXT);

-- 1a. Duplicate invoices: same supplier and the same invoice number once spacing, case and
--     punctuation are stripped. Suppliers don't reuse numbers, so the amount can differ (a keying typo).
INSERT INTO exceptions
SELECT 'duplicate_invoice', 'invoice', b.invoice_id, b.period_posted, b.amount,
       CASE WHEN b.amount = a.amount THEN 'Same number and amount as invoice ' ELSE 'Same number as invoice ' END
       || a.invoice_id || ' (' || a.invoice_no || ')'
FROM invoices a
JOIN invoices b ON b.supplier_id = a.supplier_id AND b.invoice_no_norm = a.invoice_no_norm AND b.invoice_id > a.invoice_id;

-- 1b. Re-sent invoices: new number, but same supplier and amount within 14 days.
--     A wider window starts catching regular monthly charges.
INSERT INTO exceptions
SELECT 'duplicate_invoice', 'invoice', b.invoice_id, b.period_posted, b.amount,
       'Same amount as invoice ' || a.invoice_id || ', '
       || CAST(julianday(b.invoice_date) - julianday(a.invoice_date) AS INTEGER) || ' days apart'
FROM invoices a
JOIN invoices b ON b.supplier_id = a.supplier_id AND b.amount = a.amount
               AND b.invoice_no_norm <> a.invoice_no_norm AND b.invoice_id > a.invoice_id
               AND julianday(b.invoice_date) - julianday(a.invoice_date) BETWEEN 0 AND 14
WHERE b.invoice_id NOT IN (SELECT record_id FROM exceptions WHERE check_name = 'duplicate_invoice');

-- 2. Miscoded expenses: coded somewhere other than the supplier's usual account,
--    unless the recode carries an approval note.
INSERT INTO exceptions
SELECT 'miscoded_expense', 'invoice', i.invoice_id, i.period_posted, i.amount,
       'Coded ' || i.nominal || ', supplier default ' || s.default_nominal || ' (' || i.description || ')'
FROM invoices i JOIN suppliers s USING (supplier_id)
WHERE i.nominal <> s.default_nominal AND i.description NOT LIKE 'Approved recode%';

-- 3. Missed accruals: service delivered in an earlier period than the invoice was posted to,
--    with no accrual journal for that invoice in the service period.
INSERT INTO exceptions
SELECT 'missed_accrual', 'invoice', i.invoice_id, substr(i.service_to, 1, 7), i.amount,
       'Service to ' || i.service_to || ', posted in ' || i.period_posted
FROM invoices i
WHERE i.period_posted > substr(i.service_to, 1, 7)
  AND NOT EXISTS (SELECT 1 FROM manual_journals j WHERE j.ref = i.invoice_no AND j.period = substr(i.service_to, 1, 7));

-- 4. Marketplace fee audit: fees charged above the contracted rate on net sales.
INSERT INTO exceptions
SELECT 'fee_overcharge', 'marketplace_line', line_id, substr(date, 1, 7),
       ROUND(fees - :fee_rate * (gross_sales - refunds), 2),
       settlement_id || ' on ' || date
FROM marketplace
WHERE fees - :fee_rate * (gross_sales - refunds) > 1.00;
