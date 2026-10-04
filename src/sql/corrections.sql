-- Correcting journals, driven only by what the checks found (never by the answer key).

-- Duplicates: reverse the second posting.
INSERT INTO gl
SELECT 'ADJ-DUP-' || g.ref, g.date, g.period, g.account, g.channel, -g.amount, 'adjustment', g.ref
FROM gl g JOIN exceptions e ON e.check_name = 'duplicate_invoice' AND g.source = 'invoice' AND g.ref = e.record_id;

-- Miscoding: move the cost to the supplier's default account.
INSERT INTO gl
SELECT 'ADJ-CODE-' || i.invoice_id, i.invoice_date, i.period_posted, i.nominal, i.channel, -i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN exceptions e ON e.check_name = 'miscoded_expense' AND e.record_id = i.invoice_id
UNION ALL
SELECT 'ADJ-CODE-' || i.invoice_id, i.invoice_date, i.period_posted, s.default_nominal,
       CASE WHEN s.default_nominal = '6200' THEN 'Website' ELSE 'Central' END, i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN suppliers s USING (supplier_id)
JOIN exceptions e ON e.check_name = 'miscoded_expense' AND e.record_id = i.invoice_id;

-- Missed accruals: accrue in the month the service was delivered, release in the month it was posted.
INSERT INTO gl
SELECT 'ADJ-ACR-' || i.invoice_id, i.service_to, substr(i.service_to, 1, 7), i.nominal, i.channel, i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN exceptions e ON e.check_name = 'missed_accrual' AND e.record_id = i.invoice_id
UNION ALL SELECT 'ADJ-ACR-' || i.invoice_id, i.service_to, substr(i.service_to, 1, 7), '2100', 'Central', -i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN exceptions e ON e.check_name = 'missed_accrual' AND e.record_id = i.invoice_id
UNION ALL SELECT 'ADJ-ACR-' || i.invoice_id, i.invoice_date, i.period_posted, i.nominal, i.channel, -i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN exceptions e ON e.check_name = 'missed_accrual' AND e.record_id = i.invoice_id
UNION ALL SELECT 'ADJ-ACR-' || i.invoice_id, i.invoice_date, i.period_posted, '2100', 'Central', i.amount, 'adjustment', i.invoice_id
FROM invoices i JOIN exceptions e ON e.check_name = 'missed_accrual' AND e.record_id = i.invoice_id;

-- Fee overcharges: take the excess out of fees and raise a claim against the marketplace.
INSERT INTO gl
SELECT 'ADJ-FEE-' || record_id, period || '-28', period, '6100', 'Marketplace', -amount, 'adjustment', record_id
FROM exceptions WHERE check_name = 'fee_overcharge'
UNION ALL SELECT 'ADJ-FEE-' || record_id, period || '-28', period, '1150', 'Central', amount, 'adjustment', record_id
FROM exceptions WHERE check_name = 'fee_overcharge';
