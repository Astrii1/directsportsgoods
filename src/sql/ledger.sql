-- Posts every cleaned source table to the general ledger.
-- One row per journal line. amount is signed: debit positive, credit negative.

CREATE TABLE accounts (account TEXT PRIMARY KEY, name TEXT, pl_line TEXT);
INSERT INTO accounts VALUES
  ('1000', 'Bank', NULL),
  ('1050', 'Card settlement clearing', NULL),
  ('1100', 'Marketplace receivable', NULL),
  ('1150', 'Marketplace fee claims', NULL),
  ('1200', 'Inventory', NULL),
  ('1500', 'Fixed assets at cost', NULL),
  ('1510', 'Accumulated depreciation', NULL),
  ('2000', 'Trade payables', NULL),
  ('2100', 'Accruals', NULL),
  ('2300', 'Payroll control', NULL),
  ('3000', 'Opening reserves', NULL),
  ('4000', 'Sales', 'Gross sales'),
  ('4100', 'Customer returns', 'Returns'),
  ('5000', 'Cost of goods sold', 'Cost of goods sold'),
  ('6000', 'Delivery and fulfilment', 'Delivery'),
  ('6100', 'Marketplace fees', 'Marketplace fees'),
  ('6200', 'Marketing', 'Marketing'),
  ('7000', 'Payroll', 'Payroll'),
  ('7100', 'Rent', 'Other overheads'),
  ('7110', 'Software', 'Other overheads'),
  ('7120', 'Utilities', 'Other overheads'),
  ('7130', 'Professional fees', 'Other overheads'),
  ('7140', 'Insurance', 'Other overheads'),
  ('7150', 'Repairs and maintenance', 'Other overheads'),
  ('7500', 'Depreciation', 'Depreciation');

CREATE TABLE gl (journal TEXT, date TEXT, period TEXT, account TEXT, channel TEXT, amount REAL, source TEXT, ref TEXT);

-- Opening balances at 31 January 2023, balanced against reserves.
INSERT INTO gl SELECT 'OPEN', '2023-01-31', '2023-01', account, 'Central', amount, 'opening', NULL FROM opening;
INSERT INTO gl SELECT 'OPEN', '2023-01-31', '2023-01', '3000', 'Central', -SUM(amount), 'opening', NULL FROM opening;

-- Website sales. Takings net of refunds go to card clearing until the processor pays out.
INSERT INTO gl
SELECT 'WEB-' || date, date, substr(date, 1, 7), '1050', 'Website', gross_sales - refunds, 'web_sales', date FROM web_sales
UNION ALL SELECT 'WEB-' || date, date, substr(date, 1, 7), '4000', 'Website', -gross_sales, 'web_sales', date FROM web_sales
UNION ALL SELECT 'WEB-' || date, date, substr(date, 1, 7), '4100', 'Website', refunds, 'web_sales', date FROM web_sales;

-- Marketplace statement lines. The marketplace keeps its fees and pays the rest per settlement.
INSERT INTO gl
SELECT 'MKT-' || line_id, date, substr(date, 1, 7), '1100', 'Marketplace', gross_sales - refunds - fees, 'marketplace', line_id FROM marketplace
UNION ALL SELECT 'MKT-' || line_id, date, substr(date, 1, 7), '4000', 'Marketplace', -gross_sales, 'marketplace', line_id FROM marketplace
UNION ALL SELECT 'MKT-' || line_id, date, substr(date, 1, 7), '4100', 'Marketplace', refunds, 'marketplace', line_id FROM marketplace
UNION ALL SELECT 'MKT-' || line_id, date, substr(date, 1, 7), '6100', 'Marketplace', fees, 'marketplace', line_id FROM marketplace;

-- Cost of sales from the warehouse dispatch report.
INSERT INTO gl
SELECT 'COGS-' || period || '-' || channel, period || '-28', period, '5000', channel, cost_of_goods_dispatched, 'dispatch', period FROM dispatch
UNION ALL SELECT 'COGS-' || period || '-' || channel, period || '-28', period, '1200', 'Central', -cost_of_goods_dispatched, 'dispatch', period FROM dispatch;

-- Supplier invoices, posted to the period the AP team chose.
INSERT INTO gl
SELECT 'INV-' || invoice_id, invoice_date, period_posted, nominal, channel, amount, 'invoice', invoice_id FROM invoices
UNION ALL SELECT 'INV-' || invoice_id, invoice_date, period_posted, '2000', 'Central', -amount, 'invoice', invoice_id FROM invoices;

-- Payroll cost by department, cleared when the payroll run is paid.
INSERT INTO gl
SELECT 'PAY-' || period, period || '-28', period, '7000', 'Central', SUM(total_cost), 'payroll', period FROM payroll GROUP BY period
UNION ALL SELECT 'PAY-' || period, period || '-28', period, '2300', 'Central', -SUM(total_cost), 'payroll', period FROM payroll GROUP BY period;

-- Month-end manual journals: depreciation and accruals.
INSERT INTO gl
SELECT 'MJ-' || period || '-' || journal || ref, period || '-28', period, debit, channel, amount, 'manual', journal FROM manual_journals
UNION ALL SELECT 'MJ-' || period || '-' || journal || ref, period || '-28', period, credit, channel, -amount, 'manual', journal FROM manual_journals;

-- Bank statement. Each line clears the balance it settles.
INSERT INTO gl
SELECT 'BANK-' || rowid, date, substr(date, 1, 7), '1000', 'Central', amount, 'bank', description FROM bank
UNION ALL
SELECT 'BANK-' || rowid, date, substr(date, 1, 7),
       CASE WHEN description LIKE 'CARDPAY%' THEN '1050'
            WHEN description LIKE 'MKTPLACE%' THEN '1100'
            WHEN description LIKE 'PAYROLL%' THEN '2300'
            ELSE '2000' END,
       'Central', -amount, 'bank', description FROM bank;
