# Fraud Review Agent

A class project for our Advanced Fraud graduate course. The app reads a company's records from one Google Sheet, runs fraud checks on five areas of the business, and then connects the clues to show which **people** show up in more than one area.

**All data used with this app is fake.** Never connect a sheet with real employee or company information.

## Pages

| Page | What it checks |
|------|----------------|
| Overview: connected clues | Links every red flag to the people involved and ranks who shows up in 2+ areas |
| Payroll | Paid after termination, shared bank accounts, ghost employee signs, unapproved raises, unusual hours |
| Vendor payment tracing | Three-way match (PO → receiving report → invoice → payment), duplicate payments, invoices just under the approval limit, vendors that match an employee, sequential invoice numbers, self-approval, Benford's Law |
| Bank activity | Structuring (cash just under $10,000), round-dollar wires, after-hours/weekend transfers, payments to employees' personal accounts, Benford's Law |
| Check register | Missing check numbers, voided checks that were cashed, altered amounts, checks to Cash, self-approved checks |
| Expense reports | Duplicate receipts, split purchases, weekend charges, expenses after termination, self-approved reports, claims just under the receipt limit |

Flags are leads for review, not proof of fraud.

## Data

One Google Sheet with these tabs (names must match exactly): Payroll, Vendors, Purchase Orders, Receiving, Invoices, Vendor Payments, Checks, Bank, Expenses. The app reads each tab by name and re-downloads them every 60 seconds. You can also upload the Excel workbook instead.

## Files

- `app.py`: the app. Settings (sheet link, limits) are at the top.
- `requirements.txt`: the free Python tools the app needs.
