# =====================================================================
#  FRAUD REVIEW AGENT  -  Streamlit app
#  Reads a company's records from one Google Sheet (one tab per record
#  type), runs fraud checks on each area, and connects the clues across
#  areas to show which people deserve a closer look.
# =====================================================================

import io
import re
import urllib.parse
import urllib.request
from datetime import datetime
from math import log10, sqrt

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# SETTINGS  (the only part your group should need to change)
# ---------------------------------------------------------------------

# Paste your Google Sheet's link between the quotes.
# The sheet must be shared as "Anyone with the link - Viewer".
SHEET_URL = ""

PAY_PERIOD_END = "2026-10-16"   # last day of the payroll period being reviewed
HOURS_LIMIT = 90                # payroll: flag anyone above this many hours
APPROVAL_LIMIT = 5000           # vendor invoices at or above this need the controller's approval
OVER_PO_TOLERANCE = 0.05        # invoices may run up to 5% over the PO (freight, tax)
RECEIPT_LIMIT = 75              # expenses at or above this need a receipt
PURCHASE_LIMIT = 500            # single expense purchases above this need pre-approval
CASH_REPORT_LIMIT = 10000       # banks report cash transactions of $10,000 or more
REFRESH_SECONDS = 60            # how often the app re-downloads the sheet

# The tab names the app looks for, and the columns each tab must have
TABS = {
    "Payroll": ["Employee ID", "Name", "Department", "Job Title", "Manager", "Hire Date", "Termination Date",
                "SSN on File", "Home Address", "Bank Acct (last 4)", "Prior Rate ($/hr)", "Current Rate ($/hr)",
                "Rate Change Approved", "Hours Worked", "Gross Pay ($)"],
    "Vendors": ["Vendor ID", "Vendor Name", "Address", "Tax ID on File", "Bank Acct (last 4)", "Date Added", "Created By"],
    "Purchase Orders": ["PO Number", "PO Date", "Vendor ID", "Vendor Name", "Description", "Qty Ordered",
                        "Unit Price ($)", "PO Total ($)", "Approved By"],
    "Receiving": ["Receiving Report", "PO Number", "Date Received", "Qty Received", "Received By"],
    "Invoices": ["Invoice Number", "Vendor ID", "Vendor Name", "Invoice Date", "PO Number", "Qty Billed",
                 "Invoice Amount ($)"],
    "Vendor Payments": ["Payment ID", "Payment Date", "Vendor ID", "Vendor Name", "Invoice Number", "Amount ($)",
                        "Method", "Check Number", "Approved By"],
    "Checks": ["Check Number", "Check Date", "Payee", "Amount ($)", "Memo", "Status", "Written By", "Approved By"],
    "Bank": ["Date/Time", "Type", "Description", "Direction", "Amount ($)", "Check Number",
             "Other Party Acct (last 4)", "Initiated By"],
    "Expenses": ["Report ID", "Employee ID", "Employee Name", "Expense Date", "Merchant", "Category", "Amount ($)",
                 "Receipt ID", "Approved By", "Date Submitted"],
}


# =====================================================================
# PART 1: LOADING THE DATA
# =====================================================================

def sheet_id_from(link):
    """Pull the sheet's ID out of a Google Sheet link."""
    match = re.search(r"/spreadsheets/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", link)
    return match.group(1) if match else None


@st.cache_data(ttl=REFRESH_SECONDS)
def load_tab_from_google(sheet_id, tab_name):
    """Download one tab (by its name) from a public Google Sheet."""
    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&headers=1"
           f"&sheet={urllib.parse.quote(tab_name)}")
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=20) as response:
        text = response.read().decode("utf-8-sig")
    if text.lstrip().lower().startswith(("<!doctype", "<html")):
        raise PermissionError("Google sent back a sign-in page, so the sheet isn't public yet.")
    return pd.read_csv(io.StringIO(text), dtype=str)


def load_tabs_from_workbook(uploaded_file):
    """Read every tab from an uploaded Excel workbook."""
    return pd.read_excel(uploaded_file, sheet_name=None, dtype=str)


def tidy(df):
    """Trim column names and drop completely empty rows."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)


def has_columns(df, tab):
    return df is not None and all(col in df.columns for col in TABS[tab])


# Small helpers that turn text from the sheet into the right kind of value
def money(col):
    return pd.to_numeric(col.astype(str).str.replace(r"[$,\s]", "", regex=True), errors="coerce").fillna(0.0)


def number(col):
    return pd.to_numeric(col.astype(str).str.replace(r"[,\s]", "", regex=True), errors="coerce")


def dates(col):
    return pd.to_datetime(col, format="mixed", errors="coerce")


def text(col):
    return col.fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})


def acct(col):
    """Bank account last-4: keep leading zeros, blanks stay blank."""
    t = text(col).str.replace(r"\.0$", "", regex=True)
    return t.where(t == "", t.str.zfill(4))


def prepare(raw):
    """Clean every tab so the checks can use it. Missing tabs become None."""
    d = {tab: (tidy(raw[tab]) if raw.get(tab) is not None else None) for tab in TABS}
    d = {tab: (df if has_columns(df, tab) else None) for tab, df in d.items()}

    if d["Payroll"] is not None:
        p = d["Payroll"]
        p = p[text(p["Employee ID"]) != ""].copy()
        for c in ["Prior Rate ($/hr)", "Current Rate ($/hr)", "Hours Worked", "Gross Pay ($)"]:
            p[c] = money(p[c])
        for c in ["Hire Date", "Termination Date"]:
            p[c] = dates(p[c])
        for c in ["Employee ID", "Name", "Department", "Job Title", "Manager", "SSN on File", "Home Address",
                  "Rate Change Approved"]:
            p[c] = text(p[c])
        p["Bank Acct (last 4)"] = acct(p["Bank Acct (last 4)"])
        d["Payroll"] = p.reset_index(drop=True)

    if d["Vendors"] is not None:
        v = d["Vendors"]
        for c in ["Vendor ID", "Vendor Name", "Address", "Tax ID on File", "Created By"]:
            v[c] = text(v[c])
        v["Bank Acct (last 4)"] = acct(v["Bank Acct (last 4)"])
        v["Date Added"] = dates(v["Date Added"])

    if d["Purchase Orders"] is not None:
        po = d["Purchase Orders"]
        for c in ["PO Number", "Vendor ID", "Vendor Name", "Description", "Approved By"]:
            po[c] = text(po[c])
        po["PO Date"] = dates(po["PO Date"])
        for c in ["Qty Ordered", "Unit Price ($)", "PO Total ($)"]:
            po[c] = money(po[c])

    if d["Receiving"] is not None:
        r = d["Receiving"]
        for c in ["Receiving Report", "PO Number", "Received By"]:
            r[c] = text(r[c])
        r["Date Received"] = dates(r["Date Received"])
        r["Qty Received"] = money(r["Qty Received"])

    if d["Invoices"] is not None:
        i = d["Invoices"]
        for c in ["Invoice Number", "Vendor ID", "Vendor Name", "PO Number"]:
            i[c] = text(i[c])
        i["Invoice Date"] = dates(i["Invoice Date"])
        i["Qty Billed"] = money(i["Qty Billed"])
        i["Invoice Amount ($)"] = money(i["Invoice Amount ($)"])

    if d["Vendor Payments"] is not None:
        vp = d["Vendor Payments"]
        for c in ["Payment ID", "Vendor ID", "Vendor Name", "Invoice Number", "Method", "Approved By"]:
            vp[c] = text(vp[c])
        vp["Payment Date"] = dates(vp["Payment Date"])
        vp["Amount ($)"] = money(vp["Amount ($)"])
        vp["Check Number"] = number(vp["Check Number"])

    if d["Checks"] is not None:
        c = d["Checks"]
        for col in ["Payee", "Memo", "Status", "Written By", "Approved By"]:
            c[col] = text(c[col])
        c["Check Number"] = number(c["Check Number"])
        c["Check Date"] = dates(c["Check Date"])
        c["Amount ($)"] = money(c["Amount ($)"])

    if d["Bank"] is not None:
        b = d["Bank"]
        for col in ["Type", "Description", "Direction", "Initiated By"]:
            b[col] = text(b[col])
        b["Date/Time"] = dates(b["Date/Time"])
        b["Amount ($)"] = money(b["Amount ($)"])
        b["Check Number"] = number(b["Check Number"])
        b["Other Party Acct (last 4)"] = acct(b["Other Party Acct (last 4)"])

    if d["Expenses"] is not None:
        e = d["Expenses"]
        for col in ["Report ID", "Employee ID", "Employee Name", "Merchant", "Category", "Receipt ID", "Approved By"]:
            e[col] = text(e[col])
        e["Expense Date"] = dates(e["Expense Date"])
        e["Date Submitted"] = dates(e["Date Submitted"])
        e["Amount ($)"] = money(e["Amount ($)"])
    return d


# =====================================================================
# PART 2: THE FRAUD CHECKS
# Every check returns "findings". A finding records:
#   area     - which part of the business (Payroll, Vendors, ...)
#   check    - which test found it
#   subject  - who or what it's about (an employee, a vendor, a check #)
#   people   - employees it points to (used to connect clues across areas)
#   points   - how serious it is
#   detail   - a plain-English explanation
# =====================================================================

def finding(area, check, subject, points, detail, people=()):
    return {"area": area, "check": check, "subject": subject, "points": points, "detail": detail,
            "people": [p for p in people if p]}


def fmt(x):
    return f"${x:,.2f}"


def day(x):
    return x.strftime("%b %d, %Y") if pd.notna(x) else "?"


# ---------------- PAYROLL ----------------
def payroll_checks(p, period_end, hours_limit):
    out = []
    A = "Payroll"
    p = p.copy()
    p["who"] = p["Name"] + " (" + p["Employee ID"] + ")"
    # 1. Paid after termination
    hits = p[(p["Termination Date"] < period_end) & (p["Gross Pay ($)"] > 0)]
    for _, r in hits.iterrows():
        out.append(finding(A, "Paid after termination", r["who"], 3,
                           f"Paid {fmt(r['Gross Pay ($)'])} after being terminated on {day(r['Termination Date'])}",
                           [r["Name"]]))
    # 2. Shared bank account
    counts = p.groupby("Bank Acct (last 4)")["Employee ID"].transform("count")
    for account, grp in p[counts > 1].groupby("Bank Acct (last 4)"):
        for _, r in grp.iterrows():
            others = ", ".join(grp[grp["Employee ID"] != r["Employee ID"]]["who"])
            out.append(finding(A, "Shared bank account", r["who"], 3,
                               f"Shares bank account ending {account} with {others}", [r["Name"]]))
    # 3. Ghost-employee warning signs
    for _, r in p.iterrows():
        if r["Manager"] == "":
            out.append(finding(A, "Ghost employee signs", r["who"], 2, "No manager listed", [r["Name"]]))
        if r["SSN on File"] != "Y":
            out.append(finding(A, "Ghost employee signs", r["who"], 2, "No SSN on file", [r["Name"]]))
        if "po box" in r["Home Address"].lower():
            out.append(finding(A, "Ghost employee signs", r["who"], 2, "PO Box instead of a home address", [r["Name"]]))
    # 4. Unapproved rate change
    hits = p[(p["Current Rate ($/hr)"] != p["Prior Rate ($/hr)"]) & (p["Rate Change Approved"] != "Y")]
    for _, r in hits.iterrows():
        out.append(finding(A, "Unapproved rate change", r["who"], 3,
                           f"Pay rate changed from {fmt(r['Prior Rate ($/hr)'])} to {fmt(r['Current Rate ($/hr)'])}/hr "
                           f"with no approval", [r["Name"]]))
    # 5. Unusual hours
    typical = p["Hours Worked"].median()
    for _, r in p[p["Hours Worked"] > hours_limit].iterrows():
        out.append(finding(A, "Unusual hours", r["who"], 1,
                           f"Logged {r['Hours Worked']:.0f} hours (typical is {typical:.0f})", [r["Name"]]))
    return out


# ---------------- VENDOR PAYMENT TRACING ----------------
def vendor_checks(d):
    out = []
    A = "Vendors"
    v, po, rec, inv, vp = d["Vendors"], d["Purchase Orders"], d["Receiving"], d["Invoices"], d["Vendor Payments"]
    payroll, chk = d["Payroll"], d["Checks"]
    vname = dict(zip(v["Vendor ID"], v["Vendor Name"]))
    name_of = lambda vid: vname.get(vid, vid)

    # V1. Payment with no invoice behind it
    inv_keys = set(zip(inv["Vendor ID"], inv["Invoice Number"]))
    for _, r in vp.iterrows():
        if (r["Vendor ID"], r["Invoice Number"]) not in inv_keys:
            out.append(finding(A, "Payment with no invoice", name_of(r["Vendor ID"]), 3,
                               f"{r['Payment ID']}: paid {fmt(r['Amount ($)'])} on {day(r['Payment Date'])} for invoice "
                               f"{r['Invoice Number'] or '(none)'}, but there is no such invoice on file"))

    # V2. Invoices with no purchase order (or a PO number that doesn't exist)
    po_numbers = set(po["PO Number"])
    no_po = inv[(inv["PO Number"] == "") | (~inv["PO Number"].isin(po_numbers))]
    for vid, grp in no_po.groupby("Vendor ID"):
        out.append(finding(A, "Invoice with no purchase order", name_of(vid), 2,
                           f"{len(grp)} invoice(s) totaling {fmt(grp['Invoice Amount ($)'].sum())} have no matching "
                           f"purchase order"))

    # V3 + V4. Three-way match: PO vs receiving report vs invoice
    received = rec.groupby("PO Number")["Qty Received"].sum()
    po_by_no = po.set_index("PO Number")
    for _, r in inv[inv["PO Number"].isin(po_numbers)].iterrows():
        got = received.get(r["PO Number"], 0)
        if got == 0:
            out.append(finding(A, "Goods not received", name_of(r["Vendor ID"]), 3,
                               f"Invoice {r['Invoice Number']} ({fmt(r['Invoice Amount ($)'])}) bills for "
                               f"{r['PO Number']}, but no receiving report shows the goods ever arrived"))
        elif r["Qty Billed"] > got:
            out.append(finding(A, "Billed more than received", name_of(r["Vendor ID"]), 3,
                               f"Invoice {r['Invoice Number']} bills for {r['Qty Billed']:.0f} units on "
                               f"{r['PO Number']}, but only {got:.0f} were received"))
        po_total = po_by_no.loc[r["PO Number"], "PO Total ($)"]
        po_total = po_total.iloc[0] if isinstance(po_total, pd.Series) else po_total
        if po_total > 0 and r["Invoice Amount ($)"] > po_total * (1 + OVER_PO_TOLERANCE):
            over = r["Invoice Amount ($)"] / po_total - 1
            out.append(finding(A, "Invoice over PO amount", name_of(r["Vendor ID"]), 2,
                               f"Invoice {r['Invoice Number']} is {fmt(r['Invoice Amount ($)'])}, {over:.0%} more "
                               f"than the approved PO ({fmt(po_total)})"))

    # V5. Same invoice paid twice (inside Vendor Payments, or again by a separate check)
    for (vid, no), grp in vp.groupby(["Vendor ID", "Invoice Number"]):
        if len(grp) > 1 and no:
            out.append(finding(A, "Duplicate payment", name_of(vid), 3,
                               f"Invoice {no} was paid {len(grp)} times ({', '.join(grp['Payment ID'])})"))
    if chk is not None:
        for _, c in chk[chk["Status"].str.lower() != "void"].iterrows():
            for no in re.findall(r"[A-Z]{2,4}-\d{3,}", c["Memo"]):
                earlier = vp[vp["Invoice Number"] == no]
                other = earlier[~((earlier["Method"].str.lower() == "check") & (earlier["Check Number"] == c["Check Number"]))]
                if not earlier.empty and not other.empty:
                    first = other.iloc[0]
                    out.append(finding(A, "Duplicate payment", name_of(first["Vendor ID"]), 3,
                                       f"Invoice {no} was paid by {first['Method']} on {day(first['Payment Date'])} "
                                       f"({fmt(first['Amount ($)'])}), then paid again with check "
                                       f"#{c['Check Number']:.0f} on {day(c['Check Date'])} ({fmt(c['Amount ($)'])})",
                                       [c["Written By"]]))

    # V6. Pattern of invoices just under the approval limit
    near = inv[(inv["Invoice Amount ($)"] >= APPROVAL_LIMIT * 0.9) & (inv["Invoice Amount ($)"] < APPROVAL_LIMIT)]
    for vid, grp in near.groupby("Vendor ID"):
        if len(grp) >= 3:
            out.append(finding(A, "Just under approval limit", name_of(vid), 3,
                               f"{len(grp)} invoices fall between {fmt(APPROVAL_LIMIT * 0.9)} and {fmt(APPROVAL_LIMIT)}, "
                               f"right under the amount that needs the controller's approval"))

    # V7. Vendor shares an address or bank account with an employee
    if payroll is not None:
        addr_owner = {a.lower(): n for a, n in zip(payroll["Home Address"], payroll["Name"]) if a}
        bank_owner = payroll[payroll["Bank Acct (last 4)"] != ""].groupby("Bank Acct (last 4)")["Name"].apply(list).to_dict()
        for _, r in v.iterrows():
            who = addr_owner.get(r["Address"].lower())
            if who:
                out.append(finding(A, "Vendor matches an employee", r["Vendor Name"], 4,
                                   f"Vendor address ({r['Address']}) is the home address of employee {who}", [who]))
            owners = bank_owner.get(r["Bank Acct (last 4)"])
            if owners:
                out.append(finding(A, "Vendor matches an employee", r["Vendor Name"], 4,
                                   f"Vendor is paid into bank account ending {r['Bank Acct (last 4)']}, the same account "
                                   f"{', '.join(owners)} gets paid into on payroll", owners))

    # V8. Sequential invoice numbers (a sign the company is the vendor's only customer)
    for vid, grp in inv.groupby("Vendor ID"):
        nums = sorted(int(m.group()) for m in (re.search(r"\d+$", x) for x in grp["Invoice Number"]) if m)
        if len(nums) >= 4 and max(b - a for a, b in zip(nums, nums[1:])) <= 2:
            out.append(finding(A, "Sequential invoice numbers", name_of(vid), 2,
                               f"{len(nums)} invoices numbered {nums[0]} through {nums[-1]} with no gaps. A real "
                               f"supplier bills many customers, so its invoice numbers usually jump around"))

    # V9. Same person set up the vendor AND approved its payments
    creator = dict(zip(v["Vendor ID"], v["Created By"]))
    self_ok = vp[vp.apply(lambda r: r["Approved By"] != "" and r["Approved By"] == creator.get(r["Vendor ID"]), axis=1)]
    for vid, grp in self_ok.groupby("Vendor ID"):
        who = creator[vid]
        out.append(finding(A, "Creator approved own vendor", name_of(vid), 3,
                           f"{who} added this vendor to the system and also approved {len(grp)} of its payments "
                           f"({fmt(grp['Amount ($)'].sum())}). No one else checked their work", [who]))

    # Extra info: vendor missing a tax ID (needed for a 1099)
    for _, r in v[v["Tax ID on File"] != "Y"].iterrows():
        out.append(finding(A, "No tax ID on file", r["Vendor Name"], 1,
                           f"No tax ID on file. Real businesses provide one; fake vendors often can't"))
    return out


# ---------------- BANK ACTIVITY ----------------
def bank_checks(d):
    out = []
    A = "Bank"
    b = d["Bank"].copy()
    payroll = d["Payroll"]
    b["who"] = b["Initiated By"].replace("", "Unknown")
    outgoing = b[b["Direction"].str.lower() == "out"]

    # B1. Structuring: 2+ cash transactions just under $10,000 within 7 days
    band = b[b["Type"].str.lower().str.contains("cash") & (b["Amount ($)"] >= CASH_REPORT_LIMIT * 0.8)
             & (b["Amount ($)"] < CASH_REPORT_LIMIT)].sort_values("Date/Time")
    for who, grp in band.groupby("who"):
        times = list(grp["Date/Time"])
        if any((times[j] - times[i]).days <= 7 for i in range(len(times)) for j in range(i + 1, len(times))):
            items = ", ".join(f"{fmt(a)} ({t:%b %d})" for a, t in zip(grp["Amount ($)"], grp["Date/Time"]))
            out.append(finding(A, "Possible structuring", who, 4,
                               f"{len(grp)} cash transactions just under {fmt(CASH_REPORT_LIMIT)} within a week: {items}. "
                               f"Splitting cash like this avoids the bank's $10,000 report", [who]))

    # B2. Round-dollar wires and cash withdrawals (skipping ones that repeat every month, like rent)
    risky = outgoing[outgoing["Type"].str.lower().isin(["wire out", "cash withdrawal"])].copy()
    risky["month"] = risky["Date/Time"].dt.month
    months_seen = risky.groupby(["Description", "Amount ($)"])["month"].transform("nunique")
    rounds = risky[(risky["Amount ($)"] >= 1000) & (risky["Amount ($)"] % 100 == 0) & (months_seen < 2)]
    for who, grp in rounds.groupby("who"):
        out.append(finding(A, "Round-dollar amounts", who, 1,
                           f"{len(grp)} round-dollar wire(s)/withdrawal(s) totaling {fmt(grp['Amount ($)'].sum())}. "
                           f"Real business bills rarely come out to an exact hundred", [who]))

    # B3. Money moved on a weekend or outside business hours (7 AM - 8 PM)
    moved = outgoing[outgoing["Type"].str.lower().isin(["wire out", "cash withdrawal", "ach out"])]
    odd = moved[(moved["Date/Time"].dt.weekday >= 5) | (moved["Date/Time"].dt.hour < 7) | (moved["Date/Time"].dt.hour >= 20)]
    for _, r in odd.iterrows():
        out.append(finding(A, "After-hours or weekend", r["who"], 2,
                           f"{r['Type']} of {fmt(r['Amount ($)'])} on {r['Date/Time']:%A, %b %d at %I:%M %p} "
                           f"(\"{r['Description']}\")", [r["Initiated By"]]))

    # B4. Company money sent to an account that belongs to an employee (not payroll)
    if payroll is not None:
        bank_owner = payroll[payroll["Bank Acct (last 4)"] != ""].groupby("Bank Acct (last 4)")["Name"].apply(list).to_dict()
        to_emp = outgoing[(outgoing["Type"].str.lower() != "payroll") & outgoing["Other Party Acct (last 4)"].isin(bank_owner)]
        for acc, grp in to_emp.groupby("Other Party Acct (last 4)"):
            owners = bank_owner[acc]
            senders = sorted(set(grp["Initiated By"]) - {""})
            out.append(finding(A, "Payment to an employee's account", f"Account ending {acc}", 4,
                               f"{len(grp)} non-payroll payment(s) totaling {fmt(grp['Amount ($)'].sum())} went to "
                               f"account ending {acc}, the personal paycheck account of {', '.join(owners)}. "
                               f"Sent by: {', '.join(senders) or 'unknown'}", owners + senders))
    return out


# ---------------- CHECK REGISTER vs BANK ----------------
def check_checks(d):
    out = []
    A = "Checks"
    c = d["Checks"].dropna(subset=["Check Number"]).copy()
    bank = d["Bank"]
    cleared = bank[bank["Check Number"].notna()].groupby("Check Number")["Amount ($)"].sum() if bank is not None else pd.Series(dtype=float)
    nums = set(c["Check Number"].astype(int))

    # C1 / C3. Missing check numbers (and whether they cleared the bank anyway)
    for n in range(min(nums), max(nums) + 1):
        if n not in nums:
            if n in cleared.index:
                out.append(finding(A, "Cleared but not in register", f"Check #{n}", 4,
                                   f"Check #{n} is missing from the check register, but the bank paid it for "
                                   f"{fmt(cleared[n])}. No one recorded who it was written to"))
            else:
                out.append(finding(A, "Missing check number", f"Check #{n}", 2,
                                   f"Check #{n} is missing from the register. Unused checks should be marked Void, "
                                   f"not skipped"))
    payroll_names = set(d["Payroll"]["Name"]) if d["Payroll"] is not None else set()
    for _, r in c.iterrows():
        n = int(r["Check Number"])
        tag = f"Check #{n}"
        paid = cleared.get(n)
        # C2. Marked void, but the bank paid it
        if r["Status"].lower() == "void" and paid is not None:
            out.append(finding(A, "Voided check was cashed", tag, 4,
                               f"Marked VOID in the register (payee: {r['Payee']}, memo: \"{r['Memo']}\"), but the bank "
                               f"paid it for {fmt(paid)}",
                               [r["Written By"]] + ([r["Payee"]] if r["Payee"] in payroll_names else [])))
        # C4. Bank paid a different amount than the register shows (altered check)
        if r["Status"].lower() != "void" and paid is not None and abs(paid - r["Amount ($)"]) > 0.01:
            out.append(finding(A, "Amount doesn't match bank", tag, 4,
                               f"Register says {fmt(r['Amount ($)'])} to {r['Payee']}, but the bank paid {fmt(paid)}. "
                               f"The check may have been altered", [r["Written By"]]))
        # C5. Made out to "Cash"
        if r["Payee"].strip().lower() == "cash":
            out.append(finding(A, "Check made out to Cash", tag, 2,
                               f"{fmt(r['Amount ($)'])} payable to \"Cash\" on {day(r['Check Date'])}. Anyone can "
                               f"cash it, so there's no record of who got the money", [r["Written By"]]))
        # C6. Same person wrote and approved it
        if r["Written By"] and r["Written By"] == r["Approved By"]:
            out.append(finding(A, "Wrote and approved own check", tag, 2,
                               f"{r['Written By']} both wrote and approved this check ({fmt(r['Amount ($)'])} to "
                               f"{r['Payee']})", [r["Written By"]]))
    return out


# ---------------- EXPENSE REPORTS ----------------
def expense_checks(d):
    out = []
    A = "Expenses"
    e = d["Expenses"]
    payroll = d["Payroll"]

    # E1. Same receipt used more than once
    with_receipt = e[e["Receipt ID"] != ""]
    for rid, grp in with_receipt.groupby("Receipt ID"):
        if grp["Report ID"].nunique() > 1:
            r = grp.iloc[0]
            out.append(finding(A, "Duplicate receipt", r["Employee Name"], 3,
                               f"Receipt {rid} ({r['Merchant']}, {fmt(r['Amount ($)'])}, {day(r['Expense Date'])}) was "
                               f"claimed on {grp['Report ID'].nunique()} different reports: {', '.join(grp['Report ID'].unique())}",
                               [r["Employee Name"]]))

    # E2. One purchase split into smaller ones to stay under the limit
    for (who, merchant, when), grp in e.groupby(["Employee Name", "Merchant", "Expense Date"]):
        if len(grp) >= 2 and (grp["Amount ($)"] < PURCHASE_LIMIT).all() and grp["Amount ($)"].sum() > PURCHASE_LIMIT:
            out.append(finding(A, "Split purchase", who, 3,
                               f"{len(grp)} charges at {merchant} on {day(when)} ({', '.join(fmt(a) for a in grp['Amount ($)'])}), "
                               f"each under {fmt(PURCHASE_LIMIT)} but {fmt(grp['Amount ($)'].sum())} total", [who]))

    # E3. Weekend charges
    wk = e[e["Expense Date"].dt.weekday >= 5]
    for who, grp in wk.groupby("Employee Name"):
        out.append(finding(A, "Weekend charges", who, 1,
                           f"{len(grp)} charge(s) on a Saturday or Sunday ({fmt(grp['Amount ($)'].sum())}). Often fine "
                           f"(travel, trade shows) but worth a look", [who]))

    # E4. Expenses after the employee left the company
    if payroll is not None:
        term = dict(zip(payroll["Employee ID"], payroll["Termination Date"]))
        late = e[e.apply(lambda r: pd.notna(term.get(r["Employee ID"])) and r["Expense Date"] > term[r["Employee ID"]], axis=1)]
        for who, grp in late.groupby("Employee Name"):
            t = term[grp.iloc[0]["Employee ID"]]
            out.append(finding(A, "Expenses after termination", who, 4,
                               f"{len(grp)} expense(s) totaling {fmt(grp['Amount ($)'].sum())} dated after their "
                               f"termination on {day(t)}", [who]))

    # E5. Approved their own report
    for rep, grp in e[e["Approved By"] == e["Employee Name"]].groupby("Report ID"):
        who = grp.iloc[0]["Employee Name"]
        out.append(finding(A, "Self-approved report", who, 2,
                           f"{who} approved their own report {rep} ({fmt(grp['Amount ($)'].sum())})", [who]))

    # E6. Pattern of no-receipt claims just under the receipt limit
    near = e[(e["Receipt ID"] == "") & (e["Amount ($)"] >= RECEIPT_LIMIT * 0.93) & (e["Amount ($)"] < RECEIPT_LIMIT)]
    for who, grp in near.groupby("Employee Name"):
        if len(grp) >= 3:
            out.append(finding(A, "Just under receipt limit", who, 2,
                               f"{len(grp)} claims between {fmt(RECEIPT_LIMIT * 0.93)} and {fmt(RECEIPT_LIMIT)} with no "
                               f"receipt ({', '.join(fmt(a) for a in grp['Amount ($)'])})", [who]))
    return out


# ---------------- BENFORD'S LAW ----------------
def benford(amounts):
    """Compare how often each first digit (1-9) shows up vs. what Benford's Law expects."""
    a = pd.Series(amounts)
    a = a[a >= 10]
    first = a.astype(str).str.lstrip("0.").str[0].astype(int)
    n = len(first)
    rows = []
    for dgt in range(1, 10):
        exp = log10(1 + 1 / dgt)
        obs = (first == dgt).mean() if n else 0
        z = (abs(obs - exp) - 1 / (2 * n)) / sqrt(exp * (1 - exp) / n) if n else 0
        rows.append({"Digit": str(dgt), "Actual %": round(obs * 100, 1), "Expected %": round(exp * 100, 1),
                     "Count": int((first == dgt).sum()), "Unusual": bool(z > 2.5 and obs > exp)})
    return pd.DataFrame(rows), n


# ---------------- PUTTING IT TOGETHER ----------------
def risk_level(points):
    return "HIGH" if points >= 6 else "MEDIUM" if points >= 3 else "LOW"


def rank(findings):
    """Group findings by subject, add up points, and sort most serious first."""
    groups = {}
    for f in findings:
        g = groups.setdefault(f["subject"], {"subject": f["subject"], "points": 0, "items": []})
        g["points"] += f["points"]
        g["items"].append(f)
    ranked = sorted(groups.values(), key=lambda g: (-g["points"], g["subject"]))
    for g in ranked:
        g["risk"] = risk_level(g["points"])
    return ranked


def connect_people(all_findings, payroll):
    """The 'agent' step: link every finding to the people involved, across all areas."""
    people = {}
    for f in all_findings:
        for name in set(f["people"]):
            g = people.setdefault(name, {"name": name, "points": 0, "areas": set(), "items": []})
            g["points"] += f["points"]
            g["areas"].add(f["area"])
            g["items"].append(f)
    role = dict(zip(payroll["Name"], payroll["Job Title"])) if payroll is not None else {}
    ranked = sorted(people.values(), key=lambda g: (-len(g["areas"]), -g["points"], g["name"]))
    for g in ranked:
        g["role"] = role.get(g["name"], "Not on payroll")
        g["risk"] = "HIGH" if g["points"] >= 8 or len(g["areas"]) >= 3 else "MEDIUM" if g["points"] >= 4 or len(g["areas"]) >= 2 else "LOW"
    return ranked


def run_everything(d, period_end, hours_limit):
    results = {}
    if d["Payroll"] is not None:
        results["Payroll"] = payroll_checks(d["Payroll"], period_end, hours_limit)
    if all(d[t] is not None for t in ["Vendors", "Purchase Orders", "Receiving", "Invoices", "Vendor Payments"]):
        results["Vendors"] = vendor_checks(d)
    if d["Bank"] is not None:
        results["Bank"] = bank_checks(d)
    if d["Checks"] is not None:
        results["Checks"] = check_checks(d)
    if d["Expenses"] is not None:
        results["Expenses"] = expense_checks(d)
    return results


# =====================================================================
# PART 3: THE SCREEN
# =====================================================================
if __name__ == "__main__":
    st.set_page_config(page_title="Fraud Review Agent", page_icon="🔎", layout="wide")
    ICON = {"HIGH": "🔴", "MEDIUM": "🟠", "LOW": "🟡"}
    PAGES = {
        "🧭 Overview: connected clues": "Overview",
        "👥 Payroll": "Payroll",
        "🧾 Vendor payment tracing": "Vendors",
        "🏦 Bank activity": "Bank",
        "✍️ Check register": "Checks",
        "💳 Expense reports": "Expenses",
    }
    AREA_NAME = {v: k.split(" ", 1)[1] for k, v in PAGES.items()}

    # ---------- sidebar ----------
    with st.sidebar:
        st.header("🔎 Fraud Review Agent")
        page = PAGES[st.radio("Go to", list(PAGES))]
        st.divider()
        st.subheader("Data source")
        link = st.text_input("Google Sheet link", value=SHEET_URL,
                             help="Share the sheet as 'Anyone with the link - Viewer', then paste the link here.")
        uploaded = st.file_uploader("…or upload the Excel workbook", type=["xlsx"])
        if st.button("🔄 Refresh data now"):
            load_tab_from_google.clear()
        with st.expander("Settings"):
            period_end = pd.to_datetime(st.date_input("Payroll period end", pd.to_datetime(PAY_PERIOD_END)))
            hours_limit = st.number_input("Flag payroll hours above", value=HOURS_LIMIT, step=1)

    # ---------- load ----------
    raw, source, problems = {}, None, []
    if uploaded is not None:
        try:
            raw = load_tabs_from_workbook(uploaded)
            source = f"Uploaded workbook: {uploaded.name}"
        except Exception as problem:
            st.error(f"Couldn't read that workbook. Details: {problem}")
            st.stop()
    elif link.strip():
        sid = sheet_id_from(link)
        if sid is None:
            st.error("That doesn't look like a Google Sheet link. It should contain /spreadsheets/d/…")
            st.stop()
        for tab in TABS:
            try:
                raw[tab] = load_tab_from_google(sid, tab)
            except Exception as problem:
                problems.append(f"{tab}: {type(problem).__name__}: {problem}")
        if len(problems) == len(TABS):
            st.error("Couldn't read the sheet. Check that it's shared as **Anyone with the link - Viewer**.")
            st.code("Details:\n" + "\n".join(problems))
            st.stop()
        source = "Google Sheet (refreshes automatically)"
    else:
        st.title("🔎 Fraud Review Agent")
        st.info("👈 Paste your Google Sheet link (or upload the Excel workbook) in the sidebar to start.")
        st.stop()

    data = prepare(raw)
    missing = [t for t in TABS if data[t] is None]
    results = run_everything(data, period_end, hours_limit)
    everyone = [f for fs in results.values() for f in fs]

    def caption():
        st.caption(f"Source: {source}  •  Last checked {datetime.now().strftime('%b %d, %Y %I:%M %p')}  •  "
                   f"Flags are leads for review, not proof of fraud.")

    def need(*tabs):
        gone = [t for t in tabs if data[t] is None]
        if gone:
            st.warning(f"This page needs these tabs: **{', '.join(gone)}**. Check the tab names and column headers "
                       f"match the sample workbook exactly.")
            st.stop()

    def show_cards(ranked, unit="item"):
        if not ranked:
            st.success("✅ No red flags found.")
            return
        show = st.multiselect("Show risk levels", ["HIGH", "MEDIUM", "LOW"], default=["HIGH", "MEDIUM", "LOW"],
                              key=f"filter_{page}")
        for i, g in enumerate(ranked, start=1):
            if g["risk"] not in show:
                continue
            with st.container(border=True):
                left, right = st.columns([3, 1])
                left.markdown(f"**#{i}  {g['subject']}**")
                right.markdown(f"{ICON[g['risk']]} **{g['risk']}** · {g['points']} pts")
                for f in g["items"]:
                    st.markdown(f"- **{f['check']}:** {f['detail']}")

    def metrics(n_records, label, ranked):
        counts = pd.Series([g["risk"] for g in ranked]).value_counts()
        c = st.columns(5)
        c[0].metric(label, n_records)
        c[1].metric("Flagged", len(ranked))
        c[2].metric("🔴 High", int(counts.get("HIGH", 0)))
        c[3].metric("🟠 Medium", int(counts.get("MEDIUM", 0)))
        c[4].metric("🟡 Low", int(counts.get("LOW", 0)))

    def by_check(findings):
        if not findings:
            st.write("Nothing found.")
            return
        df = pd.DataFrame(findings)
        for check, grp in df.groupby("check", sort=False):
            with st.expander(f"⚠️ {check}: {len(grp)} finding(s)"):
                st.dataframe(grp[["subject", "points", "detail"]].rename(
                    columns={"subject": "Who / what", "points": "Points", "detail": "Finding"}),
                    hide_index=True, width="stretch")

    def download(findings, name):
        if findings:
            df = pd.DataFrame(findings)
            df["people"] = df["people"].str.join("; ")
            st.download_button("⬇️ Download findings (CSV)", df.to_csv(index=False), file_name=name, mime="text/csv")

    def benford_panel(amounts, what):
        table, n = benford(amounts)
        st.markdown(f"**Benford's Law** says that in natural financial data, about 30% of amounts start with 1 and "
                    f"under 5% start with 9. Made-up numbers usually don't follow that pattern. Tested on **{n} {what}**.")
        st.bar_chart(table.set_index("Digit")[["Actual %", "Expected %"]], stack=False)
        odd = table[table["Unusual"]]
        if odd.empty:
            st.success("✅ The first digits look natural.")
        else:
            for _, r in odd.iterrows():
                st.warning(f"⚠️ Amounts starting with **{r['Digit']}** show up {r['Actual %']}% of the time "
                           f"(expected {r['Expected %']}%). Look at which transactions start with {r['Digit']}.")
        st.dataframe(table, hide_index=True, width="stretch")

    def scoring(text_md):
        with st.expander("How the scoring works"):
            st.markdown(text_md + "\n\n**Risk:** 6+ points = HIGH, 3–5 = MEDIUM, 1–2 = LOW.")

    # ---------- OVERVIEW ----------
    if page == "Overview":
        st.title("🧭 Connected clues")
        st.write("Each area of the business is checked on its own pages. This page is the agent's big-picture step: "
                 "it links every red flag back to the **people** involved, so you can see who keeps showing up "
                 "in more than one place.")
        caption()
        if missing:
            st.warning(f"Missing or unreadable tabs: {', '.join(missing)}. Those areas are skipped.")
        cols = st.columns(len(results) or 1)
        for col, (area, fs) in zip(cols, results.items()):
            col.metric(AREA_NAME[area], f"{len(rank(fs))} flagged")
        st.divider()
        linked = connect_people(everyone, data["Payroll"])
        multi = [g for g in linked if len(g["areas"]) >= 2]
        st.subheader(f"People who show up in 2 or more areas ({len(multi)})")
        if not multi:
            st.success("✅ No one shows up in more than one area.")
        for i, g in enumerate(multi, start=1):
            with st.container(border=True):
                left, right = st.columns([3, 1])
                left.markdown(f"**#{i}  {g['name']}** · {g['role']}")
                right.markdown(f"{ICON[g['risk']]} **{g['risk']}** · {g['points']} pts")
                st.markdown("Shows up in: " + " · ".join(f"**{AREA_NAME[a]}**" for a in results if a in g["areas"]))
                for area in [a for a in results if a in g["areas"]]:
                    for f in [f for f in g["items"] if f["area"] == area]:
                        st.markdown(f"- *{AREA_NAME[area]}* · **{f['check']}:** {f['detail']}")
        single = [g for g in linked if len(g["areas"]) == 1]
        with st.expander(f"People flagged in only one area ({len(single)})"):
            st.dataframe(pd.DataFrame([{"Name": g["name"], "Role": g["role"], "Area": AREA_NAME[next(iter(g["areas"]))],
                                        "Points": g["points"], "Flags": "; ".join(f["check"] for f in g["items"])}
                                       for g in single]), hide_index=True, width="stretch")
        st.caption("Risk here: HIGH = 8+ points or 3+ areas, MEDIUM = 4+ points or 2 areas.")

    # ---------- PAYROLL ----------
    elif page == "Payroll":
        need("Payroll")
        st.title("👥 Payroll")
        st.write("Paid after termination · shared bank accounts · ghost employee signs · unapproved raises · unusual hours")
        caption()
        fs = results["Payroll"]; ranked = rank(fs)
        metrics(len(data["Payroll"]), "Employees reviewed", ranked)
        t1, t2, t3 = st.tabs(["📋 Report", "🧪 By check", "📄 Data"])
        with t1:
            show_cards(ranked)
            download(fs, "payroll_findings.csv")
            scoring("- **3:** paid after termination, shared bank account, unapproved rate change\n"
                    "- **2:** each ghost sign (no manager, no SSN, PO Box)\n- **1:** unusual hours")
        with t2: by_check(fs)
        with t3: st.dataframe(data["Payroll"], hide_index=True, width="stretch")

    # ---------- VENDORS ----------
    elif page == "Vendors":
        need("Vendors", "Purchase Orders", "Receiving", "Invoices", "Vendor Payments")
        st.title("🧾 Vendor payment tracing")
        st.write("Follows every payment back through its paperwork: **purchase order → receiving report → invoice → "
                 "payment**, and checks each vendor against the employee list.")
        caption()
        fs = results["Vendors"]; ranked = rank(fs)
        metrics(len(data["Vendors"]), "Vendors reviewed", ranked)
        t1, t2, t3, t4 = st.tabs(["📋 Report", "🧪 By check", "📊 Benford's Law", "📄 Data"])
        with t1:
            show_cards(ranked)
            download(fs, "vendor_findings.csv")
            scoring("- **4:** vendor matches an employee's address or bank account\n"
                    "- **3:** payment with no invoice, goods not received / billed more than received, duplicate payment, "
                    "pattern just under the approval limit, vendor creator approved its payments\n"
                    "- **2:** invoice with no PO, invoice more than 5% over PO, sequential invoice numbers\n"
                    "- **1:** no tax ID on file")
        with t2: by_check(fs)
        with t3: benford_panel(data["Invoices"]["Invoice Amount ($)"], "vendor invoices")
        with t4:
            for tab in ["Vendors", "Purchase Orders", "Receiving", "Invoices", "Vendor Payments"]:
                with st.expander(tab):
                    st.dataframe(data[tab], hide_index=True, width="stretch")

    # ---------- BANK ----------
    elif page == "Bank":
        need("Bank")
        st.title("🏦 Bank activity")
        st.write("Cash just under $10,000 (structuring) · round-dollar wires · after-hours and weekend transfers · "
                 "company money sent to employees' personal accounts · Benford's Law")
        caption()
        fs = results["Bank"]; ranked = rank(fs)
        metrics(len(data["Bank"]), "Transactions reviewed", ranked)
        t1, t2, t3, t4 = st.tabs(["📋 Report", "🧪 By check", "📊 Benford's Law", "📄 Data"])
        with t1:
            show_cards(ranked)
            download(fs, "bank_findings.csv")
            scoring("- **4:** possible structuring, payment to an employee's personal account\n"
                    "- **2:** each after-hours or weekend transfer\n- **1:** round-dollar wires/withdrawals")
        with t2: by_check(fs)
        with t3:
            b = data["Bank"]
            benford_panel(b[(b["Direction"].str.lower() == "out") & (b["Type"].str.lower() != "payroll")]["Amount ($)"],
                          "outgoing payments (payroll left out)")
        with t4: st.dataframe(data["Bank"], hide_index=True, width="stretch")

    # ---------- CHECKS ----------
    elif page == "Checks":
        need("Checks", "Bank")
        st.title("✍️ Check register")
        st.write("Compares the company's check register to what the bank actually paid: missing check numbers, "
                 "voided checks that got cashed, altered amounts, checks to Cash, and people approving their own checks.")
        caption()
        fs = results["Checks"]; ranked = rank(fs)
        metrics(len(data["Checks"]), "Checks reviewed", ranked)
        c = data["Checks"]; b = data["Bank"]
        outstanding = c[(c["Status"].str.lower() != "void") & (~c["Check Number"].isin(b["Check Number"]))]
        st.caption(f"{len(outstanding)} issued check(s) haven't cleared the bank yet (normal for recent checks).")
        t1, t2, t3 = st.tabs(["📋 Report", "🧪 By check", "📄 Data"])
        with t1:
            show_cards(ranked)
            download(fs, "check_findings.csv")
            scoring("- **4:** cleared the bank but not in the register, voided check was cashed, amount doesn't match bank\n"
                    "- **2:** missing check number, check made out to Cash, wrote and approved own check")
        with t2: by_check(fs)
        with t3: st.dataframe(c, hide_index=True, width="stretch")

    # ---------- EXPENSES ----------
    elif page == "Expenses":
        need("Expenses")
        st.title("💳 Expense reports")
        st.write("Duplicate receipts · split purchases · weekend charges · expenses after someone left · "
                 "self-approved reports · claims just under the receipt limit")
        caption()
        fs = results["Expenses"]; ranked = rank(fs)
        metrics(data["Expenses"]["Employee Name"].nunique(), "Employees reviewed", ranked)
        t1, t2, t3 = st.tabs(["📋 Report", "🧪 By check", "📄 Data"])
        with t1:
            show_cards(ranked)
            download(fs, "expense_findings.csv")
            scoring("- **4:** expenses after termination\n- **3:** duplicate receipt, split purchase\n"
                    "- **2:** self-approved report, 3+ claims just under the receipt limit\n- **1:** weekend charges")
        with t2: by_check(fs)
        with t3: st.dataframe(data["Expenses"], hide_index=True, width="stretch")
