# =====================================================================
#  PAYROLL FRAUD AGENT  -  Streamlit app
#  Reads payroll data from a Google Sheet, runs 5 fraud checks,
#  and shows a ranked report of who to investigate first.
# =====================================================================

import io
import re
import urllib.request
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# SETTINGS  (the only part your group should need to change)
# ---------------------------------------------------------------------

# Paste your Google Sheet's link between the quotes.
# The sheet must be shared as "Anyone with the link - Viewer".
SHEET_URL = ""

# Last day of the pay period being reviewed (year-month-day)
PAY_PERIOD_END = "2026-10-16"

# Anyone over this many hours in the pay period gets flagged
HOURS_LIMIT = 90

# How often (in seconds) the app re-downloads the sheet on its own
REFRESH_SECONDS = 60


# ---------------------------------------------------------------------
# PAGE SETUP
# ---------------------------------------------------------------------
st.set_page_config(page_title="Payroll Fraud Agent", page_icon="🔎", layout="wide")


# ---------------------------------------------------------------------
# STEP 1: GET THE DATA
# ---------------------------------------------------------------------

def sheet_link_to_csv_link(link):
    """Turn a normal Google Sheet link into a 'download as CSV' link."""
    match = re.search(r"/spreadsheets/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", link)
    if not match:
        return None
    sheet_id = match.group(1)
    tab = re.search(r"[#&?]gid=(\d+)", link)          # which tab, if given
    gid = tab.group(1) if tab else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


@st.cache_data(ttl=REFRESH_SECONDS)                    # re-download every REFRESH_SECONDS
def load_from_sheet(csv_link):
    # Download the sheet like a normal web browser would
    request = urllib.request.Request(csv_link, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=20) as response:
        text = response.read().decode("utf-8-sig")
    # If Google sent back a web page instead of data, the sheet isn't public
    if text.lstrip().lower().startswith(("<!doctype", "<html")):
        raise PermissionError("Google sent back a sign-in page instead of the data, so the sheet isn't public yet.")
    return pd.read_csv(io.StringIO(text), header=None, dtype=str)


def clean_up(raw):
    """Find the header row, drop title/total rows, and fix number and date columns."""
    # Find the row that starts with "Employee ID" and use it as the column names
    header_row = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Employee ID"][0]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row]]
    df = df[df["Employee ID"].notna() & df["Employee ID"].astype(str).str.strip().ne("")]

    # Numbers: remove $ signs and commas, then turn into real numbers
    for col in ["Prior Rate ($/hr)", "Current Rate ($/hr)", "Hours Worked", "Gross Pay ($)"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(r"[$,]", "", regex=True), errors="coerce").fillna(0)

    # Dates: blank termination date = still employed
    df["Termination Date"] = pd.to_datetime(df["Termination Date"], errors="coerce")

    # Text: blanks become empty, bank accounts keep leading zeros
    for col in ["Manager", "SSN on File", "Home Address", "Rate Change Approved"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["Bank Acct (last 4)"] = df["Bank Acct (last 4)"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(4)
    return df.reset_index(drop=True)


# ---------------------------------------------------------------------
# STEP 2: THE FIVE FRAUD CHECKS
# Each check returns a list of findings: (employee ID, points, reason)
# ---------------------------------------------------------------------

def check_paid_after_termination(df, period_end):
    hits = df[(df["Termination Date"] < period_end) & (df["Gross Pay ($)"] > 0)]
    return [(p["Employee ID"], 3, f"Paid ${p['Gross Pay ($)']:,.2f} after being terminated on {p['Termination Date'].date()}")
            for _, p in hits.iterrows()]


def check_shared_bank(df):
    findings = []
    counts = df.groupby("Bank Acct (last 4)")["Employee ID"].transform("count")
    for account, group in df[counts > 1].groupby("Bank Acct (last 4)"):
        for _, p in group.iterrows():
            others = ", ".join(group[group["Employee ID"] != p["Employee ID"]]["Name"])
            findings.append((p["Employee ID"], 3, f"Shares bank account ending {account} with {others}"))
    return findings


def check_ghost_signs(df):
    findings = []
    for _, p in df.iterrows():
        if p["Manager"] == "":
            findings.append((p["Employee ID"], 2, "No manager listed"))
        if p["SSN on File"] != "Y":
            findings.append((p["Employee ID"], 2, "No SSN on file"))
        if "po box" in p["Home Address"].lower():
            findings.append((p["Employee ID"], 2, "PO Box instead of a home address"))
    return findings


def check_unapproved_rate(df):
    hits = df[(df["Current Rate ($/hr)"] != df["Prior Rate ($/hr)"]) & (df["Rate Change Approved"] != "Y")]
    return [(p["Employee ID"], 3, f"Pay rate changed from ${p['Prior Rate ($/hr)']:.2f} to ${p['Current Rate ($/hr)']:.2f}/hr with no approval")
            for _, p in hits.iterrows()]


def check_unusual_hours(df, limit):
    typical = df["Hours Worked"].median()
    hits = df[df["Hours Worked"] > limit]
    return [(p["Employee ID"], 1, f"Logged {p['Hours Worked']:.0f} hours (typical is {typical:.0f})")
            for _, p in hits.iterrows()]


def risk_level(points):
    if points >= 6:
        return "HIGH"
    if points >= 3:
        return "MEDIUM"
    return "LOW"


def run_all_checks(df, period_end, hours_limit):
    """Run every check, add up points per person, and rank them."""
    checks = {
        "Paid after termination": check_paid_after_termination(df, period_end),
        "Shared bank account": check_shared_bank(df),
        "Ghost employee signs": check_ghost_signs(df),
        "Unapproved rate change": check_unapproved_rate(df),
        "Unusual hours": check_unusual_hours(df, hours_limit),
    }
    people = {}
    for check_name, findings in checks.items():
        for emp_id, points, reason in findings:
            person = people.setdefault(emp_id, {"points": 0, "reasons": []})
            person["points"] += points
            person["reasons"].append(reason)

    rows = []
    for emp_id, info in people.items():
        p = df[df["Employee ID"] == emp_id].iloc[0]
        rows.append({
            "Employee ID": emp_id, "Name": p["Name"], "Job Title": p["Job Title"],
            "Department": p["Department"], "Points": info["points"],
            "Risk": risk_level(info["points"]), "Reasons": info["reasons"],
        })
    report = pd.DataFrame(rows)
    if not report.empty:
        report = report.sort_values(["Points", "Name"], ascending=[False, True]).reset_index(drop=True)
    return report, checks


# ---------------------------------------------------------------------
# STEP 3: THE SCREEN
# ---------------------------------------------------------------------

st.title("🔎 Payroll Fraud Agent")
st.caption("Reviews payroll for common fraud red flags and ranks who to look at first. "
           "Flags are leads for review, not proof of fraud.")

# Sidebar: where the data comes from + settings
with st.sidebar:
    st.header("Data source")
    link = st.text_input("Google Sheet link", value=SHEET_URL,
                         help="Share the sheet as 'Anyone with the link - Viewer', then paste the link here.")
    uploaded = st.file_uploader("…or upload a CSV instead", type="csv")
    st.header("Settings")
    period_end = pd.to_datetime(st.date_input("Pay period end", pd.to_datetime(PAY_PERIOD_END)))
    hours_limit = st.number_input("Flag hours above", value=HOURS_LIMIT, step=1)
    if st.button("🔄 Refresh data now"):
        load_from_sheet.clear()

# Load the data
raw, source = None, None
try:
    if uploaded is not None:
        raw, source = pd.read_csv(uploaded, header=None, dtype=str), f"Uploaded file: {uploaded.name}"
    elif link.strip():
        csv_link = sheet_link_to_csv_link(link)
        if csv_link is None:
            st.error("That doesn't look like a Google Sheet link. It should contain /spreadsheets/d/…")
            st.stop()
        raw, source = load_from_sheet(csv_link), "Google Sheet (refreshes automatically)"
except Exception as problem:
    st.error("Couldn't read the sheet. Check that it's shared as **Anyone with the link - Viewer**, "
             "and that the payroll tab is the one in the link.")
    # Show the technical reason so we can figure out what went wrong
    st.code(f"Details: {type(problem).__name__}: {problem}\nTried: {csv_link if 'csv_link' in dir() else link}")
    st.stop()

if raw is None:
    st.info("👈 Paste your Google Sheet link (or upload a CSV) in the sidebar to start.")
    st.stop()

try:
    payroll = clean_up(raw)
except Exception:
    st.error("The sheet loaded, but I couldn't find the payroll columns. "
             "Make sure the first column header is **Employee ID** and the column names match the sample file.")
    st.stop()

report, checks = run_all_checks(payroll, period_end, hours_limit)

st.caption(f"Source: {source}  •  Pay period ending {period_end.date()}  •  "
           f"Last checked {datetime.now().strftime('%b %d, %Y %I:%M %p')}")

# Summary numbers across the top
risk_counts = report["Risk"].value_counts() if not report.empty else pd.Series(dtype=int)
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Employees reviewed", len(payroll))
c2.metric("Employees flagged", len(report))
c3.metric("🔴 High risk", int(risk_counts.get("HIGH", 0)))
c4.metric("🟠 Medium risk", int(risk_counts.get("MEDIUM", 0)))
c5.metric("🟡 Low risk", int(risk_counts.get("LOW", 0)))

st.divider()

tab_report, tab_checks, tab_data = st.tabs(["📋 Fraud report", "🧪 Results by check", "📄 Payroll data"])

# Tab 1: ranked report, one card per person
with tab_report:
    if report.empty:
        st.success("✅ No red flags found this pay period.")
    else:
        icons = {"HIGH": "🔴", "MEDIUM": "🟠", "LOW": "🟡"}
        show = st.multiselect("Show risk levels", ["HIGH", "MEDIUM", "LOW"], default=["HIGH", "MEDIUM", "LOW"])
        for rank, row in report[report["Risk"].isin(show)].iterrows():
            with st.container(border=True):
                left, right = st.columns([3, 1])
                left.markdown(f"**#{rank + 1}  {row['Name']}** ({row['Employee ID']}) · {row['Job Title']}, {row['Department']}")
                right.markdown(f"{icons[row['Risk']]} **{row['Risk']}** · {row['Points']} pts")
                for reason in row["Reasons"]:
                    st.markdown(f"- {reason}")

        download = report.assign(Reasons=report["Reasons"].str.join("; "))
        st.download_button("⬇️ Download report (CSV)", download.to_csv(index=False),
                           file_name=f"fraud_report_{period_end.date()}.csv", mime="text/csv")

    with st.expander("How the scoring works"):
        st.markdown(
            "- **3 points:** paid after termination, shared bank account, or unapproved rate change\n"
            "- **2 points:** each ghost-employee warning sign (no manager, no SSN, PO Box address)\n"
            "- **1 point:** unusual hours (could be real overtime)\n\n"
            "**Risk:** 6+ points = HIGH, 3–5 = MEDIUM, 1–2 = LOW."
        )

# Tab 2: what each check found on its own
with tab_checks:
    for check_name, findings in checks.items():
        names = {e: payroll.loc[payroll["Employee ID"] == e, "Name"].iloc[0] for e, _, _ in findings}
        label = f"{check_name}: {len(set(names))} employee(s)"
        with st.expander(("⚠️ " if findings else "✅ ") + label, expanded=bool(findings)):
            if findings:
                st.dataframe(pd.DataFrame([{"Employee ID": e, "Name": names[e], "Finding": r} for e, _, r in findings]),
                             hide_index=True, width="stretch")
            else:
                st.write("Nothing found.")

# Tab 3: the raw payroll data
with tab_data:
    st.dataframe(payroll, hide_index=True, width="stretch")
