"""Import bank statement PDFs into the Monkey Money workbook.

  python statement_import.py import   # read PDFs in .\statements -> 'Import Review' sheet
  python statement_import.py post     # move reviewed rows (Post? = Yes) into Budget Tracking
  python statement_import.py dry-run  # show what would be imported, touch nothing

Optional: --workbook "path\\to\\file.xlsx"  --inbox "path\\to\\statements"

If the workbook is already open in Excel the rows are written straight into it
(you save it yourself). Otherwise it is opened in the background and saved.
"""
import argparse
import collections
import datetime as dt
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from statement_parsers import parse_file  # noqa: E402

DEFAULT_WORKBOOK = os.path.join(ROOT, "Monkey Money.xlsx")
DEFAULT_INBOX = os.path.join(ROOT, "statements")

NAVY, ACCENT, INPUT, WHITE, YELLOW = 0x352B22, 0xDBA98E, 0xF2E1D9, 0xFFFFFF, 0x99FFFF
SOURCES = ["Any", "Maybank Savings", "Maybank Card", "CIMB Savings", "CIMB Card"]
ACTIONS = "Add,Skip,Savings"
CARD_LUMP_CATEGORIES = ("MAYBANK Shopee Credit Card", "CIMB Travel Credit Card")

# Contains (| = or), Source, Direction, Action, Type, Category, Note.  First match wins.
# Action: Add = record it | Skip = ignore | Savings = only record the money moving in/out of savings
DEFAULT_RULES = [
    ("THANK YOU", "CIMB Card", "In", "Skip", "", "", "Card bill payment - card spending is imported line by line instead"),
    ("PYMT@MAYBANK2U", "Maybank Card", "In", "Skip", "", "", "Card bill payment - card spending is imported line by line instead"),
    ("MAYBANK VISA CARD", "Maybank Savings", "Out", "Skip", "", "", "Card bill payment - card spending is imported line by line instead"),
    # --- personal rules: replace the CHANGE-ME text with what appears on YOUR statements ---
    ("CHANGE-ME EMPLOYER NAME", "Any", "In", "Add", "Income", "Salary", "Your salary - put your employer's name as it appears on the statement"),
    ("CHANGE-ME LANDLORD NAME", "Any", "Out", "Add", "Expenses", "Housing", "Rent / mortgage - put the payee name"),
    ("CHANGE-ME CAR PLATE", "Any", "Out", "Add", "Expenses", "Transportation", "Car loan - put the reference you use for the instalment"),
    # --- general rules ---
    ("MTRUSTEE", "Any", "In", "Add", "Income", "Investments", "Withdrawal from a fund held by MTrustee (e.g. Bursa Gold Dinar)"),
    ("CREDIT INTEREST|HIBAH|PROFIT PAID", "Any", "In", "Add", "Income", "Others", "Bank interest"),
    ("I-PYMT TO CCARD", "CIMB Savings", "Out", "Savings", "", "", "CIMB card bill paid from savings - the card's own lines carry the spending"),
    ("PAYLATER|SPAYLATER", "Any", "Out", "Add", "Expenses", "Debt Payments", "Buy-now-pay-later repayment"),
    ("RETAIL INTEREST|FINANCE CHARGE|LATE CHARGE|ANNUAL FEE", "Any", "Out", "Add", "Expenses", "Debt Payments", "Card charges"),
    ("[FOREIGN]", "Any", "Out", "Add", "Expenses", "Entertainment", "Spent abroad - trips go under Entertainment"),
    ("AIRASIA|AGODA|BOOKING.COM|HOTEL|MALAYSIA AIRLINES", "Any", "Out", "Add", "Expenses", "Entertainment", "Travel"),
    ("APPLE.COM/BILL|MAXIS|CELCOM|DIGI|TIMEDOTCOM|TNB|PERBADANAN BEKALAN|AIR SELANGOR|INDAH WATER|HOTLINK|UNIFI|NETFLIX|SPOTIFY|PRIMEVIDEO|XBOX|YOUTUBE", "Any", "Out", "Add", "Expenses", "Utilities", "Bills and subscriptions"),
    ("MCDONALD|KFC|NASI KANDAR|NASIKANDAR|SUSHI|JOLLIBEE|JOLIBEE|MIXUE|JAYA GROCER|SPEEDMART|KK MART|MYNEWS|TACO BELL|RESTORAN|CAFE|STARBUCKS|ZUS COFFEE|ZUSPRESSO|TEALIVE|BOOST JUICE|7-ELEVEN|FAMILYMART|LOTUS|AEON|MYDIN|ECONSAVE|VILLAGE GROCER|GRABFOOD|FOODPANDA", "Any", "Out", "Add", "Expenses", "Groceries", "Food and groceries"),
    ("SHELL|PETRONAS|PETRON |CALTEX|BHPETROL|GRAB RIDES|PARKING|TOUCH N GO|EXPRESS RAIL LINK|RAPID KL|PLUS MALAYSIA", "Any", "Out", "Add", "Expenses", "Transportation", "Fuel, rides, tolls, parking"),
    ("UNIQLO|PUMA|NIKE|H & M|ZARA|WATSON|GUARDIAN|DECATHLON|DENTAL|CLINIC|PHARMACY|HOSPITAL", "Any", "Out", "Add", "Expenses", "Personal Care", "Clothes, health, grooming"),
    ("GSC|TGV|STEAM|PLAYSTATION|NINTENDO|MUSIC", "Any", "Out", "Add", "Expenses", "Entertainment", ""),
    ("MAE QR", "Maybank Savings", "In", "Add", "Income", "Debt Payments", "Small transfer in - usually a friend paying you back"),
]


# ---------------------------------------------------------------- rules
def rule_matches(rule, txn, text):
    contains, source, direction = rule[0], rule[1] or "Any", rule[2] or "Any"
    if source != "Any" and source != txn["source"]:
        return False
    if direction == "In" and txn["amount"] <= 0:
        return False
    if direction == "Out" and txn["amount"] >= 0:
        return False
    alts = [a.strip().upper() for a in str(contains).split("|") if a.strip()]
    return any(a in text for a in alts)


def to_rows(txn, rules, own_name):
    """Turn one bank transaction into 0-2 Budget Tracking rows."""
    text = (txn["match_text"] + (" [FOREIGN]" if txn.get("foreign") else "")).upper()
    amt = txn["amount"]
    base = dict(date=txn["date"], details=txn["details"][:60], source=txn["source"], file=txn["file"],
                bank_text=f"{txn['description']} | {txn['details']} {txn.get('memo', '')}".strip())
    hit = next((r for r in rules if rule_matches(r, txn, text)), None)
    own = bool(own_name) and own_name.upper() in text

    if hit and hit[3] == "Skip":
        return [], f"skipped: {hit[6] or hit[0]}"

    if hit and hit[3] == "Savings":
        # only the movement of money into (+) / out of (-) savings
        value = round(amt if txn["source"] == "CIMB Savings" else -amt, 2)
        return [dict(base, type="Savings", category="Personal Savings", amount=value, post="Yes",
                     note=hit[6] or "Savings movement")], hit[0]

    if hit is None and own:
        # transfer between your own accounts
        if txn["source"] == "CIMB Savings":
            return [dict(base, type="Savings", category="Personal Savings", amount=round(amt, 2),
                         post="Yes", note="Own transfer into/out of CIMB savings")], "own-transfer"
        # Not matched to a CIMB line (those are skipped automatically) -> another account of yours
        # (Wise, MooMoo, TnG...). 'Savings' would be wrong: in this workbook Savings = the CIMB account.
        if amt > 0:
            typ, cat = "Income", "Investments"
            note = "CHECK: money in from another account in your name (Wise/MooMoo/TnG?) - Income/Investments"
        else:
            typ, cat = "Expenses", "Investment"
            note = "CHECK: money out to another account in your name (Wise/MooMoo/TnG?) - Expenses/Investment"
        return [dict(base, type=typ, category=cat, amount=round(abs(amt), 2), post="",
                     note=note + ". Use Savings only for CIMB.", own_transfer=True)], "own-transfer?"

    if hit:
        typ, cat, note, post = hit[4], hit[5], hit[6], "Yes"
    elif txn["source"] == "Maybank Card" and amt > 0:
        typ, cat, note, post = "Expenses", "", "CHECK: card refund - pick the category of the original purchase", ""
    else:
        typ, cat, note, post = ("Income" if amt > 0 else "Expenses"), "", "CHECK: no rule matched - pick a category", ""

    value = round(amt if typ == "Income" else -amt, 2)  # expenses positive, refunds negative
    rows = [dict(base, type=typ, category=cat, amount=value, post=post, note=note)]
    if txn["source"] == "CIMB Savings" and typ in ("Income", "Expenses"):
        # money landed in / left the savings account directly: mirror it in the Savings column
        rows.append(dict(base, type="Savings", category="Personal Savings", amount=round(amt, 2), post=post,
                         note="Paired with the row above (happened inside CIMB savings)"))
    return rows, (hit[0] if hit else "no rule")


# ---------------------------------------------------------------- Excel helpers
def open_workbook(path):
    import urllib.parse
    import pythoncom
    import win32com.client as w
    path = os.path.abspath(path)
    # Is it already open in any Excel window? OneDrive files are registered under
    # https://d.docs.live.net/... so compare the last folder + file name.
    tail = "/".join(path.replace("\\", "/").split("/")[-2:]).lower()
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for mk in rot:
            try:
                name = urllib.parse.unquote(mk.GetDisplayName(ctx, None)).replace("\\", "/").lower()
            except Exception:
                continue
            if name.endswith(tail):
                obj = rot.GetObject(mk)
                return w.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch)), False
    except Exception:
        pass
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl.Workbooks.Open(path), True


def close_workbook(wb, opened_here, save=True):
    if opened_here:
        xl = wb.Application
        name = wb.Name
        if save:
            wb.Save()
        xl.Workbooks(name).Close(False)
        xl.Quit()


def sheet(wb, name):
    try:
        return wb.Worksheets(name)
    except Exception:
        return None


def make_table(ws, top_left_row, headers, name, widths):
    for i, h in enumerate(headers):
        ws.Cells(top_left_row, 2 + i).Value = h
    rng = ws.Range(ws.Cells(top_left_row, 2), ws.Cells(top_left_row + 1, 1 + len(headers)))
    lo = ws.ListObjects.Add(1, rng, None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleLight9"
    for i, wd in enumerate(widths):
        ws.Columns(2 + i).ColumnWidth = wd
    return lo


def title(ws, text, sub, last_col):
    ws.Columns(1).ColumnWidth = 2
    r = ws.Range(ws.Cells(2, 2), ws.Cells(2, last_col))
    r.Interior.Color = NAVY
    r.Font.Color = WHITE
    r.Font.Bold = True
    r.Font.Size = 20
    ws.Rows(2).RowHeight = 34
    ws.Cells(2, 2).Value = text
    ws.Cells(3, 2).Value = sub
    ws.Cells(3, 2).Font.Color = 0x808080


def ensure_sheets(wb):
    after = wb.Worksheets("Budget Tracking")
    rules_ws = sheet(wb, "Import Rules")
    if rules_ws is None:
        rules_ws = wb.Worksheets.Add(After=wb.Worksheets("Setting"))
        rules_ws.Name = "Import Rules"
        title(rules_ws, "Import Rules", "How bank lines become categories. Checked top to bottom - the first match wins. "
              "'Contains' can list alternatives separated by |. Add your own rows; the importer reads this table every time.", 8)
        rules_ws.Range("B5").Value = "Your name as it appears on transfers:"
        rules_ws.Range("B5").Font.Bold = True
        rules_ws.Range("C5").Value = "YOUR NAME AS ON STATEMENTS"
        rules_ws.Range("C5").Interior.Color = INPUT
        rules_ws.Range("D5").Value = "Used to spot transfers between your own accounts."
        rules_ws.Range("D5").Font.Color = 0x808080
        lo = make_table(rules_ws, 7, ["Contains", "Source", "Direction", "Action", "Type", "Category", "Note"],
                        "ImportRules", [48, 17, 10, 9, 10, 26, 60])
        for r in DEFAULT_RULES[1:]:
            lo.ListRows.Add()
        lo.DataBodyRange.Value = [list(r) for r in DEFAULT_RULES]
        rules_ws.Tab.Color = 0xBFBFBF
    else:
        upgrade_rules(rules_ws.ListObjects("ImportRules"))
    lo = rules_ws.ListObjects("ImportRules")
    for col, lst in (("Source", ",".join(SOURCES)), ("Direction", "Any,In,Out"), ("Action", ACTIONS),
                     ("Type", "Income,Expenses,Savings")):
        rng = lo.ListColumns(col).DataBodyRange
        rng.Validation.Delete()
        rng.Validation.Add(3, 1, 1, lst)

    rev = sheet(wb, "Import Review")
    if rev is None:
        rev = wb.Worksheets.Add(After=after)
        rev.Name = "Import Review"
        title(rev, "Import Review", "New bank lines land here first. Fix any yellow row (pick Type + Category, set Post? "
              "to Yes), set Post? to No for anything you don't want, then run 'Post reviewed rows'.", 11)
        lo = make_table(rev, 5, ["Post?", "Date", "Type", "Category", "Amount", "Details", "Source", "Rule / note",
                                 "Bank text", "Statement"], "ImportReview", [7, 11, 10, 24, 11, 30, 15, 46, 50, 30])
        lo.ListColumns("Date").DataBodyRange.NumberFormat = "d mmm yyyy"
        lo.ListColumns("Amount").DataBodyRange.NumberFormat = "#,##0.00;[Red]-#,##0.00"
        rev.Tab.Color = NAVY

    log = sheet(wb, "Import Log")
    if log is None:
        log = wb.Worksheets.Add(After=wb.Worksheets("Import Rules"))
        log.Name = "Import Log"
        title(log, "Import Log", "Statements already imported (a file listed here is not read again).", 8)
        make_table(log, 5, ["File", "Account", "Imported on", "Bank lines", "Rows created", "Skipped as duplicate", "Check",
                            "Fingerprint"], "ImportLog", [40, 17, 18, 11, 13, 19, 60, 20])
        log.Tab.Color = 0xBFBFBF
    llo = log.ListObjects("ImportLog")
    if llo.ListColumns.Count < 8:  # older log without fingerprints
        llo.ListColumns.Add().Name = "Fingerprint"
        log.Columns(9).ColumnWidth = 20
    tidy_tabs(wb)
    # moving a sheet invalidates COM handles to it - fetch them again
    return sheet(wb, "Import Rules"), sheet(wb, "Import Review"), sheet(wb, "Import Log")


def tidy_tabs(wb):
    """Keep the tab bar clean: Import Review right after the main sheets, helper sheets hidden."""
    try:
        anchor = next((n for n in ("Investment", "Budget Tracking") if sheet(wb, n) is not None), None)
        rev = sheet(wb, "Import Review")
        if anchor and rev is not None and rev.Index != wb.Worksheets(anchor).Index + 1:
            rev.Move(None, wb.Worksheets(anchor))  # positional (Before, After); a bad call would move it to a new workbook
        for n in ("Import Rules", "Import Log", "Bank Check"):
            s = sheet(wb, n)
            if s is not None and s.Visible == -1 and wb.ActiveSheet.Name != n:
                s.Visible = 0
    except Exception as e:  # cosmetic only - never block an import
        print("(tab tidy skipped:", e, ")")


def upgrade_rules(lo):
    """Bring an existing rules table up to date without touching rules the user added or edited."""
    rows = table_rows(lo)
    have = {str(r[0]).strip().upper() for r in rows if r[0]}
    body = lo.DataBodyRange
    for i, r in enumerate(rows, start=1):
        # the CIMB card bill used to be booked as a lump expense; it is now a savings movement only
        if str(r[0]).strip().upper() == "I-PYMT TO CCARD" and r[3] == "Add" and r[5] == "CIMB Travel Credit Card":
            new = next(d for d in DEFAULT_RULES if d[0] == "I-PYMT TO CCARD")
            lo.Parent.Range(body.Cells(i, 1), body.Cells(i, 7)).Value = [list(new)]
    missing = [d for d in DEFAULT_RULES if d[0].upper() not in have]
    for d in missing:  # new defaults go to the top so they win over broad rules
        lo.ListRows.Add(1)
        lo.Parent.Range(lo.DataBodyRange.Cells(1, 1), lo.DataBodyRange.Cells(1, 7)).Value = [list(d)]
        print(f"Import Rules: added rule '{d[0]}' ({d[1]})")


def fingerprint(path):
    import hashlib
    with open(path, "rb") as f:
        return hashlib.sha1(f.read()).hexdigest()[:16]


def table_rows(lo):
    if lo.DataBodyRange is None:
        return []
    vals = lo.DataBodyRange.Value
    return [list(r) for r in vals]


def serial(d):
    """Excel date serial - avoids pywin32 shifting naive datetimes by the timezone offset."""
    return float((d - dt.date(1899, 12, 30)).days)


def to_date(v):
    if isinstance(v, (int, float)) and v > 30000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if hasattr(v, "year"):
        return dt.date(v.year, v.month, v.day)
    return None


def append_rows(lo, rows):
    """Append rows (lists) to a ListObject. Returns the first new row index (1-based within table)."""
    if not rows:
        return
    existing = lo.ListRows.Count
    first_blank = existing == 1 and all(v in (None, "") for v in lo.DataBodyRange.Value[0])
    start = 1 if first_blank else existing + 1
    need = start + len(rows) - 1 - existing
    for _ in range(max(0, need)):
        lo.ListRows.Add()
    ncol = len(rows[0])
    body = lo.DataBodyRange
    rng = lo.Parent.Range(body.Cells(start, 1), body.Cells(start + len(rows) - 1, ncol))
    rng.Value2 = rows
    return start


def category_lists(wb):
    """Category names per Type, read from the Budget Planning tables Income / Expenses / Savings."""
    out = {}
    for t in ("Income", "Expenses", "Savings"):
        vals = []
        for ws in wb.Worksheets:
            try:
                lo = ws.ListObjects(t)
            except Exception:
                continue
            vals = [str(v[0]).strip() for v in lo.ListColumns(1).DataBodyRange.Value
                    if v[0] not in (None, "") and not str(v[0]).startswith("Enter ")]
            break
        out[t] = vals
    return out


# ---------------------------------------------------------------- duplicates
class DupIndex:
    """Finds rows that are already in Budget Tracking / Import Review.

    Imported rows carry the merchant in Details, so they only match a row with the same date, type,
    amount AND details (two real RM12.90 meals on the same day at different places stay separate).
    Rows you typed by hand usually have no Details, so those match on date + type + amount.
    """

    def __init__(self):
        self.strict, self.loose = collections.Counter(), collections.Counter()

    @staticmethod
    def _norm(s):
        return re.sub(r"\s+", " ", str(s or "")).strip().upper()[:60]

    def add(self, d, typ, amt, details):
        if not d or amt in (None, ""):
            return
        key = (d, typ, round(float(amt), 2))
        det = self._norm(details)
        if det:
            self.strict[key + (det,)] += 1
        else:
            self.loose[key] += 1

    def add_tracking(self, rows):
        for r in rows:
            self.add(to_date(r[0]), r[1], r[3], r[4] if len(r) > 4 else None)

    def take(self, d, typ, amt, details):
        key = (d, typ, round(float(amt), 2))
        sk = key + (self._norm(details),)
        if self.strict[sk] > 0:
            self.strict[sk] -= 1
            return True
        if self.loose[key] > 0:
            self.loose[key] -= 1
            return True
        return False


# ---------------------------------------------------------------- bank check
BANK_ACCOUNTS = {"Maybank Savings": "Balance", "CIMB Savings": "Savings"}


def ensure_bank_check(wb):
    ws = sheet(wb, "Bank Check")
    if ws is not None:
        return ws.ListObjects("BankCheck")
    ws = wb.Worksheets.Add(After=wb.Worksheets("Import Review"))
    ws.Name = "Bank Check"
    title(ws, "Bank Check", "Does the tracker agree with the bank? Bank balances are read from your statements; "
          "the tracker figure is Budget Tracking on the same date.", 9)
    lo = make_table(ws, 5, ["Account", "Tracker column", "As of", "Bank balance", "Tracker balance", "Difference",
                            "Status", "Read from"], "BankCheck", [17, 15, 12, 14, 15, 13, 16, 34])
    lo.ListRows.Add()
    lo.DataBodyRange.Value = [[a, c, None, None, None, None, None, None] for a, c in BANK_ACCOUNTS.items()]
    m = 'MATCH(MAXIFS(Tracking[Date],Tracking[Date],"<="&[@[As of]]),Tracking[Date],0)'
    lo.ListColumns("Tracker balance").DataBodyRange.Formula = (
        f'=IF([@[As of]]="","",IF([@[Tracker column]]="Savings",INDEX(Tracking[Savings],{m}),INDEX(Tracking[Balance],{m})))')
    lo.ListColumns("Difference").DataBodyRange.Formula = '=IF([@[As of]]="","",ROUND([@[Tracker balance]]-[@[Bank balance]],2))'
    lo.ListColumns("Status").DataBodyRange.Formula = (
        '=IF([@[As of]]="","no statement yet",IF(ABS([@Difference])<0.01,"Matches",'
        'IF([@Difference]>0,"Tracker too high","Tracker too low")))')
    lo.ListColumns("As of").DataBodyRange.NumberFormat = "d mmm yyyy"
    for c in ("Bank balance", "Tracker balance", "Difference"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0.00;[Red]-#,##0.00"
    st = lo.ListColumns("Status").DataBodyRange
    for word, color in (("Matches", 0xCEEFC6), ("too", 0xCEC7FF)):
        fc = st.FormatConditions.Add(2, None, f'=ISNUMBER(SEARCH("{word}",{st.Cells(1, 1).Address.replace("$", "")}))')
        fc.Interior.Color = color
    notes = [
        "If it doesn't match, the usual reasons are:",
        "1. Rows still waiting in Import Review (not posted yet) on or before the 'As of' date.",
        "2. A transfer posted with the wrong Type - 'Savings' means money in/out of CIMB only.",
        "   Transfers to Wise / MooMoo / TnG are Expenses > Investment; money back from them is Income > Investments.",
        "3. Maybank only: card spending you have not paid yet (the tracker counts card spending when it happens).",
        "4. A transaction typed by hand AND imported (duplicate) - search Budget Tracking for the amount.",
    ]
    for i, n in enumerate(notes):
        ws.Cells(10 + i, 2).Value = n
        ws.Cells(10 + i, 2).Font.Color = 0x808080
    ws.Cells(10, 2).Font.Bold = True
    ws.Tab.Color = NAVY
    tidy_tabs(wb)
    return lo


def refresh_bank_check(wb, inbox, new_results):
    """Store the newest closing balance per account (from new files and any statements still in the inbox)."""
    lo = ensure_bank_check(wb)
    closings = {}
    seen = {f for f, *_ in new_results}
    files = [(f, k, ts) for f, k, ts, _, _ in new_results]
    for f in os.listdir(inbox):
        if f.lower().endswith(".pdf") and f not in seen:
            kind, ts, _, _ = parse_file(os.path.join(inbox, f))
            if kind in BANK_ACCOUNTS:
                files.append((f, kind, ts))
    for f, kind, ts in files:
        c = getattr(ts, "closing", None)
        if kind in BANK_ACCOUNTS and c and (kind not in closings or c[0] > closings[kind][0]):
            closings[kind] = (c[0], c[1], f)
    body = lo.DataBodyRange
    for i, r in enumerate(table_rows(lo), start=1):
        acc = r[0]
        if acc not in closings:
            continue
        asof, bal, f = closings[acc]
        stored = to_date(r[2])
        if stored and stored > asof:
            continue  # keep a newer figure from a statement you've since deleted
        body.Cells(i, 3).Value2 = serial(asof)
        body.Cells(i, 4).Value = bal
        body.Cells(i, 8).Value = f


def print_bank_check(wb):
    wb.Application.Calculate()
    lo = sheet(wb, "Bank Check").ListObjects("BankCheck")
    print("\nBank check (tracker vs bank):")
    for r in table_rows(lo):
        if r[2] in (None, ""):
            print(f"  {r[0]:16} no statement yet")
            continue
        d = to_date(r[2])
        print(f"  {r[0]:16} {d:%d %b %Y}: bank {float(r[3]):>10,.2f} | tracker {float(r[4]):>10,.2f} | {r[6]}"
              + ("" if r[6] == "Matches" else f" (difference {float(r[5]):,.2f})"))


# ---------------------------------------------------------------- commands
def collect(inbox, done_prints):
    """Parse every PDF whose content has not been imported before (banks reuse names like CIMBClicks.pdf)."""
    results = []
    for f in sorted(os.listdir(inbox)):
        if not f.lower().endswith(".pdf"):
            continue
        fp = fingerprint(os.path.join(inbox, f))
        if fp in done_prints:
            continue
        kind, txns, ok, check = parse_file(os.path.join(inbox, f))
        results.append((f, kind, txns, ok, check))
        done_prints[fp] = f
    return results


def build_candidates(results, rules, own_name):
    cimb = [t for _, k, ts, _, _ in results if k == "CIMB Savings" for t in ts
            if own_name and own_name.upper() in t["match_text"].upper()]
    out, skipped = [], collections.Counter()
    card_payments = []
    for f, kind, txns, ok, check in results:
        made = 0
        for t in txns:
            rows, why = to_rows(t, rules, own_name)
            if not rows and t["source"].endswith("Card") and t["amount"] > 0:
                card_payments.append(t)
            for r in rows:
                if r.pop("own_transfer", False):
                    # Maybank-side own transfer that mirrors a CIMB line -> CIMB side already records it
                    if any(abs(c["amount"] + t["amount"]) < 0.005 and abs((c["date"] - t["date"]).days) <= 2 for c in cimb):
                        skipped[f] += 1
                        continue
                out.append(r)
                made += 1
            if not rows:
                skipped[f] += 1
        yield_info = (f, kind, len(txns), made, ok, check)
        out.append(("__file__", yield_info))
    return out, skipped, card_payments


def retire_card_lumps(rev_lo, card_payments):
    """A card bill that was booked as one lump expense is replaced by the card's own lines once
    that card's history is imported - mark the lump 'No' so nothing is counted twice."""
    rows = table_rows(rev_lo)
    body = rev_lo.DataBodyRange
    n = 0
    for i, r in enumerate(rows, start=1):
        d, typ, cat, amt = to_date(r[1]), r[2], r[3], r[4]
        if typ != "Expenses" or cat not in CARD_LUMP_CATEGORIES or str(r[0]).lower() == "no" or amt in (None, ""):
            continue
        for p in card_payments:
            if abs(float(amt) - p["amount"]) < 0.005 and abs((p["date"] - d).days) <= 3:
                body.Cells(i, 1).Value = "No"
                body.Cells(i, 8).Value = f"Replaced by the itemised lines from {p['file']} - not posted"
                body.Rows(i).Interior.Pattern = -4142
                n += 1
                break
    return n


def run_import(args, dry=False):
    os.makedirs(args.inbox, exist_ok=True)
    if dry:
        rules, own, done, existing_keys = DEFAULT_RULES, "", {}, DupIndex()
        wb = None
    else:
        wb, opened_here = open_workbook(args.workbook)
    success = False
    try:
        if wb is not None:
            rules_ws, rev, log = ensure_sheets(wb)
            rules = [tuple("" if v is None else str(v) for v in r) for r in table_rows(rules_ws.ListObjects("ImportRules"))
                     if r[0]]
            own = str(rules_ws.Range("C5").Value or "").strip()
            llo = log.ListObjects("ImportLog")
            done = {}
            for i, r in enumerate(table_rows(llo), start=1):
                if not r[0]:
                    continue
                fp = r[7] if len(r) > 7 else None
                p = os.path.join(args.inbox, str(r[0]))
                if not fp and os.path.exists(p):
                    # log entry from before fingerprints existed: fill it in, but only if the file with that
                    # name is still the same statement (a new CIMBClicks.pdf may have overwritten it)
                    kind, txns, _, _ = parse_file(p)
                    if kind == r[1] and len(txns) == int(r[3] or -1):
                        fp = fingerprint(p)
                        llo.DataBodyRange.Cells(i, 8).Value = fp
                if fp:
                    done[str(fp)] = r[0]
            existing_keys = DupIndex()
            existing_keys.add_tracking(table_rows(wb.Worksheets("Budget Tracking").ListObjects("Tracking")))
            for r in table_rows(rev.ListObjects("ImportReview")):
                existing_keys.add(to_date(r[1]), r[2], r[4], r[5])

        results = collect(args.inbox, done)
        if wb is not None:
            refresh_bank_check(wb, args.inbox, results)
        if not results:
            print("No new PDF statements in", args.inbox)
            if wb is not None:
                print_bank_check(wb)
            success = True
            return
        prints = {f: fp for fp, f in done.items()}
        cands, skipped, card_payments = build_candidates(results, rules, own)
        new_rows, log_rows, dupes = [], [], collections.Counter()
        cur_file_rows = []
        for c in cands:
            if isinstance(c, tuple):
                f, kind, n, made, ok, check = c[1]
                kept = 0
                for r in cur_file_rows:
                    if existing_keys.take(r["date"], r["type"], r["amount"], r["details"]):
                        dupes[f] += 1
                        continue
                    new_rows.append(r)
                    kept += 1
                log_rows.append([f, kind or "?", dt.datetime.now().strftime("%d %b %Y %H:%M"), n, kept, dupes[f],
                                 ("OK - " if ok else "CHECK FAILED - ") + check, prints.get(f, "")])
                cur_file_rows = []
            else:
                cur_file_rows.append(c)

        order = {"Maybank Savings": 0, "CIMB Savings": 1, "Maybank Card": 2, "CIMB Card": 3}
        new_rows.sort(key=lambda r: (r["date"], order.get(r["source"], 9)))
        for lr in log_rows:
            print(f"{lr[0]}: {lr[1]}, {lr[3]} bank lines -> {lr[4]} new rows, {lr[5]} already tracked | {lr[6]}")
        todo = sum(1 for r in new_rows if r["post"] != "Yes")
        print(f"\n{len(new_rows)} rows for review, {todo} need a category.")
        if dry:
            for r in new_rows:
                print(f"  {'  ' if r['post']=='Yes' else '??'} {r['date']}  {r['type']:8} {r['category'][:24]:24} {r['amount']:>10,.2f}  {r['details'][:30]:30} | {r['note'][:50]}")
            return
        if any(not lr[6].startswith("OK") for lr in log_rows):
            print("\nWARNING: a statement failed its total check - compare it with the PDF before posting.")

        rev_lo = rev.ListObjects("ImportReview")
        retired = retire_card_lumps(rev_lo, card_payments)
        if retired:
            print(f"{retired} lump card-bill row(s) in Import Review set to 'No' - replaced by the card's own lines.")
        start = append_rows(rev_lo, [[r["post"], serial(r["date"]), r["type"], r["category"], r["amount"],
                                      r["details"], r["source"], r["note"], r["bank_text"][:120], r["file"]] for r in new_rows])
        if new_rows:
            body = rev_lo.DataBodyRange
            body.Columns(2).NumberFormat = "d mmm yyyy"
            for i, r in enumerate(new_rows):
                if r["post"] != "Yes":
                    body.Rows(start + i).Interior.Color = YELLOW
            rev_lo.ListColumns("Post?").DataBodyRange.Validation.Delete()
            rev_lo.ListColumns("Post?").DataBodyRange.Validation.Add(3, 1, 1, "Yes,No,Force")
            rev_lo.ListColumns("Type").DataBodyRange.Validation.Delete()
            rev_lo.ListColumns("Type").DataBodyRange.Validation.Add(3, 1, 1, "Income,Expenses,Savings")
            cat = rev_lo.ListColumns("Category").DataBodyRange
            cat.Validation.Delete()
            first = rev_lo.ListColumns("Type").DataBodyRange.Cells(1, 1).Address.replace("$", "")
            # same dropdown as Budget Tracking (Category list depends on Type in the cell to its left)
            trk_cat = wb.Worksheets("Budget Tracking").ListObjects("Tracking").ListColumns("Category").DataBodyRange.Cells(1, 1)
            try:
                trk_cat.Copy()
                cat.PasteSpecial(6)  # xlPasteValidation
                wb.Application.CutCopyMode = False
            except Exception as e:
                print("(could not add the Category dropdown:", e, ")")
        append_rows(log.ListObjects("ImportLog"), log_rows)
        rev.Activate()
        rev.Range("A1").Select()
        print_bank_check(wb)
        print("  (Rows still in Import Review are not in the tracker yet, so a difference is expected until you post.)")
        success = True
    finally:
        if wb is not None:
            close_workbook(wb, opened_here, save=success)
    print("\nDone. Open the 'Import Review' sheet, fix yellow rows, then run 'Post reviewed rows'.")
    if wb is not None and not opened_here:
        print("(Your workbook was open, so rows were written into it - save it in Excel.)")


def run_post(args):
    wb, opened_here = open_workbook(args.workbook)
    success = False
    try:
        rev = sheet(wb, "Import Review")
        if rev is None:
            print("No 'Import Review' sheet yet - run import first.")
            return
        cats = category_lists(wb)
        rev_lo = rev.ListObjects("ImportReview")
        rows = table_rows(rev_lo)
        post, keep_idx, problems = [], [], []
        already = DupIndex()
        already.add_tracking(table_rows(wb.Worksheets("Budget Tracking").ListObjects("Tracking")))
        for i, r in enumerate(rows, start=1):
            flag, d, typ, cat, amt, details = r[0], to_date(r[1]), r[2], r[3], r[4], r[5]
            if all(v in (None, "") for v in r):
                continue
            f = str(flag).strip().lower()
            if f == "no":
                keep_idx.append((i, "drop"))
                continue
            if f not in ("yes", "force"):
                keep_idx.append((i, "keep"))
                continue
            if not d or typ not in cats or amt in (None, "") or (cats[typ] and cat not in cats[typ]):
                problems.append(f"row {i}: {d} {typ} '{cat}' {amt} - fix Type/Category")
                keep_idx.append((i, "keep"))
                continue
            if f == "yes" and already.take(d, typ, amt, details):
                # e.g. you typed it into Budget Tracking by hand after importing
                problems.append(f"row {i}: {d:%d %b} {typ} {float(amt):,.2f} {details} - already in Budget Tracking "
                                f"(set Post? to Force to post anyway, or No to drop it)")
                rev_lo.DataBodyRange.Cells(i, 8).Value = "POSSIBLE DUPLICATE of a Budget Tracking row - Force to post anyway, No to drop"
                rev_lo.DataBodyRange.Rows(i).Interior.Color = YELLOW
                keep_idx.append((i, "keep"))
                continue
            post.append((i, [serial(d), typ, cat, float(amt), details or ""]))
        if problems:
            print("Not posted (fix these and run again):\n  " + "\n  ".join(problems))
        if not post:
            print("Nothing marked Post? = Yes.")
        else:
            trk = wb.Worksheets("Budget Tracking").ListObjects("Tracking")
            n0 = trk.ListRows.Count
            for _ in post:
                trk.ListRows.Add()
            body = trk.DataBodyRange
            ncols = body.Columns.Count
            # carry formula columns (Balance, Effective Date, Savings, ...) down from the last existing row
            tws = trk.Parent
            if ncols > 5:
                src = tws.Range(body.Cells(n0, 6), body.Cells(n0, ncols))
                dst = tws.Range(body.Cells(n0 + 1, 6), body.Cells(n0 + len(post), ncols))
                src.Copy(dst)
            rng = tws.Range(body.Cells(n0 + 1, 1), body.Cells(n0 + len(post), 5))
            rng.Value2 = [p[1] for p in post]
            dcol = tws.Range(body.Cells(n0 + 1, 1), body.Cells(n0 + len(post), 1))
            dcol.NumberFormat = body.Cells(n0, 1).NumberFormat
            print(f"Posted {len(post)} rows to Budget Tracking.")
        # remove posted and 'No' rows from the review table (bottom-up)
        remove = sorted([i for i, _ in post] + [i for i, k in keep_idx if k == "drop"], reverse=True)
        for i in remove:
            rev_lo.ListRows(i).Delete()
        left = rev_lo.ListRows.Count
        print(f"{left} rows still waiting in Import Review.")
        wb.Application.CalculateFull()
        refresh_bank_check(wb, args.inbox, [])
        print_bank_check(wb)
        success = True
    finally:
        close_workbook(wb, opened_here, save=success)
    if not opened_here:
        print("(Your workbook was open, so changes were made in it - save it in Excel.)")


def run_check(args):
    wb, opened_here = open_workbook(args.workbook)
    success = False
    try:
        refresh_bank_check(wb, args.inbox, [])
        print_bank_check(wb)
        success = True
    finally:
        close_workbook(wb, opened_here, save=success)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("command", choices=["import", "post", "check", "dry-run"], nargs="?", default="import")
    ap.add_argument("--workbook", default=DEFAULT_WORKBOOK)
    ap.add_argument("--inbox", default=DEFAULT_INBOX)
    args = ap.parse_args()
    if args.command == "post":
        run_post(args)
    elif args.command == "check":
        run_check(args)
    else:
        run_import(args, dry=args.command == "dry-run")


if __name__ == "__main__":
    main()
