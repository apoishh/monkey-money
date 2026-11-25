"""Read the Monkey Money workbook into plain JSON for the local dashboard.

Read-only: the workbook is copied to a temp file first (works while it is open in Excel,
using the last saved version) and never written to.
"""
import datetime as dt
import os
import shutil
import subprocess
import tempfile
import warnings

import openpyxl

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(os.path.dirname(os.path.dirname(HERE)), "Monkey Money.xlsx")

PLATFORMS = [  # table name, display name, kind
    ("kwsp", "KWSP / EPF", "Retirement"),
    ("moomoo", "MooMoo (S&P 500 ETF)", "Stocks / ETF"),
    ("bursagold", "Bursa Gold Dinar", "Gold"),
    ("wise", "Wise", "Multi-currency"),
    ("asb", "ASB", "Unit trust"),
    ("tng", "Touch 'n Go GO+", "Money market"),
    ("stashaway", "StashAway", "ETF (robo)"),
]
CARD_LUMPS = ("MAYBANK Shopee Credit Card", "CIMB Travel Credit Card")

_cache = {"mtime": None, "data": None}


def _copy_locked(src, dst):
    try:
        shutil.copyfile(src, dst)
    except PermissionError:
        # Excel can hold a lock that Python's open() trips over; Windows' own copy shares it fine
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        f"Copy-Item -LiteralPath '{src}' -Destination '{dst}' -Force"], check=True, capture_output=True)


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def _num(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def _table(ws, name):
    ref = ws.tables[name].ref
    rows = list(ws[ref])
    hdr = [c.value for c in rows[0]]
    return hdr, [[c.value for c in r] for r in rows[1:]]


def _named(wb, name):
    dest = list(wb.defined_names[name].destinations)[0]
    return wb[dest[0]][dest[1].replace("$", "")].value


def load(path=WORKBOOK):
    mtime = os.path.getmtime(path)
    if _cache["mtime"] == mtime:
        return _cache["data"]
    tmp = os.path.join(tempfile.gettempdir(), "pbp_dashboard_copy.xlsx")
    _copy_locked(path, tmp)
    wb = openpyxl.load_workbook(tmp, data_only=True)
    data = build(wb)
    data["source"] = {"file": os.path.basename(path),
                      "saved": dt.datetime.fromtimestamp(mtime).strftime("%d %b %Y %H:%M")}
    _cache.update(mtime=mtime, data=data)
    return data


def build(wb):
    today = dt.date.today()

    # ---------------- transactions
    hdr, rows = _table(wb["Budget Tracking"], "Tracking")
    ix = {h: i for i, h in enumerate(hdr)}
    tx = []
    for r in rows:
        d = _d(r[ix["Date"]])
        amt = _num(r[ix["Amount"]])
        if not d or amt is None or not r[ix["Type"]]:
            continue
        tx.append({"date": d.isoformat(), "type": r[ix["Type"]], "category": r[ix["Category"]] or "",
                   "amount": amt, "details": str(r[ix["Details"]] or ""),
                   "balance": _num(r[ix["Balance"]]), "savings": _num(r[ix["Savings"]])})
    tx.sort(key=lambda t: t["date"])
    last = tx[-1] if tx else None
    last_date = last["date"] if last else None
    # Balance / Savings columns are cumulative per date: take any row on the latest date
    cash = last["balance"] if last else 0
    savings = last["savings"] if last else 0

    # ---------------- budget plan (Budget Planning: one row per category, one column per month)
    bp = wb["Budget Planning"]
    budget = {}
    for tname in ("Income", "Expenses", "Savings"):
        ref = bp.tables[tname].ref
        first_row = bp[ref][0][0].row
        last_row = bp[ref][-1][0].row
        col_of = {}
        for c in bp[first_row]:
            if isinstance(c.value, dt.datetime):
                col_of[c.value.strftime("%Y-%m")] = c.column
        for rr in range(first_row + 1, last_row + 1):
            cat = bp.cell(rr, 3).value
            if not cat or str(cat).startswith("Enter "):
                continue
            for ym, col in col_of.items():
                v = _num(bp.cell(rr, col).value)
                if v:
                    budget.setdefault(tname, {}).setdefault(ym, {})[str(cat).strip()] = v

    # ---------------- investments
    inv_ws = wb["Investment"]
    platforms = []
    for tname, label, kind in PLATFORMS:
        if tname not in inv_ws.tables:
            continue
        hdr, rows = _table(inv_ws, tname)
        series, last_invested, last_value, last_upd = [], None, None, None
        for r in rows:
            d = _d(r[0])
            if not d:
                continue
            total = _num(r[2])
            value = _num(r[3])
            if total is not None:
                last_invested = total
            if value is not None:
                last_value = value
                last_upd = d
            series.append({"date": d.isoformat(), "moved": _num(r[1]), "invested": total, "value": value})
        closed = (last_value or 0) == 0 and (last_invested or 0) <= 0
        platforms.append({"id": tname, "name": label, "kind": kind, "series": series,
                          "invested": last_invested or 0, "value": last_value or 0,
                          "updated": last_upd.isoformat() if last_upd else None, "closed": closed})

    # ---------------- assets & debts (Loans / Net Worth sheets)
    lo = wb["Loans"]
    loans = []
    for r in (3, 4):
        name = str(lo[f"B{r}"].value or "").strip()
        if not name:
            continue
        loans.append({"name": name[:1].upper() + name[1:],
                      "monthly": _num(lo[f"F{r}"].value), "outstanding": _num(lo[f"K{r}"].value) or 0,
                      "market_value": _num(lo[f"L{r}"].value) or 0})
    nw = wb["Net Worth"]
    possessions = [{"name": str(nw[f"B{r}"].value), "value": _num(nw[f"C{r}"].value)}
                   for r in range(6, 12) if nw[f"B{r}"].value and _num(nw[f"C{r}"].value)]
    cards = _num(nw["C20"].value) or 0

    # ---------------- settings & goals
    s = {k: _named(wb, k) for k in ("dob", "monthly_salary", "planned_monthly_saving", "target_savings_rate",
                                    "ef_target_months", "ef_target_date", "lt_growth", "epf_dividend")}
    dob = _d(s["dob"])
    sv = wb["Saving"]
    goals_extra = []
    for r in (3, 4):
        name = sv[f"I{r}"].value
        if name and _num(sv[f"K{r}"].value):
            goals_extra.append({"name": str(name).encode("ascii", "ignore").decode().strip(),
                                "target": _num(sv[f"K{r}"].value), "saved": _num(sv[f"L{r}"].value) or 0,
                                "date": _d(sv[f"P{r}"].value).isoformat() if _d(sv[f"P{r}"].value) else None})

    # ---------------- bank check
    bank = []
    if "Bank Check" in wb.sheetnames:
        hdr, rows = _table(wb["Bank Check"], "BankCheck")
        for r in rows:
            if r[0]:
                bank.append({"account": r[0], "asof": _d(r[2]).isoformat() if _d(r[2]) else None,
                             "bank": _num(r[3]), "tracker": _num(r[4]), "diff": _num(r[5]), "status": r[6]})

    # ---------------- excel's own net worth (for a consistency check on the page)
    excel_nw = None
    goals = []
    if "Wealth Summary" in wb.sheetnames:
        ws = wb["Wealth Summary"]
        excel_nw = _num(ws["C27"].value)
        # goals table on Wealth Summary (rows 48-52): name, measured by, target, ..., target date
        for r in range(48, 53):
            name, measured, target, when = ws[f"B{r}"].value, ws[f"C{r}"].value, _num(ws[f"D{r}"].value), _d(ws[f"G{r}"].value)
            if not name or not target or not when:
                continue
            m = str(measured or "").lower()
            goals.append({"name": str(name), "target": target, "date": when.isoformat(),
                          "measure": "nw" if "net worth" in m else "pot",
                          "kind": "emergency" if "emergency" in str(name).lower()
                          else ("stacked" if "saving sheet" in m else "plain")})

    return {
        "today": today.isoformat(),
        "transactions": tx,
        "last_transaction": last_date,
        "cash": cash, "savings": savings,
        "budget": budget,
        "platforms": platforms,
        "loans": loans, "possessions": possessions, "cards": cards,
        "settings": {
            "dob": dob.isoformat() if dob else None,
            "monthly_salary": _num(s["monthly_salary"]),
            "planned_monthly_saving": _num(s["planned_monthly_saving"]),
            "target_savings_rate": s["target_savings_rate"],
            "ef_target_months": s["ef_target_months"],
            "ef_target_date": _d(s["ef_target_date"]).isoformat() if _d(s["ef_target_date"]) else None,
            "lt_growth": s["lt_growth"], "epf_dividend": s["epf_dividend"],
        },
        "goals_extra": goals_extra,
        "goals": goals,
        "bank_check": bank,
        "card_lump_categories": list(CARD_LUMPS),
        "excel_net_worth": excel_nw,
    }


if __name__ == "__main__":
    import json
    d = load()
    print(json.dumps({k: (v if k not in ("transactions", "platforms", "budget") else f"<{len(v)} items>")
                      for k, v in d.items()}, indent=1, default=str))
    for p in d["platforms"]:
        print(f"  {p['name']:24} invested {p['invested']:>10,.2f} value {p['value']:>10,.2f} updated {p['updated']} closed={p['closed']}")
    print("budget months:", {k: len(v) for k, v in d["budget"].items()}, "sample", d["budget"]["Expenses"].get("2026-10"))
