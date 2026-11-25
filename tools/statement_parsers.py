"""Read bank statement PDFs into plain transaction records.

Supported layouts (auto-detected from the PDF text):
  - Maybank / Maybank Islamic savings account statement   -> source "Maybank Savings"
  - Maybank credit card statement                          -> source "Maybank Card"
  - CIMB Clicks "Account Details and Transaction History"  -> source "CIMB Savings"

Each parser also returns a checksum result so a mis-read line is caught
instead of silently dropped.
"""
import datetime as dt
import os
import re
import shutil
import subprocess

PDFTOTEXT_CANDIDATES = [
    shutil.which("pdftotext"),
    r"C:\Program Files\Git\mingw64\bin\pdftotext.exe",
    r"C:\Program Files (x86)\Git\mingw64\bin\pdftotext.exe",
]


def _pdftotext():
    for p in PDFTOTEXT_CANDIDATES:
        if p and os.path.exists(p):
            return p
    raise RuntimeError("pdftotext not found. It ships with Git for Windows (mingw64\\bin).")


def pdf_text(path, mode="-table"):
    out = subprocess.run([_pdftotext(), mode, "-enc", "UTF-8", path, "-"],
                         capture_output=True, check=True)
    return out.stdout.decode("utf-8", errors="replace")


def money(s):
    return round(float(s.replace(",", "").replace(" ", "")), 2)


class Txn(dict):
    """date, amount (+ in / - out of that account), description, details, source, file, extra"""


class TxnList(list):
    """List of Txn; .closing = (as-of date, bank balance) when the statement shows one."""
    closing = None


def detect(text):
    if "STATEMENT OF CREDIT CARD ACCOUNT" in text:
        return "Maybank Card"
    if "Account Details and Transaction History" in text and "Posting Date" in text and "Credit Card" in text:
        return "CIMB Card"
    if "Account Details and Transaction History" in text and "MYR" in text:
        return "CIMB Savings"
    if "Maybank" in text and ("SAVINGS ACCOUNT" in text or "ACCOUNT TRANSACTIONS" in text):
        return "Maybank Savings"
    return None


# ---------------------------------------------------------------- Maybank savings
_MBB_NOISE = re.compile(
    r"(Maybank Islamic|15th Floor|IBS |MUKA|TARIKH|STATEMENT DATE|NOMBOR|ACCOUNT$|NUMBER|"
    r"PIDM|Perhation|Semua|tempoh|All items|Sila|Please|Faedah|Interest refers|URUSNIAGA|ENTRY DATE|"
    r"TRANSACTION DESCRIPTION|^\s*:|^\s*\d{1,2}\s*$|^\s*\"\s*$|\s{5,}:\s*\d|\s{5,}:\s*$|[\u3000-\u9fff])")
_MBB_TX = re.compile(r"(\d{2}/\d{2}/\d{2})\s+(\S.*?)\s{2,}([\d,]+\.\d{2})([+-])(?:\s+([\d,]+\.\d{2}))?\s*$")
_DETAIL_JUNK = re.compile(r"^(SALE DEBIT|MBB CT-?|DUITNOW QR-?|MAE QR|Transfer|Fund Transfer|fund transfer|Interbank GIRO|"
                          r"QR\d+|\d{6,}Q?|(?=[A-Z]*\d)[A-Z0-9_]{12,}|\d{10,}_\w*|00|XXXX.*)$", re.I)


def parse_maybank_savings(path):
    text = pdf_text(path)
    lines = text.splitlines()
    txns, cur = [], None
    begin = end = tot_cr = tot_dr = None
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        m = re.search(r"BEGINNING BALANCE\s+([\d,]+\.\d{2})", line)
        if m:
            begin = money(m.group(1)); continue
        m = re.search(r"ENDING BALANCE\s*:\s*([\d,]+\.\d{2})", line)
        if m:
            end = money(m.group(1)); cur = None; continue
        m = re.search(r"TOTAL CREDIT\s*:\s*([\d,]+\.\d{2})", line)
        if m:
            tot_cr = money(m.group(1)); continue
        m = re.search(r"TOTAL DEBIT\s*:\s*([\d,]+\.\d{2})", line)
        if m:
            tot_dr = money(m.group(1)); continue
        m = _MBB_TX.search(line)
        if m:
            d = dt.datetime.strptime(m.group(1), "%d/%m/%y").date()
            amt = money(m.group(3)) * (1 if m.group(4) == "+" else -1)
            cur = Txn(date=d, amount=amt, description=m.group(2).strip(), lines=[],
                      balance=money(m.group(5)) if m.group(5) else None)
            txns.append(cur)
            continue
        if cur is not None and not _MBB_NOISE.search(line) and line.startswith(" "):
            cur["lines"].append(line.strip())
    for t in txns:
        useful = [l.rstrip(" *").strip() for l in t["lines"] if not _DETAIL_JUNK.match(l.rstrip(" *").strip())]
        useful = [u for u in useful if not re.search(r",\s*MYS$", u, re.I)]
        t["details"] = useful[0] if useful else t["description"].title()
        if not useful and any(l.startswith("MBBQR") for l in t["lines"]):
            t["details"] = "QR payment (merchant not named)"
        t["memo"] = " ".join(useful[1:])
        t["match_text"] = " ".join([t["description"]] + t["lines"])
        t["source"] = "Maybank Savings"
        t["file"] = os.path.basename(path)
        del t["lines"]
    cr = round(sum(t["amount"] for t in txns if t["amount"] > 0), 2)
    dr = round(-sum(t["amount"] for t in txns if t["amount"] < 0), 2)
    ok = (tot_cr is None or abs(cr - tot_cr) < 0.005) and (tot_dr is None or abs(dr - tot_dr) < 0.005)
    if begin is not None and end is not None:
        ok = ok and abs(begin + cr - dr - end) < 0.005
    check = f"credits {cr:,.2f} vs {tot_cr}, debits {dr:,.2f} vs {tot_dr}, balance {begin} -> {end}"
    m = re.search(r"_(\d{8})", os.path.basename(path))
    asof = dt.datetime.strptime(m.group(1), "%Y%m%d").date() if m else (max(t["date"] for t in txns) if txns else None)
    txns = TxnList(txns)
    if end is not None and asof:
        txns.closing = (asof, end)
    return txns, ok, check


# ---------------------------------------------------------------- Maybank credit card
_CC_TX = re.compile(r"^\s*(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(\S.*?)\s*$")
_CC_AMT = re.compile(r"([\d,]+\.\d{2})(CR)?\s*$")


def parse_maybank_card(path):
    text = pdf_text(path)
    m = re.search(r"(\d{1,2} (?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC) \d{2})\s", text)
    stmt = dt.datetime.strptime(m.group(1).title(), "%d %b %y").date() if m else None
    card = None
    txns = []
    tot_cr = tot_dr = None
    for raw in text.splitlines():
        line = raw.rstrip()
        cm = re.search(r"([A-Z ]*(?:VISA|MASTERCARD|AMEX|AMERICAN EXPRESS)[A-Z ]*)\s*:\s*[\d ]{12,}", line)
        if cm:
            card = cm.group(1).strip()
        mm = re.search(r"TOTAL CREDIT THIS MONTH\s+([\d,]+\.\d{2})", line)
        if mm:
            tot_cr = money(mm.group(1)); continue
        mm = re.search(r"TOTAL DEBIT THIS MONTH\s+([\d,]+\.\d{2})", line)
        if mm:
            tot_dr = money(mm.group(1)); continue
        fx = re.search(r"TRANSACTED AMOUNT\s+([A-Z]{3})\s+([\d,]+\.\d{2})", line)
        if fx and txns:
            txns[-1]["memo"] = f"{fx.group(1)} {fx.group(2)}"
            txns[-1]["foreign"] = True
            continue
        m = _CC_TX.match(line)
        if not m:
            continue
        rest = m.group(3)
        am = _CC_AMT.search(rest)
        if not am:
            continue  # e.g. "RETAIL INTEREST RATE = 15.00%"
        body = rest[:am.start()].rstrip()
        parts = re.split(r"\s{2,}", body)
        desc = re.sub(r"\s+AP$", "", parts[0]).strip()
        where = " ".join(p for p in parts[1:] if p).strip()
        tday, tmon = int(m.group(2)[:2]), int(m.group(2)[3:])
        year = stmt.year if stmt else dt.date.today().year
        if stmt and tmon > stmt.month:
            year -= 1
        amt = money(am.group(1))
        # card view: purchases are money OUT (negative), credits (payments/refunds) are IN
        signed = amt if am.group(2) else -amt
        txns.append(Txn(date=dt.date(year, tmon, tday), amount=signed, description=desc, details=desc,
                        memo=where, match_text=f"{desc} {where}", source="Maybank Card",
                        file=os.path.basename(path), card=card, foreign=False))
    cr = round(sum(t["amount"] for t in txns if t["amount"] > 0), 2)
    dr = round(-sum(t["amount"] for t in txns if t["amount"] < 0), 2)
    ok = tot_cr is not None and tot_dr is not None and abs(cr - tot_cr) < 0.005 and abs(dr - tot_dr) < 0.005
    check = f"credits {cr:,.2f} vs {tot_cr}, debits {dr:,.2f} vs {tot_dr}"
    return txns, ok, check


# ---------------------------------------------------------------- CIMB Clicks
_CIMB_TX = re.compile(r"^\s*(\d{2} [A-Z][a-z]{2} \d{4})\s{2,}(\S.*?)\s{2,}(MYR.*)$")


def parse_cimb(path):
    text = pdf_text(path, "-layout")
    text = re.sub(r"MYR\s*([\d,]+)\.\s*(\d{2})", r"MYR \1.\2", text)
    txns, cur = [], None
    for raw in text.splitlines():
        line = raw.rstrip()
        m = _CIMB_TX.match(line)
        if m:
            amts = [money(a) for a in re.findall(r"MYR ([\d,]+\.\d{2})", m.group(3))]
            if len(amts) < 2:
                cur = None
                continue
            cur = Txn(date=dt.datetime.strptime(m.group(1), "%d %b %Y").date(), raw_amount=amts[0],
                      balance=amts[-1], description=m.group(2).strip(), lines=[])
            txns.append(cur)
            continue
        if cur is not None and line.strip() and line.startswith("   ") and "Protected by" not in line:
            cur["lines"].append(line.strip())
        elif not line.strip():
            cur = None
    # newest first in the PDF -> sign from balance movement
    ok = True
    for i, t in enumerate(txns):
        older = txns[i + 1]["balance"] if i + 1 < len(txns) else None
        if older is not None:
            delta = round(t["balance"] - older, 2)
            if abs(abs(delta) - t["raw_amount"]) > 0.005:
                ok = False
            t["amount"] = delta
        else:
            inflow = "INTEREST" in t["description"] or "FROM" in t["description"]
            t["amount"] = t["raw_amount"] if inflow else -t["raw_amount"]
        useful = [l for l in t["lines"] if not _DETAIL_JUNK.match(l) and not re.fullmatch(r"\d+", l)]
        t["details"] = useful[0] if useful else t["description"].title()
        t["memo"] = " ".join(useful[1:])
        t["match_text"] = " ".join([t["description"]] + t["lines"])
        t["source"] = "CIMB Savings"
        t["file"] = os.path.basename(path)
        del t["lines"], t["raw_amount"]
    txns.reverse()
    check = "balance chain consistent" if ok else "balance chain BROKEN"
    txns = TxnList(txns)
    if txns:
        txns.closing = (txns[-1]["date"], txns[-1]["balance"])
    return txns, ok, check


# ---------------------------------------------------------------- CIMB Clicks credit card history
_CIMBCC_TX = re.compile(r"^\s*(\d{2} [A-Z][a-z]{2} \d{4})\s+(\d{2} [A-Z][a-z]{2} \d{4})\s+(\S.*?)\s{2,}MYR ([\d,]+\.\d{2})\s*$")
_CREDIT_WORDS = re.compile(r"THANK YOU|PAYMENT|REFUND|REVERSAL|CASHBACK|REBATE", re.I)


def parse_cimb_card(path):
    text = pdf_text(path, "-layout")
    text = re.sub(r"MYR\s*([\d,]+)\.\s*(\d{2})", r"MYR \1.\2", text)
    debit_col = credit_col = None
    txns, cur = [], None
    for raw in text.splitlines():
        line = raw.rstrip()
        if "Posting Date" in line and "Debit" in line and "Credit" in line:
            debit_col, credit_col = line.index("Debit"), line.index("Credit")
            continue
        m = _CIMBCC_TX.match(line)
        if m:
            amt = money(m.group(4))
            if debit_col is not None:
                pos = line.rindex("MYR")
                is_credit = abs(pos - credit_col) < abs(pos - debit_col)
            else:
                is_credit = bool(_CREDIT_WORDS.search(m.group(3)))
            desc = re.sub(r"\s{2,}", " ", m.group(3)).strip()
            cur = Txn(date=dt.datetime.strptime(m.group(2), "%d %b %Y").date(), amount=amt if is_credit else -amt,
                      description=desc, details=desc[:40], lines=[], source="CIMB Card",
                      file=os.path.basename(path), foreign=False)
            txns.append(cur)
            continue
        if cur is not None and line.strip() and line.startswith("   "):
            cur["lines"].append(line.strip())
        elif not line.strip():
            cur = None
    for t in txns:
        extra = [l for l in t["lines"] if l not in ("MY", "Payment Desc")]
        t["memo"] = " ".join(extra)
        t["match_text"] = " ".join([t["description"]] + t["lines"])
        del t["lines"]
    txns.reverse()
    dr = round(-sum(t["amount"] for t in txns if t["amount"] < 0), 2)
    cr = round(sum(t["amount"] for t in txns if t["amount"] > 0), 2)
    ok = len(txns) > 0 and debit_col is not None
    check = f"debits {dr:,.2f}, credits {cr:,.2f} (history export - no bank totals to compare)"
    return txns, ok, check


PARSERS = {"Maybank Savings": parse_maybank_savings, "Maybank Card": parse_maybank_card, "CIMB Savings": parse_cimb,
           "CIMB Card": parse_cimb_card}


def parse_file(path):
    kind = detect(pdf_text(path, "-layout"))
    if not kind:
        return None, [], False, "unrecognised statement layout"
    txns, ok, check = PARSERS[kind](path)
    return kind, txns, ok, check


if __name__ == "__main__":
    import sys
    for p in sys.argv[1:]:
        kind, txns, ok, check = parse_file(p)
        print(f"\n### {os.path.basename(p)} -> {kind}: {len(txns)} txns, {'OK' if ok else 'CHECK FAILED'} ({check})")
        for t in txns:
            print(f"  {t['date']}  {t['amount']:>10,.2f}  {t['description'][:28]:28} | {t.get('details','')[:30]:30} | {t.get('memo','')[:30]}")
