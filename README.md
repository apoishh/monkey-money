# Monkey Money 🐒💰

A personal-finance kit for Malaysia, built on Excel and Windows:

- **An Excel workbook** for daily spending, a budget plan, investments, loans and net worth
- **A bank-statement importer**: drop your Maybank / CIMB PDFs in a folder, review, post. No typing.
- **A bank check** that tells you whether the tracker still agrees with your real bank balances
- **A local dashboard website**: spending analytics, investments, and a "what if" future planner

Everything runs **on your own PC**. Nothing is uploaded anywhere.

> `Monkey Money.xlsx` ships with **fictional sample data** for a made-up "Demo User" (about RM7M net worth), so you
> can try everything straight away. Replace it with your own numbers when you're ready (see *Start with your own data*).

---

## Requirements

- Windows with **Microsoft Excel** (desktop)
- **Python 3.10+** – then `pip install -r requirements.txt`
- **Git for Windows**, which includes the `pdftotext` tool used to read the PDF statements
- Internet the first time you open the dashboard (the chart library loads from cdnjs)

## Quick start

1. Download or clone this folder.
2. `pip install -r requirements.txt`
3. Double-click **`4 - Open dashboard.bat`**. Your browser opens http://localhost:8765 with the sample data.
4. Open **`Monkey Money.xlsx`** to see the workbook.

---

## The workbook

| Tab | Use it for |
|---|---|
| **Dashboard** | Monthly view inside Excel (pick a year and month at the top right) |
| **Budget Planning** | Your plan: how much per category, per month |
| **Budget Tracking** | Every transaction (typed by hand or posted from statements) |
| **Investment** | Money in or out of each platform, plus a current-value row whenever you check |
| Import Review | Created on the first import: where new bank lines wait for your OK |
| Setting | Date of birth, salary, savings targets, growth assumptions (rows 47–57) |

Helper sheets are **hidden** (right-click a tab → *Unhide…*):
- **Net Worth**: card balances and valuables
- **Loans**: property and car loans, market values
- **Saving**: medium and long-term goals
- **Wealth Summary**: net worth, cash flow, emergency fund, goals
- Import Rules, Import Log, Bank Check, Calculations, Dropdown Data

## Monthly routine

1. Save your statement PDFs into the **`statements`** folder:
   - Maybank savings e-statement and Maybank credit-card e-statement (Maybank2u)
   - CIMB savings and CIMB credit-card **Transaction History** PDFs (CIMB Clicks)
2. Double-click **`1 - Import statements.bat`**. New lines go to **Import Review**, already categorised where a rule matches.
   Each statement is checked against the bank's own totals.
3. On **Import Review**, fix the yellow rows (choose Type and Category, set **Post?** to `Yes`). Set **Post?** to `No`
   for anything you don't want.
4. Double-click **`2 - Post reviewed rows.bat`**. The rows move into Budget Tracking. Save the workbook.
5. Check the bank check output (or run **`3 - Check bank balances.bat`**). Maybank should match the tracker's
   *Balance* and CIMB should match its *Savings*.

### How money is counted

| Thing | Recorded as |
|---|---|
| Card purchase | One expense per purchase, on the day you spend |
| Paying the card bill | Not an expense (it would be counted twice): skipped, or only a savings movement if paid from CIMB |
| Maybank ↔ CIMB transfer | Savings › Personal Savings (+ into CIMB, – out of it) |
| Transfer to Wise / MooMoo / TnG | Expenses › Investment (out), Income › Investments (back) |

**Savings = the CIMB account only.** That way the bank check can compare it with a real balance.

### Teach it your merchants (Import Rules sheet)

Each rule says: *if the bank line contains X (from this source, in this direction), record it as Type / Category.*
Rules are checked top to bottom and the first match wins. Use `|` to give alternatives (`ZUS COFFEE|TEALIVE`).

The default rules include three **`CHANGE-ME`** rows: your employer, your landlord or mortgage, and your car-loan
reference. Change them to the text on *your* statements. Also set **your name** (cell C5) so transfers between your own
accounts are recognised.

### Duplicates and safety

- Every PDF is fingerprinted, so a statement is never imported twice, even if it's renamed (both CIMB downloads are
  called `CIMBClicks.pdf`). You can delete PDFs after importing.
- Rows already in Budget Tracking are skipped. Imported rows match on date + type + amount + merchant; hand-typed rows
  match on date + type + amount.
- Posting checks again. Possible duplicates are flagged and not posted unless you set **Post?** to `Force`.
- The bank check is the final safety net.

---

## Dashboard

`4 - Open dashboard.bat` → http://localhost:8765. It's read-only and reads the last **saved** version of the workbook,
even while the file is open in Excel. Save in Excel, then refresh the browser.

| Tab | What you get |
|---|---|
| **Overview** | Net worth (checked against Excel), accounts, bank check, how old each data source is |
| **Spending** | Money in vs out by month, categories, budget vs actual, top merchants, searchable transactions |
| **Investments** | Value over time per platform, allocation, gain/loss |
| **Future plans** | Sliders for income, raises, inflation, growth and spending cuts, showing when each goal is reached |

Goals come from the **Wealth Summary** sheet (rows 48–52). Growth rates are your own assumptions, not forecasts.

---

## Start with your own data

1. Keep a copy of `Monkey Money.xlsx` as a reference, then in your working copy:
   - **Budget Tracking**: delete the sample rows, then add one *Income › Others* "Opening balance" row equal to your
     main-account balance, and one *Savings* row equal to your savings-account balance
   - **Investment**: replace each platform's rows with your own
   - **Budget Planning**: your monthly plan
   - **Setting** (rows 47–57), **Loans**, **Net Worth**, **Saving**: your details and goals
   - **Wealth Summary** rows 48–52: your goal names, targets and dates
2. **Never commit your real workbook or statements.** The `.gitignore` already blocks `statements/`, PDFs and
   `* - private.xlsx`. A safe habit is to name your real file `Monkey Money - private.xlsx` and point the tools at it:
   `python tools\statement_import.py import --workbook "Monkey Money - private.xlsx"`. For the dashboard, change
   `WORKBOOK` in `tools\dashboard\data.py`.

## Supported statements

| Bank | Document | Check |
|---|---|---|
| Maybank / Maybank Islamic savings | Monthly e-statement PDF | Credits, debits and closing balance vs the statement |
| Maybank credit card | Monthly e-statement PDF | Total credit and debit vs the statement |
| CIMB savings | CIMB Clicks transaction history PDF | Running balance chain |
| CIMB credit card | CIMB Clicks transaction history PDF | Basic (the export has no totals) |

Other banks need a new reader in `tools/statement_parsers.py`. Each reader turns a PDF into
`date / amount / description` lines.

## Project layout

```
Monkey Money.xlsx            sample workbook (fictional data)
1 - Import statements.bat    read PDFs in statements\ -> Import Review
2 - Post reviewed rows.bat   Import Review -> Budget Tracking
3 - Check bank balances.bat  tracker vs bank
4 - Open dashboard.bat       local website
statements\                  drop PDFs here (git-ignored)
tools\statement_parsers.py   PDF readers per bank
tools\statement_import.py    rules, review, post, duplicate and bank checks
tools\dashboard\             FastAPI server + static page (read-only)
```

*Not financial advice. The projections are arithmetic on your own assumptions.*
