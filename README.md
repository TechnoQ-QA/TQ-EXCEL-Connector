# TQ Connector for NetSuite (prototype)

Excel formulas that pull live general ledger balances from NetSuite, with
drill-down to the transactions behind a number.

```
=TQ.BALANCE(account, fromPeriod, [toPeriod], [subsidiary], [department], [class], [location], [filters])
=TQ.ACCOUNTNAME(account)
```

- Balance sheet accounts return the balance at the end of the period.
- Income and expense accounts return activity from `fromPeriod` to `toPeriod`.
- Income, liabilities and equity come back positive.
- Periods are NetSuite period names (`Jan 2026`). Segments accept a name or an internal ID.

## Installing at a customer (one file per PC)

Techno Q builds one file, `build\TQConnectorSetup.exe`, and the customer runs it on each PC. No Python,
no administrator rights, nothing else to install.

```
powershell -ExecutionPolicy Bypass -File installer\build.ps1
```

(runs the tests, makes `TQConnector.exe` with PyInstaller kept in `build\venv`, and makes the setup file with
Inno Setup 6 from `installer\TQConnector.iss`; install Inno Setup once with
`winget install --id JRSoftware.InnoSetup -e`). The setup file is a normal Windows installer: no PowerShell,
installs for the signed-in user only. Running it on a PC:

1. copies the connector to `%LOCALAPPDATA%\TQ Connector` (an update keeps its `data` folder);
2. makes **that PC's own certificate** for https://localhost (not a certificate authority, valid for
   localhost only) and trusts it for the user: Windows asks the user to confirm;
3. registers the add-in with desktop Excel, starts the connector now and at every sign-in, and adds
   it to Settings > Apps so it can be removed;
4. on a first install, asks for the NetSuite account ID, the RESTlet URL and the sign-in client ID, and
   for the licence file (all optional: they can be entered later on the setup page). To roll out to
   several PCs without typing, put `TQConnector.settings.json` (a prepared config.json) and
   `licence.json` next to the setup file: those pages are then skipped and the files copied;
5. opens the **setup page** (also reachable from the panel's Setup link), where the details, the
   company logo, seats and updates are managed. Secrets typed there are saved on the PC and never
   shown again.

Uninstall from Settings > Apps: it stops the connector, removes the add-in registration, the start-up
shortcut and exactly this PC's certificate (by its thumbprint, in `data\certs\localhost.crt.sha1`), and
keeps the `data` folder (settings, licence).

**Signing.** The build signs `TQConnector.exe`, the setup file and its uninstaller with Techno Q's own
certificate (made 19 Nov 2026, in this PC's certificate store; `installer\sign.ps1`), and writes the setup's
SHA-256 next to it with `TechnoQ.cer`. A client trusts that certificate once; see
[docs/CLIENT-INSTALL.md](docs/CLIENT-INSTALL.md), which also covers antivirus, silent installs and backing up the
key. With a bought certificate later, set `TQ_SIGNTOOL` before building, with `$f` where
the file goes, for example
`$env:TQ_SIGNTOOL = 'signtool sign /fd sha256 /tr http://timestamp.digicert.com /td sha256 /sha1 <thumbprint> $f'`
(or the cloud signing tool's command). `TQConnector.exe`, the setup file and its uninstaller are then signed
and time-stamped. Unsigned, antivirus such as Kaspersky may flag the setup file (it trusts a certificate and
adds a start-up entry).

**Licence.** Connecting to NetSuite needs a licence file signed by Techno Q, for that NetSuite account
and until a date. Make one with

```
python scripts\make_licence.py --customer "World of Hospitality" --account 1234567 --users 5 --expires 2027-12-31
```

and send the file; the customer pastes it on the setup page. The signing key is in
`%USERPROFILE%\TQ Licence Key\private-key.json` on the Techno Q PC that made it: keep a safe copy, never
ship it. Sample data needs no licence. The number of users is enforced per PC (see Seats per PC). Without a valid licence only the connection test reaches NetSuite.

Checked: the exe was built and run from its staging folder (sample data, setup page, filters). **Not
yet tried**: running the setup file on a clean PC (it would clash with this development copy on port
3000 and the add-in ID), Excel loading the add-in from the installed copy, antivirus reactions to an
unsigned exe. For customers the exe should be code-signed. The Inno Setup file was run on the development PC
(Nov 2026, silent install over an earlier copy): it kept the data folder, reused the trusted certificate without
asking again, registered the add-in, added the start-up shortcut and the Settings > Apps entry, removed the old
installer's entry, and the installed connector answered from NetSuite.

## Parts

| Folder | What it is |
|---|---|
| `addin/` | The Excel add-in: formulas, side panel (Reports tab: filters, layout, data tables, lists; Chat tab), setup page, manifest |
| `server/` | Local server: serves the add-in, forwards requests to NetSuite (`server.py`), adds totals (`ledger.py`), lays out reports (`reports.py`) and designed layouts (`layouts.py`), runs data queries (`queries.py`), reads chat and Designer requests (`chat.py`) |
| `netsuite/` | The three scripts to upload into the NetSuite account (RESTlet with data queries, Suitelet, balance logic) |
| `scripts/` | Certificate setup for development, licence maker, icon generator |
| `installer/` | Build script and the per-PC install / uninstall scripts |
| `tests/` | Server tests: `python -m unittest discover -s tests` |

## 1. Try it with sample data (no NetSuite needed)

1. One-time setup (Windows asks you to confirm trusting the certificate). It creates and
   trusts the localhost certificate and registers the add-in with desktop Excel.
   Add `-Remove` to undo both.

   ```
   powershell -ExecutionPolicy Bypass -File scripts\setup-excel.ps1
   ```

2. Start the server and leave it running.

   ```
   python server\server.py
   ```

3. Load the add-in into Excel:
   - **Excel desktop:** close Excel completely, reopen it, then Home > Add-ins. The add-in
     is listed under Developer Add-ins.
   - **Excel on the web:** Home > Add-ins > More Add-ins > My Add-ins > Upload My Add-in,
     then pick `addin\manifest.xml`.

4. Type `=TQ.BALANCE(4000, "Jan 2026")` in a cell. Sample accounts are 1000-6200 and
   sample periods are Jan 2025 to Dec 2026 (use the panel's "Accounts" and "Periods" buttons).

## 2. Connect to a NetSuite sandbox

**The quick way: the NetSuite package.** `netsuite/sdf` installs the scripts, their deployments (the Suitelet
running as Administrator) and the role TQ Excel User in one deploy. See `netsuite/sdf/README.md`. The steps below
are the same thing done by hand.

In NetSuite:

1. Enable features (Setup > Company > Enable Features > SuiteCloud): Server SuiteScript,
   REST Web Services, Token-Based Authentication.
2. Upload the three files in `netsuite\` to one File Cabinet folder under SuiteScripts,
   `tq_excel_core.js` first (the Suitelet loads it).
3. Create two script records (Customization > Scripting > Scripts > New). The IDs matter,
   because the RESTlet finds the Suitelet by them:
   - `tq_excel_elevated.js` (Suitelet): script ID `_tq_excel_elevated`, deployment ID
     `_tq_excel_elevated`, Status Released, **Execute As Role: Administrator**, audience
     all internal roles, "Available Without Login" left off.
   - `tq_excel_restlet.js` (RESTlet): deploy to the roles that will use it. Copy its
     **External URL**.
4. Setup > Integration > Manage Integrations > New: tick Token-Based Authentication,
   untick the OAuth 2.0 options. Copy the consumer key and secret (shown once).
5. Setup > Users/Roles > Access Tokens > New: pick the integration, a user and a role.
   Copy the token ID and secret (shown once).

On this PC:

6. Copy `server\config.example.json` to `server\config.json` and fill in the six NetSuite values.
   The account ID is in the URL prefix, e.g. `1234567_SB1`.
7. Restart the server. The panel should show "Connected to NetSuite account ...".

`server\config.json` holds secrets and is excluded from git.

## 3. Reports and chat

The panel's Chat tab builds reports from plain requests. It is a built-in command reader
(`server\chat.py`), not an AI service: nothing to install, nothing to pay, and nothing is
sent anywhere except the usual lookups in NetSuite. It understands a fixed set of wordings,
so it is predictable but will not follow free conversation. Type `help` in the chat for the list.

| Say | You get |
|---|---|
| `P&L for 2026 by month` | Income statement, one column per month plus a total |
| `income statement Q1 2026 vs last year` | Two periods side by side with change and change % |
| `P&L Jan to Jun 2026 by department` | One column per department that has a balance, "No department", and the total |
| `balance sheet as at Oct 2026 vs Dec 2025` | Balance sheet with comparison |
| `balance sheet by quarter for 2026, summary` | One row per section instead of per account |
| `trial balance Oct 2026` | Debit and credit columns |
| `create layout` | A sample layout sheet to turn into your own report format (see below) |
| `custom P&L for 2026 by month` | The report in your own format, built from the layout sheet |
| `add column Q1 2025` / `remove line Rent` | Changes the layout and rebuilds the report on the same sheet |
| `what was rent in Mar 2026?` / `net income this year` | The number, in the chat |
| `transactions for 6100 in Jan 2026` | The posting lines on a new sheet |
| `list accounts` / `list periods 2026` / `list departments` | The list on a new sheet (add `here` for the selected cell) |

After a report, a short follow-up changes it: `by quarter`, `by month`, `summary`, `detailed`,
`compare with last year`, `no comparison`, `for department Sales`, `no filters`, `now for 2025`,
`all accounts` (keep accounts with no balance).

### Filters

Every report has filter cells at the top: **B3 Other filters**, then B4:B7 subsidiary, department,
class, location. In any of them:

| Type | Means |
|---|---|
| `Sales` | one value (name or internal ID) |
| `Sales, Bar` | several values added together |
| `not Sales` | everything except these |
| `101 Food*` | that value and everything under it |
| `Parent Company*` (subsidiary) | the parent and its subsidiaries, **consolidated** in the parent's currency |

B3 takes any other field as `Field: values; Field: values`, for example `Brand: Premium, Basic; Name: Al Noor`.
Fields are the custom segments of the account (by their name, such as Brand or R&H Revenue Center),
Name (customer, vendor or employee) and Item. The panel's **Filters** section does all this by
clicking: pick a field, tick values, Apply. "Keep in the layout" saves it in a layout so a rebuild keeps
it. In the chat: `P&L for 2026 filter Brand: Premium`, `balance sheet for subsidiary Parent consolidated`.
A layout's column can have its own Other filters too (a column per brand).

**Dropdowns in Excel.** B4:B7 always pick from the NetSuite lists (a value can still be typed, so
`A, B` and `not A` work). Any other field can get a dropdown too: in the Filters section pick the field
and press **Add as a dropdown on the sheet**, or say `add dropdown filter Brand, Name` in the chat. The
report gets a "More filters" block at E3:F7, one row per field (up to 4), whose value cell picks from
that field's values. Custom segments list every active value set up in NetSuite; Name and Item are
typed (their lists are too long). B3 then becomes a formula that joins the block, so every number
follows the dropdowns after **Refresh all**. The rows are kept with the report: in a layout they are
saved on the layout sheet ("Dropdown filters" row) and come back on every rebuild; on a standard report
they carry over to follow-ups such as `by month`. `no dropdown filters` removes them.

Budgets can be filtered by the four segments only.

### Consolidation

In a OneWorld account whose subsidiaries use more than one currency, amounts of several subsidiaries
are translated with NetSuite's consolidated exchange rates: to the named parent for `Parent*`, otherwise
to the top subsidiary. Income and expenses use each posting period's rate, balance sheet accounts the
rate of the report date, as NetSuite's own consolidated reports do. A consolidated balance sheet and
trial balance get a **Currency translation (consolidation)** line (`calc:Translation`) so they balance.
One subsidiary on its own stays in its own currency. Budgets are not translated.

Filters and consolidation need `tq_excel_core.js` version 2 in NetSuite. With the older script the
connector says so instead of giving wrong numbers. Checked on the sandbox (Nov 2026, version 2
installed): two locations, `not` a location, a parent location with `*`, the same through B3, and
`Name:` with and without `not` all matched the single-value figures and the transaction lists to the
cent. Name matches what NetSuite shows (`1 Techno Q`), the ID or the name alone. **Not checked against
real data**: consolidation (the sandbox is not OneWorld; checked with a stand-in only) and custom
segment filters (the sandbox's custom segments have no postings).

How the reports are built:

- Every number is a live `TQ.BALANCE` formula that points at the period names in the header
  and at four filter cells (B4:B7: subsidiary, department, class, location). Change a period
  or type a filter, then **Refresh all**; there is no need to rebuild the report.
- Only accounts with a balance in the report's periods are listed. Each report ends with a
  grey **check** row that should show 0: it compares the listed accounts with NetSuite's
  own total, so an account that gains its first posting later does not go unnoticed.
- A detailed report on a large chart of accounts takes a few seconds to lay out; a split by
  department or class takes longer (about 25 seconds for 60 departments in the sandbox).

### Your own report format (Report layout section and layouts)

Each customer's report can have its own lines and its own columns. The format is kept on an
ordinary sheet in the workbook whose name starts with `TQ Layout`, so it is saved with the
file and can be copied to another workbook. There are three ways to change it, and they all
edit the same sheet:

- **Report layout section** in the panel: New layout, then add, remove and reorder lines and columns
  by clicking. Accounts are picked from a search box; formulas are built by clicking names.
- **Chat**: `add column Q1 2025`, `add column Jan to Jun 2026 for department Sales called Sales H1`,
  `add column Change = [2026]-[2025]`, `remove column Total`,
  `add line Rent accounts 6100, 6200-6299 negative in section Staff Cost`,
  `add line Margin = [Gross Profit]/[Net Sales] percent`,
  `add subtotal Total Staff Cost in section Staff Cost`, `remove line Rent`, `update report`.
  `custom P&L for 2026 by month` (or `vs last year`, `by department`) replaces all the columns.
- **The layout sheet itself**: type into its two tables, then say `update report`.

The report is written to a sheet named after the report title and is **rewritten in place**
on every change, so formulas on other sheets that point at it keep working. Filters typed
into the report's B4:B7 are overwritten by a rebuild; to keep a filter, set it in the layout
(`for department Sales` in the chat, or the filter rows at the top of the layout sheet).

Lines table:

| Section | Line | Type | Accounts or formula | Sign | Format |
|---|---|---|---|---|---|
| Revenue | Sales | accounts | `4000, 4010-4099` | 1 | |
| Revenue | Net Sales | subtotal | | | |
| Cost of Sales | Food Cost | accounts | `5000-5099` | -1 | |
| Cost of Sales | Total Cost of Sales | subtotal | | | |
| Gross Profit | Gross Profit | formula | `=[Net Sales]+[Total Cost of Sales]` | | bold |
| Gross Profit | Gross Margin | formula | `=[Gross Profit]/[Net Sales]` | | percent |
| KPI | Average Sales Daily | formula | `=[Net Sales]/days` | | |

- **accounts**: account numbers separated by commas, ranges (`4000-4999`), or `type:Income`.
  Sign `-1` shows the amount as a negative, so costs can be added to revenue.
- **subtotal**: adds the accounts lines above it in the same section.
- **formula**: any Excel formula. `[Line name]` is that line in the same column; `days` is the
  number of days in the column's periods.

Columns table (to the right of the lines on the same sheet):

| Column | Column type | From | To | Department | Column formula | Column format |
|---|---|---|---|---|---|---|
| YTD | period | Jan 2026 | Aug 2026 | | | |
| Last year | period | Jan 2025 | Aug 2025 | | | |
| Food YTD | period | Jan 2026 | Aug 2026 | 101 Food | | |
| Change | formula | | | | `=[YTD]-[Last year]` | |
| Change % | formula | | | | `=[Change]/ABS([Last year])` | percent |
| Total | total | Jan 2026 | Dec 2026 | | `=SUM([Jan 2026]:[Dec 2026])` | |

- **period**: balances for From..To; optionally for one subsidiary, department, class or
  location, which overrides the report's filter for that column only.
- **formula**: an Excel formula over other columns, applied to every line.
- **total**: like formula, but a ratio line is worked out again on the totals (a year's
  margin is the year's profit over the year's sales, not a sum of monthly margins).

**One sheet per location.** A layout can also produce a pack: the summary sheet, then one
sheet for each location (or department, class, subsidiary) that has a balance, each filtered
to that value. Choose it under Sheets in the Report layout section (the sheets are made as soon as you
pick it; there is no need to press Update report), say `one sheet per location` in
the chat, or type the word in the "One sheet per" row of the layout sheet. `one sheet only`
turns it off. Every sheet is live formulas, so opening the workbook later and pressing
**Refresh all** updates the whole pack; **Update report** is only needed after the layout
changes or to pick up a location that has its first posting.

**Starting a layout.** In the Report layout section, "Start from" offers ready-made templates: the
management P&L (account by account), a P&L summary, a balance sheet and an indirect cash flow.
Their lines use account types, so they fit any NetSuite account. On the sandbox the balance
sheet template's check line is 0.00 and the cash flow's net change equals the movement in the
bank accounts (2025).

"My own format, typed on a sheet" reads a format from a worksheet: type the report's lines in
one column, one per row. A row becomes a line when it has account numbers in a column next to
it (`4000`, `4100-4199`), starts with an account number, is a total, is a result such as Gross
Profit, is a ratio, has an account's name, or names a group such as Sales or Expenses. A name
on a row of its own is a section heading. Rows that match nothing are listed and left out,
never guessed. The matching is by rules, so check each line's accounts before relying on it.

**Budget and variance.** `custom P&L for 2026 by quarter vs budget` gives actual, budget,
variance and variance % for each period. A single budget column is a period column whose
Data is `budget` (`add column budget 2026`, or the tick box in the Designer).

**GL accounts under each line.** `show account detail` (or the tick box in the Designer)
lists, under every line that covers several accounts, the accounts that have a balance, so
a figure can be traced to its GL codes. `hide account detail` removes them.

**Summary sheet.** A one-sheet-per report also gets a sheet "<title> - Summary": one row per
location (or company, department, class), with the first key line in every column and the
other key lines at the total. Key lines are the subtotals and bold formula lines, or the
lines whose Format includes `key`. Each name is a link to that sheet.

**Cash flow.** A line can use `move:` in front of its accounts for the change in balance
sheet accounts over the column's periods, which is what an indirect cash flow is made of.

The report also has:

- a **Year** cell (B8), with a dropdown of the years NetSuite has, that every period in the
  header is built from: pick another year and use **Refresh all**. It appears when the
  period names contain the year.
- **dropdown lists** on the filter cells B4:B7, fed by lists written to the right of the report.
- a line in the chat reply naming any profit and loss account that has a balance in the
  report's periods but is in no line of the layout.

Totals that one formula can ask for (handled by `server\ledger.py`, no change in NetSuite):

```
=TQ.BALANCE("type:Income", "Jan 2026", "Mar 2026")       all accounts of a NetSuite account type
=TQ.BALANCE("type:Bank,AcctRec", "Oct 2026")             several types together
=TQ.BALANCE("calc:NetIncome", "Jan 2026", "Mar 2026")    income less costs
=TQ.BALANCE("calc:NetIncomeYTD", "Oct 2026")             the same, from the start of the fiscal year
=TQ.BALANCE("calc:RetainedEarnings", "Oct 2026")         net income of all earlier fiscal years
=TQ.BALANCE("4000, 4010, 4100-4199", "Jan 2026")         a list of accounts and ranges added together
=TQ.BALANCE("calc:Days", "Jan 2026", "Mar 2026")         days in the periods (90)
=TQ.BALANCE("budget:4000-4999", "Jan 2026", "Mar 2026")  the NetSuite budget, not the postings
=TQ.BALANCE("budget(Plan A):type:Income", "Jan 2026")    the budget of one budget category
=TQ.BALANCE("move:1100, 1200", "Jan 2026", "Mar 2026")   change in balance sheet accounts (for a cash flow)
```

The fiscal year starts in January unless `"fiscalYearStartMonth"` (1 to 12) is set in
`server\config.json`.

**Drill-down** (the Drill down button) works on any balance cell. On a single
account it lists the transactions. On a total (a list, a range, a type, net income, a budget)
it first lists the accounts behind the cell with each one's amount; clicking an account lists
its transactions. A budget has no transactions behind it.

**Accounts without an account number** (NetSuite's own system accounts, for example its
default Accounts Receivable) are known as `ID` plus their internal id, such as `ID123`.
They are included in every type total, so the balance sheet balances.

### Month-end pack: roll forward, send, budget

**Roll forward with one cell.** The template "Month and year to date P&L (rolls forward)", or the
chat's `monthly P&L for Oct 2026`, gives the month and the year to date, each against last year,
with the change and change %. Cell B8 holds the month: pick another from its dropdown and every
column moves. A layout column's From and To can be relative to that month: `CM` (the month),
`CM-1`, `LY` (the same month last year), `FYS`/`FYE` (first/last month of its fiscal year), `QS`/`QE`,
`FYS-12` (start of last year), `FYS-1` (last year end). The layout keeps the month in its "Report
month" row. The Balance sheet template now uses it too: as at the month, the same month last year
and last year end.

**Relative periods in any formula.** `=TQ.BALANCE("4000","FYS","CM")` is this year to date;
`=TQ.PERIOD(B8,-1)`, `=TQ.PERIOD(B8,"FYS")`, `=TQ.PERIOD(B8,"LY")` work out a period name from a cell.

**Insert a formula.** In the panel, pick the account (or a total such as Income), the period
(including This month, Year to date), the filters and actual or budget; the formula goes into the
selected cell. Read selected cell loads an existing formula back into the pickers. Tick "Take the
period from cell B8" for formulas that roll forward.

**Drill-down to NetSuite.** Each transaction's number opens the record in NetSuite, in the panel
and on the sheet (a link).

**Refresh this sheet** recalculates only the active sheet; Refresh all still does the workbook.

**Send the report.** Copy as values makes copies of the report (or the whole pack: report,
summary and location sheets) with every formula replaced by its value, on grey tabs, for people
without the add-in (right-click the tabs, Move or Copy, to a new book). Set up for PDF makes every
sheet of the pack landscape, one page wide, with the headings repeated and page numbers; then File >
Export > Create PDF.

**Budget upload.** Budget upload in the panel: choose the fiscal year and budget category and Make
budget sheet (one row per profit and loss account, the year's months across, optionally filled with
last year's actuals). Change the numbers, add rows (each row may have its own subsidiary,
department, class, location), press Check (writes nothing, says what would be created or updated
and which rows have a problem), then Upload to NetSuite, which is enabled only after a check of the
same cells passed. Amounts are in their natural sign (income and expenses positive). Uploading needs
a NetSuite role that may set up budgets, or Administrator; it needs tq_excel_core.js version 3.

Checked: 70 server tests; against the sandbox, relative periods, the monthly P&L (year to date equal
to NetSuite's net income), the rolling balance sheet (difference 0.00 in every column), drill-down
links, the budget sheet and the budget check, and a real upload (Nov 2026): one line for FY 2026,
Budget - B, account 110250 was created (budget id 121, 100 to 1,200 by month, read back by TQ.BALANCE
as 7,800), then updated in place by a second upload. **Not yet tried inside Excel itself.**

### More (tq_excel_core.js version 4)

**New report wizard.** The Reports tab opens with a card for the sheet you are on. On a report it shows the
month picker (when the report rolls forward), Refresh sheet, Filters, Edit layout and Send. On a layout,
budget or journal sheet it shows what that sheet needs; elsewhere it shows New report. New report asks one
question per step:
1. What: Monthly P&L, P&L summary, P&L account by account, Budget against actual, Balance sheet, Cash
   flow, or your own format.
2. Periods.
3. Filters, and which fields get a dropdown.
4. Sheets: one per location and the like, GL accounts, comments, highlights.
5. A one-sentence check.

Every report it builds is a layout, so it can be changed and rebuilt in place.

**Ribbon.** A "TQ Connector" tab in Excel with these buttons: Open panel, New report, Refresh all,
Refresh sheet, Refresh selection, Drill down, Insert formula, Copy as values and Set up for PDF.
Excel picks up the new tab after the add-in is reinstalled or Excel's add-in cache is cleared.

**Variance highlights and comments.** For a layout report: `highlight changes over 10%` (or Highlight in
Report layout) turns change % and variance % cells at or above it green, and at or below minus it red.
`add comments column` adds a Comments column; what is typed there is put back on the same line whenever
the report is rebuilt.

**Your own highlight rules.** On a layout report, a rule colours the amounts it names. Each rule is:
what (all, an account such as 6100, a range such as 4000-4999, type:Expense, or a line name), optionally
`in` a column, then a condition (`>` `<` `>=` `<=` or `between ... and ...`), a value, and a colour. The
colours are red, yellow, amber, green, blue and grey. Examples:
- `highlight 6100 > 5000 red`
- `highlight Net Profit < 0 red`
- `highlight all between -100 and 100 yellow`
- `highlight 4000-4999 in Month < 10000 amber`
- `highlight Gross Margin in Change % below -5% red`

A value written with % applies to the % columns. Rules can also be added and removed in Report layout
(Your own highlight rules), or typed in the layout's "Highlight rules" row. Amounts are compared as shown,
so a cost shown as a negative is compared as a negative. A line is coloured when all its accounts are in
the rule (with Show accounts on, also the account's own row). A rule that matches nothing on the report
is reported in the reply.

**Company logo and name on the PDF** (tq_excel_core.js version 5). Set up for PDF now fills the page
headers: on the left the company name from NetSuite (for a report on one subsidiary, that subsidiary's
legal name, with the company on the right), and the report title in the middle. The logo comes from
NetSuite (Setup > Company > Company Information, Company Logo (Forms); for one subsidiary, its own logo)
and goes at the top right of each sheet of the pack, because Excel does not let an add-in put a picture
in the page header itself. Untick "Company logo and name" to leave them off.

**Dashboards.** Use Dashboard on the ribbon, Dashboards and charts in the panel, the wizard, or the chat
(`dashboard for Oct 2026`, `sales dashboard`, `cash dashboard`). Each one builds a sheet of tiles and
native Excel charts over live TQ formulas. The sheet has a report's filter cells (B3:B7) and month cell
(B8), so it filters, rolls forward with the month and refreshes like a report. Its charts print, go into
the PDF and stay on Copy as values.

| Dashboard | Tiles | Charts |
|---|---|---|
| Executive | Sales, gross profit and net income for the year to date; cash at month end. Each with its change against last year, green or red. | Sales by month, this year and last; sales by location (or department, class, subsidiary, whichever has sales); a pie of costs (the six largest accounts and Other); cash by month. |
| Sales | Sales for the year to date and the month; gross profit; gross margin. | Sales by month; gross margin by month; sales by location. |
| Cash | Cash, receivables and payables at month end; net income for the year to date. | Cash by month; receivables and payables by month; cash movement by month; bank accounts. |

**Finance dashboards.** Three more, from the panel's Finance buttons, the wizard, or the chat:

| Dashboard | Chat | Tiles | Charts |
|---|---|---|---|
| Profit and loss | `finance dashboard`, `p&l dashboard` | Revenue for the year to date; gross margin and net margin (change in points); operating expenses (up shown red) | Revenue, costs and net income by month; gross and net margin by month; a pie of costs; a bridge from revenue to net income |
| Financial position | `financial position dashboard` | Working capital, current ratio, debt to equity and debtor days, each against a year earlier | Working capital by month; a pie of what the company owns; assets, liabilities and equity against a year earlier; receivables and payables by month |
| Budget vs actual | `budget dashboard` | Revenue, costs and net income for the year to date against the budget; revenue for the month | Revenue by month (actual and budget); net income by month; year to date by group; cumulative revenue against budget |

Notes on the figures:
- Equity is worked out as assets less liabilities, so it always balances.
- Quick ratio and creditor days are in the data block.
- Budget formulas filter by subsidiary, department, class and location only.

The tiles and charts read a data block to the right of the sheet. The monthly series stop at the chosen
month. Which locations and cost accounts are listed is decided when the dashboard is built; their numbers
stay live.

**Chart the selection.** Select cells on any sheet (names in the first row or column) and choose columns,
bars, line, pie, doughnut or area: an Excel chart is added beside them.

**Ribbon icons.** Each button on the TQ Connector tab has its own icon (scripts/make_icons.py draws them).
The add-in's own icon, Open panel and the setup file use the TQ mark, a white T and Q on Techno Q's orange
(244, 121, 32, taken from the company logo). Icon addresses carry a version tag (`?v=112`), so Excel
fetches them again after an update instead of showing old cached ones. Raise the tag when icons change.
A Chat button opens the panel on the Chat tab, and New report has a wizard's wand.

**Statistical accounts and non-posting data.** A statistical account (covers, headcount) works like any
account: its number or `type:Stat` in a line or a formula gives the units for the period, never
translated. A formula line such as `=[Net Sales]/[Covers]` then gives sales per cover. For data that does
not post to the ledger, `=TQ.SQL("SELECT COUNT(*) FROM transaction WHERE type = 'SalesOrd' AND
trandate >= TO_DATE(?, 'YYYY-MM-DD')", "2026-01-01")` gives the first value of a SuiteQL query. It runs
with the user's own NetSuite role.

**Accounting book and reporting currency.** In the other filters (B3 or a formula's last argument):
`Book: Secondary Book` reads a non-primary book (Multi-Book Accounting). `Currency: USD` translates
with NetSuite's consolidated rates into the currency of the highest subsidiary that uses it (OneWorld).

**Journal upload.** In the panel, Journal upload, then Make journal sheet. Rows with the same Entry
number become one journal entry; each row is one line, with a Debit or a Credit, and optionally a
memo, a Name, a department, class or location. Check writes nothing: it reports entries whose debits
and credits differ, closed periods, and unknown accounts or names. Upload creates the entries and lists
their numbers. It needs a role with Make Journal Entry, or Administrator.

**Seats.** The licence's number of users is enforced in `tq_seats.json`, a file next to the NetSuite
script. With Sign in with NetSuite, a seat belongs to the **NetSuite user** (the same person on two PCs
takes one seat); with token keys it belongs to the PC. A new user or PC is refused once all seats are
taken, and one not used for 30 days frees its seat. The setup page's Seats section lists who holds a
seat and when it was last used; an administrator can Release one at once (tq_excel_core.js version 6).

**Company logo.** The setup page's Logo section takes a PNG or JPEG (up to 1 MB) that replaces the logo
from NetSuite on PDFs, the panel header and new dashboards. Remove goes back to the NetSuite logo.

**Automatic updates.** On the setup page, under Updates, enter the address Techno Q gives you. Check
for an update compares versions; Install it downloads the new setup file and runs it. The download must
match the description Techno Q signed with the licence key (scripts/sign_update.py), so a tampered
update is refused. An update keeps the settings, the licence and the seat. The connector also looks
for an update a minute after it starts and then weekly; when a newer signed version is there the panel
shows one line with Install. `"autoUpdateCheck": false` in config.json turns the weekly look off.

Checked on the sandbox (Nov 19 2026): version 6 deployed with SDF; the seat list showed the per-user seat
and the old per-PC one, and Release freed it; the logo route and the status update field answered from
the reinstalled connector; 83 server tests. Not yet tried by a customer: the installer's NetSuite and
licence pages on a fresh PC, and an update actually offered by the weekly check.

Checked: 75 server tests. Against the sandbox:
- the seat registration
- type:Stat
- clear messages for Book and Currency (the sandbox has neither multi-book nor OneWorld)
- TQ.SQL
- the journal check (an unbalanced entry is refused)
- the wizard, with location sheets, highlights and comments
- the panel and the wizard in a browser, against a stand-in for Excel

**Not done:** posting a real journal entry; reporting currency and book on an account that has them;
the ribbon and everything above inside Excel itself.

Checked so far: 63 server tests on sample data (they recalculate the generated reports and
compare with the ledger). Against the sandbox (Nov 2026): the reports, chat edit commands and
designer requests; the balance sheet for Oct 2026 now balances to 0.00 with the 34 accounts
that have no number included; budget totals matched the budget records by category; the
drill-down breakdown, account detail, report pack and summary sheet were built. The Designer
and Data tables sections were driven in a browser against a stand-in for Excel. The standard reports,
custom layouts, the Designer and the Data tables section were confirmed working in desktop Excel
(Nov 2026). **Not yet tried inside Excel**: budget columns, account detail rows, the summary
sheet, the year dropdown and the drill-down on a total.

### AI add-on (optional)

The chat's built-in reader is free and works offline. It now also drops polite wording ("can
you please show me..."), puts right spelling slips in command words ("ballance sheet", "by
quater") and says how it read them, takes several requests in one message (`custom P&L for 2026
by quarter then show account detail; remove column Total`) and offers the nearest commands when
it does not understand.

For free wording and for **screenshots of a report format**, a customer can switch on the AI
add-on with their own Anthropic API key. Add this to `server\config.json` and restart the server:

```
"ai": { "apiKey": "<the customer's own key>", "model": "claude-opus-5-5" }
```

- Off unless the key is there. The panel's status line then shows "AI add-on on".
- The AI is used only when a picture is attached (Picture button, or paste into the message
  box), when a message starts with `ai:`, or when the built-in reader does not understand.
- The AI never writes to the workbook. It returns commands for the built-in reader and, for a
  picture, the report's rows with account numbers; both go through the same checks as typed
  commands. An account number that is not in NetSuite is left out and listed.
- **What leaves the PC** when it is used: the message, the picture, the chart of accounts
  (number, name, type), the names of periods, departments, classes, locations and subsidiaries,
  and the lines and columns of the workbook's layouts. No balances and no transactions.
- The customer pays their AI provider for each use. The key is in `config.json` in plain text,
  like the NetSuite token, so the file must stay private.
- **Not yet tried against the real AI service**: it is tested with a stand-in that returns
  answers in the service's format.

Known gaps:

- In sample mode the balance sheet check is not 0, because the sample numbers are random.
- No parent-account sub-totals yet (see below).
- Budgets are read as entered in NetSuite. A budget saved for one department is left out when
  the report is filtered to another, and a budget with no department is left out of a
  department-filtered column.
- Custom layouts cover period amounts (a P&L). A layout line can name balance sheet accounts
  and will show their closing balance, but there is no balance check row for a custom layout.
- Custom layouts read ledger balances only. Other NetSuite records come in through the Data
  tables section (section 4) as tables on their own sheets; a layout line cannot use them directly.
- The chat is a command reader, not an AI: it cannot read a screenshot.
- The sandbox returns no subsidiaries, so the entity dropdown has not been seen working.

## 4. Data queries (any NetSuite data)

The **Data tables** section of the Reports tab brings tables of NetSuite data into the workbook, not only ledger balances:

- **A saved search**, exactly as it is defined in NetSuite: find it by name and click it.
- **A query built by clicking**: pick a record (customers, vendors, items, transactions,
  transaction lines, GL lines, stock by location and others, or type any table name such as
  a custom record), tick the fields, optionally bring fields from a linked record, add
  filters and sorting. "Read all fields of this record" lists every field NetSuite returns.
- **Query text (SuiteQL)**: the builder writes a SELECT that can be edited, or one can be
  typed from scratch, so anything SuiteQL can reach can be brought in.

Each query fills its own sheet, rewritten in place on every run, so pivot tables and
formulas that point at it keep working. The queries are listed on a sheet called
`TQ Queries` and so are saved with the workbook; **Refresh data** (or `refresh data` in the
chat) runs them all again. A query can use a parameter as `{from}`; its value is set once
under Parameters and applies to every query that uses it. In the chat:
`list saved searches invoice`, `load search Open Invoices`, `refresh data`.

**Ready-made queries.** "Start from a ready-made query" offers receivables aging, payables
aging, open invoices, purchases by vendor, purchase orders, sales by customer, sales by item,
stock on hand, and budgets by account and period. Each one fills the query text, where it can
be changed before it is added. All nine were run against the sandbox.

Limits: 5,000 rows per query unless Max rows is raised (up to 20,000). A new query is
refused if the workbook already has a sheet with its name, so it cannot wipe that sheet.

**Access.** Queries run in NetSuite as the user's own role, not in the elevated Suitelet
used for balances. A user gets only the records and fields their role can see in NetSuite.

**NetSuite script for this.** The query code lives in `tq_excel_restlet.js`, next to the
part that forwards balance requests to the Suitelet. An account installed before the Data tables section
existed needs that one file replaced in the File Cabinet (open the file record, Edit, paste
or choose the new file); the script record and its deployment stay as they are.
`tq_excel_core.js` and `tq_excel_elevated.js` are unchanged. Until then the Data tables section answers
that the account's scripts are older than the connector; everything else works.

A parameter whose value is a date written as `2026-01-31` is sent to NetSuite as a date, so
`trandate >= {from}` works as typed.

Checked so far: 4 server tests on sample data; the Data tables section driven in a browser against a
stand-in for Excel; and against the sandbox (Nov 2026, Administrator role) listing saved
searches, reading a record's fields, running a saved search, a built query with a join,
totals and a date parameter, and stock by location. **Not yet tried**: inside Excel, and
with a role that has limited permissions.

## Not in this prototype yet

- **Per-user sign-in.** One access token is shared by everyone using this server, so
  everyone sees what that token's role sees. The product needs OAuth 2.0 per user.
- **Statistical accounts, parent-account roll-ups, non-primary accounting books.**
- **Shared server and central deployment.** Each PC runs its own connector (see Installing).
  A shared server for many users and deployment from the Microsoft 365 admin centre are not built.
- **Automatic updates.** Running a newer setup file over an installed copy updates it.
- **Excel on the web** has not been tried. Desktop Excel (Microsoft 365, Windows) loaded
  the add-in and worked against the sandbox in Nov 2026.

## How access works (settled in sandbox testing, Nov 2026)

A RESTlet always runs as the caller's own role, and a query run that way silently
leaves out transaction types the role cannot view. In the sandbox a finance role
without Vendor Prepayment, Assembly Build and Inventory Worksheet access got balances
that differed from NetSuite's by exactly those transactions, with no error.

So the RESTlet (`tq_excel_restlet.js`) only forwards the request to a Suitelet
(`tq_excel_elevated.js`) deployed as Administrator, which runs the logic in
`tq_excel_core.js` with full data access. With this in place all 44 accounts with a
balance for Oct 2026, and all income/expense activity for 2025, matched NetSuite's
transaction search to the cent.

Who may ask is decided in `assertReportAccess`: the user must hold at least one role
with the Financial Statements, Trial Balance or General Ledger report permission
(or Administrator).

The caller's role itself only needs **Log in using Access Tokens** (Setup), and must be
a role actually assigned to the user. The token form also lists roles the user does
not have, and those tokens fail to log in.

Consequences to design for before a client rollout:

- **Drill-down shows every transaction** on the account, including types the user's
  role cannot open in NetSuite (payroll journals, for example).
- **Role restrictions are not applied.** A role limited to one subsidiary or department
  still gets company-wide balances.
- **The refusal path is untested.** Only a user who passes the access check was tried.
- NetSuite reports the caller's role inside the Suitelet even though data access is
  elevated, so the code cannot detect a Suitelet deployed without Execute As Role. Run
  `scripts\check_netsuite.py` and compare with the Trial Balance after every install.
