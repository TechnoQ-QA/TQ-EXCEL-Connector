"""TQ Excel Connector - starting points for a layout.

Two ways to start a design without adding its lines one by one:

    build(ledger, template id, name)   a ready-made layout (P&L, balance sheet, cash flow). The lines
                                       use account types, so they fit any NetSuite account.
    from_sheet(ledger, name, rows)     the customer's own format, typed on a worksheet: the row
                                       labels become the lines and are matched to the accounts.
"""

import re

import layouts
from ledger import LedgerError, PL_SIGNS
from reports import ReportError, sheet_name

MAX_IMPORT_ROWS = 400
MAX_IMPORT_COLS = 30

TEMPLATES = [
    {"id": "monthytd", "name": "Month and year to date P&L (rolls forward)",
     "about": "The month and the year to date, each against last year, with the change and change %. Pick "
              "another month in B8 and every column moves with it: the monthly management pack."},
    {"id": "management", "name": "Management P&L, account by account",
     "about": "One line for each profit and loss account that has a balance, with gross profit, net profit, "
              "margins and average daily sales."},
    {"id": "summary", "name": "P&L summary, one line per group",
     "about": "Sales, cost of sales, gross profit, operating expenses, operating profit, other income and "
              "expenses, net profit, with margins."},
    {"id": "balance", "name": "Balance sheet",
     "about": "Assets, liabilities and equity by group, this period against the same period last year, "
              "with a check line that must be zero."},
    {"id": "cashflow", "name": "Cash flow, indirect",
     "about": "Net income plus the movement in each balance sheet group, month by month, checked against "
              "the movement in the bank accounts."},
]

# (section, line, type, accounts or formula, sign, format)
_SUMMARY = [
    ("Revenue", "Sales", "accounts", "type:Income", "", ""),
    ("Cost of Sales", "Cost of Sales", "accounts", "type:COGS", "-1", ""),
    ("Gross Profit", "Gross Profit", "formula", "=[Sales]+[Cost of Sales]", "", "bold"),
    ("Gross Profit", "Gross Margin", "formula", "=[Gross Profit]/[Sales]", "", "percent"),
    ("Operating Expenses", "Operating Expenses", "accounts", "type:Expense", "-1", ""),
    ("Operating Profit", "Operating Profit", "formula", "=[Gross Profit]+[Operating Expenses]", "", "bold"),
    ("Operating Profit", "Operating Margin", "formula", "=[Operating Profit]/[Sales]", "", "percent"),
    ("Other Income and Expenses", "Other Income", "accounts", "type:OthIncome", "", ""),
    ("Other Income and Expenses", "Other Expenses", "accounts", "type:OthExpense", "-1", ""),
    ("Result", "Net Profit", "formula", "=[Operating Profit]+[Other Income]+[Other Expenses]", "", "bold"),
    ("Result", "Net Margin", "formula", "=[Net Profit]/[Sales]", "", "percent"),
    ("Result", "Average Sales Daily", "formula", "=[Sales]/days", "", ""),
]
_BALANCE = [
    ("Assets", "Bank and Cash", "accounts", "type:Bank", "", ""),
    ("Assets", "Receivables", "accounts", "type:AcctRec,UnbilledRec", "", ""),
    ("Assets", "Other Current Assets", "accounts", "type:OthCurrAsset,DeferExpense", "", ""),
    ("Assets", "Fixed Assets", "accounts", "type:FixedAsset", "", ""),
    ("Assets", "Other Assets", "accounts", "type:OthAsset", "", ""),
    ("Assets", "Total Assets", "subtotal", "", "", ""),
    ("Liabilities", "Payables", "accounts", "type:AcctPay", "", ""),
    ("Liabilities", "Credit Cards", "accounts", "type:CredCard", "", ""),
    ("Liabilities", "Other Current Liabilities", "accounts", "type:OthCurrLiab,DeferRevenue", "", ""),
    ("Liabilities", "Long Term Liabilities", "accounts", "type:LongTermLiab", "", ""),
    ("Liabilities", "Total Liabilities", "subtotal", "", "", ""),
    ("Equity", "Equity", "accounts", "type:Equity", "", ""),
    ("Equity", "Retained Earnings", "accounts", "calc:RetainedEarnings", "", ""),
    ("Equity", "Net Income This Year", "accounts", "calc:NetIncomeYTD", "", ""),
    ("Equity", "Total Equity", "subtotal", "", "", ""),
    ("Check", "Total Liabilities and Equity", "formula", "=[Total Liabilities]+[Total Equity]", "", "bold"),
    ("Check", "Difference (must be zero)", "formula",
     "=[Total Assets]-[Total Liabilities]-[Total Equity]", "", ""),
]
_CASHFLOW = [
    ("Operating Activities", "Net Income", "accounts", "calc:NetIncome", "", ""),
    ("Operating Activities", "Change in Receivables", "accounts", "move:type:AcctRec,UnbilledRec", "-1", ""),
    ("Operating Activities", "Change in Other Current Assets", "accounts",
     "move:type:OthCurrAsset,DeferExpense", "-1", ""),
    ("Operating Activities", "Change in Payables", "accounts", "move:type:AcctPay,CredCard", "", ""),
    ("Operating Activities", "Change in Other Current Liabilities", "accounts",
     "move:type:OthCurrLiab,DeferRevenue", "", ""),
    ("Operating Activities", "Cash from Operating Activities", "subtotal", "", "", ""),
    ("Investing Activities", "Fixed Assets, net of depreciation", "accounts", "move:type:FixedAsset", "-1", ""),
    ("Investing Activities", "Other Assets", "accounts", "move:type:OthAsset", "-1", ""),
    ("Investing Activities", "Cash from Investing Activities", "subtotal", "", "", ""),
    ("Financing Activities", "Long Term Liabilities", "accounts", "move:type:LongTermLiab", "", ""),
    ("Financing Activities", "Equity", "accounts", "move:type:Equity", "", ""),
    ("Financing Activities", "Cash from Financing Activities", "subtotal", "", "", ""),
    ("Cash", "Net Change in Cash", "formula",
     "=[Cash from Operating Activities]+[Cash from Investing Activities]+[Cash from Financing Activities]",
     "", "bold"),
    ("Cash", "Movement in Bank Accounts", "accounts", "move:type:Bank", "", ""),
    ("Cash", "Difference (must be zero)", "formula", "=[Net Change in Cash]-[Movement in Bank Accounts]", "", ""),
]
_LINES = {"summary": ("P&L Summary", _SUMMARY), "balance": ("Balance Sheet", _BALANCE),
          "cashflow": ("Cash Flow", _CASHFLOW)}


def _design(name, title, rows):
    design = {"name": name, "title": title, "filters": {}, "pages": "", "detail": False, "lines": [], "columns": []}
    for section, label, kind, source, sign, fmt in rows:
        design["lines"].append(layouts.new_line(label, kind, source, sign, fmt, section))
    return design


def _balance_columns(ledger):
    """As at the report's month, the same month last year and last year end, with the change since then.
    The columns follow the month cell (B8), so the balance sheet rolls forward with it."""
    new = layouts.new_column
    return [new("As at month", "period", "CM", "CM"),
            new("Same month last year", "period", "LY", "LY"),
            new("Last year end", "period", "FYS-1", "FYS-1"),
            new("Change since year end", "formula", "", "", None, "=[As at month]-[Last year end]")]


def month_ytd_columns():
    """The month and the year to date, each against last year: they follow the month cell (B8)."""
    new = layouts.new_column
    return [new("Month", "period", "CM", "CM"),
            new("Month last year", "period", "LY", "LY"),
            new("Change", "formula", "", "", None, "=[Month]-[Month last year]"),
            new("Change %", "formula", "", "", None, "=[Change]/ABS([Month last year])", "percent"),
            new("Year to date", "period", "FYS", "CM"),
            new("YTD last year", "period", "FYS-12", "LY"),
            new("Change YTD", "formula", "", "", None, "=[Year to date]-[YTD last year]"),
            new("Change % YTD", "formula", "", "", None, "=[Change YTD]/ABS([YTD last year])", "percent")]


def build(ledger, template, name):
    """A ready-made design. Raises ReportError for an unknown template."""
    template = (template or "management").strip().lower()
    if template == "management":
        return layouts.sample(ledger, name)
    if template == "monthytd":
        design = layouts.sample(ledger, name)
        design["title"] = "Monthly P&L"
        design["columns"] = month_ytd_columns()
        return design
    if template not in _LINES:
        raise ReportError('There is no template called "%s".' % template)
    title, rows = _LINES[template]
    design = _design(name, title, rows)
    try:
        design["columns"] = _balance_columns(ledger) if template == "balance" else layouts.default_columns(ledger)
    except LedgerError as err:
        raise ReportError(str(err))
    return design


# --------------------------------------------------------- a format on a sheet --

_GROUPS = [  # what a row label usually means when it is not an account's own name
    (r"other\s+(income|revenue)s?|non[- ]operating\s+income", "type:OthIncome"),
    (r"other\s+(expense|cost)s?|non[- ]operating\s+expenses?", "type:OthExpense"),
    (r"(total\s+)?(net\s+)?(sales|revenues?|income|turnover)", "type:Income"),
    (r"(total\s+)?(cost\s+of\s+(goods\s+)?(sales|sold|revenue)|cogs|direct\s+costs?)", "type:COGS"),
    (r"(total\s+)?((operating|general|admin\w*|indirect)\s+)*(expenses?|overheads?|opex|costs?)", "type:Expense"),
    (r"(cash|bank)(\s+(and|&)\s+(cash|bank)(\s+equivalents)?)?(\s+accounts?)?", "type:Bank"),
    (r"(accounts?\s+|trade\s+)?receivables?", "type:AcctRec,UnbilledRec"),
    (r"(accounts?\s+|trade\s+)?payables?", "type:AcctPay"),
    (r"fixed\s+assets?|property,?\s+plant.*", "type:FixedAsset"),
    (r"(share\s+capital|equity|capital)", "type:Equity"),
    (r"retained\s+earnings", "calc:RetainedEarnings"),
]
_SUBTOTAL = re.compile(r"^(total|sub[- ]?total|net\s+(sales|revenues?))\b", re.I)
_RESULT = re.compile(r"^(gross|net|operating)\s+(profit|income|loss|result|surplus)|\bebit(da)?\b|"
                     r"^(profit|loss|result|surplus)\b|profit\s+(before|after)", re.I)
_PERCENT = re.compile(r"%|\bmargin\b|\bpercent", re.I)
_HEADERS = {"name", "names", "account", "accounts", "description", "particulars", "line", "lines", "details"}
_LEAD_NUMBER = re.compile(r"^\s*([\w.]*\d[\w.]*)\s*[-:– ]\s*(.+)$")


def _str(value):
    if isinstance(value, bool) or value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        value = int(value)
    return str(value).strip()


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def from_sheet(ledger, name, title, rows, columns=None):
    """A design from a format typed on a worksheet. Returns (design, notes).

    The column with the most text is taken as the row labels. A label becomes a line when it is, in
    this order: next to a cell of account numbers; an account number followed by a name; a total; a
    result such as Gross Profit; a ratio; an account's name; a group such as Sales or Expenses.
    A label on a row with nothing else becomes a section heading. The rest is listed in the notes.
    """
    rows = [list(row[:MAX_IMPORT_COLS]) for row in rows[:MAX_IMPORT_ROWS] if isinstance(row, list)]
    width = max((len(row) for row in rows), default=0)
    rows = [row + [""] * (width - len(row)) for row in rows]
    is_label = lambda v: isinstance(v, str) and re.search(r"[A-Za-z]{2}", v) and not v.startswith("=")  # noqa: E731
    counts = [sum(1 for row in rows if is_label(row[c])) for c in range(width)]
    if not counts or max(counts) < 2:
        raise ReportError("That sheet has no column of row names to read. Put the report's lines in one "
                          "column, one per row, then try again.")
    at = columns[0] if columns else counts.index(max(counts))
    try:
        accounts = ledger.accounts()
        numbers = {a["number"].lower(): a for a in accounts}
        by_name = {}
        for a in accounts:
            for key in {a["name"].lower(), a["name"].split(" : ")[-1].strip().lower()}:
                by_name.setdefault(key, []).append(a["number"])

        def spec(value):
            """The cell as an accounts entry, if every part of it is a known account, range or type."""
            token = _str(value)
            if not token or len(token) > 300 or (isinstance(value, float) and value != int(value)):
                return ""
            try:
                prefix, _, rest = ledger.split_prefix(token)
                if rest.lower().startswith("calc:"):
                    return token
                return token if ledger.members(rest) else ""
            except LedgerError:
                return ""

        hits = [sum(1 for row in rows if is_label(row[at]) and spec(row[c])) if c != at else 0 for c in range(width)]
        labelled = max(counts)
        account_col = hits.index(max(hits)) if hits and max(hits) >= max(2, labelled // 5) else None
        if columns:
            account_col = columns[1]
        types = {a["number"]: a["type"] for a in accounts}

        def sign_of(source):
            try:
                members = ledger.members(ledger.split_prefix(source)[2])
            except LedgerError:
                return ""
            negative = sum(1 for n in members if PL_SIGNS.get(types.get(n)) == -1)
            return "-1" if members and negative * 2 > len(members) else ""

        design = {"name": name, "title": title, "filters": {}, "pages": "", "detail": False,
                  "lines": [], "columns": []}
        used, skipped, section, heading, after_result = set(), [], "", "", False
        has_data = any(_is_number(v) for row in rows for c, v in enumerate(row) if c != at)
        labels = [row[at].strip() if is_label(row[at]) else "" for row in rows]

        def is_account(label):
            lead = _LEAD_NUMBER.match(label)
            return label.lower() in by_name or bool(lead and lead.group(1).lower() in numbers)
        open_lines, blocks = [], []   # blocks: the lines a result such as Net Profit adds up

        def unique(label):
            label = re.sub(r"\s+", " ", label).strip(" :")[:80]
            out, n = label, 2
            while out.lower() in used:
                out, n = "%s (%d)" % (label, n), n + 1
            used.add(out.lower())
            return out

        def add(label, kind, source="", sign="", fmt=""):
            nonlocal section, heading, after_result
            if heading:
                section, heading, after_result = heading, "", False
                del open_lines[:]
            elif kind == "formula" and fmt != "percent" and section and not after_result:
                section, after_result = re.sub(r"\s+", " ", label).strip(" :")[:80], True  # a result stands apart
                del open_lines[:]
            elif kind != "formula" and after_result:
                section, after_result = "", False
            label = unique(label)
            design["lines"].append(layouts.new_line(label, kind, source, sign, fmt, section))
            return label

        def accounts_line(label, source):
            made = add(label, "accounts", source, sign_of(source))
            blocks.append(made)
            open_lines.append(made)

        for n, row in enumerate(rows):
            label = labels[n]
            if not label:
                continue
            following = next((l for l in labels[n + 1:] if l), "")
            lowered = re.sub(r"\s+", " ", label.lower()).strip(" :")
            others = [v for c, v in enumerate(row) if c != at and _str(v)]
            source = spec(row[account_col]) if account_col is not None else ""
            lead = _LEAD_NUMBER.match(label)
            if lowered in _HEADERS and not any(_is_number(v) for v in others):
                continue                 # the heading of the names column
            if source:
                accounts_line(label, source)
            elif lead and lead.group(1).lower() in numbers:
                accounts_line(lead.group(2), numbers[lead.group(1).lower()]["number"])
            elif _PERCENT.search(lowered):
                results = [l["label"] for l in design["lines"] if l["kind"] != "accounts" and l["format"] != "percent"]
                if results and blocks:
                    add(label, "formula", "=[%s]/[%s]" % (results[-1], blocks[0]), "", "percent")
                else:
                    skipped.append(label)
            elif _SUBTOTAL.match(lowered) and open_lines and not heading:
                mine = list(open_lines)
                total = add(label, "subtotal")
                blocks[:] = [b for b in blocks if b not in mine] + [total]
                del open_lines[:]
            elif _RESULT.search(lowered) and blocks:
                add(label, "formula", "=" + "+".join("[%s]" % b for b in blocks), "", "bold")
            elif lowered in by_name:
                accounts_line(label, ", ".join(by_name[lowered]))
            elif not others and (has_data or columns or is_account(following)
                                 or not any(re.fullmatch(p, lowered) for p, _ in _GROUPS)):
                heading = label          # a name on its own row: the heading of what follows
            else:
                group = next((source for pattern, source in _GROUPS if re.fullmatch(pattern, lowered)), "")
                if group and (group.startswith("calc:") or ledger.members(group)):
                    accounts_line(label, group)
                elif not others:
                    heading = label
                else:
                    skipped.append(label)
        if not design["lines"]:
            raise ReportError("None of the row names on that sheet could be matched to an account, an account "
                              "group or a total. Put the account numbers in a column next to the names, then "
                              "try again.")
        design["columns"] = layouts.default_columns(ledger)
    except LedgerError as err:
        raise ReportError(str(err))
    notes = ["%s %d lines%s." % ("Built" if columns else "Read",
        len(design["lines"]), " using the account numbers given" if columns else
        " using the account numbers in column %s" % layouts.col_letter(account_col) if account_col is not None else ", matching the names to accounts and account groups")]
    if skipped:
        shown = ", ".join(skipped[:10]) + (" and %d more" % (len(skipped) - 10) if len(skipped) > 10 else "")
        notes.append("%d row(s) could not be matched and were left out: %s. Add them in the Designer, or put "
                     "their account numbers in a column next to the names and import again." % (len(skipped), shown))
    notes.append("Check each line's accounts on the layout sheet before you rely on the numbers.")
    return design, notes


def layout_name(wanted, taken):
    """A free "TQ Layout ..." sheet name."""
    name = sheet_name(wanted)
    if name and not name.lower().startswith(layouts.SHEET_PREFIX.lower()):
        name = sheet_name("%s %s" % (layouts.SHEET_PREFIX, name))
    base = name or layouts.SHEET_PREFIX
    name, n = base, 2
    while name.lower() in taken:
        name, n = sheet_name("%s %d" % (base[:28], n)), n + 1
    return name
