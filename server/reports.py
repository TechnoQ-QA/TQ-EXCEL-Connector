"""TQ Excel Connector - report layouts.

Turns a report request (which statement, which periods, which layout) into one
block of cells for a new worksheet. The numbers are live TQ.BALANCE formulas
that point at the period names in the header and at four filter cells
(subsidiary, department, class, location), so the sheet can be changed and
refreshed without building it again.

A spec is a plain dict:
    report   "pl" | "bs" | "tb"
    start, end        positions in ledger.periods()
    compare           None or (start, end) of the periods to compare with
    layout   "total" | "month" | "quarter"
    split    None or a segment kind: one column per subsidiary/department/...
    detail   True = one row per account, False = one row per section
    zero     True = keep accounts with no balance
    filters  {segment kind: name}
    title    optional heading instead of the statement's name
    yearCell optional year: adds a Year cell that every period in the header is built from
"""

import re

from ledger import FILTER_KINDS, LedgerError, SEGMENT_KINDS

NUMBER_FORMAT = "#,##0.00;(#,##0.00)"
PERCENT_FORMAT = "0.0%"
MAX_SPLIT_COLUMNS = 60   # columns kept after the empty ones are dropped
MAX_PROBE = 80000        # balances looked up while laying out one report
MAX_CELLS = 60000
GREY = "#6B7280"
YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")
CREDIT_TYPES = {"AcctPay", "CredCard", "OthCurrLiab", "LongTermLiab", "DeferRevenue",
                "Equity", "Income", "OthIncome"}
TITLES = {"pl": "Income Statement", "bs": "Balance Sheet", "tb": "Trial Balance"}
KIND_LABEL = {"subsidiary": "Subsidiary", "department": "Department", "class": "Class", "location": "Location",
              "filters": "Other filters"}
PARAM_ROW = {"filters": 3, "subsidiary": 4, "department": 5, "class": 6, "location": 7}  # 1-based sheet rows
FILTER_NOTE = "e.g. Brand: Premium; Name: Al Noor"
TRANSLATION_KEY = ("calc:Translation", "Currency translation (consolidation)")
MAX_DROPDOWNS = 4        # dropdown filter rows beside the header (E4:F7)
BLOCK_COL = 4            # column E: the field; F: its value, picked from a dropdown
MAX_LIST = 1000          # longest list offered as a dropdown
FIELD_WORDS = {"customer": "name", "vendor": "name", "employee": "name", "entity": "name", "items": "item"}

PL_SECTIONS = [
    ("income", "Income", ["Income"]),
    ("cogs", "Cost of Sales", ["COGS"]),
    ("expense", "Expenses", ["Expense"]),
    ("othincome", "Other Income", ["OthIncome"]),
    ("othexpense", "Other Expenses", ["OthExpense"]),
]
BS_SECTIONS = [
    ("ASSETS", "Total Assets", [
        ("Current Assets", ["Bank", "AcctRec", "UnbilledRec", "OthCurrAsset", "DeferExpense"]),
        ("Fixed Assets", ["FixedAsset"]),
        ("Other Assets", ["OthAsset"]),
    ]),
    ("LIABILITIES", "Total Liabilities", [
        ("Current Liabilities", ["AcctPay", "CredCard", "OthCurrLiab", "DeferRevenue"]),
        ("Long Term Liabilities", ["LongTermLiab"]),
    ]),
    ("EQUITY", "Total Equity", [
        ("Equity", ["Equity"]),
    ]),
]
BS_EQUITY_KEYS = [("calc:RetainedEarnings", "Retained Earnings (prior years)"),
                  ("calc:NetIncomeYTD", "Net Income (current year)")]


class ReportError(Exception):
    pass


def _values(ledger, kind):
    try:
        return ledger.segments(kind)
    except LedgerError:
        return []


def dropdown_fields(ledger, names):
    """'Brand, Name' (or a list) -> [{id, label, kind}] for the dropdown filter rows. Raises ReportError."""
    wanted = names if isinstance(names, (list, tuple)) else str(names or "").split(",")
    wanted = [str(n).strip() for n in wanted if str(n).strip()]
    if not wanted:
        return []
    try:
        fields = ledger.filter_fields()
    except LedgerError as err:
        raise ReportError("The list of fields could not be read from NetSuite: %s" % err)
    out = []
    for name in wanted:
        key = FIELD_WORDS.get(name.lower(), name.lower())
        hit = next((f for f in fields if key in (f["id"].lower(), f["label"].lower())), None)
        if not hit:
            raise ReportError('"%s" is not a field reports can be filtered on. These are: %s.'
                              % (name, ", ".join(f["label"] for f in fields)))
        if hit["id"] in SEGMENT_KINDS:
            continue  # subsidiary, department, class and location have their own cells, with dropdowns
        if hit not in out:
            out.append(hit)
    if len(out) > MAX_DROPDOWNS:
        raise ReportError("A report has room for %d dropdown filters." % MAX_DROPDOWNS)
    return out


def _parts(text):
    """'Brand: Premium; Name: Al Noor' -> [('Brand', 'Premium'), ('Name', 'Al Noor')]."""
    out = []
    for part in str(text or "").split(";"):
        field, sep, value = part.partition(":")
        if sep and field.strip():
            out.append((field.strip(), value.strip()))
    return out


def add_filter_lists(action, ledger, dropdowns=(), fixed="", extra=()):
    """Dropdown lists on the report's filter cells.

    - B4:B7 (subsidiary, department, class, location) pick from their NetSuite lists.
    - dropdowns: fields such as Brand or Name get a row each in a "More filters" block at E4:F7; the value
      cell F picks from that field's values (Name and Item are typed: their lists are too long).
    - B3 then becomes a formula joining the block's choices with any fixed other filters, so every
      number in the report follows the dropdowns.
    The lists themselves go to the right of the report. Typing is still allowed in every filter cell,
    so "A, B" and "not A" work.
    """
    rows = action["rows"]
    action.setdefault("dropdowns", [])
    action.setdefault("textRanges", [])
    action.setdefault("formats", [])
    action.setdefault("widths", [])
    lists = [("%s list" % KIND_LABEL[kind], names, [(PARAM_ROW[kind] - 1, 1)])
             for kind in SEGMENT_KINDS for names in [_values(ledger, kind)] if names and len(names) <= MAX_LIST]
    lists = list(extra) + lists   # extra: (title, names, [(row, col)]) such as the report's month list
    width = len(rows[0])
    if dropdowns:
        width = max(width, BLOCK_COL + 2)
        for row in rows:
            row += [""] * (width - len(row))
        chosen = {field.lower(): value for field, value in _parts(fixed)}
        labels = {f["label"].lower() for f in dropdowns} | {f["id"].lower() for f in dropdowns}
        rest = "; ".join("%s: %s" % (field, value) for field, value in _parts(fixed) if field.lower() not in labels)
        rows[2][BLOCK_COL], rows[2][BLOCK_COL + 1] = "More filters", "Pick (or type A, B / not A)"
        for i, field in enumerate(dropdowns):
            r = 3 + i
            rows[r][BLOCK_COL] = field["label"]
            rows[r][BLOCK_COL + 1] = chosen.get(field["label"].lower(), chosen.get(field["id"].lower(), ""))
            names = _values(ledger, field["id"]) if field.get("kind") not in ("name", "item") else []
            if names and len(names) <= MAX_LIST:
                lists.append(("%s list" % field["label"], names, [(r, BLOCK_COL + 1)]))
        e, f = col_letter(BLOCK_COL), col_letter(BLOCK_COL + 1)
        joined = '"%s"' % ("; " + rest).replace('"', '""') if rest else '""'
        for i in range(len(dropdowns)):
            r = 4 + i
            joined += '&IF($%s$%d="","","; "&$%s$%d&": "&$%s$%d)' % (f, r, e, r, f, r)
        rows[2][1] = "=MID(%s,3,2000)" % joined
        for b in action["textRanges"]:          # B3 is a formula now, not text
            if b[0] <= 2 <= b[2] and b[1] <= 1 <= b[3]:
                b[0] = 3
        action["textRanges"].append([3, BLOCK_COL, 2 + len(dropdowns), BLOCK_COL + 1])
        action["formats"] += [
            {"r0": 2, "c0": BLOCK_COL, "r1": 2, "c1": BLOCK_COL + 1, "bold": True, "bottom": "thin"},
            {"r0": 3, "c0": BLOCK_COL + 1, "r1": 2 + len(dropdowns), "c1": BLOCK_COL + 1, "fill": "#FFF8E1"},
            {"r0": 2, "c0": 1, "r1": 2, "c1": 1, "color": GREY},
        ]
        action["widths"] += [{"c0": BLOCK_COL, "c1": BLOCK_COL, "width": 120},
                             {"c0": BLOCK_COL + 1, "c1": BLOCK_COL + 1, "width": 170}]
    if not lists:
        return
    height = max(len(rows), 1 + max(len(names) for _, names, _ in lists))
    rows += [[""] * width for _ in range(height - len(rows))]
    for row in rows:
        row += [""] * (width + 1 + len(lists) - len(row))
    for n, (title, names, cells) in enumerate(lists):
        col = width + 1 + n
        rows[0][col] = title
        for i, name in enumerate(names):
            rows[1 + i][col] = name
        letter = col_letter(col)
        action["textRanges"].append([1, col, len(names), col])
        for r, c in cells:
            action["dropdowns"].append({"r": r, "c": c, "free": True,
                                        "formula": "=$%s$2:$%s$%d" % (letter, letter, len(names) + 1)})
    last = width + len(lists)
    action["formats"].append({"r0": 0, "c0": width + 1, "r1": height - 1, "c1": last, "color": GREY})
    action["formats"].append({"r0": 0, "c0": width + 1, "r1": 0, "c1": last, "bold": True})
    action["widths"] += [{"c0": width, "c1": width, "width": 30}, {"c0": width + 1, "c1": last, "width": 150}]


def col_letter(index):
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def range_label(ledger, start, end):
    periods = ledger.periods()
    if start == end:
        return periods[start]["name"]
    fiscal_year = periods[start]["fy"]
    whole = (periods[end]["fy"] == fiscal_year and ledger.fy_start(start) == start
             and (end + 1 == len(periods) or periods[end + 1]["fy"] != fiscal_year))
    if whole:
        return ("%d" if ledger.fiscal_start_month == 1 else "FY %d") % fiscal_year
    return "%s to %s" % (periods[start]["name"], periods[end]["name"])


def _chunks(ledger, start, end, layout):
    """[(label, first, last)] for the columns of one period range."""
    periods = ledger.periods()
    if layout == "month":
        return [(periods[i]["name"], i, i) for i in range(start, end + 1)]
    if layout == "quarter":
        out = []
        for i in range(start, end + 1):
            key = (periods[i]["fy"], periods[i]["fm"] // 3)
            if out and out[-1][0] == key:
                out[-1][2] = i
            else:
                out.append([key, i, i])
        return [("Q%d %d" % (key[1] + 1, key[0]), first, last) for key, first, last in out]
    return [(range_label(ledger, start, end), start, end)]


def year_formula(text, base, cell):
    """'Jan 2025' with base 2026 -> ="Jan "&($B$8-1), so the name follows the Year cell."""
    parts = YEAR_RE.split(text)
    if len(parts) == 1:
        return text
    out = []
    for n, part in enumerate(parts):
        if n % 2:
            delta = int(part) - base
            out.append(cell if not delta else "(%s%+d)" % (cell, delta))
        elif part or n == 0:
            out.append('"%s"' % part.replace('"', '""'))
    return "=" + "&".join(out)


class _Sheet:
    """Collects rows, formats and the header positions while a report is laid out."""

    def __init__(self, ledger, spec, columns, single):
        self.ledger, self.spec, self.columns, self.single = ledger, spec, columns, single
        self.width = 2 + len(columns)
        self.rows, self.formats, self.text_ranges = [], [], []
        filters = spec.get("filters") or {}
        self.year = spec.get("yearCell")
        split = spec.get("split")
        self._add([spec.get("title") or TITLES[spec["report"]]], bold=True, size=14)
        self._add([spec["subtitle"]], color=GREY)
        self._add([KIND_LABEL["filters"], filters.get("filters", ""), FILTER_NOTE])
        for kind in SEGMENT_KINDS:
            note = ""
            if spec.get("split") == kind:
                note = "shown by column"
            elif kind == "subsidiary":
                note = "Pick or type. Blank = all. A, B / not A."
            self._add([KIND_LABEL[kind], "" if spec.get("split") == kind else filters.get(kind, ""), note])
        self.formats.append({"r0": 2, "c0": 2, "r1": 6, "c1": 2, "color": GREY, "italic": True})
        self.text_ranges.append([2, 1, 6, 1])
        if self.year:
            self._add(["Year", self.year, "Change the year to move every period in the header."])
            self.year_row = len(self.rows)
            self.formats.append({"r0": self.year_row - 1, "c0": 2, "r1": self.year_row - 1, "c1": 2,
                                 "color": GREY, "italic": True})
            self.formats.append({"r0": self.year_row - 1, "c0": 1, "r1": self.year_row - 1, "c1": 1,
                                 "numberFormat": "0", "align": "left", "bold": True})
        self._add([])
        periods = ledger.periods()
        if self.year:
            named = lambda name: year_formula(name, self.year, "$B$%d" % self.year_row)  # noqa: E731
        else:
            named = lambda name: name  # noqa: E731
        if not single:
            for label, key in (("From period", "from"), ("To period", "to")):
                self._add(["", label] + [named(periods[c[key]]["name"]) if c["kind"] == "value" else ""
                                         for c in columns], color=GREY)
            self.from_row, self.to_row = len(self.rows) - 1, len(self.rows)
        self._add([spec.get("accountHeader") or "Account", "Name"]
                  + [named(c["label"]) if c["kind"] == "value" and not c.get("seg") else c["label"] for c in columns],
                  bold=True, bottom="thin")
        self.header_row = len(self.rows)
        first_header = self.header_row - (0 if single else 2)
        if not self.year:
            self.text_ranges.append([first_header - 1, 2, self.header_row - 1, self.width - 1])
        elif split:  # segment names stay text; the period rows are formulas and must not be
            self.text_ranges.append([self.header_row - 1, 2, self.header_row - 1, self.width - 1])
        self.formats.append({"r0": first_header - 1, "c0": 2, "r1": self.header_row - 1, "c1": self.width - 1,
                             "align": "right"})
        self.first_data_row = self.header_row + 1

    def _add(self, cells, **style):
        self.rows.append(list(cells) + [""] * (self.width - len(cells)))
        if style:
            self.formats.append(dict(r0=len(self.rows) - 1, c0=0, r1=len(self.rows) - 1, c1=self.width - 1, **style))
        return len(self.rows)  # 1-based sheet row

    def _balance(self, account_ref, j, row=None):
        column, letter = self.columns[j], col_letter(2 + j)
        if self.single:
            period_from = "%s$%d" % (letter, self.header_row)
            period_to = '""' if self.spec["report"] == "bs" else period_from
        else:
            period_from, period_to = "%s$%d" % (letter, self.from_row), "%s$%d" % (letter, self.to_row)
        args = []
        for kind in FILTER_KINDS:
            if column.get("seg") == kind:
                args.append("%s$%d" % (letter, self.header_row))
            elif self.spec.get("split") == kind:
                args.append('""')
            else:
                args.append("$B$%d" % PARAM_ROW[kind])
        return "TQ.BALANCE(%s,%s,%s,%s)" % (account_ref, period_from, period_to, ",".join(args))

    def _derived(self, j, row):
        column = self.columns[j]
        ref = lambda n: "%s%d" % (col_letter(2 + n), row)  # noqa: E731
        if column["kind"] == "var":
            return "=%s-%s" % (ref(column["a"]), ref(column["b"]))
        if column["kind"] == "pct":
            return '=IF(%s=0,"",%s/ABS(%s))' % (ref(column["b"]), ref(column["a"]), ref(column["b"]))
        if column["kind"] == "sum":
            return "=SUM(%s:%s)" % (ref(column["a"]), ref(column["b"]))
        return "=%s-SUM(%s:%s)" % (ref(column["total"]), ref(column["a"]), ref(column["b"]))  # rest

    def value_row(self, account, label, value, derived=True, **style):
        """value(j) -> the formula for value column j of this row.

        derived: True = the usual change/total formulas, False = blank, or a function (j, row) -> cell.
        """
        row = len(self.rows) + 1
        cells = [account, label]
        for j, column in enumerate(self.columns):
            if column["kind"] == "value":
                cells.append(value(j))
            elif callable(derived):
                cells.append(derived(j, row))
            else:
                cells.append(self._derived(j, row) if derived else "")
        return self._add(cells, **style)

    def account(self, number, name):
        row = self.value_row(number, name, lambda j: "=" + self._balance("$A%d" % (len(self.rows) + 1), j))
        self.formats.append({"r0": row - 1, "c0": 1, "r1": row - 1, "c1": 1, "indent": 1})
        return row

    def key(self, key, label, **style):
        return self.value_row(key, label, lambda j: "=" + self._balance("$A%d" % (len(self.rows) + 1), j), **style)

    def total(self, label, first, last, **style):
        style = dict({"bold": True, "top": "thin"}, **style)
        return self.value_row("", label, lambda j: "=SUM(%s%d:%s%d)" % (
            col_letter(2 + j), first, col_letter(2 + j), last), **style)

    def combine(self, label, terms, **style):
        """terms: [(sign, row or None)] -> e.g. =C12-C20."""
        def formula(j):
            parts = ["%s%s%d" % ("+" if sign > 0 else "-", col_letter(2 + j), row)
                     for sign, row in terms if row]
            return "=" + ("".join(parts).lstrip("+") if parts else "0")
        return self.value_row("", label, formula, **dict({"bold": True, "top": "thin"}, **style))

    def heading(self, label):
        return self._add(["", label], bold=True)

    def blank(self):
        return self._add([])

    def action(self, sheet_name):
        last = len(self.rows) - 1
        self.text_ranges.append([self.first_data_row - 1, 0, last, 0])
        self.formats.append({"r0": self.first_data_row - 1, "c0": 0, "r1": last, "c1": 0, "color": GREY})
        self.formats.append({"r0": self.first_data_row - 1, "c0": 2, "r1": last, "c1": self.width - 1,
                             "numberFormat": NUMBER_FORMAT})
        for j, column in enumerate(self.columns):
            if column["kind"] == "pct":
                self.formats.append({"r0": self.first_data_row - 1, "c0": 2 + j, "r1": last, "c1": 2 + j,
                                     "numberFormat": PERCENT_FORMAT})
        return {
            "type": "write", "target": "new_sheet", "sheetName": sheet_name, "startCell": "", "rows": self.rows,
            "textRanges": self.text_ranges, "formats": self.formats,
            "freeze": {"rows": self.header_row, "cols": 2},
            "widths": [{"c0": 0, "c1": 0, "width": 80}, {"c0": 1, "c1": 1, "width": 230},
                       {"c0": 2, "c1": self.width - 1, "width": 92}],
        }


# ------------------------------------------------------------------ columns --

def _value(label, start, end, seg=None):
    return {"kind": "value", "label": label, "from": start, "to": end, "seg": seg}


def _columns(ledger, spec):
    start, end, compare = spec["start"], spec["end"], spec.get("compare")
    balance_sheet = spec["report"] == "bs"
    split = spec.get("split")
    if split:
        names = ledger.segments(split)
        if not names:
            raise ReportError("NetSuite returned no %s list, so the report cannot be split that way."
                              % KIND_LABEL[split].lower())
        first = start if not balance_sheet else end
        return _with_total([_value(name, first, end, split) for name in names], split)

    chunks = _chunks(ledger, start, end, spec["layout"])
    others = _chunks(ledger, compare[0], compare[1], spec["layout"]) if compare else None
    if others is not None and len(others) != len(chunks):
        chunks, others = _chunks(ledger, start, end, "total"), _chunks(ledger, compare[0], compare[1], "total")
    if balance_sheet:  # a balance sheet column is the position at the end of its chunk
        periods = ledger.periods()
        chunks = [(periods[last]["name"], last, last) for _, _, last in chunks]
        others = others and [(periods[last]["name"], last, last) for _, _, last in others]
    columns = []
    for n, (label, first, last) in enumerate(chunks):
        columns.append(_value(label, first, last))
        if others:
            columns.append(_value(others[n][0], others[n][1], others[n][2]))
            at = len(columns)
            columns.append({"kind": "var", "label": "Change", "a": at - 2, "b": at - 1})
            columns.append({"kind": "pct", "label": "Change %", "a": at, "b": at - 1})
    if not balance_sheet and not others and len(chunks) > 1:
        columns.append({"kind": "sum", "label": "Total", "a": 0, "b": len(chunks) - 1})
    return columns


def _with_total(split_columns, kind):
    """One column per segment value, then what is left over, then the unfiltered total."""
    count = len(split_columns)
    return split_columns + [
        {"kind": "rest", "label": "No %s" % KIND_LABEL[kind].lower(), "total": count + 1, "a": 0, "b": count - 1},
        _value("Total", split_columns[0]["from"], split_columns[0]["to"]),
    ]


def _request(ledger, spec, column, account):
    periods = ledger.periods()
    request = {"account": account, "fromPeriod": periods[column["from"]]["name"],
               "toPeriod": "" if spec["report"] == "bs" else periods[column["to"]]["name"]}
    for kind in FILTER_KINDS:
        if column.get("seg") == kind:
            request[kind] = column["label"]
        elif spec.get("split") != kind:
            request[kind] = (spec.get("filters") or {}).get(kind, "")
    return request


def _probe(ledger, spec, columns, types):
    """Looks the balances up once to decide what the report lists.

    Returns (accounts to list, positions of the value columns that have any balance).
    """
    accounts = [a for a in ledger.accounts() if a["type"] in types]
    values = [j for j, c in enumerate(columns) if c["kind"] == "value"]
    if not accounts or not values:
        return accounts, set(values)
    prune_rows = spec["detail"] and not spec.get("zero")
    full = prune_rows or bool(spec.get("split"))
    probe = accounts if full else accounts[:1]  # one lookup still shows a wrong filter or period
    if len(probe) * len(values) > MAX_PROBE:
        raise ReportError("That report needs %d balances to lay out, more than this builds in one go. "
                          "Ask for fewer periods or columns." % (len(probe) * len(values)))
    # Column by column, so each batch sent to NetSuite covers few periods and runs as few queries.
    results = ledger.balances([_request(ledger, spec, columns[j], a["number"]) for j in values for a in probe])
    errors = [r["error"] for r in results if "error" in r]
    if errors:
        raise ReportError(errors[0])  # a filter or period NetSuite does not know
    if not full:
        return accounts, set(values)
    size = len(probe)
    live = [abs(r.get("value") or 0) >= 0.005 for r in results]
    columns_with_balance = {j for n, j in enumerate(values) if any(live[n * size:(n + 1) * size])}
    if prune_rows:
        accounts = [a for i, a in enumerate(accounts) if any(live[n * size + i] for n in range(len(values)))]
    return accounts, columns_with_balance


# ------------------------------------------------------------------ reports --

def _income_statement(spec, sheet, accounts):
    totals = {}
    for key, label, kinds in PL_SECTIONS:
        members = [a for a in accounts if a["type"] in kinds]
        if not members:
            totals[key] = None
        elif spec["detail"]:
            sheet.heading(label)
            first = len(sheet.rows) + 1
            for a in sorted(members, key=lambda a: a["number"]):
                sheet.account(a["number"], a["name"])
            totals[key] = sheet.total("Total " + label, first, len(sheet.rows))
        else:
            totals[key] = sheet.key("type:" + ",".join(kinds), label)
        if key == "cogs" and totals["cogs"]:
            sheet.combine("Gross Profit", [(1, totals["income"]), (-1, totals["cogs"])])
        if key == "expense":
            totals["operating"] = sheet.combine(
                "Operating Profit", [(1, totals["income"]), (-1, totals["cogs"]), (-1, totals["expense"])])
        if spec["detail"] and (members or key == "expense"):
            sheet.blank()
    net = sheet.combine("Net Income", [(1, totals["operating"]), (1, totals["othincome"]), (-1, totals["othexpense"])],
                        bottom="double")
    if spec["detail"]:
        sheet.blank()
        sheet.value_row("", "Check: NetSuite net income less the line above (should be 0)",
                        lambda j: '=ROUND(%s-%s%d,2)' % (sheet._balance('"calc:NetIncome"', j), col_letter(2 + j), net),
                        derived=False, color=GREY, italic=True)
    return len(accounts)


def _balance_sheet(spec, sheet, accounts):
    section_totals = {}
    for title, total_label, groups in BS_SECTIONS:
        sheet.heading(title)
        equity = title == "EQUITY"
        section_first = len(sheet.rows) + 1
        group_totals = []
        for label, kinds in groups:
            members = sorted((a for a in accounts if a["type"] in kinds),
                             key=lambda a: (kinds.index(a["type"]), a["number"]))
            if not members:
                continue
            if not spec["detail"]:
                group_totals.append(sheet.key("type:" + ",".join(kinds), label))
                continue
            if not equity:
                sheet.heading(label)
            first = len(sheet.rows) + 1
            for a in members:
                sheet.account(a["number"], a["name"])
            if not equity:
                group_totals.append(sheet.total("Total " + label, first, len(sheet.rows)))
        keys = BS_EQUITY_KEYS + ([TRANSLATION_KEY] if spec.get("translation") else [])
        for key, key_label in (keys if equity else []):
            row = sheet.key(key, key_label)
            group_totals.append(row)
            if spec["detail"]:
                sheet.formats.append({"r0": row - 1, "c0": 1, "r1": row - 1, "c1": 1, "indent": 1})
        if spec["detail"] and equity:  # equity accounts and the two earnings rows are one block
            section_totals[title] = sheet.total(total_label, section_first, len(sheet.rows))
        else:
            section_totals[title] = sheet.combine(total_label, [(1, row) for row in group_totals])
        sheet.blank()
    both = sheet.combine("Total Liabilities and Equity",
                         [(1, section_totals["LIABILITIES"]), (1, section_totals["EQUITY"])], bottom="double")
    sheet.blank()
    sheet.value_row("", "Check: assets less liabilities and equity (should be 0)",
                    lambda j: "=ROUND(%s%d-%s%d,2)" % (col_letter(2 + j), section_totals["ASSETS"],
                                                       col_letter(2 + j), both),
                    derived=False, color=GREY, italic=True)
    return len(accounts)


def _trial_balance(ledger, spec):
    """Its own layout: debit and credit columns for one period."""
    periods = ledger.periods()
    at = spec["end"]
    year_start = ledger.fy_start(at)
    filters = spec.get("filters") or {}
    probe_spec = dict(spec, report="tb", split=None, detail=True)
    types = set(CREDIT_TYPES) | {"Bank", "AcctRec", "UnbilledRec", "OthCurrAsset", "DeferExpense", "FixedAsset",
                                 "OthAsset", "COGS", "Expense", "OthExpense"}
    pl_types = {"Income", "COGS", "Expense", "OthIncome", "OthExpense"}
    accounts = [a for a in ledger.accounts() if a["type"] in types]
    if accounts:
        requests = [dict(_request(ledger, probe_spec, _value("", year_start if a["type"] in pl_types else at, at),
                                  a["number"]),
                         toPeriod=periods[at]["name"] if a["type"] in pl_types else "")
                    for a in (accounts if not spec.get("zero") else accounts[:1])]
        results = ledger.balances(requests)
        errors = [r["error"] for r in results if "error" in r]
        if errors:
            raise ReportError(errors[0])
        if not spec.get("zero"):
            accounts = [a for a, r in zip(accounts, results) if abs(r.get("value") or 0) >= 0.005]
    rows = [[TITLES["tb"]], [spec["subtitle"]], [KIND_LABEL["filters"], filters.get("filters", "")]]
    for kind in SEGMENT_KINDS:
        rows.append([KIND_LABEL[kind], filters.get(kind, "")])
    rows += [[], ["As at", periods[at]["name"]], ["P&L from", periods[year_start]["name"]], [],
             ["Account", "Name", "Type", "Debit", "Credit", "Balance"]]
    header = len(rows)
    segs = ",".join("$B$%d" % PARAM_ROW[kind] for kind in FILTER_KINDS)
    entries = [(a["number"], a["name"], a["type"]) for a in sorted(accounts, key=lambda a: a["number"])]
    entries.append(("calc:RetainedEarnings", "Retained Earnings (prior years)", "Equity"))
    if spec.get("translation"):
        entries.append(TRANSLATION_KEY + ("Equity",))
    for number, name, kind in entries:
        row = len(rows) + 1
        periods_args = "$B$10,$B$9" if kind in pl_types else '$B$9,""'
        credit = kind in CREDIT_TYPES
        rows.append([number, name, kind,
                     "=MAX(%sF%d,0)" % ("-" if credit else "", row),
                     "=MAX(%sF%d,0)" % ("" if credit else "-", row),
                     "=TQ.BALANCE($A%d,%s,%s)" % (row, periods_args, segs)])
    last = len(rows)
    rows.append(["", "Total", "", "=SUM(D%d:D%d)" % (header + 1, last), "=SUM(E%d:E%d)" % (header + 1, last), ""])
    rows.append(["", "Check: debits less credits (should be 0)", "", "=ROUND(D%d-E%d,2)" % (last + 1, last + 1), "", ""])
    rows = [r + [""] * (6 - len(r)) for r in rows]
    formats = [
        {"r0": 0, "c0": 0, "r1": 0, "c1": 5, "bold": True, "size": 14},
        {"r0": 1, "c0": 0, "r1": 1, "c1": 5, "color": GREY},
        {"r0": header - 1, "c0": 0, "r1": header - 1, "c1": 5, "bold": True, "bottom": "thin"},
        {"r0": header - 1, "c0": 3, "r1": header - 1, "c1": 5, "align": "right"},
        {"r0": header, "c0": 3, "r1": last + 1, "c1": 5, "numberFormat": NUMBER_FORMAT},
        {"r0": header, "c0": 5, "r1": last - 1, "c1": 5, "color": GREY},
        {"r0": last, "c0": 0, "r1": last, "c1": 5, "bold": True, "top": "thin", "bottom": "double"},
        {"r0": last + 1, "c0": 0, "r1": last + 1, "c1": 5, "color": GREY, "italic": True},
    ]
    action = {
        "type": "write", "target": "new_sheet", "sheetName": "", "startCell": "", "rows": rows,
        "textRanges": [[2, 1, 9, 1], [header, 0, last - 1, 0]], "formats": formats,
        "freeze": {"rows": header, "cols": 2},
        "widths": [{"c0": 0, "c1": 0, "width": 80}, {"c0": 1, "c1": 1, "width": 230},
                   {"c0": 2, "c1": 2, "width": 90}, {"c0": 3, "c1": 5, "width": 92}],
    }
    return action, len(accounts)


def split_columns(spec, columns, with_balance):
    """For a report split by segment: keeps only the segment columns that have a balance."""
    if not spec.get("split"):
        return columns
    label = KIND_LABEL[spec["split"]].lower()
    kept = [c for j, c in enumerate(columns) if c.get("seg") and j in with_balance]
    if not kept:
        raise ReportError("No %s has a balance for these periods." % label)
    if len(kept) > MAX_SPLIT_COLUMNS:
        raise ReportError("%d %s values have a balance, too many for one column each (limit %d). "
                          "Filter by one instead." % (len(kept), label, MAX_SPLIT_COLUMNS))
    return _with_total(kept, spec["split"])


def is_single(spec, columns):
    """True when every column is one period, so the header needs no From/To rows."""
    return (not spec.get("split")) and all(c["from"] == c["to"] for c in columns if c["kind"] == "value")


def sheet_name(text):
    return re.sub(r"\s+", " ", re.sub(r"[\[\]:*?/\\']", " ", text)).strip()[:31]


def describe(ledger, spec):
    """One line saying what the report shows, used as the subtitle and in the chat reply."""
    periods = ledger.periods()
    if spec["report"] == "pl":
        parts = [range_label(ledger, spec["start"], spec["end"])]
        if spec.get("split"):
            parts.append("by " + KIND_LABEL[spec["split"]].lower())
        elif spec["layout"] != "total" and spec["start"] != spec["end"]:
            parts.append("by " + spec["layout"])
    else:
        parts = ["as at " + periods[spec["end"]]["name"]]
        if spec["report"] == "bs" and spec.get("split"):
            parts.append("by " + KIND_LABEL[spec["split"]].lower())
        elif spec["report"] == "bs" and spec["layout"] != "total" and spec["start"] != spec["end"]:
            parts = ["%s ends, %s" % (spec["layout"], range_label(ledger, spec["start"], spec["end"]))]
    if spec.get("compare") and not spec.get("split") and spec["report"] != "tb":
        other = spec["compare"]
        parts.append("compared with " + (periods[other[1]]["name"] if spec["report"] == "bs"
                                         else range_label(ledger, other[0], other[1])))
    if not spec.get("detail") and spec["report"] != "tb":
        parts.append("summary")
    return ", ".join(parts)


def build(ledger, spec):
    """Returns (write action, number of accounts listed). Raises ReportError."""
    try:
        spec = dict(spec, subtitle=describe(ledger, spec))
        if spec["report"] in ("bs", "tb") and spec.get("split") != "subsidiary":
            spec["translation"] = ledger.setup()["multiCurrency"]
        periods = ledger.periods()
        if spec["report"] == "tb":
            action, count = _trial_balance(ledger, spec)
            action["sheetName"] = sheet_name("TB " + periods[spec["end"]]["name"])
            add_filter_lists(action, ledger, dropdown_fields(ledger, spec.get("dropdowns") or []),
                             (spec.get("filters") or {}).get("filters", ""))
            return action, count
        columns = _columns(ledger, spec)
        if spec["report"] == "pl":
            types = {t for _, _, kinds in PL_SECTIONS for t in kinds}
        else:
            types = {t for _, _, groups in BS_SECTIONS for _, kinds in groups for t in kinds}
        accounts, with_balance = _probe(ledger, spec, columns, types)
        columns = split_columns(spec, columns, with_balance)
        single = is_single(spec, columns)
        sheet = _Sheet(ledger, spec, columns, single)
        if spec["report"] == "pl":
            count = _income_statement(spec, sheet, accounts)
            name = "P&L " + range_label(ledger, spec["start"], spec["end"])
        else:
            count = _balance_sheet(spec, sheet, accounts)
            name = "BS " + periods[spec["end"]]["name"]
        if len(sheet.rows) * sheet.width > MAX_CELLS:
            raise ReportError("That report would be %d cells, more than this builds in one go. "
                              "Ask for fewer periods or a summary." % (len(sheet.rows) * sheet.width))
        action = sheet.action(sheet_name(name))
        add_filter_lists(action, ledger, dropdown_fields(ledger, spec.get("dropdowns") or []),
                         (spec.get("filters") or {}).get("filters", ""))
        return action, count
    except LedgerError as err:
        raise ReportError(str(err))
