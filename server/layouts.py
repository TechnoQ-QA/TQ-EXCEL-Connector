"""TQ Excel Connector - report designs (custom layouts).

A design is an ordinary worksheet whose name starts with "TQ Layout". It holds
two tables side by side: the report's lines (rows) and its columns. Both are
free-form, so each customer's report can have its own shape:

    Section        Line                  Type      Accounts or formula          Sign  Format
    Revenue        Sales                 accounts  4000, 4010-4099                1
    Revenue        Net Sales             subtotal
    Cost of Sales  Food Cost             accounts  5000-5099                     -1
    Cost of Sales  Total Cost of Sales   subtotal
    Gross Profit   Gross Profit          formula   =[Net Sales]+[Total Cost of Sales]   bold
    Gross Profit   Gross Margin          formula   =[Gross Profit]/[Net Sales]          percent
    KPI            Average Sales Daily   formula   =[Net Sales]/days

    Column     Column type  From      To        Department  Column formula            Column format
    Jan 2026   period       Jan 2026  Jan 2026
    YTD        period       Jan 2026  Aug 2026
    Last year  period       Jan 2025  Aug 2025
    Food YTD   period       Jan 2026  Aug 2026  101 Food
    Change     formula                                      =[YTD]-[Last year]
    Change %   formula                                      =[Change]/ABS([Last year])  percent
    Total      total        Jan 2026  Dec 2026              =SUM([Jan 2026]:[Dec 2026])

Lines
    accounts  account numbers, ranges and type:X added together (see ledger.py);
              Sign -1 shows the amount as a negative
    subtotal  adds the accounts lines above it in the same section
    formula   any Excel formula; [Line name] is that line in the same column and
              days is the number of days in the column's periods
Columns
    period    balances for From..To, optionally for one subsidiary/department/class/location;
              Data = budget shows the NetSuite budget instead (budget(Plan A) for one category)
    formula   an Excel formula over other columns ([Column name]), applied to every line
    total     like formula, but ratio lines are worked out again on the totals

The panel sends the sheet's cells with each request. parse() reads them into a
design, the edit functions change it, emit() writes the sheet back and
build_report() writes the report. Both are written in place, on the same sheets.
"""

import re

import ledger as ledger_module
import reports
from ledger import FILTER_KINDS, LedgerError, PL_SIGNS, SEGMENT_KINDS
from reports import GREY, KIND_LABEL, PARAM_ROW, ReportError, YEAR_RE, col_letter

SHEET_PREFIX = "TQ Layout"
LINE_HEADERS = ["Section", "Line", "Type", "Accounts or formula", "Sign", "Format"]
COLUMN_HEADERS = ["Column", "Column type", "From", "To", "Subsidiary", "Department", "Class", "Location",
                  "Other filters", "Column formula", "Column format", "Data"]
COLUMN_AT = len(LINE_HEADERS) + 1   # the columns table starts one blank column after the lines table
LAYOUT_WIDTH = COLUMN_AT + len(COLUMN_HEADERS)
MAX_LINES = 400
MAX_COLUMNS = 80
MAX_LIST = 1000          # longest segment list offered as a dropdown
MAX_PAGES = 40           # most extra sheets for "one sheet per location"
MAX_SAMPLE_LINES = 40    # accounts listed one by one per section in the sample layout
REF_RE = re.compile(r"\[([^\[\]]+)\]")
DAYS_RE = re.compile(r"\bdays\b", re.I)
NAVY, MAROON, PALE, WHITE = "#1F3A5F", "#6B1F24", "#E8EDF3", "#FFFFFF"
LINE_FORMATS = {"percent": "0.00%", "integer": "#,##0;(#,##0)"}
LINE_KINDS = {"accounts": "accounts", "account": "accounts", "subtotal": "subtotal", "total": "subtotal",
              "sum": "subtotal", "formula": "formula", "calc": "formula", "calculation": "formula"}
COLUMN_TYPES = ("period", "formula", "total")
NOTES = [
    "Lines (left): one row per report line, grouped under its Section. Columns (right): one row per report column.",
    "accounts line: account numbers separated by commas, ranges such as 4000-4999, or type:Income. "
    "Sign -1 shows the amount as a negative.",
    "subtotal line: adds the accounts lines above it in its section. formula line: starts with = and uses "
    "[Line name] and days (days in the period).",
    "period column: From and To period, and optionally one department, class, location or subsidiary. "
    "formula column: = with [Column name]. total column: a formula column where ratios are worked out again.",
    "Format of a line: percent, integer, bold, key (key lines appear on the summary sheet of a one-sheet-per report). "
    "Data of a column: budget.",
    "Edit here, in the Reports tab of the panel, or in the chat (add column Q1 2025). Then say: update report.",
]


def _cell(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))  # an account number typed as a number arrives as 4000.0
    return str(value).strip()


def _format(text):
    words = set(re.split(r"[,\s]+", text.lower()))
    return ("percent" if words & {"percent", "%", "pct", "percentage"} else
            "integer" if words & {"integer", "whole", "0"} else "number"), "bold" in words


def _formula(text):
    return "=" + text.lstrip("=").strip()


# ------------------------------------------------------------------ reading --

def parse(name, rows, first_row=0):
    """The cells of a layout sheet -> design. Raises ReportError naming the sheet row."""
    grid = [[_cell(c) for c in row] if isinstance(row, list) else [] for row in rows]
    header = next((n for n, row in enumerate(grid)
                   if "line" in [c.lower() for c in row]
                   and any(c.lower().startswith(("accounts", "formula", "source")) for c in row)), None)
    if header is None:
        raise ReportError('Sheet "%s" has no header row. It needs the columns: %s.' % (name, ", ".join(LINE_HEADERS)))
    line_cols, column_cols = {}, {}
    for at, cell in enumerate(c.lower() for c in grid[header]):
        if cell in ("column", "column type", "from", "to", "column formula", "column format", "data",
                    "other filters") or cell in SEGMENT_KINDS:
            key = "label" if cell == "column" else "filters" if cell == "other filters" else cell.replace("column ", "")
            column_cols.setdefault(key, at)
        else:
            key = "source" if cell.startswith(("accounts", "formula", "source")) else cell
            if key in ("section", "line", "type", "source", "sign", "format"):
                line_cols.setdefault(key, at)
    design = {"name": name, "title": "", "filters": {}, "pages": "", "detail": False, "lines": [], "columns": [],
              "dropdowns": [], "month": "", "highlight": 0.0, "comments": False, "rules": [], "commentAt": 0.0,
              "check": True}
    for row in grid[:header]:
        key = row[0].lower() if row else ""
        key = "filters" if key in ("other filters", "other filter", "filters") else key.replace(" filter", "")
        value = row[1] if len(row) > 1 else ""
        if key in ("report title", "title") and value:
            design["title"] = value
        elif key == "one sheet per":
            design["pages"] = page_kind(value)
        elif key in ("dropdown filters", "dropdowns"):
            design["dropdowns"] = [v.strip() for v in value.split(",") if v.strip()][:reports.MAX_DROPDOWNS]
        elif key == "show accounts":
            design["detail"] = value.lower() in ("yes", "y", "true", "1", "on")
        elif key in ("report month", "month"):
            design["month"] = value
        elif key in ("highlight changes over", "highlight"):
            design["highlight"] = highlight_value(value)
        elif key in ("comments column", "comments"):
            design["comments"] = value.lower() in ("yes", "y", "true", "1", "on")
        elif key in ("check line", "check"):
            design["check"] = value.lower() not in ("no", "n", "false", "0", "off")
        elif key in ("comment needed over", "ask for a comment over"):
            design["commentAt"] = highlight_value(value)
        elif key in ("highlight rules", "rules"):
            design["rules"] = [parse_rule(t)["text"] for t in value.split(";") if t.strip()]
        elif key in FILTER_KINDS and value:
            design["filters"][key] = value
    design["title"] = design["title"] or name
    section = ""
    for n in range(header + 1, len(grid)):
        row = grid[n]
        where = 'Sheet "%s", row %d' % (name, first_row + n + 1)
        get = lambda key: row[line_cols[key]] if key in line_cols and line_cols[key] < len(row) else ""  # noqa: E731
        if any(get(key) for key in line_cols):
            section = get("section") or section
            design["lines"].append(new_line(get("line"), get("type"), get("source"), get("sign"), get("format"),
                                            section, where))
        get = lambda key: row[column_cols[key]] if key in column_cols and column_cols[key] < len(row) else ""  # noqa: E731
        if any(get(key) for key in column_cols):
            design["columns"].append(new_column(get("label"), get("type"), get("from"), get("to"),
                                                {k: get(k) for k in FILTER_KINDS if get(k)},
                                                get("formula"), get("format"), where, get("data")))
    if len(design["lines"]) > MAX_LINES or len(design["columns"]) > MAX_COLUMNS:
        raise ReportError('Sheet "%s" is over the limit of %d lines and %d columns.' % (name, MAX_LINES, MAX_COLUMNS))
    return design


def highlight_value(value):
    """'10%', '10', '0.1' -> 10.0 (percent); blank or 'no' -> 0 (no highlights)."""
    text = _cell(value).replace("%", "").strip().lower()
    if not text or text in ("no", "off", "none"):
        return 0.0
    try:
        number = float(text)
    except ValueError:
        raise ReportError('Highlight changes over "%s": give a percentage such as 10%%.' % _cell(value))
    number = number * 100 if 0 < number < 1 else number
    if not 0 < number <= 1000:
        raise ReportError('Highlight changes over "%s": give a percentage between 1%% and 1000%%.' % _cell(value))
    return round(number, 2)


COLOURS = {   # font, fill
    "red": ("#B3261E", "#FCE8E6"), "yellow": ("#7A5C00", "#FFF4C2"), "amber": ("#8A4B00", "#FFE0B2"),
    "orange": ("#8A4B00", "#FFE0B2"), "green": ("#1A7F4B", "#E6F4EA"), "blue": ("#0B5CAD", "#E3EEFB"),
    "grey": ("#444444", "#EEEEEE"), "gray": ("#444444", "#EEEEEE"),
}
OPS = {">": ">", "<": "<", ">=": ">=", "<=": "<=", "=": "=", "<>": "<>", "over": ">", "above": ">",
       "under": "<", "below": "<", "between": "between"}
RULE_RE = re.compile(r"^\s*(?P<target>.+?)\s+(?:in\s+(?P<col>.+?)\s+)?(?P<op>>=|<=|<>|>|<|=|between|over|above|under|below)"
                     r"\s*(?P<v1>-?[\d,]*\.?\d+\s*%%?)(?:\s+and\s+(?P<v2>-?[\d,]*\.?\d+\s*%%?))?\s+(?P<colour>%s)\s*$"
                     % "|".join(COLOURS), re.I)


def parse_rule(text):
    """'6100 > 5000 red', 'Net Profit < 0 red', 'all between -100 and 100 yellow', '4000-4999 in Month < 1000 amber'
    -> {target, column, op, v1, v2, percent, colour, text}. Raises ReportError."""
    m = RULE_RE.match(_cell(text))
    if not m:
        raise ReportError('Highlight rule "%s" is not understood. Write it as: what, (in column), condition, value, '
                          'colour, e.g. 6100 > 5000 red, Net Profit < 0 red, all between -100 and 100 yellow.' % _cell(text))
    op = OPS[m.group("op").lower()]
    number = lambda v: float(v.replace(",", "").replace("%", "").strip()) / (100 if "%" in v else 1)  # noqa: E731
    v1 = number(m.group("v1"))
    v2 = number(m.group("v2")) if m.group("v2") else None
    if op == "between" and v2 is None:
        raise ReportError('Highlight rule "%s": between needs two values, e.g. between -100 and 100.' % _cell(text))
    return {"target": m.group("target").strip(), "column": (m.group("col") or "").strip(), "op": op, "v1": v1, "v2": v2,
            "percent": "%" in m.group("v1"), "colour": m.group("colour").lower(), "text": " ".join(_cell(text).split())}


def page_kind(value):
    """'Location', 'locations' -> 'location'; anything else -> '' (no extra sheets)."""
    word = _cell(value).lower()
    return next((kind for kind in SEGMENT_KINDS if word and (word == kind or word == KIND_LABEL[kind].lower()
                                                             or word.rstrip("s") == kind or word == kind[:-1] + "ies"
                                                             or word == kind + "es")), "")


def new_line(label, kind, source, sign="", fmt="", section="", where=None):
    label, source, kind = _cell(label), _cell(source), _cell(kind).lower()
    where = where or 'Line "%s"' % label
    if not label:
        raise ReportError("%s: the Line name is missing." % where)
    if kind and kind not in LINE_KINDS:
        raise ReportError('%s: Type "%s" is not one of accounts, subtotal, formula.' % (where, kind))
    kind = LINE_KINDS.get(kind) or ("formula" if source.startswith("=") else "accounts" if source else "")
    if not kind or (kind != "subtotal" and not source.lstrip("=").strip()):
        raise ReportError('%s: "%s" needs accounts or a formula.' % (where, label))
    number_format, bold = _format(_cell(fmt))
    line_filter = ""
    if kind == "accounts" and "|" in source:   # "type:Income | Department: Rooms*": the line's own filter
        source, _, line_filter = source.partition("|")
        source, line_filter = source.strip(), re.sub(r"\s+", " ", line_filter).strip()
        if line_filter and ":" not in line_filter:
            raise ReportError('%s: after | comes a filter such as Department: Rooms (a field, a colon, the values).' % where)
    return {"section": _cell(section), "label": label, "kind": kind, "where": where, "filter": line_filter,
            "key": "key" in re.split(r"[,\s]+", _cell(fmt).lower()),
            "base": "base" in re.split(r"[,\s]+", _cell(fmt).lower()),   # [Column@base] in the lines below it
            "source": _formula(source) if kind == "formula" else "" if kind == "subtotal" else source,
            "sign": -1 if _cell(sign).replace(" ", "").lower() in ("-1", "-", "minus", "negative", "neg") else 1,
            "format": number_format, "bold": bold}


def new_column(label, kind, start, end="", filters=None, formula="", fmt="", where=None, data=""):
    label, kind, start, end, formula = _cell(label), _cell(kind).lower(), _cell(start), _cell(end), _cell(formula)
    where = where or 'Column "%s"' % label
    if not label:
        raise ReportError("%s: the Column name is missing." % where)
    if kind and kind not in COLUMN_TYPES:
        raise ReportError('%s: Column type "%s" is not one of period, formula, total.' % (where, kind))
    kind = kind or ("period" if not formula else "total" if start else "formula")
    if kind == "period" and not start:
        raise ReportError('%s: column "%s" needs a From period.' % (where, label))
    if kind != "period" and not formula.lstrip("=").strip():
        raise ReportError('%s: column "%s" needs a formula.' % (where, label))
    data = re.sub(r"\s+", " ", _cell(data)).strip()
    if kind != "period" or data.lower() in ("actual", "actuals"):
        data = ""
    if data and not re.fullmatch(r"budget(\([^()\"]*\))?", data, re.I):
        raise ReportError('%s: Data "%s" is not known. Leave it blank for actuals, or use budget or '
                          'budget(Category name).' % (where, data))
    return {"label": label, "type": kind, "from": start, "to": end or start, "where": where, "data": data,
            "filters": {k: _cell(v) for k, v in (filters or {}).items() if k in FILTER_KINDS and _cell(v)}
            if kind == "period" else {},
            "formula": _formula(formula) if kind != "period" else "", "format": _format(_cell(fmt))[0]}


def public(design):
    """The design as the panel's Designer tab shows it."""
    return {
        "name": design["name"], "title": design["title"], "filters": design["filters"],
        "pages": design.get("pages", ""), "detail": bool(design.get("detail")),
        "dropdowns": list(design.get("dropdowns") or []), "month": design.get("month", ""),
        "highlight": design.get("highlight") or 0, "comments": bool(design.get("comments")),
        "rules": list(design.get("rules") or []), "check": design.get("check", True) is not False,
        "lines": [dict({k: line.get(k, False) for k in ("section", "label", "kind", "source", "sign", "format", "bold", "key", "base")},
                       source=line["source"] + (" | " + line["filter"] if line.get("filter") else ""))
                  for line in design["lines"]],
        "commentAt": design.get("commentAt") or 0,
        "columns": [{k: c[k] for k in ("label", "type", "from", "to", "filters", "formula", "format", "data")}
                    for c in design["columns"]],
    }


def check(ledger, design):
    """Resolves accounts, periods and references. Raises ReportError saying where the mistake is."""
    lines, columns = design["lines"], design["columns"]
    by_label = {}
    for line in lines:
        by_label.setdefault(line["label"].lower(), []).append(line)
    open_lines = {}  # section -> accounts lines since the last subtotal
    for line in lines:
        if line["kind"] == "accounts":
            try:
                prefix, _, token = ledger.split_prefix(line["source"])
                calc = token.lower().startswith("calc:")
                line["special"] = bool(prefix) or calc   # a movement, a budget or a calculated total
                if calc:
                    ledger._expand({"account": line["source"], "fromPeriod": ledger.periods()[0]["name"]})
                line["accounts"] = [] if calc else ledger.members(token)
            except LedgerError as err:
                raise ReportError("%s: %s." % (line["where"], err))
            open_lines.setdefault(line["section"].lower(), []).append(line)
        elif line["kind"] == "subtotal":
            line["adds"] = open_lines.pop(line["section"].lower(), [])
            if not line["adds"]:
                raise ReportError('%s: subtotal "%s" has no accounts lines above it in its section.'
                                  % (line["where"], line["label"]))
        else:
            for ref in REF_RE.findall(line["source"]):
                hits = by_label.get(ref.strip().lower(), [])
                if len(hits) != 1:
                    raise ReportError('%s: the formula uses [%s], %s.' % (
                        line["where"], ref, "but %d lines have that name" % len(hits) if hits
                        else "which is not a line in the layout"))
                if hits[0] is line:
                    raise ReportError("%s: the formula refers to its own line." % line["where"])
    labels = [c["label"].lower() for c in columns]
    periods = ledger.periods()
    month = None   # the report's month (B8), when a column is relative to it: CM, LY, FYS...
    if design.get("month"):
        try:
            month = ledger.period_index(design["month"])
            design["month"] = periods[month]["name"]
        except LedgerError:
            raise ReportError('Report month "%s" is not a NetSuite period. Use a name such as %s.'
                              % (design["month"], periods[ledger.current_index()]["name"]))
    for at, column in enumerate(columns):
        if labels.count(column["label"].lower()) > 1:
            raise ReportError('%s: two columns are called "%s". Column names must differ.'
                              % (column["where"], column["label"]))
        try:
            column["relative"] = False
            if column["from"]:
                if is_relative(ledger, column["from"]) or is_relative(ledger, column["to"]):
                    if month is None:   # the first relative column gives the report a month: this one
                        month = ledger.current_index()
                        design["month"] = periods[month]["name"]
                    column["relative"] = True
                    column["from"], column["to"] = column["from"].upper(), column["to"].upper()
                    try:
                        first = ledger.resolve_period(column["from"], month)
                        last = ledger.resolve_period(column["to"], month)
                    except LedgerError:
                        # e.g. last year's month before the first period in NetSuite: the column shows nothing for
                        # this month and fills in when B8 moves to a month that has one
                        column["at"] = None
                        continue
                else:
                    first, last = ledger.period_index(column["from"]), ledger.period_index(column["to"])
                    column["from"], column["to"] = periods[min(first, last)]["name"], periods[max(first, last)]["name"]
                column["at"] = (min(first, last), max(first, last))
        except LedgerError as err:
            raise ReportError("%s: %s." % (column["where"], err))
        for kind, value in column["filters"].items():
            if kind not in SEGMENT_KINDS or ledger_module.NEW_FILTER_RE.search(value):
                continue   # several values, "not", "Parent*" and other filters are checked by NetSuite
            known = {s.lower(): s for s in _segments(ledger, kind)}
            if known and value.lower() not in known:
                raise ReportError('%s: %s "%s" was not found in NetSuite.' % (column["where"], kind, value))
            column["filters"][kind] = known.get(value.lower(), value)
        for ref in REF_RE.findall(column["formula"]):
            ref, _, at_line = ref.partition("@")
            if at_line.strip() and at_line.strip().lower() != "base" and at_line.strip().lower() not in by_label:
                raise ReportError('%s: the formula uses [%s@%s], but there is no line "%s".'
                                  % (column["where"], ref, at_line, at_line.strip()))
            if ref.strip().lower() not in labels:
                raise ReportError('%s: the formula uses [%s], which is not a column in the layout.'
                                  % (column["where"], ref))
            if labels.index(ref.strip().lower()) == at:
                raise ReportError("%s: the formula refers to its own column." % column["where"])


def is_relative(ledger, value):
    """True for CM, CM-1, LY, FYS, FYS-12... (a period worked out from the report's month), not a period name."""
    value = _cell(value)
    if not value or not ledger_module.PERIOD_EXPR_RE.fullmatch(value.replace(" ", "")):
        return False
    try:
        ledger.period_index(value)
        return False   # a period really called that
    except LedgerError:
        return bool(re.match(r"\s*[A-Za-z]", value))


def _segments(ledger, kind):
    try:
        return ledger.segments(kind)
    except LedgerError:
        return []


# ------------------------------------------------------------------ editing --

def _find(items, label, what):
    wanted = _cell(label).strip('"\' ').lower()
    hits = [n for n, item in enumerate(items) if item["label"].lower() == wanted]
    if not hits:
        hits = [n for n, item in enumerate(items) if wanted and wanted in item["label"].lower()]
    if len(hits) != 1:
        names = ", ".join(item["label"] for item in items[:30]) or "none yet"
        raise ReportError('%s %s "%s". The layout has: %s.' % (
            "There is no" if not hits else "More than one", what, _cell(label), names))
    return hits[0]


def add_line(design, line):
    lines = design["lines"]
    if any(l["label"].lower() == line["label"].lower() for l in lines):
        raise ReportError('The layout already has a line called "%s".' % line["label"])
    if not line["section"] and lines:
        line["section"] = lines[-1]["section"]
    same = [n for n, l in enumerate(lines) if l["section"].lower() == line["section"].lower()]
    if same:
        line["section"] = lines[same[-1]]["section"]
    lines.insert(same[-1] + 1 if same else len(lines), line)


def add_column(design, column):
    if any(c["label"].lower() == column["label"].lower() for c in design["columns"]):
        raise ReportError('The layout already has a column called "%s".' % column["label"])
    design["columns"].append(column)


def remove(design, what, label):
    items = design["lines" if what == "line" else "columns"]
    return items.pop(_find(items, label, what))


def move(design, what, label, by):
    items = design["lines" if what == "line" else "columns"]
    at = _find(items, label, what)
    to = max(0, min(len(items) - 1, at + by))
    items.insert(to, items.pop(at))
    if what == "line":  # a line moved past the edge of its section joins the neighbouring section
        neighbour = items[to - 1] if by > 0 and to > 0 else items[to + 1] if by < 0 and to + 1 < len(items) else None
        if neighbour and to != at:
            items[to]["section"] = neighbour["section"]


def columns_from_spec(ledger, spec):
    """The columns a chat request describes ("2026 by month", "vs last year", "by department")."""
    spec = dict(spec, report="pl", detail=False)
    columns = reports._columns(ledger, spec)
    if spec.get("split"):
        _, with_balance = reports._probe(ledger, spec, columns, set(PL_SIGNS))
        columns = reports.split_columns(spec, columns, with_balance)
    periods = ledger.periods()
    names, out = [], []
    for column in columns:
        label, n = column["label"], 2
        if column["kind"] == "var":
            label = "Change" if sum(c["kind"] == "var" for c in columns) == 1 else "Change " + names[column["a"]]
        elif column["kind"] == "pct":
            label = "Change %" if sum(c["kind"] == "pct" for c in columns) == 1 else "Change % " + names[columns[column["a"]]["a"]]
        while label.lower() in [name.lower() for name in names]:
            label, n = "%s (%d)" % (column["label"], n), n + 1
        names.append(label)
    for column, label in zip(columns, names):
        ref = lambda n: "[%s]" % names[n]  # noqa: E731
        if column["kind"] == "value":
            out.append(new_column(label, "period", periods[column["from"]]["name"], periods[column["to"]]["name"],
                                  {column["seg"]: column["label"]} if column.get("seg") else None))
        elif column["kind"] == "var":
            out.append(new_column(label, "formula", "", "", None, "=%s-%s" % (ref(column["a"]), ref(column["b"]))))
        elif column["kind"] == "pct":
            out.append(new_column(label, "formula", "", "", None,
                                  "=%s/ABS(%s)" % (ref(column["a"]), ref(column["b"])), "percent"))
        else:
            values = [c for c in columns if c["kind"] == "value"]
            start, end = periods[values[0]["from"]]["name"], periods[values[-1]["to"]]["name"]
            formula = "=SUM(%s:%s)" % (ref(column["a"]), ref(column["b"]))
            if column["kind"] == "rest":
                formula = "=%s-SUM(%s:%s)" % (ref(column["total"]), ref(column["a"]), ref(column["b"]))
            out.append(new_column(label, "total", start, end, None, formula))
    return out


def with_budget(columns):
    """Each period column followed by its budget, the variance and the variance %."""
    out = []
    for column in columns:
        if column["type"] != "period" or column.get("data"):
            continue  # totals of the plain columns would no longer add the right cells
        label = column["label"]
        budget = new_column("Budget " + label, "period", column["from"], column["to"], column["filters"], data="budget")
        out += [column, budget,
                new_column("Var " + label, "formula", "", "", None, "=[%s]-[%s]" % (label, budget["label"])),
                new_column("Var % " + label, "formula", "", "", None,
                           "=[Var %s]/ABS([%s])" % (label, budget["label"]), "percent")]
    return out


def default_columns(ledger):
    current = ledger.current_index()
    return columns_from_spec(ledger, {"start": ledger.fy_start(current), "end": current, "layout": "month",
                                      "split": None, "compare": None, "filters": {}})


# ------------------------------------------------------------------ writing --

def emit(design, activate=True):
    """The write action that puts the design back on its layout sheet."""
    rows = [[SHEET_PREFIX], ["Report title", design["title"]]]
    rows += [[KIND_LABEL[kind] + " filter", design["filters"].get(kind, "")] for kind in SEGMENT_KINDS]
    rows += [["Other filters", design["filters"].get("filters", ""), "",
              "Optional: any other filter, e.g. Brand: Premium; Name: Al Noor."]]
    rows += [["Dropdown filters", ", ".join(design.get("dropdowns") or []), "",
              "Optional: fields with a dropdown on the report, e.g. Brand, Name (up to %d)." % reports.MAX_DROPDOWNS]]
    rows += [["One sheet per", design.get("pages", ""), "",
              "Optional: location, department, class or subsidiary. Adds one report sheet for each, after the summary."]]
    rows += [["Show accounts", "yes" if design.get("detail") else "no", "",
              "yes lists, under each line, the accounts behind it that have a balance."]]
    rows += [["Report month", design.get("month", ""), "",
              "Optional: the month the report is for. Columns whose From/To is CM, CM-1, LY, FYS, FYE... follow it. "
              "Change it in B8 of the report to roll forward."]]
    rows += [["Highlight changes over", "%g%%" % design["highlight"] if design.get("highlight") else "", "",
              "Optional: e.g. 10%. Change % and Var % cells at or above it turn green, at or below minus it red."]]
    rows += [["Comments column", "yes" if design.get("comments") else "no", "",
              "yes adds a Comments column at the right of the report; what you type there is kept when it is rebuilt."]]
    rows += [["Comment needed over", "%g%%" % design["commentAt"] if design.get("commentAt") else "", "",
              "Optional, with a Comments column: e.g. 10%. A line whose change % is at least this (up or down) and has "
              "no comment yet is marked yellow, so variances get explained."]]
    rows += [["Highlight rules", "; ".join(design.get("rules") or []), "",
              "Optional, separated by ;  e.g. 6100 > 5000 red; Net Profit < 0 red; all between -100 and 100 yellow; "
              "4000-4999 in Month < 10000 amber. Colours: red, yellow, amber, green, blue, grey."]]
    rows += [["Check line", "yes" if design.get("check", True) else "no", "",
              "yes adds a last line with the accounts that no line includes (should be 0; red when it is not)."]]
    rows += [[note] for note in NOTES] + [[], LINE_HEADERS + [""] + COLUMN_HEADERS]
    header = len(rows) - 1
    count = max(len(design["lines"]), len(design["columns"]), 1)
    for n in range(count):
        row = [""] * LAYOUT_WIDTH
        if n < len(design["lines"]):
            line = design["lines"][n]
            words = (([line["format"]] if line["format"] != "number" else []) + (["bold"] if line["bold"] else [])
                     + (["key"] if line.get("key") else []) + (["base"] if line.get("base") else []))
            row[:len(LINE_HEADERS)] = [line["section"], line["label"], line["kind"],
                                       line["source"] + (" | " + line["filter"] if line.get("filter") else ""),
                                       str(line["sign"]) if line["kind"] == "accounts" else "", " ".join(words)]
        if n < len(design["columns"]):
            c = design["columns"][n]
            row[COLUMN_AT:] = ([c["label"], c["type"], c["from"], c["to"] if c["from"] else ""]
                               + [c["filters"].get(kind, "") for kind in FILTER_KINDS]
                               + [c["formula"], c["format"] if c["format"] != "number" else "", c.get("data", "")])
        rows.append(row)
    rows = [row + [""] * (LAYOUT_WIDTH - len(row)) for row in rows]
    last = len(rows) - 1
    wide = LAYOUT_WIDTH - 1
    return {
        "type": "write", "target": "sheet", "sheetName": design["name"], "startCell": "", "rows": rows,
        "activate": activate,
        # Everything is text, so a formula stays as typed instead of being worked out on this sheet.
        "textRanges": [[0, 0, last, wide]],
        "formats": [
            {"r0": 0, "c0": 0, "r1": 0, "c1": wide, "bold": True, "size": 14},
            {"r0": 1, "c0": 1, "r1": 15, "c1": 1, "bold": True},
            {"r0": 6, "c0": 3, "r1": 15, "c1": 3, "color": GREY, "italic": True},
            {"r0": 16, "c0": 0, "r1": header - 2, "c1": wide, "color": GREY, "italic": True},
            {"r0": header, "c0": 0, "r1": header, "c1": len(LINE_HEADERS) - 1, "bold": True, "fill": MAROON,
             "color": WHITE},
            {"r0": header, "c0": COLUMN_AT, "r1": header, "c1": wide, "bold": True, "fill": NAVY, "color": WHITE},
        ],
        "freeze": {"rows": header + 1, "cols": 0},
        "widths": [{"c0": 0, "c1": 0, "width": 130}, {"c0": 1, "c1": 1, "width": 210}, {"c0": 2, "c1": 2, "width": 65},
                   {"c0": 3, "c1": 3, "width": 230}, {"c0": 4, "c1": 5, "width": 55}, {"c0": 6, "c1": 6, "width": 25},
                   {"c0": COLUMN_AT, "c1": COLUMN_AT, "width": 120}, {"c0": COLUMN_AT + 1, "c1": COLUMN_AT + 7, "width": 80},
                   {"c0": COLUMN_AT + 8, "c1": COLUMN_AT + 8, "width": 130},
                   {"c0": COLUMN_AT + 9, "c1": COLUMN_AT + 9, "width": 220},
                   {"c0": COLUMN_AT + 10, "c1": wide, "width": 90}],
    }


CHECK_LABEL = "Check: accounts not in this report (should be 0)"
PL_TYPES_SPEC = "type:Income,OthIncome,COGS,Expense,OthExpense"


def shown(expr):
    """A worked-out cell: a division by zero or a blank in the sum shows as blank, but a figure NetSuite could not
    give (#N/A from TQ.BALANCE: a timeout, no seat, the connector not running) stays #N/A, so a total, a variance
    or the check line never looks complete when it is not."""
    return 'IFERROR(%s,IF(ISNA(%s),NA(),""))' % (expr, expr)


def check_spec(ledger, design):
    """(every account of the report's kind, the accounts its lines take) as account lists for TQ.BALANCE, or None.

    A profit and loss layout is checked against every P&L account, a balance sheet layout against every balance
    sheet account; a new account that no line takes shows up in the check line. No check for layouts with
    movement or budget lines (a cash flow), with lines limited by their own filter (a hotel P&L by department), with
    no accounts lines, or when the layout says Check line: no."""
    if not design.get("check", True):
        return None
    tokens, covered = [], set()
    for line in design["lines"]:
        if line["kind"] != "accounts":
            continue
        prefix, _, token = ledger.split_prefix(line["source"])
        if prefix or line.get("filter"):   # movement/budget lines, or a line limited to one department: no check
            return None
        if token.lower().startswith("calc:"):
            continue
        tokens.append(token)
        covered.update(line.get("accounts") or [])
    if not tokens or not covered:
        return None
    types = {a["number"]: a["type"] for a in ledger.accounts()}
    kinds = {types.get(n) for n in covered}
    if kinds & set(PL_SIGNS):
        every = PL_TYPES_SPEC
    else:
        bs = [t for t in ledger_module.ACCOUNT_TYPES if t not in PL_SIGNS and t not in ledger_module.STAT_TYPES]
        if not kinds <= set(bs):
            return None
        every = "type:" + ",".join(t for t in bs if t in set(types.values()))
    return every, ", ".join(tokens)


def _coverage(ledger, design):
    """P&L accounts with a balance in the report's periods that no line includes."""
    spans = [c["at"] for c in design["columns"] if c.get("at")]
    if not spans:
        return []
    covered = {number for line in design["lines"] for number in line.get("accounts", [])}
    accounts = [a for a in ledger.accounts() if a["type"] in PL_SIGNS]
    if not any(a["number"] in covered for a in accounts):
        return None  # not a profit and loss layout (a balance sheet, a cash flow): nothing to compare
    missing = [a for a in accounts if a["number"] not in covered]
    probe = missing or accounts[:1]  # one lookup still shows a wrong filter
    if not probe:
        return []
    periods = ledger.periods()
    base = dict({"fromPeriod": periods[min(s[0] for s in spans)]["name"],
                 "toPeriod": periods[max(s[1] for s in spans)]["name"]}, **design["filters"])
    results = ledger.balances([dict(base, account=a["number"]) for a in probe])
    errors = [r["error"] for r in results if "error" in r]
    if errors:
        raise ReportError(errors[0])
    return [a for a, r in zip(missing, results) if abs(r.get("value") or 0) >= 0.005]


def _detail(ledger, design):
    """For "show accounts": the accounts behind each accounts line that have a balance in the report's periods."""
    spans = [c["at"] for c in design["columns"] if c.get("at")]
    lines = [line for line in design["lines"] if line["kind"] == "accounts"]
    for line in lines:
        line["detail"] = []
    wanted = [(line, number) for line in lines if len(line["accounts"]) > 1 and not line.get("special")
              for number in line["accounts"]]
    if not design.get("detail") or not spans or not wanted:
        return
    if len(wanted) > reports.MAX_PROBE:
        raise ReportError("Listing the accounts would need %d balances, more than this does in one go." % len(wanted))
    periods = ledger.periods()
    names = {a["number"]: a["name"] for a in ledger.accounts()}
    base = dict({"fromPeriod": periods[min(s[0] for s in spans)]["name"],
                 "toPeriod": periods[max(s[1] for s in spans)]["name"]}, **design["filters"])
    results = ledger.balances([dict(base, account=number) for _, number in wanted])
    if any(c.get("data") for c in design["columns"]):  # an account with only a budget is listed too
        budgets = ledger.balances([dict(base, account="budget:" + number) for _, number in wanted])
        results = [r if abs(r.get("value") or 0) >= 0.005 else b for r, b in zip(results, budgets)]
    for (line, number), result in zip(wanted, results):
        if "error" in result:
            raise ReportError(result["error"])
        if abs(result.get("value") or 0) >= 0.005:
            line["detail"].append((number, names.get(number, "")))


def report_sheet(design):
    return reports.sheet_name(design["title"]) or "TQ Report"


def page_sheet(design, value, values=None):
    """The sheet for one location (or other segment value) of a report: "Doha - Management P&L".

    Excel allows 31 characters. When "<value> - <title>" fits for every value of the report, that is the name;
    otherwise every sheet is just the value's own name ("Technoq Salwa Road Office"), never a cut-off title."""
    short = lambda v: reports.sheet_name(v.split(" : ")[-1])  # noqa: E731
    title = reports.sheet_name(design["title"])
    if all(len(short(v)) + 3 + len(title) <= 31 for v in (values or [value])):
        return short(value) + " - " + title
    return short(value)[:31]


def page_values(ledger, design):
    """For "one sheet per location": the locations with a balance in the report's periods."""
    kind = design.get("pages")
    spans = [c["at"] for c in design["columns"] if c.get("at")]
    values = _segments(ledger, kind) if kind else []
    if not values or not spans:
        if kind:
            raise ReportError("NetSuite returned no %s list, so there is nothing to make a sheet for each of."
                              % KIND_LABEL[kind].lower())
        return []
    numbers = list(dict.fromkeys(n for line in design["lines"] for n in line.get("accounts", [])))
    if len(numbers) * len(values) > reports.MAX_PROBE:
        raise ReportError("Checking %d accounts for %d %s values is more than this does in one go."
                          % (len(numbers), len(values), kind))
    periods = ledger.periods()
    rest = {k: v for k, v in design["filters"].items() if k != kind}

    def with_balance(first, last):
        base = dict({"fromPeriod": periods[first]["name"], "toPeriod": periods[last]["name"]}, **rest)
        results = ledger.balances([dict(base, account=n, **{kind: value}) for value in values for n in numbers])
        errors = [r["error"] for r in results if "error" in r]
        if errors:
            raise ReportError(errors[0])
        size = len(numbers)
        return [value for i, value in enumerate(values)
                if any(abs(r.get("value") or 0) >= 0.005 for r in results[i * size:(i + 1) * size])]

    design["pagesWide"] = False
    live = with_balance(min(s[0] for s in spans), max(s[1] for s in spans))
    if not live:
        # Nothing in the report's own periods (e.g. this year has no postings with a location yet). The
        # report follows the year in B8, so make the sheets for the values with a balance in any period.
        live = with_balance(0, len(periods) - 1)
        design["pagesWide"] = bool(live)
    if len(live) > MAX_PAGES:
        raise ReportError("%d %s values have a balance, more than the limit of %d sheets."
                          % (len(live), kind, MAX_PAGES))
    return live


ALL_WORDS = ("all", "any", "every line", "everything", "all lines")
ACCOUNT_LIKE = re.compile(r"^(?:type:.+|[\w.]*\d[\w.]*(?:\s*(?:-|\.\.)\s*[\w.]*\d[\w.]*)?(?:\s*,\s*[\w.]*\d[\w.]*"
                          r"(?:\s*(?:-|\.\.)\s*[\w.]*\d[\w.]*)?)*)$", re.I)


def _runs(numbers):
    """[3, 4, 5, 9] -> [(3, 5), (9, 9)]"""
    out = []
    for n in sorted(set(numbers)):
        if out and n == out[-1][1] + 1:
            out[-1] = (out[-1][0], n)
        else:
            out.append((n, n))
    return out


def rule_formats(ledger, design, columns, lines, detail_rows):
    """Conditional formats for the layout's highlight rules. Rows: the lines (and account rows under them) the rule
    names; columns: the one it names, else the number columns (or the % columns, for a value written with %)."""
    out = []
    for text in design.get("rules") or []:
        rule = parse_rule(text)
        target = rule["target"]
        if target.lower() in ALL_WORDS:
            rows = [l["row"] for l in lines] + [r for _, r in detail_rows]
        elif ACCOUNT_LIKE.match(target):
            try:
                codes = set(ledger.members(target))
            except LedgerError as err:
                raise ReportError('Highlight rule "%s": %s.' % (text, err))
            rows = [l["row"] for l in lines if l.get("accounts") and set(l["accounts"]) <= codes]
            rows += [r for number, r in detail_rows if number in codes]
        else:
            hits = [l["row"] for l in lines if l["label"].lower() == target.lower()]
            if not hits:
                raise ReportError('Highlight rule "%s": there is no line or account "%s".' % (text, target))
            rows = hits
        if rule["column"]:
            at = [j for j, c in enumerate(columns) if c["label"].lower() == rule["column"].lower()]
            if not at:
                raise ReportError('Highlight rule "%s": there is no column "%s".' % (text, rule["column"]))
        else:
            at = [j for j, c in enumerate(columns) if (c["format"] == "percent") == rule["percent"]]
        font, fill = COLOURS[rule["colour"]]
        if not rows:   # said in the reply, rather than doing nothing without a word
            design.setdefault("misses", []).append(text)
        for r0, r1 in _runs(rows):
            for c0, c1 in _runs(at):
                cell = "%s%d" % (col_letter(2 + c0), r0)
                test = ("%s>=%r,%s<=%r" % (cell, rule["v1"], cell, rule["v2"]) if rule["op"] == "between"
                        else "%s%s%r" % (cell, rule["op"], rule["v1"]))
                out.append({"r0": r0 - 1, "c0": 2 + c0, "r1": r1 - 1, "c1": 2 + c1,
                            "formula": "=AND(ISNUMBER(%s),%s)" % (cell, test), "font": font, "fill": fill})
    return out


def summary_sheet(design):
    return reports.sheet_name("%s - Summary" % design["title"][:21])


def _summary(design, main, pages):
    """One row per location (or other value), with the key lines side by side, read from that value's sheet.

    main, pages: (sheet name, {line label: row}, header row) of the report and of each value's sheet.
    """
    lines, columns = design["lines"], design["columns"]
    keys = [l for l in lines if l.get("key")] or \
           [l for l in lines if l["kind"] == "subtotal" or (l["kind"] == "formula" and l["bold"])]
    if not keys:
        return None
    last = next((j for j in range(len(columns) - 1, -1, -1) if columns[j]["type"] == "total"), len(columns) - 1)
    # The first key line across every column, then each other key line at the total (or last) column.
    picks = [(keys[0], j) for j in range(len(columns))] + [(line, last) for line in keys[1:]]
    ref = lambda sheet, j, row: "='%s'!%s%d" % (sheet, col_letter(2 + j), row)  # noqa: E731
    width = 1 + len(picks)
    head = [KIND_LABEL[design["pages"]]]
    for n, (line, j) in enumerate(picks):
        if n < len(columns):
            head.append(ref(main[0], j, main[2]))               # the report's own column heading, live
        else:
            head.append(line["label"])
    rows = [[design["title"] + " - summary"] + [""] * (width - 1),
            ["%s in every column, then the other key lines for %s. Each name opens its sheet."
             % (keys[0]["label"], columns[last]["label"])] + [""] * (width - 1),
            [""] * width, head]
    for label, (sheet, row_of, _) in [("All", main)] + pages:
        rows.append(['=HYPERLINK("#\'%s\'!A1","%s")' % (sheet, label.replace('"', '""'))]
                    + [ref(sheet, j, row_of[line["label"].lower()]) for line, j in picks])
    end = len(rows) - 1
    formats = [
        {"r0": 0, "c0": 0, "r1": 0, "c1": width - 1, "bold": True, "size": 14},
        {"r0": 1, "c0": 0, "r1": 1, "c1": width - 1, "color": GREY},
        {"r0": 3, "c0": 0, "r1": 3, "c1": width - 1, "bold": True, "fill": MAROON, "color": WHITE},
        {"r0": 3, "c0": 1, "r1": 3, "c1": width - 1, "align": "right"},
        {"r0": 4, "c0": 0, "r1": 4, "c1": width - 1, "bold": True, "fill": PALE},
        {"r0": 4, "c0": 1, "r1": end, "c1": width - 1, "numberFormat": reports.NUMBER_FORMAT},
    ]
    for n, (line, j) in enumerate(picks):
        fmt = columns[j]["format"] if columns[j]["format"] in LINE_FORMATS else line["format"]
        if fmt in LINE_FORMATS:
            formats.append({"r0": 4, "c0": 1 + n, "r1": end, "c1": 1 + n, "numberFormat": LINE_FORMATS[fmt]})
    return {
        "type": "write", "target": "sheet", "sheetName": summary_sheet(design), "startCell": "", "rows": rows,
        "activate": False, "textRanges": [], "formats": formats, "freeze": {"rows": 4, "cols": 1},
        "widths": [{"c0": 0, "c1": 0, "width": 190}, {"c0": 1, "c1": width - 1, "width": 100}],
    }


def build_pack(ledger, design):
    """The report, plus a summary and one sheet per location when the layout asks for it. Returns (actions, info)."""
    action, info = build_report(ledger, design)
    actions, info["pages"], taken = [action], [], {action["sheetName"].lower()}
    kind = design.get("pages")
    pages = []
    values = page_values(ledger, design)
    for value in values:
        page = dict(design, filters=dict(design["filters"], **{kind: value}), pageValues=values)
        sheet, page_info = build_report(ledger, page, value)
        if sheet["sheetName"].lower() in taken:
            raise ReportError('Two sheets would be called "%s". Shorten the report title.' % sheet["sheetName"])
        taken.add(sheet["sheetName"].lower())
        actions.append(sheet)
        info["pages"].append(value)
        pages.append((value, (sheet["sheetName"], page_info["rows"], page_info["header"])))
    info["summary"] = ""
    if pages:
        summary = _summary(design, (action["sheetName"], info["rows"], info["header"]), pages)
        if summary and summary["sheetName"].lower() not in taken:
            actions.insert(1, summary)
            info["summary"] = summary["sheetName"]
    return actions, info


def build_report(ledger, design, page=None):
    """Returns (write action for the report sheet, info). The sheet is rewritten in place.

    page: a segment value, when this is that value's own sheet of a "one sheet per" report.
    """
    try:
        lines, columns = design["lines"], design["columns"]
        if not lines:
            raise ReportError('Layout "%s" has no lines yet. Add one, e.g. add line Sales accounts 4000-4999.'
                              % design["name"])
        if not columns:
            raise ReportError('Layout "%s" has no columns yet. Add one, e.g. add column Q1 2026.' % design["name"])
        check(ledger, design)
        unmapped = [] if page else _coverage(ledger, design)
        _detail(ledger, design)
        if report_sheet(design).lower().startswith(SHEET_PREFIX.lower()):
            raise ReportError('The report title cannot start with "%s"; that name is kept for layout sheets.'
                              % SHEET_PREFIX)
        width = 2 + len(columns)
        rows, formats, text_ranges = [], [], []

        def add(cells, **style):
            rows.append(list(cells) + [""] * (width - len(cells)))
            if style:
                formats.append(dict(r0=len(rows) - 1, c0=0, r1=len(rows) - 1, c1=width - 1, **style))
            return len(rows)  # 1-based sheet row

        add([design["title"]], bold=True, size=14)
        add(['%sBuilt from layout sheet "%s". Change the layout, then say: update report.'
             % ("%s: %s. " % (KIND_LABEL[design["pages"]], page) if page else "", design["name"])], color=GREY)
        add([KIND_LABEL["filters"], design["filters"].get("filters", ""), reports.FILTER_NOTE])
        for kind in SEGMENT_KINDS:
            add([KIND_LABEL[kind], design["filters"].get(kind, ""),
                 "Pick or type. Blank = all. A, B / not A." if kind == "subsidiary" else ""])
        formats.append({"r0": 2, "c0": 2, "r1": 6, "c1": 2, "color": GREY, "italic": True})
        text_ranges.append([2, 1, 6, 1])
        month = design.get("month") if any(c.get("relative") for c in columns) else ""
        names = [c[key] for c in columns if c["from"] for key in ("from", "to")]
        year = int(YEAR_RE.search(names[0]).group(1)) if names and all(YEAR_RE.search(n) for n in names) else None
        if month:
            year = None
            add(["Month", month, "Pick another month to roll every column forward, then use Refresh all."])
            formats.append({"r0": 7, "c0": 2, "r1": 7, "c1": 2, "color": GREY, "italic": True})
            formats.append({"r0": 7, "c0": 1, "r1": 7, "c1": 1, "bold": True, "fill": "#FFF8E1"})
            text_ranges.append([7, 1, 7, 1])
        elif year:
            add(["Year", year, "Change the year to move every period in the header."])
            formats.append({"r0": 7, "c0": 2, "r1": 7, "c1": 2, "color": GREY, "italic": True})
            formats.append({"r0": 7, "c0": 1, "r1": 7, "c1": 1, "numberFormat": "0", "align": "left", "bold": True})
        else:
            add([])
        add([])

        def named(text):
            return reports.year_formula(text, year, "$B$8") if year else text

        def period_cell(column, key):
            if not month or not column["from"]:
                return named(column[key])
            if column.get("relative"):   # worked out from the month in B8
                return '=TQ.PERIOD($B$8,"%s")' % column[key]
            return '="%s"' % column[key].replace('"', '""')   # a fixed period, kept as text

        from_row = add(["", "From period"] + [period_cell(c, "from") for c in columns], color=GREY)
        to_row = add(["", "To period"] + [period_cell(c, "to") for c in columns], color=GREY)
        kind_row = {}
        for kind in FILTER_KINDS:
            if any(c["filters"].get(kind) for c in columns):
                kind_row[kind] = add(["", KIND_LABEL[kind]] + [c["filters"].get(kind, "") for c in columns], color=GREY)
        labels = [named(c["label"]) for c in columns]
        header_row = add(["Accounts", "Name"] + labels, bold=True, bottom="thin")
        first_data = header_row + 1
        if not year and not month:
            text_ranges.append([from_row - 1, 2, to_row - 1, width - 1])
        if kind_row:
            text_ranges.append([to_row, 2, header_row - 2, width - 1])
        text_ranges += [[header_row - 1, 2 + j, header_row - 1, 2 + j]
                        for j, label in enumerate(labels) if not label.startswith("=")]
        formats.append({"r0": from_row - 1, "c0": 2, "r1": header_row - 1, "c1": width - 1, "align": "right"})

        # Where every line lands, so formulas can point at lines further down.
        row, section = first_data, None
        for line in lines:
            if line["section"].lower() != section:
                row += 0 if section is None else 1          # blank row between sections
                section = line["section"].lower()
                row += 1 if section else 0                  # the section heading
            line["row"] = row
            row += 1 + len(line.get("detail") or [])
        uses_days = any(line["kind"] == "formula" and DAYS_RE.search(REF_RE.sub("", line["source"])) for line in lines)
        days_row = row + 1
        row_of = {line["label"].lower(): line["row"] for line in lines}
        base_row = None   # the line marked "base" that [Column@base] divides by, until the next one
        for line in lines:
            base_row = line["row"] if line.get("base") else base_row
            line["base_row"] = base_row
        col_of = {c["label"].lower(): j for j, c in enumerate(columns)}
        letter = lambda j: col_letter(2 + j)  # noqa: E731

        def balance(account, j, line_filter=""):
            args = []
            for kind in FILTER_KINDS:
                whole = "$B$%d" % PARAM_ROW[kind]
                own = "%s$%d" % (letter(j), kind_row[kind]) if kind in kind_row else None
                if own and kind == "filters":   # both apply: the sheet's other filters and the column's
                    arg = 'IF(%s="",%s,IF(%s="",%s,%s&"; "&%s))' % (own, whole, whole, own, whole, own)
                else:
                    arg = 'IF(%s="",%s,%s)' % (own, whole, own) if own else whole
                if kind == "filters" and line_filter:   # the line's own filter, added to the others
                    lf = line_filter.replace('"', '""')
                    arg = 'IF(%s="","%s",%s&"; %s")' % (arg, lf, arg, lf)
                args.append(arg)
            data = columns[j].get("data")
            if data and account != '"calc:Days"':  # "budget:" in front of the account asks for the budget
                account = '"%s:"&%s' % (data.lower(), account)
            return "TQ.BALANCE(%s,%s$%d,%s$%d,%s)" % (account, letter(j), from_row, letter(j), to_row, ",".join(args))

        def across(j, at, base=None):   # a column's formula, for the line on sheet row `at`
            def ref(text):   # [Month] this line; [Month@base] the line marked base above it; [Month@Line] that line
                name, _, where = text.partition("@")
                where = where.strip().lower()
                row = at if not where else base if where == "base" else row_of.get(where)
                return "%s%d" % (letter(col_of[name.strip().lower()]), row) if row else '""'
            parts = REF_RE.split(columns[j]["formula"][1:])
            return "".join(ref(p) if n % 2 else p for n, p in enumerate(parts))

        def down(line, j):   # a line's formula, in column j
            parts = REF_RE.split(line["source"][1:])
            return "".join("%s%d" % (letter(j), row_of[p.strip().lower()]) if n % 2
                           else DAYS_RE.sub("%s$%d" % (letter(j), days_row), p) for n, p in enumerate(parts))

        def cell(line, j):
            column, at = columns[j], line["row"]
            if column["type"] == "formula":
                if column["format"] == "percent" and line["format"] == "percent":
                    return ""  # no "change %" of a percentage
                return '=' + shown(across(j, at, line.get("base_row")))
            if line["kind"] == "formula":
                return '=' + shown(down(line, j))
            if column["type"] == "total":
                return '=' + shown(across(j, at, line.get("base_row")))
            if line["kind"] == "accounts":
                return "=%s%s" % ("-" if line["sign"] < 0 else "", balance("$A%d" % at, j, line.get("filter", "")))
            added = [l["row"] for l in line["adds"]]
            if added == list(range(added[0], added[-1] + 1)):
                return "=SUM(%s%d:%s%d)" % (letter(j), added[0], letter(j), added[-1])
            return "=" + "+".join("%s%d" % (letter(j), r) for r in added)

        extra, groups, section, first = [], [], None, None
        detail_rows = []   # (account number, sheet row) of the account rows under the lines

        def close_group():
            if section and first is not None and len(rows) > first:
                groups.append([first, len(rows) - 1])

        for line in lines:
            if line["section"].lower() != section:
                close_group()
                if section is not None:
                    add([])
                section = line["section"].lower()
                if section:
                    at = add(["", line["section"]], bold=True)
                    extra.append({"r0": at - 1, "c0": 0, "r1": at - 1, "c1": width - 1, "fill": NAVY, "color": WHITE})
                first = len(rows)
            style = {"bold": True, "top": "thin"} if line["kind"] == "subtotal" else {"bold": True} if line["bold"] else {}
            at = add([line["source"] if line["kind"] == "accounts" else "", line["label"]]
                     + [cell(line, j) for j in range(len(columns))], **style)
            assert at == line["row"]
            if line["kind"] == "subtotal":
                extra.append({"r0": at - 1, "c0": 0, "r1": at - 1, "c1": width - 1, "fill": PALE})
            else:
                extra.append({"r0": at - 1, "c0": 1, "r1": at - 1, "c1": 1, "indent": 1})
            if line["format"] in LINE_FORMATS:
                extra.append({"r0": at - 1, "c0": 2, "r1": at - 1, "c1": width - 1,
                              "numberFormat": LINE_FORMATS[line["format"]]})
            for number, name in line.get("detail") or []:   # the accounts behind the line, one row each
                row = len(rows) + 1
                detail_rows.append((number, row))
                cells = []
                for j, column in enumerate(columns):
                    if column["type"] == "period":
                        cells.append("=%s%s" % ("-" if line["sign"] < 0 else "", balance("$A%d" % row, j, line.get("filter", ""))))
                    else:
                        cells.append('=' + shown(across(j, row, line.get("base_row"))))
                add([number, name] + cells, color=GREY)
                extra.append({"r0": row - 1, "c0": 1, "r1": row - 1, "c1": 1, "indent": 3})
            if line.get("detail"):
                groups.append([at, at + len(line["detail"]) - 1])
        close_group()
        last_line = len(rows)
        if uses_days:
            add([])
            at = add(["", "Days in period"] + ["=" + balance('"calc:Days"', j) if c["from"] else ""
                                               for j, c in enumerate(columns)], color=GREY, italic=True)
            assert at == days_row
        check_row = None
        spec = check_spec(ledger, design)
        if spec:   # what no line includes: P&L (or balance sheet) accounts minus the lines' accounts, per column
            every, covered = ('"%s"' % s.replace('"', '""') for s in spec)
            add([])
            check_row = len(rows) + 1
            cells = []
            for j, c in enumerate(columns):
                if c["type"] == "period" and c["from"]:
                    cells.append("=" + shown("%s-%s" % (balance(every, j), balance(covered, j))))
                elif c["type"] == "total":
                    cells.append('=' + shown(across(j, check_row)))
                else:
                    cells.append("")
            add(["", CHECK_LABEL] + cells, italic=True, color=GREY, top="thin")
        if len(rows) * width > reports.MAX_CELLS:
            raise ReportError("That report would be %d cells, more than this builds in one go. "
                              "Use fewer columns." % (len(rows) * width))
        last = len(rows) - 1
        text_ranges.append([first_data - 1, 0, last, 0])
        formats.append({"r0": first_data - 1, "c0": 0, "r1": last, "c1": 0, "color": GREY})
        formats.append({"r0": first_data - 1, "c0": 2, "r1": last, "c1": width - 1,
                        "numberFormat": reports.NUMBER_FORMAT})
        formats += extra
        formats += [{"r0": first_data - 1, "c0": 2 + j, "r1": last_line - 1, "c1": 2 + j,
                     "numberFormat": LINE_FORMATS[c["format"]]}
                    for j, c in enumerate(columns) if c["format"] in LINE_FORMATS]
        if uses_days:
            formats.append({"r0": days_row - 1, "c0": 2, "r1": days_row - 1, "c1": width - 1, "numberFormat": "0"})
        formats.append({"r0": header_row - 1, "c0": 0, "r1": header_row - 1, "c1": width - 1,
                        "fill": MAROON, "color": WHITE})
        action = {
            "type": "write", "target": "sheet", "startCell": "", "rows": rows, "activate": not page,
            "sheetName": page_sheet(design, page, design.get("pageValues")) if page else report_sheet(design),
            "textRanges": text_ranges, "formats": formats, "groups": groups,
            "freeze": {"rows": header_row, "cols": 2},
            "widths": [{"c0": 0, "c1": 0, "width": 80}, {"c0": 1, "c1": 1, "width": 230},
                       {"c0": 2, "c1": width - 1, "width": 92}],
        }
        action["dropdowns"] = []
        if design.get("highlight"):   # change % / var % cells: green at or above the threshold, red at or below minus it
            limit = design["highlight"] / 100.0
            action["conditional"] = [{"r0": first_data - 1, "c0": 2 + j, "r1": last_line - 1, "c1": 2 + j, "limit": limit,
                                      "cell": "%s%d" % (col_letter(2 + j), first_data)}   # the range's top cell, relative
                                     for j, c in enumerate(columns) if c["type"] != "period" and c["format"] == "percent"]
        if design.get("rules"):
            action.setdefault("conditional", [])
            action["conditional"] += rule_formats(ledger, design, columns, lines, detail_rows)
        if design.get("comments"):    # typed notes, kept by the panel when the sheet is rebuilt (matched by line name)
            col = len(rows[0])
            for r in rows:
                r.append("")
            rows[header_row - 1][col] = "Comments"
            formats += [{"r0": header_row - 1, "c0": col, "r1": header_row - 1, "c1": col, "fill": MAROON,
                         "color": WHITE, "bold": True},
                        {"r0": first_data - 1, "c0": col, "r1": last, "c1": col, "color": "#1F3A5F", "wrap": True}]
            text_ranges.append([first_data - 1, col, last, col])
            action["widths"].append({"c0": col, "c1": col, "width": 260})
            action["keep"] = {"header": header_row - 1, "title": "Comments", "key": 1}
            pct = next((j for j, c in enumerate(columns) if c["type"] != "period" and c["format"] == "percent"), None)
            if design.get("commentAt") and pct is not None:   # a big variance with no comment yet: marked yellow
                limit = design["commentAt"] / 100.0
                at, me = col_letter(2 + pct), col_letter(col)
                action.setdefault("conditional", []).append({
                    "r0": first_data - 1, "c0": col, "r1": last_line - 1, "c1": col,
                    "formula": '=AND(ISNUMBER($%s%d),ABS($%s%d)>=%g,LEN(TRIM(%s%d))=0,LEN($B%d)>0)' % (
                        at, first_data, at, first_data, limit, me, first_data, first_data),
                    "fill": "#FFF4CE", "font": "#7A5B00"})
        if check_row:   # the check turns red when it is not zero
            first_col = col_letter(2)
            action.setdefault("conditional", []).append({
                "r0": check_row - 1, "c0": 2, "r1": check_row - 1, "c1": width - 1,
                "formula": "=AND(ISNUMBER(%s%d),ABS(%s%d)>=0.005)" % (first_col, check_row, first_col, check_row),
                "fill": "#FDE7E9", "font": "#A4262C"})
        years = sorted({int(y) for p in ledger.periods() for y in YEAR_RE.findall(p["name"])})
        if year and 1 < len(years) <= 30:
            action["dropdowns"].append({"r": 7, "c": 1, "formula": ",".join(str(y) for y in years)})
        extra = []
        if month:   # B8 picks from the months (adjustment periods and quarters left out)
            extra.append(("Month list", [p["name"] for p in ledger.periods() if (p["end"] - p["start"]).days > 20
                                         and (p["end"] - p["start"]).days < 40], [(7, 1)]))
        reports.add_filter_lists(action, ledger, reports.dropdown_fields(ledger, design.get("dropdowns") or []),
                                 design["filters"].get("filters", ""), extra)
        return action, {
            "rows": {line["label"].lower(): line["row"] for line in lines}, "header": header_row,
            "title": design["title"], "lines": len(lines), "columns": len(columns), "unmapped": unmapped,
            "year": year, "month": month, "sheet": action["sheetName"],
            "empty": [line["label"] for line in lines if line["kind"] == "accounts" and not line["accounts"]
                      and not (line.get("special") and not line.get("accounts") and line["source"].lower().startswith("calc:"))],
        }
    except LedgerError as err:
        raise ReportError(str(err))


# ------------------------------------------------------------ sample layout --

SAMPLE_SECTIONS = [  # (section, account type, sign, subtotal label, name of a one-line section)
    ("Revenue", "Income", 1, "Net Sales", "Sales"),
    ("Cost of Sales", "COGS", -1, "Total Cost of Sales", "Cost of Sales"),
    ("Operating Expenses", "Expense", -1, "Total Operating Expenses", "Expenses"),
    ("Other Income and Expenses", "OthIncome", 1, None, "Other Income"),
    ("Other Income and Expenses", "OthExpense", -1, "Total Other", "Other Expenses"),
]


def sample(ledger, name=SHEET_PREFIX):
    """A starting design made from this NetSuite account's own P&L accounts."""
    try:
        periods = ledger.periods()
        current = ledger.current_index()
        start = ledger.fy_start(current)
        start = ledger.fy_start(start - 1) if start > 0 else start  # this fiscal year and the one before
        accounts = [a for a in ledger.accounts() if a["type"] in PL_SIGNS]
        results = ledger.balances([{"account": a["number"], "fromPeriod": periods[start]["name"],
                                    "toPeriod": periods[current]["name"]} for a in accounts])
        errors = [r["error"] for r in results if "error" in r]
        if errors:
            raise ReportError(errors[0])
        active = [a for a, r in zip(accounts, results) if abs(r.get("value") or 0) >= 0.005]
        design = {"name": name, "title": "Management P&L", "filters": {}, "pages": "", "detail": False,
                  "lines": [], "columns": []}
        subtotals = []
        for section, kind, sign, subtotal, one_line in SAMPLE_SECTIONS:
            members = sorted((a for a in active if a["type"] == kind), key=lambda a: a["number"])
            if 0 < len(members) <= MAX_SAMPLE_LINES:
                for a in members:
                    label, n = a["name"], 2
                    while any(l["label"].lower() == label.lower() for l in design["lines"]):
                        label, n = "%s (%d)" % (a["name"], n), n + 1
                    design["lines"].append(new_line(label, "accounts", a["number"], sign, "", section))
            elif any(a["type"] == kind for a in accounts):  # too many to list, or none active: one line for the type
                design["lines"].append(new_line(one_line, "accounts", "type:" + kind, sign, "", section))
            if subtotal and any(l["section"] == section for l in design["lines"]):
                design["lines"].append(new_line(subtotal, "subtotal", "", "", "", section))
                subtotals.append(subtotal)
                if subtotal == "Total Cost of Sales" and "Net Sales" in subtotals:
                    design["lines"].append(new_line("Gross Profit", "formula", "=[Net Sales]+[Total Cost of Sales]",
                                                    "", "bold", "Gross Profit"))
                    design["lines"].append(new_line("Gross Margin", "formula", "=[Gross Profit]/[Net Sales]",
                                                    "", "percent", "Gross Profit"))
        if not design["lines"]:
            raise ReportError("NetSuite returned no profit and loss accounts, so there is nothing to lay out.")
        design["lines"].append(new_line("Net Profit", "formula", "=" + "+".join("[%s]" % s for s in subtotals),
                                        "", "bold", "Result"))
        if "Net Sales" in subtotals:
            design["lines"].append(new_line("Net Margin", "formula", "=[Net Profit]/[Net Sales]", "", "percent", "Result"))
            design["lines"].append(new_line("Average Sales Daily", "formula", "=[Net Sales]/days", "", "", "Result"))
        design["columns"] = default_columns(ledger)
        return design
    except LedgerError as err:
        raise ReportError(str(err))
