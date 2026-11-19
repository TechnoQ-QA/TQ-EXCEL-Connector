"""TQ Excel Connector - data queries (anything in NetSuite, not only balances).

A query brings a table of NetSuite data into a worksheet: a saved search as it
is defined in NetSuite, or a SuiteQL SELECT over any record tables. Queries
are kept on a sheet called "TQ Queries", so they are saved with the workbook:

    Query          Type    Source                          Output sheet   Max rows  Last run  Rows
    Open invoices  search  customsearch_open_invoices      Open invoices  5000
    Customers      sql     SELECT id, companyname FROM ... Customers      5000

    Parameter  Value
    from       2026-01-01          used in a query as {from}

Each query writes to its own output sheet, which is rewritten in place every
time it runs, so pivots and formulas pointing at it keep working. "Refresh
data" runs them all again.

Queries run in NetSuite as the user's own role (tq_excel_restlet.js), so a user
gets only the records that role can see.
"""

import datetime
import re

import reports
from ledger import LedgerError

SHEET = "TQ Queries"
HEADERS = ["Query", "Type", "Source", "Output sheet", "Max rows", "Last run", "Rows"]
PARAM_HEADERS = ["Parameter", "Value"]
PARAM_AT = len(HEADERS) + 1
WIDTH = PARAM_AT + len(PARAM_HEADERS)
MAX_QUERIES = 100
MAX_ROWS = 20000
DEFAULT_ROWS = 5000
MAX_CELL = 32000
NOTES = [
    "One row per query. Type search: Source is a saved search ID (customsearch...). Type sql: Source is a SuiteQL SELECT.",
    "A query can use a parameter as {name}; give its value in the Parameter table on the right.",
    "Add and run queries under Data tables in the Reports tab. Each query fills its own output sheet; Refresh data runs them all.",
]
IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
FIELD = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?$")
PARAM = re.compile(r"\{(\w+)\}")
NUMBER = re.compile(r"^-?(?:0|[1-9]\d{0,14})(?:\.\d+)?$")
TOTALS = {"sum": "SUM", "count": "COUNT", "min": "MIN", "max": "MAX", "avg": "AVG"}
COMPARE = ("=", "<>", ">", ">=", "<", "<=")

# Common record tables, to start a query from by clicking. Any other table name can be typed.
CATALOG = [
    {"table": "customer", "label": "Customers",
     "fields": ["id", "entityid", "companyname", "email", "phone", "salesrep", "category", "datecreated", "isinactive"]},
    {"table": "vendor", "label": "Vendors",
     "fields": ["id", "entityid", "companyname", "email", "phone", "category", "datecreated", "isinactive"]},
    {"table": "employee", "label": "Employees",
     "fields": ["id", "entityid", "firstname", "lastname", "email", "title", "department", "location",
                "supervisor", "hiredate", "isinactive"]},
    {"table": "item", "label": "Items",
     "fields": ["id", "itemid", "displayname", "itemtype", "description", "isinactive"]},
    {"table": "transaction", "label": "Transactions (headers)",
     "fields": ["id", "tranid", "type", "trandate", "entity", "status", "foreigntotal", "currency",
                "postingperiod", "duedate", "memo", "createddate"],
     "links": [["transaction.entity", "customer", "id"], ["transaction.entity", "vendor", "id"],
               ["transaction.postingperiod", "accountingperiod", "id"]]},
    {"table": "transactionline", "label": "Transaction lines",
     "fields": ["transaction", "id", "linesequencenumber", "item", "quantity", "rate", "netamount",
                "foreignamount", "department", "class", "location", "subsidiary", "memo", "mainline"],
     "links": [["transactionline.transaction", "transaction", "id"], ["transactionline.item", "item", "id"],
               ["transactionline.department", "department", "id"], ["transactionline.location", "location", "id"]]},
    {"table": "transactionaccountingline", "label": "GL lines",
     "fields": ["transaction", "transactionline", "account", "amount", "debit", "credit", "posting"],
     "links": [["transactionaccountingline.transaction", "transaction", "id"],
               ["transactionaccountingline.account", "account", "id"]]},
    {"table": "account", "label": "Accounts", "fields": ["id", "acctnumber", "fullname", "accttype", "isinactive"]},
    {"table": "accountingperiod", "label": "Accounting periods",
     "fields": ["id", "periodname", "startdate", "enddate", "closed"]},
    {"table": "department", "label": "Departments", "fields": ["id", "name", "fullname", "isinactive"]},
    {"table": "classification", "label": "Classes", "fields": ["id", "name", "fullname", "isinactive"]},
    {"table": "location", "label": "Locations", "fields": ["id", "name", "fullname", "isinactive"]},
    {"table": "subsidiary", "label": "Subsidiaries", "fields": ["id", "name", "fullname", "isinactive"]},
    {"table": "currency", "label": "Currencies", "fields": ["id", "name", "symbol"]},
    {"table": "aggregateitemlocation", "label": "Stock by item and location",
     "fields": ["item", "location", "quantityonhand", "quantityavailable", "quantityonorder"],
     "links": [["aggregateitemlocation.item", "item", "id"], ["aggregateitemlocation.location", "location", "id"]]},
]


def _aging(kind):
    bucket = "SUM(CASE WHEN %s THEN t.foreignamountunpaid ELSE 0 END) AS %s"
    return "\n".join([
        "SELECT BUILTIN.DF(t.entity) AS name,",
        "  " + bucket % ("NVL(t.daysoverduesearch, 0) <= 0", "not_due") + ",",
        "  " + bucket % ("t.daysoverduesearch BETWEEN 1 AND 30", "days_1_30") + ",",
        "  " + bucket % ("t.daysoverduesearch BETWEEN 31 AND 60", "days_31_60") + ",",
        "  " + bucket % ("t.daysoverduesearch BETWEEN 61 AND 90", "days_61_90") + ",",
        "  " + bucket % ("t.daysoverduesearch > 90", "over_90") + ",",
        "  SUM(t.foreignamountunpaid) AS total",
        "FROM transaction t",
        "WHERE t.type = '%s' AND t.foreignamountunpaid > 0" % kind,
        "GROUP BY BUILTIN.DF(t.entity)",
        "ORDER BY SUM(t.foreignamountunpaid) DESC"])


# Ready-made queries for the Data tab. Each was run against a NetSuite sandbox; amounts are in the
# transaction's own currency, and a user sees only the records their role can see.
TEMPLATES = [
    {"name": "Receivables aging", "sql": _aging("CustInvc"),
     "note": "Unpaid customer invoices by customer, in aging buckets by days overdue today."},
    {"name": "Payables aging", "sql": _aging("VendBill"),
     "note": "Unpaid vendor bills by vendor, in aging buckets by days overdue today."},
    {"name": "Open invoices", "note": "Every unpaid customer invoice with its due date and days overdue.",
     "sql": "SELECT BUILTIN.DF(t.entity) AS customer, t.tranid AS invoice, t.trandate AS invoice_date, "
            "t.duedate AS due_date, t.daysoverduesearch AS days_overdue, t.foreigntotal AS amount, "
            "t.foreignamountunpaid AS unpaid\nFROM transaction t\n"
            "WHERE t.type = 'CustInvc' AND t.foreignamountunpaid > 0\nORDER BY t.duedate"},
    {"name": "Purchases by vendor", "note": "Vendor bills between {from} and {to}, totalled by vendor.",
     "sql": "SELECT BUILTIN.DF(t.entity) AS vendor, COUNT(t.id) AS bills, SUM(ABS(t.foreigntotal)) AS total\n"
            "FROM transaction t\nWHERE t.type = 'VendBill' AND t.trandate >= {from} AND t.trandate <= {to}\n"
            "GROUP BY BUILTIN.DF(t.entity)\nORDER BY SUM(ABS(t.foreigntotal)) DESC"},
    {"name": "Purchase orders", "note": "Purchase orders dated between {from} and {to}, with their status.",
     "sql": "SELECT t.tranid AS purchase_order, t.trandate AS order_date, BUILTIN.DF(t.entity) AS vendor, "
            "BUILTIN.DF(t.status) AS status, ABS(t.foreigntotal) AS amount\nFROM transaction t\n"
            "WHERE t.type = 'PurchOrd' AND t.trandate >= {from} AND t.trandate <= {to}\nORDER BY t.trandate DESC"},
    {"name": "Sales by customer", "note": "Invoices and cash sales between {from} and {to}, totalled by customer.",
     "sql": "SELECT BUILTIN.DF(t.entity) AS customer, COUNT(t.id) AS documents, SUM(t.foreigntotal) AS total\n"
            "FROM transaction t\nWHERE t.type IN ('CustInvc', 'CashSale') AND t.trandate >= {from} "
            "AND t.trandate <= {to}\nGROUP BY BUILTIN.DF(t.entity)\nORDER BY SUM(t.foreigntotal) DESC"},
    {"name": "Sales by item", "note": "Quantity and amount sold per item between {from} and {to}.",
     "sql": "SELECT BUILTIN.DF(tl.item) AS item, SUM(-NVL(tl.quantity, 0)) AS quantity, SUM(-NVL(tl.netamount, 0)) AS amount\n"
            "FROM transactionline tl\nINNER JOIN transaction t ON t.id = tl.transaction\n"
            "WHERE t.type IN ('CustInvc', 'CashSale') AND tl.mainline = 'F' AND tl.taxline = 'F' "
            "AND tl.item IS NOT NULL AND t.trandate >= {from} AND t.trandate <= {to}\n"
            "GROUP BY BUILTIN.DF(tl.item)\nORDER BY SUM(-NVL(tl.netamount, 0)) DESC"},
    {"name": "Stock on hand", "note": "Quantity on hand and available for each item and location.",
     "sql": "SELECT BUILTIN.DF(a.item) AS item, BUILTIN.DF(a.location) AS location, a.quantityonhand AS on_hand, "
            "a.quantityavailable AS available\nFROM aggregateitemlocation a\nWHERE a.quantityonhand <> 0\n"
            "ORDER BY BUILTIN.DF(a.item)"},
    {"name": "Budget by account and period", "note": "The NetSuite budgets, one row per account, category and period.",
     "sql": "SELECT BUILTIN.DF(b.account) AS account, BUILTIN.DF(b.category) AS category, "
            "BUILTIN.DF(b.subsidiary) AS subsidiary, BUILTIN.DF(b.department) AS department, "
            "BUILTIN.DF(bm.period) AS period, bm.amount AS amount\nFROM budgets b\n"
            "INNER JOIN budgetsmachine bm ON bm.budget = b.id\nINNER JOIN accountingperiod ap ON ap.id = bm.period\n"
            "WHERE ap.isquarter = 'F' AND ap.isyear = 'F'\nORDER BY BUILTIN.DF(b.account), ap.startdate"},
]


class QueryError(Exception):
    pass


def _cell(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def _rows_limit(value):
    try:
        return max(1, min(MAX_ROWS, int(float(_cell(value)))))
    except ValueError:
        return DEFAULT_ROWS


# ---------------------------------------------------------- the queries sheet --

def parse(rows, first_row=0):
    """The cells of the "TQ Queries" sheet -> {"queries": [...], "params": {name: value}}."""
    grid = [[_cell(c) for c in row] if isinstance(row, list) else [] for row in (rows or [])]
    state = {"queries": [], "params": {}}
    header = next((n for n, row in enumerate(grid)
                   if "query" in [c.lower() for c in row] and "source" in [c.lower() for c in row]), None)
    if header is None:
        return state
    cols = {}
    for at, cell in enumerate(c.lower() for c in grid[header]):
        cols.setdefault({"output sheet": "sheet", "max rows": "max", "last run": "last"}.get(cell, cell), at)
    for n in range(header + 1, len(grid)):
        row = grid[n]
        get = lambda key: row[cols[key]] if key in cols and cols[key] < len(row) else ""  # noqa: E731
        if get("query") or get("source"):
            where = 'Sheet "%s", row %d' % (SHEET, first_row + n + 1)
            query = new_query(get("query"), get("type"), get("source"), get("sheet"), get("max"), where)
            query.update(last=get("last"), rows=get("rows"))
            state["queries"].append(query)
        if get("parameter"):
            state["params"][get("parameter").strip("{} ").lower()] = get("value")
    if len(state["queries"]) > MAX_QUERIES:
        raise QueryError('Sheet "%s" has more than %d queries.' % (SHEET, MAX_QUERIES))
    return state


def new_query(name, kind, source, sheet="", max_rows="", where=None):
    name, kind, source = _cell(name), _cell(kind).lower(), _cell(source)
    where = where or 'Query "%s"' % name
    if not name:
        raise QueryError("%s: the query needs a name." % where)
    if not source:
        raise QueryError('%s: "%s" has no source (a saved search ID or a SELECT).' % (where, name))
    if kind not in ("search", "sql"):
        kind = "sql" if re.match(r"\s*(select|with)\b", source, re.I) else "search"
    if kind == "sql" and not re.match(r"\s*(select|with)\b", source, re.I):
        raise QueryError('%s: a SuiteQL query must start with SELECT.' % where)
    if len(source) > MAX_CELL:
        raise QueryError("%s: the query text is too long for one cell." % where)
    sheet = reports.sheet_name(_cell(sheet) or name)
    if sheet.lower() == SHEET.lower() or sheet.lower().startswith("tq layout"):
        raise QueryError('%s: "%s" cannot be the output sheet. Give the query another name.' % (where, sheet))
    return {"name": name, "type": kind, "source": source, "sheet": sheet, "max": _rows_limit(max_rows),
            "last": "", "rows": ""}


def public(state):
    return {"queries": [dict(q) for q in state["queries"]],
            "params": [{"name": k, "value": v} for k, v in state["params"].items()]}


def emit(state):
    """The write action that puts the queries back on their sheet."""
    rows = [[SHEET]] + [[note] for note in NOTES] + [[], HEADERS + [""] + PARAM_HEADERS]
    header = len(rows) - 1
    params = list(state["params"].items())
    for n in range(max(len(state["queries"]), len(params), 1)):
        row = [""] * WIDTH
        if n < len(state["queries"]):
            q = state["queries"][n]
            row[:len(HEADERS)] = [q["name"], q["type"], q["source"], q["sheet"], str(q["max"]), q["last"], str(q["rows"])]
        if n < len(params):
            row[PARAM_AT:] = list(params[n])
        rows.append(row)
    rows = [row + [""] * (WIDTH - len(row)) for row in rows]
    last = len(rows) - 1
    return {
        "type": "write", "target": "sheet", "sheetName": SHEET, "startCell": "", "rows": rows, "activate": False,
        "textRanges": [[0, 0, last, WIDTH - 1]],  # a SELECT or a date must stay as typed
        "formats": [
            {"r0": 0, "c0": 0, "r1": 0, "c1": WIDTH - 1, "bold": True, "size": 14},
            {"r0": 1, "c0": 0, "r1": header - 2, "c1": WIDTH - 1, "color": reports.GREY, "italic": True},
            {"r0": header, "c0": 0, "r1": header, "c1": len(HEADERS) - 1, "bold": True, "fill": "#6B1F24",
             "color": "#FFFFFF"},
            {"r0": header, "c0": PARAM_AT, "r1": header, "c1": WIDTH - 1, "bold": True, "fill": "#1F3A5F",
             "color": "#FFFFFF"},
        ],
        "freeze": {"rows": header + 1, "cols": 0},
        "widths": [{"c0": 0, "c1": 0, "width": 150}, {"c0": 1, "c1": 1, "width": 55}, {"c0": 2, "c1": 2, "width": 330},
                   {"c0": 3, "c1": 3, "width": 130}, {"c0": 4, "c1": 6, "width": 75}, {"c0": 7, "c1": 7, "width": 25},
                   {"c0": PARAM_AT, "c1": WIDTH - 1, "width": 110}],
    }


# ------------------------------------------------------------ building a SELECT --

def _field(text, table):
    name = _cell(text)
    if not FIELD.match(name):
        raise QueryError('"%s" is not a field name.' % name)
    return name if "." in name else "%s.%s" % (table, name)


def _literal(value):
    text = _cell(value)
    if PARAM.fullmatch(text) or NUMBER.match(text):
        return text
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return "TO_DATE('%s', 'YYYY-MM-DD')" % text
    return "'%s'" % text.replace("'", "''")


def compose(spec):
    """A query described by clicks -> a SuiteQL SELECT.

    spec: table, fields [{field, show: "name"|"", total: sum|count|min|max|avg|"", label}],
          joins [{table, left, right}], filters [{field, op, value}], sort [{field, desc}]
    """
    spec = spec if isinstance(spec, dict) else {}
    table = _cell(spec.get("table")).lower()
    if not IDENT.match(table):
        raise QueryError("Choose a table first.")
    items = lambda key: [i for i in (spec.get(key) or []) if isinstance(i, dict)]  # noqa: E731
    select, plain, used = [], [], set()
    for item in items("fields"):
        expr = _field(item.get("field"), table)
        total = _cell(item.get("total")).lower()
        name = (re.sub(r"\W+", "_", _cell(item.get("label")).lower()).strip("_")
                or expr.split(".")[-1] + ("_name" if item.get("show") == "name" else "")
                + ("_" + total if total in TOTALS else ""))
        base, n = name, 2
        while name in used:
            name, n = "%s_%d" % (base, n), n + 1
        used.add(name)
        shown = "BUILTIN.DF(%s)" % expr if item.get("show") == "name" else expr
        if total in TOTALS:
            select.append("%s(%s) AS %s" % (TOTALS[total], expr, name))
        else:
            select.append("%s AS %s" % (shown, name))
            plain.append(shown)
    if not select:
        raise QueryError("Tick at least one field.")
    sql = ["SELECT " + ", ".join(select), "FROM " + table]
    for join in items("joins"):
        other = _cell(join.get("table")).lower()
        if not IDENT.match(other) or other == table:
            raise QueryError('"%s" cannot be joined.' % other)
        sql.append("LEFT JOIN %s ON %s = %s" % (other, _field(join.get("right"), other), _field(join.get("left"), table)))
    where = []
    for item in items("filters"):
        expr, op, value = _field(item.get("field"), table), _cell(item.get("op")).lower(), _cell(item.get("value"))
        if op == "empty":
            where.append("%s IS NULL" % expr)
        elif op == "notempty":
            where.append("%s IS NOT NULL" % expr)
        elif not value:
            continue
        elif op in COMPARE:
            where.append("%s %s %s" % (expr, op, _literal(value)))
        elif op in ("contains", "starts"):
            lead = "'%' || " if op == "contains" else ""
            inner = value if PARAM.fullmatch(value) else "'%s'" % value.replace("'", "''")
            where.append("LOWER(%s) LIKE %sLOWER(%s) || '%%'" % (expr, lead, inner))
        elif op == "in":
            where.append("%s IN (%s)" % (expr, ", ".join(_literal(v) for v in value.split(",") if v.strip())))
        else:
            raise QueryError('Filter "%s" is not known.' % op)
    if where:
        sql.append("WHERE " + " AND ".join(where))
    if len(plain) < len(select) and plain:
        sql.append("GROUP BY " + ", ".join(plain))
    order = ["%s%s" % (_field(item.get("field"), table), " DESC" if item.get("desc") else "") for item in items("sort")
             if _cell(item.get("field"))]
    if order:
        sql.append("ORDER BY " + ", ".join(order))
    return "\n".join(sql)


# -------------------------------------------------------------------- running --

def _value(value):
    """A NetSuite value as a worksheet cell: numbers as numbers, the rest as text that Excel will not reinterpret."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, (int, float)):
        return value
    text = str(value)[:MAX_CELL]
    if NUMBER.match(text):
        return float(text) if "." in text else int(text)
    if text[:1] in "=+-@" or re.fullmatch(r"0\d+", text):
        return "'" + text  # a leading apostrophe keeps "=..." or "00123" as text
    return text


def _title(name):
    return re.sub(r"_+", " ", str(name)).strip().title() if str(name).islower() else str(name)


class QueryService:
    """Answers POST /api/query for the Data tab, and the chat's data commands."""

    def __init__(self, ledger):
        self.ledger = ledger

    def _call(self, payload):
        try:
            return self.ledger.call(payload)
        except LedgerError as err:
            if str(err).startswith("Unknown action"):
                raise QueryError("The scripts in this NetSuite account are older than this connector and cannot run "
                                 "data queries yet. Replace tq_excel_restlet.js in the File Cabinet with the new version "
                                 "(README, section 4).")
            raise QueryError(str(err))

    def searches(self, text=""):
        return self._call({"action": "searchList", "text": _cell(text)[:200]}).get("searches", [])

    def run(self, query, params):
        """Runs one query. Returns the write action for its output sheet and updates its status."""
        if query["type"] == "search":
            data = self._call({"action": "searchRun", "id": query["source"], "maxRows": query["max"]})
        else:
            values = []

            def bind(match):
                name = match.group(1).lower()
                if name not in params or params[name] == "":
                    raise QueryError('Query "%s" uses {%s}, which has no value. Add it under Parameters.'
                                     % (query["name"], match.group(1)))
                values.append(params[name])
                # NetSuite will not compare a date field with text, so a date value is sent as a date,
                # unless the query already wraps the parameter in TO_DATE(...) itself.
                wrapped = re.search(r"TO_DATE\(\s*$", query["source"][:match.start()], re.I)
                if re.fullmatch(r"\d{4}-\d{2}-\d{2}", params[name]) and not wrapped:
                    return "TO_DATE(?, 'YYYY-MM-DD')"
                return "?"
            text = PARAM.sub(bind, query["source"])
            data = self._call({"action": "sql", "query": text, "params": values, "maxRows": query["max"]})
        columns = [_title(c) for c in data.get("columns") or []]
        found = data.get("rows") or []
        width = max([len(columns)] + [len(r) for r in found if isinstance(r, list)] + [1])
        rows = [columns + [""] * (width - len(columns))] if columns else [["No rows"] + [""] * (width - 1)]
        rows += [[_value(v) for v in row] + [""] * (width - len(row)) for row in found if isinstance(row, list)]
        query["rows"] = len(found)
        query["last"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        query["truncated"] = bool(data.get("truncated"))
        return {
            "type": "write", "target": "sheet", "sheetName": query["sheet"], "startCell": "", "rows": rows,
            "textRanges": [], "autofit": True, "freeze": {"rows": 1, "cols": 0},
            "formats": [{"r0": 0, "c0": 0, "r1": 0, "c1": width - 1, "bold": True, "fill": "#1F3A5F", "color": "#FFFFFF"}],
        }

    def run_all(self, state, only=None):
        """Runs every query (or the one named). Returns (actions, reply lines); one failure does not stop the rest."""
        actions, lines = [], []
        wanted = [q for q in state["queries"] if only is None or q["name"].lower() == only.lower()]
        if not wanted:
            raise QueryError('There is no query called "%s".' % only if only else
                             "This workbook has no data queries yet. Add one under Data tables in the Reports tab.")
        for query in wanted:
            try:
                action = self.run(query, state["params"])
            except QueryError as err:
                lines.append('%s: failed. %s' % (query["name"], err))
                query["last"], query["rows"] = "failed", ""
                continue
            action["activate"] = len(wanted) == 1
            actions.append(action)
            lines.append('%s: %d row(s) on sheet "%s"%s.' % (
                query["name"], query["rows"], query["sheet"],
                ", cut at the limit of %d (raise Max rows to get more)" % query["max"] if query.pop("truncated") else ""))
        return [emit(state)] + actions, lines

    def handle(self, payload):
        context = payload.get("context") if isinstance(payload.get("context"), dict) else {}
        args = payload.get("args") if isinstance(payload.get("args"), dict) else {}
        op = _cell(payload.get("op"))
        arg = lambda key: _cell(args.get(key))  # noqa: E731
        sheet = context.get("queries") if isinstance(context.get("queries"), dict) else {}
        try:
            if op == "catalog":
                return {"ok": True, "catalog": CATALOG, "templates": TEMPLATES}
            if op == "searches":
                return {"ok": True, "searches": self.searches(arg("text"))}
            if op == "columns":
                data = self._call({"action": "tableColumns", "table": arg("table")})
                return {"ok": True, "columns": data.get("columns") or [], "empty": bool(data.get("empty"))}
            if op == "compose":
                return {"ok": True, "sql": compose(args.get("spec"))}
            state = parse(sheet.get("rows"), sheet.get("row") if isinstance(sheet.get("row"), int) else 0)
            if op == "get":
                return {"ok": True, "state": public(state), "actions": [], "reply": ""}
            if op == "save":
                query = new_query(arg("name"), arg("type"), arg("source"), arg("sheet"), arg("max"))
                others = [q for q in state["queries"] if q["name"].lower() != query["name"].lower()]
                clash = next((q for q in others if q["sheet"].lower() == query["sheet"].lower()), None)
                if clash:
                    raise QueryError('Query "%s" already writes to sheet "%s". Give this one another name.'
                                     % (clash["name"], clash["sheet"]))
                at = next((n for n, q in enumerate(state["queries"]) if q["name"].lower() == query["name"].lower()),
                          len(state["queries"]))
                existing = [_cell(n).lower() for n in context.get("sheets") or [] if isinstance(n, str)]
                if at == len(state["queries"]) and query["sheet"].lower() in existing:
                    # A new query must not wipe a sheet the user already has under that name.
                    raise QueryError('This workbook already has a sheet called "%s". Give the query another name.'
                                     % query["sheet"])
                state["queries"][at:at + 1] = [query]
                actions, lines = self.run_all(state, query["name"])
            elif op == "run":
                actions, lines = self.run_all(state, arg("name") or None)
            elif op == "remove":
                before = len(state["queries"])
                state["queries"] = [q for q in state["queries"] if q["name"].lower() != arg("name").lower()]
                if len(state["queries"]) == before:
                    raise QueryError('There is no query called "%s".' % arg("name"))
                actions, lines = [emit(state)], ['Removed query "%s". Its output sheet is still in the workbook.'
                                                 % arg("name")]
            elif op == "setParam":
                name = arg("name").strip("{} ").lower()
                if not re.fullmatch(r"\w+", name):
                    raise QueryError("A parameter name is one word, e.g. from or year.")
                if arg("value") == "":
                    state["params"].pop(name, None)
                else:
                    state["params"][name] = arg("value")[:500]
                actions, lines = [emit(state)], ["Parameter {%s} saved. Use Refresh data to apply it." % name]
            else:
                return {"ok": False, "error": "Unknown data request"}
            return {"ok": True, "state": public(state), "actions": actions, "reply": "\n".join(lines)}
        except (QueryError, LedgerError, reports.ReportError) as err:
            return {"ok": False, "error": str(err)}
