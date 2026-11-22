"""TQ Excel Connector - dashboards: a sheet of tiles and native Excel charts over live TQ formulas.

    build(ledger, kind, month, filters, title, taken)  -> (write action, info)

kind: "executive", "sales" or "cash". The sheet has the same filter cells as a report (B3 other filters,
B4:B7 segments, B8 the month), so the month picker, the dropdowns and Refresh work the same way. A data
block to the right holds the formulas the tiles and charts read; the charts are Excel's own, so they
print, go into the PDF and survive Copy as values. Pick another month in B8 and everything moves.

Which locations and expense accounts the bar and pie charts list is decided when the dashboard is built
(those with a balance in the year to date); the numbers stay live.
"""

import reports
from ledger import LedgerError, SEGMENT_KINDS
from reports import GREY, KIND_LABEL, ReportError, col_letter

KINDS = {
    "executive": "Executive dashboard",
    "sales": "Sales dashboard",
    "cash": "Cash dashboard",
    "pl": "Profit and loss dashboard",
    "position": "Financial position",
    "budget": "Budget vs actual dashboard",
}
FINANCE = ("pl", "position", "budget")
SEG = "$B$4,$B$5,$B$6,$B$7,$B$3"   # the sheet's filters, in TQ.BALANCE's order
DATA = 9                            # column J: the data block starts here
WIDTH = DATA + 7
CHART_AT = [("A14", "D30"), ("E14", "H30"), ("A32", "D48"), ("E32", "H48")]
TILE_COLS = [0, 2, 4, 6]
TOP = 6                             # bars or pie slices before "Other"
NAVY, MAROON = "#1F3A5F", "#6B1F24"


def _q(text):
    return '"%s"' % str(text).replace('"', '""')


def _bal(account, start, end, seg=SEG):
    return "TQ.BALANCE(%s,%s,%s,%s)" % (account, start, end, seg)


class Sheet:
    """The cells, formats and charts of one dashboard, filled block by block."""

    def __init__(self):
        self.cells, self.formats, self.text, self.charts, self.conditional = {}, [], [], [], []
        self.data_row = 2   # next free row (0-based) in the data block

    def put(self, r, c, value, **style):
        self.cells[(r, c)] = value
        if style:
            self.formats.append(dict(r0=r, c0=c, r1=r, c1=c, **style))

    def block(self, title, head, rows):
        """A table in the data block: a title line, a header row, then rows. Returns (first, last) 0-based rows
        of the header and the last data row, so a chart can take header + rows."""
        r = self.data_row
        self.put(r, DATA, title, bold=True, color=NAVY)
        for j, h in enumerate(head):
            self.put(r + 1, DATA + j, h, bold=True)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                self.put(r + 2 + i, DATA + j, v)
        self.text.append([r + 1, DATA, r + 1, DATA + len(head) - 1])
        first, last = r + 1, r + 1 + len(rows)
        self.formats.append({"r0": first + 1, "c0": DATA + 1, "r1": last, "c1": DATA + len(head) - 1,
                             "numberFormat": reports.NUMBER_FORMAT})
        self.data_row = last + 3
        return first, last

    def ref(self, r, c):
        return "$%s$%d" % (col_letter(c), r + 1)

    def chart(self, kind, title, first, last, c0, c1, slot):
        tl, br = CHART_AT[slot]
        self.charts.append({"type": kind, "title": title, "slot": slot, "topLeft": tl, "bottomRight": br,
                            "range": "%s%d:%s%d" % (col_letter(c0), first + 1, col_letter(c1), last + 1)})

    def tile(self, n, label, value, change=None, note=None, fmt="#,##0;(#,##0)", change_fmt="+0.0%;-0.0%;0.0%",
             invert=False):
        """A tile: label, the number (big), and its change (green when better, red when worse; invert for costs,
        debtor days and the like, where up is worse)."""
        c = TILE_COLS[n]
        self.put(9, c, label, color=GREY, bold=True)
        self.put(10, c, '=IFERROR(%s,"")' % value, bold=True, size=16, color=NAVY, numberFormat=fmt, merge=True,
                 align="left")
        self.formats[-1]["c1"] = c + 1
        if change:
            self.put(11, c, '=IFERROR(%s,"")' % change, numberFormat=change_fmt, merge=True, align="left")
            self.formats[-1]["c1"] = c + 1
            cell = "%s12" % col_letter(c)
            good, bad = ("<0", ">0") if invert else (">0", "<0")
            for test, font, fill in ((good, "#1A7F4B", "#E6F4EA"), (bad, "#B3261E", "#FCE8E6")):
                self.conditional.append({"r0": 11, "c0": c, "r1": 11, "c1": c, "font": font, "fill": fill,
                                         "formula": "=AND(ISNUMBER(%s),%s%s)" % (cell, cell, test)})
        if note:
            self.put(12, c, note, color=GREY, italic=True)
        self.formats.append({"r0": 9, "c0": c, "r1": 12, "c1": c + 1, "fill": "#F3F6FA"})


def _top(ledger, requests, names, limit=TOP):
    """The names whose request has the largest balance (absolute), largest first; [] when NetSuite says no."""
    if not requests:
        return []
    try:
        results = ledger.balances(requests)
    except LedgerError:
        return []
    scored = [(abs(r.get("value") or 0), n) for n, r in zip(names, results) if "value" in r]
    return [n for v, n in sorted(scored, reverse=True) if v >= 0.005][:limit]


def _segment_split(ledger, start, end, filters):
    """(kind, values): the first segment (location, department, class, subsidiary) with sales in the period."""
    for kind in ("location", "department", "class", "subsidiary"):
        try:
            values = ledger.segments(kind)
        except LedgerError:
            continue
        if not values or len(values) > 400:
            continue
        base = {k: v for k, v in filters.items() if k != kind}
        found = _top(ledger, [dict(base, account="type:Income", fromPeriod=start, toPeriod=end, **{kind: v})
                              for v in values], values, 8)
        if found:
            return kind, found
    return None, []


def build(ledger, kind, month="", filters=None, title="", taken=()):
    kind = kind if kind in KINDS else "executive"
    filters = {k: v for k, v in (filters or {}).items() if v}
    periods = ledger.periods()
    try:
        month = ledger.period_name(month) if month else ledger.period_name("CM-1")
        start = ledger.shift_period(month, "FYS")
    except LedgerError as err:
        raise ReportError(str(err))
    title = reports.sheet_name(title or KINDS[kind])
    name, n = title, 2
    while name.lower() in [t.lower() for t in taken]:   # never write over a sheet that is already there
        name, n = reports.sheet_name("%s %d" % (title, n)), n + 1
    s = Sheet()
    s.put(0, 0, name, bold=True, size=16, color=NAVY)
    s.put(1, 0, "Live from NetSuite: pick another month in B8 (or change the filters in B3:B7), then Refresh. "
                "The tiles and charts read the data block on the right.", color=GREY)
    s.put(2, 0, KIND_LABEL["filters"])
    s.put(2, 1, filters.get("filters", ""))
    s.put(2, 2, reports.FILTER_NOTE, color=GREY, italic=True)
    for i, k in enumerate(SEGMENT_KINDS):
        s.put(3 + i, 0, KIND_LABEL[k])
        s.put(3 + i, 1, filters.get(k, ""))
    s.put(7, 0, "Month")
    s.put(7, 1, month, bold=True, fill="#FFF8E1")
    s.text += [[2, 1, 7, 1]]

    # The periods every formula uses: the month, last year's month, and the starts of both years.
    s.put(0, DATA, "Data for the tiles and charts (worked out from B3:B8 - do not type here)", bold=True, color=MAROON)
    s.put(1, DATA, "Periods")
    s.put(1, DATA + 1, "=$B$8")
    s.put(1, DATA + 2, '=TQ.PERIOD($B$8,"LY")')
    s.put(1, DATA + 3, '=TQ.PERIOD($B$8,"FYS")')
    s.put(1, DATA + 4, '=IFERROR(TQ.PERIOD($B$8,"FYS-12"),"")')
    M, LY, FYS, LYS = (s.ref(1, DATA + j) for j in (1, 2, 3, 4))
    if kind in FINANCE:
        import finance_dash   # imported here: it uses this module's building blocks
        finance_dash.build(s, kind, (M, LY, FYS, LYS))
        return _finish(s, ledger, periods, filters, name, kind, month)

    def four(account):
        """[month, month last year, year to date, last year to date] for an account or total."""
        return ["=IFERROR(%s,NA())" % _bal(account, M, M), "=IFERROR(%s,NA())" % _bal(account, LY, LY),
                "=IFERROR(%s,NA())" % _bal(account, FYS, M), "=IFERROR(%s,NA())" % _bal(account, LYS, LY)]

    head = ["", "Month", "Month LY", "Year to date", "YTD LY"]
    if kind == "cash":
        kpi = [["Cash and bank"] + four('"type:Bank"'), ["Receivables"] + four('"type:AcctRec,UnbilledRec"'),
               ["Payables"] + four('"type:AcctPay"'), ["Net income"] + four('"calc:NetIncome"')]
    else:
        kpi = [["Sales"] + four('"type:Income"'), ["Cost of sales"] + four('"type:COGS"'),
               ["Net income"] + four('"calc:NetIncome"'), ["Cash and bank"] + four('"type:Bank"')]
    first, last = s.block("Key figures", head, kpi)
    kr = lambda row, col: s.ref(first + 1 + row, DATA + col)   # noqa: E731  key figure cell
    gp_row = None
    if kind != "cash":
        gp = ["Gross profit"] + ["=%s-%s" % (kr(0, j), kr(1, j)) for j in (1, 2, 3, 4)]
        for j, v in enumerate(gp):
            s.put(last + 1, DATA + j, v)
        s.formats.append({"r0": last + 1, "c0": DATA + 1, "r1": last + 1, "c1": DATA + 4, "numberFormat": reports.NUMBER_FORMAT})
        gp_row = last + 1
        s.data_row = last + 4
    gpr = lambda col: s.ref(gp_row, DATA + col)   # noqa: E731
    change = lambda now, before: "%s/ABS(%s)-1" % (now, before)   # noqa: E731

    if kind == "executive":
        s.tile(0, "Sales, year to date", kr(0, 3), change(kr(0, 3), kr(0, 4)), "vs last year")
        s.tile(1, "Gross profit, year to date", gpr(3), change(gpr(3), gpr(4)), "vs last year")
        s.tile(2, "Net income, year to date", kr(2, 3), change(kr(2, 3), kr(2, 4)), "vs last year")
        s.tile(3, "Cash and bank at month end", kr(3, 1), change(kr(3, 1), kr(3, 2)), "vs a year earlier")
    elif kind == "sales":
        s.tile(0, "Sales, year to date", kr(0, 3), change(kr(0, 3), kr(0, 4)), "vs last year")
        s.tile(1, "Sales, this month", kr(0, 1), change(kr(0, 1), kr(0, 2)), "vs the same month last year")
        s.tile(2, "Gross profit, year to date", gpr(3), change(gpr(3), gpr(4)), "vs last year")
        s.put(9, 6, "Gross margin, year to date", color=GREY, bold=True)
        s.put(10, 6, '=IFERROR(%s/%s,"")' % (gpr(3), kr(0, 3)), bold=True, size=16, color=NAVY, numberFormat="0.0%",
              merge=True, align="left")
        s.formats[-1]["c1"] = 7
        s.put(11, 6, '=IFERROR(%s/%s-%s/%s,"")' % (gpr(3), kr(0, 3), gpr(4), kr(0, 4)), numberFormat='+0.0%;-0.0%;0.0%')
        s.put(12, 6, "change in points vs last year", color=GREY, italic=True)
        s.formats.append({"r0": 9, "c0": 6, "r1": 12, "c1": 7, "fill": "#F3F6FA"})
    else:
        s.tile(0, "Cash and bank at month end", kr(0, 1), change(kr(0, 1), kr(0, 2)), "vs a year earlier")
        s.tile(1, "Receivables at month end", kr(1, 1), change(kr(1, 1), kr(1, 2)), "vs a year earlier")
        s.tile(2, "Payables at month end", kr(2, 1), change(kr(2, 1), kr(2, 2)), "vs a year earlier")
        s.tile(3, "Net income, year to date", kr(3, 3), change(kr(3, 3), kr(3, 4)), "vs last year")

    # Month by month through the fiscal year: shown up to the chosen month, empty (a gap) after it.
    months = ['=TQ.PERIOD($B$8,"FYS%s")' % ("+%d" % n if n else "") for n in range(12)]
    r0 = s.data_row + 2
    so_far = s.ref(s.data_row, DATA + 5)
    s.put(s.data_row, DATA + 4, "Months so far")
    s.put(s.data_row, DATA + 5, "=MATCH($B$8,%s:%s,0)" % (s.ref(r0, DATA), s.ref(r0 + 11, DATA)))

    def by_month(n, account, last_year=False):
        p = s.ref(r0 + n, DATA)
        if last_year:   # last year: every month, for comparison
            p = "TQ.PERIOD(%s,-12)" % p
            return "=IFERROR(%s,NA())" % _bal(account, p, p)
        acc = account
        return "=IF(%d<=%s,IFERROR(%s,NA()),NA())" % (n + 1, so_far, _bal(acc, p, p)) if not last_year else \
            "=IFERROR(%s,NA())" % _bal(acc, p, p)

    def series(title, head, columns):
        rows = [[months[n]] + [col(n) for col in columns] for n in range(12)]
        first, last = s.block(title, ["Month"] + head, rows)
        s.text.append([first + 1, DATA, last, DATA])
        return first, last

    slot = 0
    if kind in ("executive", "sales"):
        f, l = series("Sales by month", ["This year", "Last year"],
                      [lambda n: by_month(n, '"type:Income"'), lambda n: by_month(n, '"type:Income"', True)])
        s.chart("Line", "Sales by month: this year and last", f, l, DATA, DATA + 2, slot)
        slot += 1
    if kind == "sales":
        sales_first = f
        income = lambda n: s.ref(sales_first + 1 + n, DATA + 1)   # noqa: E731
        f2, l2 = series("Gross margin by month", ["Gross margin"],
                        [lambda n: '=IFERROR((%s-%s)/%s,NA())' % (income(n), _bal('"type:COGS"', s.ref(r0 + n, DATA),
                                                                                  s.ref(r0 + n, DATA)), income(n))])
        s.formats.append({"r0": f2 + 1, "c0": DATA + 1, "r1": l2, "c1": DATA + 1, "numberFormat": "0.0%"})
        s.chart("Line", "Gross margin by month", f2, l2, DATA, DATA + 1, slot)
        slot += 1
    if kind == "cash":
        f, l = series("Cash and bank by month", ["Cash and bank"], [lambda n: by_month(n, '"type:Bank"')])
        s.chart("Area", "Cash and bank at each month end", f, l, DATA, DATA + 1, slot)
        slot += 1
        f, l = series("Receivables and payables by month", ["Receivables", "Payables"],
                      [lambda n: by_month(n, '"type:AcctRec,UnbilledRec"'), lambda n: by_month(n, '"type:AcctPay"')])
        s.chart("Line", "Receivables and payables at each month end", f, l, DATA, DATA + 2, slot)
        slot += 1
        f, l = series("Cash movement by month", ["Movement"], [lambda n: by_month(n, '"move:type:Bank"')])
        s.chart("ColumnClustered", "Cash in less cash out, by month", f, l, DATA, DATA + 1, slot)
        slot += 1
        banks = [a for a in ledger.accounts() if a["type"] == "Bank"]
        top = _top(ledger, [dict(filters, account=a["number"], fromPeriod=month) for a in banks],
                   [a["number"] for a in banks])
        if top:
            names = {a["number"]: a["name"].split(" : ")[-1] for a in banks}
            f, l = s.block("Bank accounts at month end", ["Account", "Balance"],
                           [[names[t], "=IFERROR(%s,NA())" % _bal(_q(t), M, M)] for t in top])
            s.chart("BarClustered", "Bank accounts at month end", f, l, DATA, DATA + 1, slot)
            slot += 1

    if kind in ("executive", "sales"):
        split, values = _segment_split(ledger, start, month, filters)
        if split:
            arg = {k: "$B$%d" % reports.PARAM_ROW[k] for k in SEGMENT_KINDS}
            rows = []
            for v in values:
                parts = dict(arg, **{split: _q(v)})
                seg = ",".join(parts[k] for k in SEGMENT_KINDS) + ",$B$3"
                rows.append([v.split(" : ")[-1], "=IFERROR(%s,NA())" % _bal('"type:Income"', FYS, M, seg)])
            f, l = s.block("Sales by %s, year to date" % split, [KIND_LABEL[split], "Sales"], rows)
            s.text.append([f + 1, DATA, l, DATA])
            s.chart("BarClustered", "Sales by %s, year to date" % split, f, l, DATA, DATA + 1, slot)
            slot += 1
    if kind == "executive":
        costs = [a for a in ledger.accounts() if a["type"] in ("Expense", "COGS")]
        top = _top(ledger, [dict(filters, account=a["number"], fromPeriod=start, toPeriod=month) for a in costs],
                   [a["number"] for a in costs])
        if top:
            names = {a["number"]: a["name"].split(" : ")[-1] for a in costs}
            rows = [[names[t], "=IFERROR(%s,NA())" % _bal(_q(t), FYS, M)] for t in top]
            rows.append(["Other costs", "=IFERROR(%s-%s,NA())" % (_bal('"type:Expense,COGS"', FYS, M),
                                                              _bal('"%s"' % ", ".join(top), FYS, M))])
            f, l = s.block("Costs, year to date", ["Cost", "Amount"], rows)
            s.text.append([f + 1, DATA, l, DATA])
            s.chart("Pie", "Where the money went, year to date", f, l, DATA, DATA + 1, slot)
            slot += 1
    if slot < 4 and kind == "executive":
        f, l = series("Cash and bank by month", ["Cash and bank"], [lambda n: by_month(n, '"type:Bank"')])
        s.chart("Area", "Cash and bank at each month end", f, l, DATA, DATA + 1, slot)
        slot += 1

    return _finish(s, ledger, periods, filters, name, kind, month)


def _finish(s, ledger, periods, filters, name, kind, month):
    height = max(r for r, _ in s.cells) + 1
    rows = [[""] * WIDTH for _ in range(max(height, 49))]
    for (r, c), v in s.cells.items():
        rows[r][c] = v
    s.formats += [{"r0": 0, "c0": DATA, "r1": len(rows) - 1, "c1": WIDTH - 1, "color": GREY},
                  {"r0": 2, "c0": 2, "r1": 6, "c1": 2, "color": GREY, "italic": True}]
    action = {
        "type": "write", "target": "sheet", "sheetName": name, "startCell": "", "rows": rows, "activate": True,
        "textRanges": s.text, "formats": s.formats, "charts": s.charts, "conditional": s.conditional,
        "widths": [{"c0": 0, "c1": 7, "width": 105}, {"c0": 8, "c1": 8, "width": 20},
                   {"c0": DATA, "c1": DATA, "width": 150}, {"c0": DATA + 1, "c1": WIDTH - 1, "width": 90}],
        "gridlines": False,
    }
    action["dropdowns"] = []
    month_list = [p["name"] for p in periods if 20 < (p["end"] - p["start"]).days < 40]
    reports.add_filter_lists(action, ledger, [], filters.get("filters", ""), [("Month list", month_list, [(7, 1)])])
    return action, {"sheet": name, "kind": kind, "month": month, "charts": len(s.charts)}
