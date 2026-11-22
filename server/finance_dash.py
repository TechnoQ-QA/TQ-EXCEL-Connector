"""TQ Excel Connector - the finance dashboards (profit and loss, financial position, budget against actual).

Called by dashboard.build with the sheet being filled and the cells holding the month, last year's month and
the starts of both fiscal years. Every number is a live TQ formula over the sheet's filters (B3:B7).
"""

from dashboard import DATA, _bal


def build(s, kind, periods):
    M, LY, FYS, LYS = periods
    pct, ratio, days = "0.0%", "0.00", "0"
    points, plus = "+0.0%;-0.0%;0.0%", "+0.00;-0.00;0.00"

    def key(title, head, rows):
        """A key-figure block; returns cell(row name, column index) for the tiles and charts to read."""
        first, _ = s.block(title, [""] + head, rows)
        names = [r[0] for r in rows]
        return lambda name, col: s.ref(first + 1 + names.index(name), DATA + col)

    # the months of the fiscal year, shown up to the chosen month (a gap after it)
    months = ['=TQ.PERIOD($B$8,"FYS%s")' % ("+%d" % n if n else "") for n in range(12)]
    state = {"r0": None, "so_far": None}

    def series(title, head, columns, upto=None):
        """columns: functions (n, period cell) -> formula; upto[i] False = that column runs the whole year."""
        r = s.data_row
        if state["r0"] is None:
            state["r0"] = r + 2
            s.put(r, DATA + 4, "Months so far")
            s.put(r, DATA + 5, "=MATCH($B$8,%s:%s,0)" % (s.ref(r + 2, DATA), s.ref(r + 13, DATA)))
            state["so_far"] = s.ref(r, DATA + 5)
        rows = []
        for n in range(12):
            p = s.ref(state["r0"] + n, DATA)
            row = [months[n]]
            for i, col in enumerate(columns):
                f = col(n, p)
                whole = upto is not None and not upto[i]
                row.append("=IFERROR(%s,NA())" % f if whole
                           else "=IF(%d<=%s,IFERROR(%s,NA()),NA())" % (n + 1, state["so_far"], f))
            rows.append(row)
        first, last = s.block(title, ["Month"] + head, rows)
        s.text.append([first + 1, DATA, last, DATA])
        return first, last

    bal = _bal
    if kind == "pl":
        _profit_and_loss(s, key, series, bal, M, LY, FYS, LYS, pct, points)
    elif kind == "position":
        _position(s, key, series, bal, M, LY, FYS, LYS, ratio, days, plus)
    else:
        _budget(s, key, series, bal, M, FYS)


def _profit_and_loss(s, key, series, bal, M, LY, FYS, LYS, pct, points):
    def four(acc):
        return ["=IFERROR(%s,NA())" % bal(acc, M, M), "=IFERROR(%s,NA())" % bal(acc, LY, LY),
                "=IFERROR(%s,NA())" % bal(acc, FYS, M), "=IFERROR(%s,NA())" % bal(acc, LYS, LY)]
    k = key("Profit and loss", ["Month", "Month LY", "Year to date", "YTD LY"], [
        ["Revenue"] + four('"type:Income"'), ["Cost of sales"] + four('"type:COGS"'),
        ["Operating expenses"] + four('"type:Expense"'), ["Other income"] + four('"type:OthIncome"'),
        ["Other expenses"] + four('"type:OthExpense"'), ["Net income"] + four('"calc:NetIncome"')])

    def gm(c):
        return "(%s-%s)/%s" % (k("Revenue", c), k("Cost of sales", c), k("Revenue", c))

    def nm(c):
        return "%s/%s" % (k("Net income", c), k("Revenue", c))

    s.tile(0, "Revenue, year to date", k("Revenue", 3), "%s/ABS(%s)-1" % (k("Revenue", 3), k("Revenue", 4)),
           "vs last year")
    s.tile(1, "Gross margin, year to date", gm(3), "%s-%s" % (gm(3), gm(4)), "change in points vs last year", pct, points)
    s.tile(2, "Operating expenses, year to date", k("Operating expenses", 3),
           "%s/ABS(%s)-1" % (k("Operating expenses", 3), k("Operating expenses", 4)), "vs last year (up is worse)",
           invert=True)
    s.tile(3, "Net margin, year to date", nm(3), "%s-%s" % (nm(3), nm(4)), "change in points vs last year", pct, points)

    f, l = series("Revenue, costs and net income by month", ["Revenue", "Costs", "Net income"], [
        lambda n, p: bal('"type:Income"', p, p),
        lambda n, p: "%s-%s" % (bal('"type:Income"', p, p), bal('"calc:NetIncome"', p, p)),
        lambda n, p: bal('"calc:NetIncome"', p, p)])
    s.chart("ColumnClustered", "Revenue, costs and net income by month", f, l, DATA, DATA + 3, 0)

    def rev(n):
        return s.ref(f + 1 + n, DATA + 1)

    def net(n):
        return s.ref(f + 1 + n, DATA + 3)

    f2, l2 = series("Margins by month", ["Gross margin", "Net margin"], [
        lambda n, p: "(%s-%s)/%s" % (rev(n), bal('"type:COGS"', p, p), rev(n)),
        lambda n, p: "%s/%s" % (net(n), rev(n))])
    s.formats.append({"r0": f2 + 1, "c0": DATA + 1, "r1": l2, "c1": DATA + 2, "numberFormat": pct})
    s.chart("Line", "Gross and net margin by month", f2, l2, DATA, DATA + 2, 1)
    f3, l3 = s.block("Costs, year to date", ["Cost", "Amount"], [
        ["Cost of sales", "=%s" % k("Cost of sales", 3)], ["Operating expenses", "=%s" % k("Operating expenses", 3)],
        ["Other expenses", "=%s" % k("Other expenses", 3)]])
    s.chart("Pie", "Where the money went, year to date", f3, l3, DATA, DATA + 1, 2)
    f4, l4 = s.block("From revenue to net income, year to date", ["Step", "Amount"], [
        ["Revenue", "=%s" % k("Revenue", 3)], ["Cost of sales", "=-%s" % k("Cost of sales", 3)],
        ["Operating expenses", "=-%s" % k("Operating expenses", 3)], ["Other income", "=%s" % k("Other income", 3)],
        ["Other expenses", "=-%s" % k("Other expenses", 3)], ["Net income", "=%s" % k("Net income", 3)]])
    s.chart("ColumnClustered", "From revenue to net income, year to date", f4, l4, DATA, DATA + 1, 3)


CURRENT_ASSETS = '"type:Bank,AcctRec,UnbilledRec,OthCurrAsset,DeferExpense"'
CURRENT_LIABILITIES = '"type:AcctPay,CredCard,OthCurrLiab,DeferRevenue"'


def _position(s, key, series, bal, M, LY, FYS, LYS, ratio, days, plus):
    def two(acc):
        return ["=IFERROR(%s,NA())" % bal(acc, M, M), "=IFERROR(%s,NA())" % bal(acc, LY, LY)]
    k = key("Financial position", ["Month end", "A year earlier"], [
        ["Bank and cash"] + two('"type:Bank"'), ["Receivables"] + two('"type:AcctRec,UnbilledRec"'),
        ["Other current assets"] + two('"type:OthCurrAsset,DeferExpense"'), ["Fixed assets"] + two('"type:FixedAsset"'),
        ["Other assets"] + two('"type:OthAsset"'), ["Payables"] + two('"type:AcctPay"'),
        ["Other current liabilities"] + two('"type:CredCard,OthCurrLiab,DeferRevenue"'),
        ["Long-term liabilities"] + two('"type:LongTermLiab"'),
        ["Sales, year to date", "=IFERROR(%s,NA())" % bal('"type:Income"', FYS, M),
         "=IFERROR(%s,NA())" % bal('"type:Income"', LYS, LY)],
        ["Cost of sales, year to date", "=IFERROR(%s,NA())" % bal('"type:COGS"', FYS, M),
         "=IFERROR(%s,NA())" % bal('"type:COGS"', LYS, LY)],
        ["Days, year to date", "=IFERROR(%s,NA())" % bal('"calc:Days"', FYS, M),
         "=IFERROR(%s,NA())" % bal('"calc:Days"', LYS, LY)]])

    def ca(c):
        return "(%s+%s+%s)" % (k("Bank and cash", c), k("Receivables", c), k("Other current assets", c))

    def cl(c):
        return "(%s+%s)" % (k("Payables", c), k("Other current liabilities", c))

    def ta(c):
        return "(%s+%s+%s)" % (ca(c), k("Fixed assets", c), k("Other assets", c))

    def tl(c):
        return "(%s+%s)" % (cl(c), k("Long-term liabilities", c))

    def eq(c):   # equity as assets less liabilities, so it always balances
        return "(%s-%s)" % (ta(c), tl(c))

    def turn(balance, flow, c):   # days of sales (or cost of sales) the balance stands for
        return '=IFERROR(%s/%s*%s,"")' % (k(balance, c), k(flow, c), k("Days, year to date", c))

    names = ["Working capital", "Current ratio", "Quick ratio", "Debt to equity", "Debtor days", "Creditor days"]
    f0, l0 = s.block("Ratios", ["", "Month end", "A year earlier"], [
        ["Working capital", "=%s-%s" % (ca(1), cl(1)), "=%s-%s" % (ca(2), cl(2))],
        ["Current ratio", '=IFERROR(%s/%s,"")' % (ca(1), cl(1)), '=IFERROR(%s/%s,"")' % (ca(2), cl(2))],
        ["Quick ratio", '=IFERROR((%s+%s)/%s,"")' % (k("Bank and cash", 1), k("Receivables", 1), cl(1)),
         '=IFERROR((%s+%s)/%s,"")' % (k("Bank and cash", 2), k("Receivables", 2), cl(2))],
        ["Debt to equity", '=IFERROR(%s/%s,"")' % (tl(1), eq(1)), '=IFERROR(%s/%s,"")' % (tl(2), eq(2))],
        ["Debtor days", turn("Receivables", "Sales, year to date", 1), turn("Receivables", "Sales, year to date", 2)],
        ["Creditor days", turn("Payables", "Cost of sales, year to date", 1),
         turn("Payables", "Cost of sales, year to date", 2)]])
    s.formats.append({"r0": f0 + 2, "c0": DATA + 1, "r1": l0, "c1": DATA + 2, "numberFormat": ratio})

    def r(name, c):
        return s.ref(f0 + 1 + names.index(name), DATA + c)

    s.tile(0, "Working capital at month end", r("Working capital", 1),
           "%s/ABS(%s)-1" % (r("Working capital", 1), r("Working capital", 2)), "vs a year earlier")
    s.tile(1, "Current ratio", r("Current ratio", 1), "%s-%s" % (r("Current ratio", 1), r("Current ratio", 2)),
           "vs a year earlier; quick ratio in the data block", ratio, plus)
    s.tile(2, "Debt to equity", r("Debt to equity", 1), "%s-%s" % (r("Debt to equity", 1), r("Debt to equity", 2)),
           "vs a year earlier (up is worse)", ratio, plus, invert=True)
    s.tile(3, "Debtor days", r("Debtor days", 1), "%s-%s" % (r("Debtor days", 1), r("Debtor days", 2)),
           "vs a year earlier (up is worse)", days, "+0;-0;0", invert=True)

    f, l = series("Current assets and liabilities by month", ["Current assets", "Current liabilities", "Working capital"], [
        lambda n, p: bal(CURRENT_ASSETS, p, p), lambda n, p: bal(CURRENT_LIABILITIES, p, p),
        lambda n, p: "%s-%s" % (bal(CURRENT_ASSETS, p, p), bal(CURRENT_LIABILITIES, p, p))])
    s.chart("Line", "Working capital by month", f, l, DATA, DATA + 3, 0)
    f2, l2 = s.block("What the company owns, at month end", ["Asset", "Amount"], [
        [n, "=%s" % k(n, 1)] for n in ("Bank and cash", "Receivables", "Other current assets", "Fixed assets",
                                       "Other assets")])
    s.chart("Pie", "What the company owns, at month end", f2, l2, DATA, DATA + 1, 1)
    f3, l3 = s.block("Assets, liabilities and equity", ["", "Month end", "A year earlier"], [
        ["Total assets", "=%s" % ta(1), "=%s" % ta(2)], ["Total liabilities", "=%s" % tl(1), "=%s" % tl(2)],
        ["Equity", "=%s" % eq(1), "=%s" % eq(2)]])
    s.chart("ColumnClustered", "Assets, liabilities and equity: now and a year earlier", f3, l3, DATA, DATA + 2, 2)
    f4, l4 = series("Receivables and payables by month", ["Receivables", "Payables"], [
        lambda n, p: bal('"type:AcctRec,UnbilledRec"', p, p), lambda n, p: bal('"type:AcctPay"', p, p)])
    s.chart("Line", "Receivables and payables by month", f4, l4, DATA, DATA + 2, 3)


def _budget(s, key, series, bal, M, FYS):
    def four(acc):
        return ["=IFERROR(%s,NA())" % bal('"%s"' % acc, M, M), "=IFERROR(%s,NA())" % bal('"budget:%s"' % acc, M, M),
                "=IFERROR(%s,NA())" % bal('"%s"' % acc, FYS, M), "=IFERROR(%s,NA())" % bal('"budget:%s"' % acc, FYS, M)]
    k = key("Actual and budget", ["Month", "Month budget", "Year to date", "YTD budget"], [
        ["Revenue"] + four("type:Income"), ["Cost of sales"] + four("type:COGS"),
        ["Operating expenses"] + four("type:Expense"), ["Net income"] + four("calc:NetIncome")])

    def cost(c):
        return "(%s+%s)" % (k("Cost of sales", c), k("Operating expenses", c))

    def vs(a, b):
        return "%s/ABS(%s)-1" % (a, b)

    s.tile(0, "Revenue, year to date", k("Revenue", 3), vs(k("Revenue", 3), k("Revenue", 4)), "vs budget")
    s.tile(1, "Costs, year to date", cost(3), vs(cost(3), cost(4)), "vs budget (up is worse)", invert=True)
    s.tile(2, "Net income, year to date", k("Net income", 3), vs(k("Net income", 3), k("Net income", 4)), "vs budget")
    s.tile(3, "Revenue, this month", k("Revenue", 1), vs(k("Revenue", 1), k("Revenue", 2)), "vs budget")
    f, l = series("Revenue by month: actual and budget", ["Actual", "Budget"], [
        lambda n, p: bal('"type:Income"', p, p), lambda n, p: bal('"budget:type:Income"', p, p)], upto=[True, False])
    s.chart("ColumnClustered", "Revenue by month: actual and budget", f, l, DATA, DATA + 2, 0)
    f2, l2 = series("Net income by month: actual and budget", ["Actual", "Budget"], [
        lambda n, p: bal('"calc:NetIncome"', p, p), lambda n, p: bal('"budget:calc:NetIncome"', p, p)],
        upto=[True, False])
    s.chart("Line", "Net income by month: actual and budget", f2, l2, DATA, DATA + 2, 1)
    f3, l3 = s.block("Year to date: actual and budget", ["", "Actual", "Budget"], [
        [n, "=%s" % k(n, 3), "=%s" % k(n, 4)] for n in ("Revenue", "Cost of sales", "Operating expenses", "Net income")])
    s.chart("BarClustered", "Year to date: actual and budget", f3, l3, DATA, DATA + 2, 2)

    def cum(col):
        return lambda n, p: "SUM(%s:%s)" % (s.ref(f + 1, DATA + col), s.ref(f + 1 + n, DATA + col))
    f4, l4 = series("Revenue so far this year: actual and budget", ["Actual", "Budget"], [cum(1), cum(2)],
                    upto=[True, False])
    s.chart("Line", "Revenue so far this year: actual and budget", f4, l4, DATA, DATA + 2, 3)
