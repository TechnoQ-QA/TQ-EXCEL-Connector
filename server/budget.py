"""TQ Excel Connector - budget sheets and the upload to NetSuite.

    sheet(ledger, year, category, fill)   a "TQ Budget" sheet: one row per profit and loss account, one column
                                          per month of the fiscal year, optionally filled with last year's actuals
    upload(ledger, rows, dry_run)         reads that sheet back and creates or updates NetSuite's budget lines

Amounts are in their natural sign, as NetSuite stores budgets: income and expenses are both positive.
A row may name its own subsidiary, department, class and location; blank cells take the sheet's defaults.
"""

import reports
from ledger import LedgerError, PL_SIGNS, SEGMENT_KINDS
from reports import GREY, ReportError, col_letter

SHEET_PREFIX = "TQ Budget"
HEAD = ["Account", "Name", "Subsidiary", "Department", "Class", "Location"]
FIRST_MONTH = len(HEAD)
HEADER_ROW = 9        # 0-based row of the column headings
CHUNK = 40            # lines per NetSuite call; the script stops early and says where if it runs short
MAX_LINES = 2000


def info(ledger):
    """Fiscal years (with their months), budget categories, and whether this user may upload."""
    data = ledger.call({"action": "budgetInfo"})
    years = [y for y in data.get("years") or [] if y.get("months")]
    if not years:
        raise ReportError("NetSuite returned no fiscal years with months.")
    return {"years": years, "categories": data.get("categories") or [], "canWrite": bool(data.get("canWrite")),
            "recordOk": data.get("recordOk", True), "recordError": data.get("recordError") or ""}


def _year(found, name):
    hit = next((y for y in found["years"] if y["name"].lower() == (name or "").strip().lower()), None)
    if not hit:
        raise ReportError('Fiscal year "%s" not found. NetSuite has: %s.' % (
            name, ", ".join(y["name"] for y in found["years"])))
    return hit


def sheet(ledger, year, category="", fill=""):
    """The write action for a budget sheet. fill: "lastyear" puts last year's actuals in as a starting point."""
    found = info(ledger)
    y = _year(found, year)
    if found["categories"] and category and category.lower() not in [c.lower() for c in found["categories"]]:
        raise ReportError('Budget category "%s" not found. NetSuite has: %s.' % (category, ", ".join(found["categories"])))
    category = category or (found["categories"][0] if found["categories"] else "")
    months = y["months"]
    pl = [a for a in ledger.accounts() if a["type"] in PL_SIGNS]
    accounts, amounts = pl, {}   # amounts: (account number, month position) -> last year's actual
    if fill == "lastyear" and _all_known(ledger, months):
        before = [ledger.shift_period(m, -12) for m in months]
        results = ledger.balances([{"account": a["number"], "fromPeriod": p} for a in pl for p in before])
        for n, r in enumerate(results):
            if isinstance(r.get("value"), (int, float)):
                amounts[(pl[n // len(before)]["number"], n % len(before))] = r["value"]
        used = [a for a in pl if any(abs(amounts.get((a["number"], m), 0)) >= 0.005 for m in range(len(before)))]
        accounts = used or pl   # only the accounts that had something last year
    width = FIRST_MONTH + len(months) + 1
    rows = [[SHEET_PREFIX + " - " + y["name"]],
            ["Budget category", category, "", "One of: %s" % (", ".join(found["categories"]) or "(Multiple Budgets is off)")],
            ["Year", y["name"], "", "A NetSuite fiscal year. The months below are its months, in order."],
            ["Subsidiary", "", "", "Defaults for rows that leave these blank. One value each (no lists, no 'not')."],
            ["Department", ""], ["Class", ""], ["Location", ""],
            ["Amounts in their natural sign: income and expenses both positive. Change the numbers, add rows for "
             "more accounts or segments, then press Check and Upload in the panel."],
            [],
            HEAD + months + ["Total"]]
    first = len(rows)
    for a in accounts:
        r = len(rows) + 1
        values = [round(amounts.get((a["number"], m), 0.0), 2) if amounts else "" for m in range(len(months))]
        rows.append([a["number"], a["name"], "", "", "", ""] + values
                    + ["=SUM(%s%d:%s%d)" % (col_letter(FIRST_MONTH), r, col_letter(FIRST_MONTH + len(months) - 1), r)])
    rows = [row + [""] * (width - len(row)) for row in rows]
    last = len(rows) - 1
    action = {
        "type": "write", "target": "sheet", "sheetName": reports.sheet_name("%s %s" % (SHEET_PREFIX, y["name"])),
        "startCell": "", "rows": rows, "activate": True,
        # Text, so Excel keeps "Jan 2027" as the period's name instead of turning it into a date.
        "textRanges": [[1, 1, 6, 1], [HEADER_ROW, 0, HEADER_ROW, width - 1], [first, 0, max(first, last), 5]],
        "formats": [
            {"r0": 0, "c0": 0, "r1": 0, "c1": width - 1, "bold": True, "size": 14},
            {"r0": 1, "c0": 1, "r1": 6, "c1": 1, "bold": True, "fill": "#FFF8E1"},
            {"r0": 1, "c0": 3, "r1": 7, "c1": width - 1, "color": GREY, "italic": True},
            {"r0": 7, "c0": 0, "r1": 7, "c1": 0, "color": GREY, "italic": True},
            {"r0": HEADER_ROW, "c0": 0, "r1": HEADER_ROW, "c1": width - 1, "bold": True, "fill": "#6B1F24",
             "color": "#FFFFFF"},
            {"r0": first, "c0": FIRST_MONTH, "r1": max(first, last), "c1": width - 1, "numberFormat": reports.NUMBER_FORMAT},
            {"r0": first, "c0": width - 1, "r1": max(first, last), "c1": width - 1, "bold": True},
        ],
        "freeze": {"rows": HEADER_ROW + 1, "cols": 2},
        "widths": [{"c0": 0, "c1": 0, "width": 80}, {"c0": 1, "c1": 1, "width": 220},
                   {"c0": 2, "c1": 5, "width": 110}, {"c0": FIRST_MONTH, "c1": width - 1, "width": 85}],
        "dropdowns": [],
    }
    if found["categories"]:
        action["dropdowns"].append({"r": 1, "c": 1, "formula": ",".join(found["categories"])[:250]})
    lists = []
    for n, kind in enumerate(SEGMENT_KINDS):
        names = reports._values(ledger, kind)
        if names and len(names) <= reports.MAX_LIST:
            cells = [(3 + n, 1), (first, 2 + n, first + len(accounts) + 200)]   # the defaults cell, the column
            lists.append(("%s list" % kind.title(), names, cells))
    _lists(action, lists)
    return action, {"year": y["name"], "category": category, "accounts": len(accounts), "months": len(months),
                    "filled": bool(amounts)}


def _all_known(ledger, months):
    try:
        for m in months:
            ledger.shift_period(m, -12)
        return True
    except LedgerError:
        return False


def _lists(action, lists):
    """Segment dropdowns: the lists go to the right of the sheet, as on reports."""
    rows = action["rows"]
    width = len(rows[0])
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
        for cell in cells:   # (row, col) or (row, col, last row)
            drop = {"r": cell[0], "c": cell[1], "free": True, "formula": "=$%s$2:$%s$%d" % (letter, letter, len(names) + 1)}
            if len(cell) > 2:
                drop["r1"] = cell[2]
            action["dropdowns"].append(drop)
    action["formats"].append({"r0": 0, "c0": width + 1, "r1": height - 1, "c1": width + len(lists), "color": GREY})


def _cell(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def parse(rows):
    """The budget sheet's cells -> {year, category, defaults, months, lines}. Raises ReportError naming the row."""
    grid = [list(r) if isinstance(r, list) else [] for r in rows or []]
    top = {}
    header = None
    for n, row in enumerate(grid):
        key = _cell(row[0]).lower() if row else ""
        if key == "account" and len(row) > 1 and _cell(row[1]).lower() == "name":
            header = n
            break
        if key:
            top[key] = _cell(row[1]) if len(row) > 1 else ""
    if header is None or not top.get("year"):
        raise ReportError("This is not a budget sheet made by TQ Connector (it needs Year and an Account / Name row). "
                          "Make one with Make budget sheet.")
    heads = [_cell(c) for c in grid[header]]
    end = heads.index("Total") if "Total" in heads else len(heads)
    months = [h for h in heads[FIRST_MONTH:end] if h]
    defaults = {k: top.get(k, "") for k in SEGMENT_KINDS}
    lines = []
    for n in range(header + 1, len(grid)):
        row = grid[n] + [""] * (end - len(grid[n]))
        account = _cell(row[0])
        if not account:
            continue
        amounts = []
        for j in range(len(months)):
            value = row[FIRST_MONTH + j]
            if value in ("", None):
                amounts.append(None)
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                amounts.append(round(float(value), 2))
            else:
                try:
                    amounts.append(round(float(str(value).replace(",", "")), 2))
                except ValueError:
                    raise ReportError('Row %d, %s: "%s" is not a number.' % (n + 1, months[j], value))
        line = {"account": account, "row": n + 1, "amounts": amounts}
        for k, kind in enumerate(SEGMENT_KINDS):
            line[kind] = _cell(row[2 + k]) or defaults[kind]
        lines.append(line)
    if not lines:
        raise ReportError("The budget sheet has no account rows.")
    if len(lines) > MAX_LINES:
        raise ReportError("The budget sheet has %d rows, more than %d in one upload." % (len(lines), MAX_LINES))
    return {"year": top["year"], "category": top.get("budget category", ""), "months": months, "lines": lines}


def handle(ledger, payload):
    """POST /api/budget: {op: info | sheet | check | upload, ...} -> {ok, ...}."""
    op = str(payload.get("op") or "")
    try:
        if op == "info":
            return dict({"ok": True}, **info(ledger))
        if op == "sheet":
            action, about = sheet(ledger, str(payload.get("year") or ""), str(payload.get("category") or ""),
                                  "lastyear" if payload.get("fill") else "")
            return {"ok": True, "actions": [action], "reply":
                    'Budget sheet "%s": %d accounts by %d months%s. Change the numbers, then press Check, then Upload.'
                    % (action["sheetName"], about["accounts"], about["months"],
                       ", filled with last year's actuals" if about["filled"] else "")}
        if op in ("check", "upload"):
            return upload(ledger, payload.get("rows"), dry_run=(op == "check"))
        return {"ok": False, "error": "Unknown budget request"}
    except (ReportError, LedgerError) as err:
        return {"ok": False, "error": str(err)}


def upload(ledger, rows, dry_run=True):
    """Checks (dry_run) or uploads the budget sheet. Returns {ok, reply, results}."""
    sheet_data = parse(rows)
    found = info(ledger)
    y = _year(found, sheet_data["year"])
    if [m.lower() for m in sheet_data["months"]] != [m.lower() for m in y["months"]]:
        raise ReportError("The month columns do not match fiscal year %s (%s ... %s). Make the sheet again."
                          % (y["name"], y["months"][0], y["months"][-1]))
    if not dry_run and not found["canWrite"]:
        raise ReportError("Your NetSuite role may not change budgets (it needs Set Up Budgets, or Administrator).")
    if not found["recordOk"]:
        raise ReportError("This NetSuite account does not let scripts write budgets: %s" % found["recordError"])
    lines = sheet_data["lines"]
    results = []
    at = 0
    while at < len(lines):
        chunk = lines[at:at + CHUNK]
        data = ledger.call({"action": "budgetSave", "year": y["name"], "category": sheet_data["category"],
                            "dryRun": dry_run,
                            "lines": [{k: line[k] for k in ("account", "amounts") + SEGMENT_KINDS} for line in chunk]})
        done = int(data.get("done") or 0)
        results += data.get("results") or []
        if done <= 0:
            raise ReportError("NetSuite stopped before saving anything. Try again in a minute.")
        at += done
    bad = [(line, r) for line, r in zip(lines, results) if not r.get("ok")]
    counts = {}
    for r in results:
        if r.get("ok"):
            counts[r["action"]] = counts.get(r["action"], 0) + 1
    words = ", ".join("%d %s" % (v, k) for k, v in sorted(counts.items())) or "nothing"
    lead = ("Checked %d line(s) for %s%s: %s would be done." if dry_run else "Uploaded %d line(s) for %s%s: %s.") % (
        len(lines), y["name"], " (%s)" % sheet_data["category"] if sheet_data["category"] else "", words)
    reply = [lead]
    if bad:
        reply.append("%d line(s) have a problem:" % len(bad))
        reply += ["Row %d (%s): %s" % (line["row"], line["account"], r.get("error")) for line, r in bad[:15]]
    elif dry_run:
        reply.append("Nothing was written yet. Press Upload to NetSuite to save it.")
    return {"ok": True, "reply": "\n".join(reply), "results": results, "problems": len(bad), "checked": dry_run}
