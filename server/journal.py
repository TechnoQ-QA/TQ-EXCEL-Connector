"""TQ Excel Connector - journal entry upload.

    sheet()                      an empty "TQ Journals" sheet with the columns to fill
    upload(ledger, rows, dry)    reads the sheet back; rows with the same Entry number become one journal entry

Check (dry) writes nothing: it says, per entry, whether debits equal credits, the period is open and every account,
name and segment exists. Upload then creates the entries and puts their numbers in the reply.
"""

import datetime

import reports
from ledger import LedgerError
from reports import GREY, ReportError, col_letter

SHEET_NAME = "TQ Journals"
HEAD = ["Entry", "Date", "Account", "Debit", "Credit", "Line memo", "Name", "Department", "Class", "Location",
        "Subsidiary", "Entry memo"]
HEADER_ROW = 4
CHUNK = 10
MAX_ENTRIES = 500


def sheet(ledger):
    width = len(HEAD)
    today = datetime.date.today().isoformat()
    rows = [["TQ Journals"],
            ["Rows with the same Entry number become one journal entry; its Date, Subsidiary and Entry memo are taken "
             "from its first row. Each line has a Debit or a Credit."],
            ["Accounts by number; Name, Department, Class, Location by their NetSuite names. Then Check, then Upload."],
            [],
            HEAD,
            ["1", today, "", "", "", "", "", "", "", "", "", ""],
            ["1", today, "", "", "", "", "", "", "", "", "", ""]]
    rows = [r + [""] * (width - len(r)) for r in rows]
    action = {
        "type": "write", "target": "sheet", "sheetName": SHEET_NAME, "startCell": "", "rows": rows, "activate": True,
        "textRanges": [[HEADER_ROW, 0, HEADER_ROW + 300, 2], [HEADER_ROW, 5, HEADER_ROW + 300, width - 1]],
        "formats": [
            {"r0": 0, "c0": 0, "r1": 0, "c1": width - 1, "bold": True, "size": 14},
            {"r0": 1, "c0": 0, "r1": 2, "c1": 0, "color": GREY, "italic": True},
            {"r0": HEADER_ROW, "c0": 0, "r1": HEADER_ROW, "c1": width - 1, "bold": True, "fill": "#6B1F24",
             "color": "#FFFFFF"},
            {"r0": HEADER_ROW + 1, "c0": 3, "r1": HEADER_ROW + 300, "c1": 4, "numberFormat": reports.NUMBER_FORMAT},
        ],
        "freeze": {"rows": HEADER_ROW + 1, "cols": 0},
        "widths": [{"c0": 0, "c1": 1, "width": 80}, {"c0": 2, "c1": 2, "width": 90}, {"c0": 3, "c1": 4, "width": 95},
                   {"c0": 5, "c1": 5, "width": 200}, {"c0": 6, "c1": 10, "width": 120}, {"c0": 11, "c1": 11, "width": 200}],
        "dropdowns": [],
    }
    lists = []
    for n, kind in enumerate(("department", "class", "location", "subsidiary")):
        names = reports._values(ledger, kind)
        if names and len(names) <= reports.MAX_LIST:
            lists.append((kind, names, 7 + n))
    if lists:
        for row in rows:
            row += [""] * (1 + len(lists))
        height = max(len(rows), 1 + max(len(names) for _, names, _ in lists))
        rows += [[""] * len(rows[0]) for _ in range(height - len(rows))]
        for i, (kind, names, column) in enumerate(lists):
            col = width + 1 + i
            rows[0][col] = "%s list" % kind.title()
            for j, name in enumerate(names):
                rows[1 + j][col] = name
            letter = col_letter(col)
            action["textRanges"].append([1, col, len(names), col])
            action["dropdowns"].append({"r": HEADER_ROW + 1, "r1": HEADER_ROW + 300, "c": column, "free": True,
                                        "formula": "=$%s$2:$%s$%d" % (letter, letter, len(names) + 1)})
    return action


def _text(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def _date(value, where):
    """An Excel date (a serial number), or text such as 2026-11-18 or 18/11/2026 -> 'YYYY-MM-DD'."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(value))).isoformat()
    text = _text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    raise ReportError('%s: "%s" is not a date. Use 2026-11-18.' % (where, text))


def _amount(value, where):
    if value in ("", None):
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(float(value), 2)
    try:
        return round(float(str(value).replace(",", "")), 2)
    except ValueError:
        raise ReportError('%s: "%s" is not a number.' % (where, value))


def parse(rows):
    """The journal sheet's cells -> [{key, row, date, subsidiary, memo, lines}] in sheet order."""
    grid = [list(r) if isinstance(r, list) else [] for r in rows or []]
    header = next((n for n, r in enumerate(grid) if [_text(c) for c in r[:3]] == HEAD[:3]), None)
    if header is None:
        raise ReportError("This is not a journal sheet made by TQ Connector (it needs Entry, Date, Account columns). "
                          "Make one with Make journal sheet.")
    entries, order = {}, []
    for n in range(header + 1, len(grid)):
        row = grid[n] + [""] * (len(HEAD) - len(grid[n]))
        if not any(_text(c) for c in row[:len(HEAD)]):
            continue
        where = "Row %d" % (n + 1)
        key = _text(row[0])
        if not key:
            raise ReportError("%s: the Entry number is missing." % where)
        if key not in entries:
            entries[key] = {"key": key, "row": n + 1, "date": _date(row[1], where), "subsidiary": _text(row[10]),
                            "memo": _text(row[11]), "lines": []}
            order.append(key)
        entries[key]["lines"].append({
            "account": _text(row[2]), "debit": _amount(row[3], where), "credit": _amount(row[4], where),
            "memo": _text(row[5]), "name": _text(row[6]), "department": _text(row[7]), "class": _text(row[8]),
            "location": _text(row[9])})
    if not order:
        raise ReportError("The journal sheet has no lines.")
    if len(order) > MAX_ENTRIES:
        raise ReportError("%d entries is more than %d in one upload." % (len(order), MAX_ENTRIES))
    return [entries[k] for k in order]


def upload(ledger, rows, dry_run=True):
    entries = parse(rows)
    results, at = [], 0
    while at < len(entries):
        chunk = entries[at:at + CHUNK]
        data = ledger.call({"action": "journalSave", "dryRun": dry_run,
                            "entries": [{k: e[k] for k in ("date", "subsidiary", "memo", "lines")} for e in chunk]})
        done = int(data.get("done") or 0)
        results += data.get("results") or []
        if done <= 0:
            raise ReportError("NetSuite stopped before saving anything. Try again in a minute.")
        at += done
        if dry_run and data.get("canWrite") is False:
            results.append({"note": "Your NetSuite role may check but not create journal entries."})
    bad = [(e, r) for e, r in zip(entries, results) if not r.get("ok")]
    good = [(e, r) for e, r in zip(entries, results) if r.get("ok")]
    if dry_run:
        lead = "Checked %d entr%s: %d ready%s." % (len(entries), "y" if len(entries) == 1 else "ies", len(good),
                                                  ", %d with a problem" % len(bad) if bad else "")
    else:
        lead = "Created %d journal entr%s%s." % (len(good), "y" if len(good) == 1 else "ies", ": " + ", ".join(
            "%s (entry %s)" % (r.get("number") or r.get("id"), e["key"]) for e, r in good[:20]) if good else "")
    reply = [lead] + ["Entry %s (row %d): %s" % (e["key"], e["row"], r.get("error")) for e, r in bad[:20]]
    reply += [r["note"] for r in results if r.get("note")]
    if dry_run and not bad:
        reply.append("Nothing was written yet. Press Upload to NetSuite to create them.")
    return {"ok": True, "reply": "\n".join(reply), "problems": len(bad), "results": results[:len(entries)]}


def handle(ledger, payload):
    op = str(payload.get("op") or "")
    try:
        if op == "sheet":
            return {"ok": True, "actions": [sheet(ledger)],
                    "reply": 'Sheet "%s" is ready: fill one row per line, then press Check.' % SHEET_NAME}
        if op in ("check", "upload"):
            return upload(ledger, payload.get("rows"), dry_run=(op == "check"))
        return {"ok": False, "error": "Unknown journal request"}
    except (ReportError, LedgerError) as err:
        return {"ok": False, "error": str(err)}
