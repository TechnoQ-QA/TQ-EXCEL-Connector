import base64
import datetime
import hashlib
import hmac
import json
import os
import re
import sys
import threading
import unittest
import urllib.parse
import time
import urllib.error
import urllib.request

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "server"))
import server  # noqa: E402

CFG = {
    "accountId": "1234567_SB1",
    "restletUrl": "https://1234567-sb1.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=12&deploy=1",
    "consumerKey": "ck",
    "consumerSecret": "cs",
    "tokenId": "tk",
    "tokenSecret": "ts",
}
TODAY = datetime.date(2026, 11, 16)


def trim(cells):
    """A row without the empty cells at its end (lists for the dropdowns sit further right)."""
    cells = list(cells)
    while cells and cells[-1] in ("", None):
        cells.pop()
    return cells


def new_ledger(**options):
    return server.ledger_module.Ledger(server.MockBackend().handle, today=TODAY, **options)


class OAuthTest(unittest.TestCase):
    def test_signature_matches_hand_built_base_string(self):
        header = server.oauth_header("POST", CFG["restletUrl"], CFG, nonce="abc", timestamp=1700000000)
        # Parameters sorted by name, URL query parameters included, per OAuth 1.0.
        params = ("deploy=1&oauth_consumer_key=ck&oauth_nonce=abc&oauth_signature_method=HMAC-SHA256"
                  "&oauth_timestamp=1700000000&oauth_token=tk&oauth_version=1.0&script=12")
        base = ("POST&https%3A%2F%2F1234567-sb1.restlets.api.netsuite.com%2Fapp%2Fsite%2Fhosting%2Frestlet.nl&"
                + params.replace("=", "%3D").replace("&", "%26"))
        expected = base64.b64encode(hmac.new(b"cs&ts", base.encode(), hashlib.sha256).digest()).decode()
        self.assertIn('oauth_signature="%s"' % server.pct(expected), header)
        self.assertTrue(header.startswith('OAuth realm="1234567_SB1", '))

    def test_config_rejects_non_netsuite_url(self):
        path = os.path.join(os.path.dirname(__file__), "_bad_config.json")
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(dict(CFG, restletUrl="https://example.com/restlet.nl"), handle)
        try:
            with self.assertRaises(ValueError):
                server.load_config(path)
        finally:
            os.remove(path)


class BusyNetSuiteTest(unittest.TestCase):
    """NetSuite refuses requests over its concurrency limit; they are tried again, one at a time."""

    BUSY = (400, '{"error" : {"code" : "SSS_REQUEST_LIMIT_EXCEEDED", "message" : "Request Limit Exceeded!"}}')

    def setUp(self):
        self.saved = (server.post_to_netsuite, server.RETRY_DELAYS)
        server.RETRY_DELAYS = (0, 0)
        self.calls = 0

    def tearDown(self):
        server.post_to_netsuite, server.RETRY_DELAYS = self.saved

    def answers(self, *replies):
        def post(cfg, body):
            self.calls += 1
            return replies[min(self.calls, len(replies)) - 1]
        server.post_to_netsuite = post

    def test_a_refused_request_is_tried_again(self):
        self.answers(self.BUSY, self.BUSY, (200, '{"ok": true, "account": "X"}'))
        self.assertEqual(server.call_netsuite(CFG, {"action": "ping"}), {"ok": True, "account": "X"})
        self.assertEqual(self.calls, 3)

    def test_it_gives_up_with_a_plain_message(self):
        self.answers(self.BUSY)
        out = server.call_netsuite(CFG, {"action": "ping"})
        self.assertIn("NetSuite is busy", out["error"])
        self.assertEqual(self.calls, 3)

    def test_other_errors_are_not_retried(self):
        self.answers((401, "INVALID_LOGIN"))
        self.assertIn("HTTP 401", server.call_netsuite(CFG, {"action": "ping"})["error"])
        self.assertEqual(self.calls, 1)


class MockTest(unittest.TestCase):
    def setUp(self):
        self.mock = server.MockBackend()

    def balance(self, **request):
        result = self.mock.handle({"action": "balances", "requests": [request]})["results"][0]
        return result.get("value", result.get("error"))

    def test_pl_range_is_sum_of_months(self):
        months = [self.balance(account="4000", fromPeriod=p) for p in ("Jan 2026", "Feb 2026", "Mar 2026")]
        total = self.balance(account="4000", fromPeriod="Jan 2026", toPeriod="Mar 2026")
        self.assertAlmostEqual(total, sum(months), places=2)

    def test_balance_sheet_is_cumulative(self):
        jan = self.balance(account="1000", fromPeriod="Jan 2026")
        feb = self.balance(account="1000", fromPeriod="Feb 2026")
        self.assertGreater(feb, jan)

    def test_errors_are_per_cell(self):
        out = self.mock.handle({"action": "balances", "requests": [
            {"account": "9999", "fromPeriod": "Jan 2026"},
            {"account": "4000", "fromPeriod": "Smarch 2026"},
            {"account": "4000", "fromPeriod": "Jan 2026"},
        ]})
        self.assertIn("not found", out["results"][0]["error"])
        self.assertIn("not found", out["results"][1]["error"])
        self.assertIn("value", out["results"][2])

    def test_drilldown_adds_up_to_balance(self):
        request = {"account": "6000", "fromPeriod": "Jan 2026", "toPeriod": "Jun 2026", "department": "Sales"}
        rows = self.mock.handle({"action": "transactions", "request": request})["rows"]
        self.assertAlmostEqual(sum(r["amount"] for r in rows), self.balance(**request), places=2)


class CareTest(unittest.TestCase):
    """Keeping the connector healthy and moving a setup between PCs (server/care.py)."""

    def setUp(self):
        import tempfile
        self.care = server.care
        self.dir = tempfile.mkdtemp()
        self._docs = self.care.documents_dir
        self.care.documents_dir = lambda: self.dir   # never write into the real Documents folder

    def tearDown(self):
        import shutil
        self.care.documents_dir = self._docs
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_filter_sets(self):
        prefs = self.care.Prefs(self.dir)
        self.assertEqual(prefs.handle({"op": "get"})["filterSets"], [])
        self.assertFalse(prefs.handle({"op": "saveSet", "name": "Empty", "cells": {}})["ok"])
        out = prefs.handle({"op": "saveSet", "name": "Riyadh hotels", "cells": {"location": "Riyadh*", "x": "no"}})
        self.assertEqual(out["filterSets"], [{"name": "Riyadh hotels", "cells": {"location": "Riyadh*"}}])
        prefs.handle({"op": "saveSet", "name": "riyadh HOTELS", "cells": {"department": "Rooms"}})   # same name: replaced
        self.assertEqual(len(prefs.handle({"op": "get"})["filterSets"]), 1)
        self.assertEqual(prefs.handle({"op": "removeSet", "name": "Riyadh Hotels"})["filterSets"], [])

    def test_secrets_never_leave(self):
        cfg = {"accountId": "123", "consumerSecret": "s3", "tokenSecret": "t3", "consumerKey": "k",
               "ai": {"apiKey": "sk-x", "model": "m"}, "_session": object(), "clientId": "public"}
        clean = self.care.redact(cfg)
        self.assertEqual(clean["accountId"], "123")
        self.assertEqual(clean["clientId"], "public")
        for k in ("consumerSecret", "tokenSecret", "consumerKey"):
            self.assertEqual(clean[k], "***")
        self.assertEqual(clean["ai"]["apiKey"], "***")
        self.assertNotIn("_session", clean)
        path = self.care.diagnostics(self.dir, "9.9", {"mode": "x"}, cfg, os.path.join(self.dir, "none.crt"),
                                     open_folder=False)
        import zipfile
        with zipfile.ZipFile(path) as z:
            text = "".join(z.read(n).decode() for n in z.namelist())
        for secret in ("s3", "t3", "sk-x"):
            self.assertNotIn('"%s"' % secret, text)

    def test_export_and_read_back(self):
        cfg = {"accountId": "123_SB1", "restletUrl": "https://x", "clientId": "c", "consumerSecret": "never"}
        path, summary = self.care.export_bundle(cfg, '{"customer": "A"}', (b"\x89PNG\r\n\x1a\nabc", "image/png"),
                                                {"filterSets": [{"name": "A", "cells": {"location": "B"}}]},
                                                [{"lines": [], "title": "Hotel"}, {"bad": 1}], [{"rows": [["a"]]}], "1.4.0", "Hotel A")
        self.assertEqual(summary, {"settings": True, "licence": True, "logo": True, "filterSets": 1, "layouts": 1, "queries": 1})
        with open(path, encoding="utf-8") as f:
            text = f.read()
        self.assertNotIn("never", text)
        bundle = self.care.read_bundle(json.loads(text))
        self.assertEqual(bundle["settings"]["accountId"], "123_SB1")
        with self.assertRaises(ValueError):
            self.care.read_bundle({"kind": "something else"})

    def test_certificate_expiry(self):
        cert, key, _ = server.tqcrypto.make_certificate(days=30)
        path = os.path.join(self.dir, "c.crt")
        with open(path, "w", encoding="ascii") as f:
            f.write(cert)
        self.assertIn(self.care.cert_days_left(path), (29, 30, 31))


class HttpTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = server.make_server(0, cfg=None, use_tls=False, today=TODAY)
        cls.base = "http://127.0.0.1:%d" % cls.httpd.server_address[1]
        threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def post(self, payload, content_type="application/json", path="/api"):
        request = urllib.request.Request(self.base + path, data=json.dumps(payload).encode(),
                                         headers={"Content-Type": content_type}, method="POST")
        with urllib.request.urlopen(request, timeout=10) as response:
            return json.loads(response.read().decode())

    def test_budget_route_and_new_files(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:   # writing goes through /api/budget only
            self.post({"action": "budgetSave", "dryRun": False, "lines": []})
        self.assertEqual(caught.exception.code, 400)
        info = self.post({"op": "info"}, path="/api/budget")
        self.assertEqual([y["name"] for y in info["years"]], ["FY2025", "FY2026"])
        self.assertEqual(self.post({"action": "period", "base": "Oct 2026", "shift": "FYS"}), {"ok": True, "name": "Jan 2026"})
        for page in ("builder.js", "pack.js", "budget.js"):
            with urllib.request.urlopen(self.base + "/" + page, timeout=10) as response:
                self.assertEqual(response.status, 200)

    def test_logo_upload(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:     # sample data: no logo yet
            urllib.request.urlopen(self.base + "/logo", timeout=10)
        self.assertEqual(caught.exception.code, 404)
        self.assertIn("PNG or JPEG", self.post({"op": "save", "data": base64.b64encode(b"hello").decode()}, path="/api/logo")["error"])
        png = open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "addin", "assets", "icon-32.png"), "rb").read()
        try:
            self.assertEqual(self.post({"op": "save", "data": base64.b64encode(png).decode()}, path="/api/logo")["logo"], "custom")
            with urllib.request.urlopen(self.base + "/logo", timeout=10) as response:
                self.assertEqual((response.headers["Content-Type"], response.read()), ("image/png", png))
            brand = self.post({"action": "branding", "subsidiary": ""})
            self.assertEqual(base64.b64decode(brand["logo"]["data"]), png)    # the uploaded logo wins on PDFs too
        finally:
            self.assertEqual(self.post({"op": "remove"}, path="/api/logo")["logo"], "netsuite")

    def test_status_and_static_files(self):
        with urllib.request.urlopen(self.base + "/api/status", timeout=10) as response:
            self.assertEqual(json.loads(response.read().decode()),
                         {"ok": True, "mode": "sample", "chat": True, "ai": False, "version": server.VERSION,
                          "licence": None, "configured": False, "seats": None, "update": None,
                          "signIn": {"mode": "sample"}})
        for name in ("taskpane.html", "setup.html", "panel.css", "nav.js", "library.js", "extras.js", "drill.html", "shortcuts.json", "functions.js", "chat.js", "designer.js", "data.js", "functions.json", "manifest.xml",
                     "assets/icon-32.png"):
            with urllib.request.urlopen(self.base + "/" + name, timeout=10) as response:
                self.assertEqual(response.status, 200, name)

    def test_api_round_trip(self):
        out = self.post({"action": "balances", "requests": [{"account": "4000", "fromPeriod": "Jan 2026"},
                                                            {"account": "type:Income", "fromPeriod": "Jan 2026"}]})
        self.assertTrue(out["ok"])
        self.assertIsInstance(out["results"][0]["value"], float)
        self.assertGreater(out["results"][1]["value"], out["results"][0]["value"])

    def test_chat_route(self):
        out = self.post({"message": "trial balance Oct 2026"}, path="/api/chat")
        self.assertEqual(out["actions"][0]["sheetName"], "TB Oct 2026")

    def test_refuses_posts_from_other_sites(self):
        def refused(headers):
            request = urllib.request.Request(self.base + "/api/setup", data=b"{}", headers=headers, method="POST")
            with self.assertRaises(urllib.error.HTTPError) as caught:
                urllib.request.urlopen(request, timeout=10)
            return caught.exception.code
        self.assertEqual(refused({"Content-Type": "text/plain; x=application/json"}), 415)
        self.assertEqual(refused({"Content-Type": "application/json", "Sec-Fetch-Site": "cross-site"}), 403)
        ok = self.post({"action": "ping"})
        self.assertTrue(ok["ok"])

    def test_rejects_non_json_posts(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.post({"action": "ping"}, content_type="text/plain")
        self.assertEqual(caught.exception.code, 415)


class LedgerTest(unittest.TestCase):
    def setUp(self):
        self.ledger = new_ledger()

    def value(self, account, start, end="", ledger=None):
        result = (ledger or self.ledger).balances([{"account": account, "fromPeriod": start, "toPeriod": end}])[0]
        return result.get("value", result.get("error"))

    def net(self, start, end):
        return (self.value("type:Income", start, end) - self.value("type:COGS", start, end)
                - self.value("type:Expense", start, end))

    def test_type_total_is_the_sum_of_its_accounts(self):
        each = self.value("4000", "Jan 2026", "Mar 2026") + self.value("4100", "Jan 2026", "Mar 2026")
        self.assertAlmostEqual(self.value("type:Income", "Jan 2026", "Mar 2026"), each, places=2)
        both = self.value("type:Bank", "Oct 2026") + self.value("type:AcctRec", "Oct 2026")
        self.assertAlmostEqual(self.value("TYPE:bank, acctrec", "Oct 2026"), both, places=2)

    def test_net_income_and_its_split_across_fiscal_years(self):
        self.assertAlmostEqual(self.value("calc:NetIncome", "Jan 2026", "Jun 2026"), self.net("Jan 2026", "Jun 2026"), places=2)
        self.assertAlmostEqual(self.value("calc:NetIncomeYTD", "Jun 2026"), self.net("Jan 2026", "Jun 2026"), places=2)
        self.assertAlmostEqual(self.value("calc:RetainedEarnings", "Jun 2026"), self.net("Jan 2025", "Dec 2025"), places=2)
        self.assertEqual(self.value("calc:RetainedEarnings", "Jun 2025"), 0)  # nothing before the first year

    def test_fiscal_year_start_month(self):
        april = new_ledger(fiscal_start_month=4)
        self.assertAlmostEqual(self.value("calc:NetIncomeYTD", "Jun 2026", ledger=april),
                               self.net("Apr 2026", "Jun 2026"), places=2)

    def test_errors_and_pass_through(self):
        self.assertIn("Unknown account type", self.value("type:Nope", "Jan 2026"))
        self.assertIn("not found", self.value("type:Income", "Smarch 2026"))
        drill = self.ledger.handle({"action": "transactions",
                                    "request": {"account": "type:Income", "fromPeriod": "Jan 2026"}})
        self.assertIn("single account", drill["error"])
        self.assertTrue(self.ledger.handle({"action": "ping"})["ok"])

    def test_budget_movement_and_breakdown(self):
        actual, budget = self.value("type:Income", "Jan 2026", "Mar 2026"), self.value("budget:4000-4999", "Jan 2026", "Mar 2026")
        self.assertEqual(budget, self.value("budget(Plan A):4000", "Jan 2026", "Mar 2026")
                         + self.value("BUDGET:4100", "Jan 2026", "Mar 2026"))
        self.assertNotEqual(actual, budget)
        # The movement of a balance sheet account is its closing balance less the one before the first period.
        self.assertAlmostEqual(self.value("move:1000", "Feb 2026", "Mar 2026"),
                               self.value("1000", "Mar 2026") - self.value("1000", "Jan 2026"), places=2)
        self.assertAlmostEqual(self.value("move:4000", "Feb 2026", "Mar 2026"), self.value("4000", "Feb 2026", "Mar 2026"))
        out = self.ledger.handle({"action": "breakdown",
                                  "request": {"account": "4000-4999, type:COGS", "fromPeriod": "Jan 2026"}})
        self.assertEqual([r["account"] for r in out["rows"]], ["4000", "4100", "5000"])
        self.assertAlmostEqual(out["total"], self.value("4000-4999, type:COGS", "Jan 2026"), places=2)
        self.assertEqual(out["rows"][0]["request"]["account"], "4000")   # what a click on the row drills into
        net = self.ledger.handle({"action": "breakdown", "request": {"account": "calc:NetIncome", "fromPeriod": "Jan 2026"}})
        self.assertAlmostEqual(net["total"], self.value("calc:NetIncome", "Jan 2026"), places=2)

    def test_list_lookups_are_answered_from_the_cache(self):
        calls = []
        mock = server.MockBackend()
        ledger = server.ledger_module.Ledger(lambda p: calls.append(p["action"]) or mock.handle(p), today=TODAY)
        for _ in range(3):
            self.assertEqual(len(ledger.handle({"action": "accounts"})["accounts"]), len(server.MOCK_ACCOUNTS))
            self.assertTrue(ledger.handle({"action": "segments", "kind": "department"})["segments"])
            ledger.accounts()
        self.assertEqual(calls, ["accounts", "segments"])

    def test_account_lists_ranges_and_days(self):
        income = self.value("type:Income", "Jan 2026", "Mar 2026")
        cogs = self.value("5000", "Jan 2026", "Mar 2026")
        self.assertAlmostEqual(self.value("4000, 4100", "Jan 2026", "Mar 2026"), income, places=2)
        self.assertAlmostEqual(self.value("4000-4999, type:COGS", "Jan 2026", "Mar 2026"), income + cogs, places=2)
        self.assertAlmostEqual(self.value("6000..6199", "Jan 2026"),
                               self.value("6000", "Jan 2026") + self.value("6100", "Jan 2026"), places=2)
        self.assertEqual(self.ledger.members("type:Bank,AcctRec, 4000, 4000-4100"), ["1000", "1100", "4000", "4100"])
        self.assertIn('"9999" not found', self.value("4000, 9999", "Jan 2026"))
        self.assertEqual(self.value("calc:Days", "Jan 2026", "Mar 2026"), 90)
        self.assertEqual(self.value("calc:Days", "Feb 2026"), 28)
        drill = self.ledger.handle({"action": "transactions",
                                    "request": {"account": "4000,4100", "fromPeriod": "Jan 2026"}})
        self.assertIn("single account", drill["error"])


class Workbook:
    """Just enough of Excel to work out a report: cell references, SUM, arithmetic, ROUND, IFERROR,
    text joined with & and TQ.BALANCE."""

    KEYS = ("account", "fromPeriod", "toPeriod", "subsidiary", "department", "class", "location", "filters")

    def __init__(self, rows, ledger, book=None):
        self.rows, self.ledger, self.book = rows, ledger, book or {}  # book: other sheets, by name

    @staticmethod
    def position(ref):
        letters, digits = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref).groups()
        col = 0
        for ch in letters:
            col = col * 26 + ord(ch) - 64
        return int(digits) - 1, col - 1

    def value(self, ref):
        row, col = self.position(ref)
        return self.evaluate(self.rows[row][col])

    def evaluate(self, content):
        if not (isinstance(content, str) and content.startswith("=")):
            return content
        formula = content[1:]
        other = re.fullmatch(r"'([^']+)'!([A-Z]+\d+)", formula)  # ='Doha - Report'!C12
        if other:
            return Workbook(self.book[other.group(1)], self.ledger, self.book).value(other.group(2))
        period = re.fullmatch(r'TQ\.PERIOD\(([^,()]+)(?:,"?([^"()]*)"?)?\)', formula)   # =TQ.PERIOD($B$8,"FYS")
        if period:
            base = period.group(1)
            base = base.strip('"') if base.startswith('"') else self.value(base)
            return self.ledger.shift_period(base, period.group(2) or "")
        if re.fullmatch(r'"[^"]*"', formula):   # ="Jan 2026"
            return formula.strip('"')
        wrapped = re.fullmatch(r'IFERROR\((.*),""\)', formula)
        if wrapped:
            try:
                return self.evaluate("=" + wrapped.group(1))
            except ZeroDivisionError:
                return ""
        if "&" in formula and "TQ." not in formula:  # ="Jan "&($B$8-1)
            def piece(part):
                if part.startswith('"'):
                    return part.strip('"')
                shifted = re.fullmatch(r"\((.+?)([+-]\d+)\)", part)
                return str(int(self.value(shifted.group(1))) + int(shifted.group(2))) if shifted \
                    else str(self.value(part))
            return "".join(piece(part) for part in formula.split("&"))

        def argument(text):
            joined = re.fullmatch(r'"([^"]*)"&(.+)', text)  # "budget:"&$A12
            if joined:
                return joined.group(1) + str(self.value(joined.group(2)))
            return text.strip('"') if text.startswith('"') else self.value(text)

        def balance(match):
            args = [argument(a) for a in match.group(1).split(",")]
            return repr(self.ledger.balances([dict(zip(self.KEYS, args))])[0]["value"])

        def total(match):
            (r0, c0), (r1, c1) = self.position(match.group(1)), self.position(match.group(2))
            return repr(sum(self.evaluate(self.rows[r][c]) or 0
                            for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)))

        def choose(match):  # IF(C$12="",$B$5,C$12): a column's own filter, else the report's
            return match.group(2) if self.value(match.group(1)) in ("", None) else match.group(3)

        formula = re.sub(r'IF\(([$A-Z0-9]+)="",([$A-Z0-9]+),([$A-Z0-9]+)\)', choose, formula)
        formula = re.sub(r"TQ\.BALANCE\(([^()]*)\)", balance, formula)
        formula = re.sub(r"SUM\(([A-Z]+\d+):([A-Z]+\d+)\)", total, formula)
        formula = re.sub(r"\$?\b[A-Z]{1,2}\$?\d+\b", lambda m: repr(self.value(m.group(0)) or 0), formula)
        formula = formula.replace("ROUND(", "round(").replace("ABS(", "abs(").replace("MAX(", "max(")
        return eval(formula, {"__builtins__": {}}, {"round": round, "abs": abs, "max": max})  # only numbers and operators remain

    def find(self, label):
        # A section heading can share a line's name; the line is the one with numbers.
        return next(n + 1 for n, row in enumerate(self.rows) if row[1] == label and row[2] != "")


class ReportTest(unittest.TestCase):
    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def build(self, message, last=None):
        out = self.chat.handle({"message": message, "context": {"last": last}})
        self.assertIn("actions", out, out)
        self.assertTrue(out["actions"], out)
        return out

    def net(self, start, end, **segments):
        request = dict(account="calc:NetIncome", fromPeriod=start, toPeriod=end, **segments)
        return self.ledger.balances([request])[0]["value"]

    def test_income_statement_by_month_adds_up(self):
        action = self.build("P&L for 2026 by month")["actions"][0]
        self.assertEqual(action["sheetName"], "P&L 2026")
        self.assertEqual(trim(action["rows"][8][2:]), ["%s 2026" % m for m in server.MONTHS] + ["Total"])
        self.assertIn([8, 2, 8, 14], action["textRanges"])  # period names must stay text in Excel
        book = Workbook(action["rows"], self.ledger)
        row = book.find("Net Income")
        self.assertAlmostEqual(book.value("E%d" % row), self.net("Mar 2026", "Mar 2026"), places=2)
        self.assertAlmostEqual(book.value("O%d" % row), self.net("Jan 2026", "Dec 2026"), places=2)
        check = book.find("Check: NetSuite net income less the line above (should be 0)")
        self.assertEqual(book.value("C%d" % check), 0)

    def test_follow_ups_change_the_last_report(self):
        first = self.build("income statement 2026")
        second = self.build("by quarter", first["spec"])
        self.assertEqual(trim(second["actions"][0]["rows"][10][2:]), ["Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026", "Total"])
        third = self.build("compare with last year, summary", second["spec"])
        rows = third["actions"][0]["rows"]
        self.assertEqual(rows[10][2:6], ["Q1 2026", "Q1 2025", "Change", "Change %"])
        book = Workbook(rows, self.ledger)
        row = book.find("Net Income")
        self.assertAlmostEqual(book.value("D%d" % row), self.net("Jan 2025", "Mar 2025"), places=2)
        self.assertAlmostEqual(book.value("E%d" % row),
                               self.net("Jan 2026", "Mar 2026") - self.net("Jan 2025", "Mar 2025"), places=2)
        fourth = self.build("no comparison, by month, detailed", third["spec"])
        self.assertEqual(fourth["spec"]["compare"], None)
        self.assertEqual(len(trim(fourth["actions"][0]["rows"][8])), 15)
        out = self.chat.handle({"message": "now for 2025", "context": {"last": third["spec"]}})
        self.assertIn("no periods a year earlier", out["reply"])  # the sample data starts in 2025

    def test_split_by_department_and_filter(self):
        out = self.build("P&L Jan to Jun 2026 by department for location Doha")
        rows = out["actions"][0]["rows"]
        self.assertEqual(trim(rows[10][2:]), ["Sales", "Operations", "Finance", "No department", "Total"])
        self.assertEqual(rows[6][:2], ["Location", "Doha"])
        self.assertEqual(out["spec"]["split"], "department")
        book = Workbook(rows, self.ledger)
        row = book.find("Net Income")
        self.assertAlmostEqual(book.value("C%d" % row),
                               self.net("Jan 2026", "Jun 2026", department="Sales", location="Doha"), places=2)
        self.assertAlmostEqual(book.value("G%d" % row), self.net("Jan 2026", "Jun 2026", location="Doha"), places=2)

    def test_balance_sheet_layout(self):
        rows = self.build("balance sheet as at Oct 2026 vs Dec 2025")["actions"][0]["rows"]
        self.assertEqual(trim(rows[8][2:]), ["Oct 2026", "Dec 2025", "Change", "Change %"])
        book = Workbook(rows, self.ledger)
        value = lambda n, p: self.ledger.balances([{"account": n, "fromPeriod": p}])[0]["value"]  # noqa: E731
        assets = sum(value(n, "Oct 2026") for n in ("1000", "1100", "1200", "1500"))
        self.assertAlmostEqual(book.value("C%d" % book.find("Total Assets")), assets, places=2)
        equity = value("3000", "Dec 2025") + self.net("Jan 2025", "Dec 2025") + value("calc:Translation", "Dec 2025")
        self.assertAlmostEqual(book.value("D%d" % book.find("Total Equity")), equity, places=2)
        summary = self.build("bs by quarter 2026 summary")["actions"][0]["rows"]
        self.assertEqual(trim(summary[8][2:]), ["Mar 2026", "Jun 2026", "Sep 2026", "Nov 2026"])  # not later than today

    def test_trial_balance(self):
        rows = self.build("trial balance Oct 2026")["actions"][0]["rows"]
        self.assertEqual(rows[8][:2], ["As at", "Oct 2026"])
        self.assertEqual(rows[9][:2], ["P&L from", "Jan 2026"])
        self.assertEqual(rows[19][:1] + rows[19][3:5], ["4000", "=MAX(-F20,0)", "=MAX(F20,0)"])  # income is a credit


LAYOUT = [
    ["TQ Layout"], ["Report title", "Restaurant P&L"], [],
    ["Section", "Line", "Type", "Accounts or formula", "Sign", "Format"],
    ["KPI", "Average Sales Daily", "formula", "=[Net Sales]/days", "", ""],
    ["Revenue", "Sales", "accounts", 4000.0, 1.0, ""],     # as Excel hands over cells typed as numbers
    ["", "Services", "", "4100", "", ""],
    ["", "Net Sales", "subtotal", "", "", ""],
    ["Cost of Sales", "Food Cost", "accounts", "5000-5999", -1.0, ""],
    ["", "Total Cost of Sales", "subtotal", "", "", ""],
    ["Gross Profit", "Gross Profit", "formula", "=[Net Sales]+[Total Cost of Sales]", "", "bold"],
    ["", "Gross Margin", "", "=[Gross Profit]/[Net Sales]", "", "percent"],
    ["Staff Cost", "Salaries", "accounts", "6000", -1, ""],
]


class LayoutTest(unittest.TestCase):
    """A layout sheet's lines and columns -> the report sheet, both rewritten in place."""

    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def context(self, rows, last=None, name="TQ Layout", active=""):
        return {"last": last, "activeSheet": active,
                "layouts": [{"name": name, "row": 0, "rows": rows}] if rows else []}

    def ask(self, message, rows=LAYOUT, last=None):
        return self.chat.handle({"message": message, "context": self.context(rows, last)})

    def test_hospitality_and_planning_templates(self):
        """Hotel (USALI) from the customer's departments, restaurant, rolling forecast, last 12 months."""
        saved = dict(server.MOCK_SEGMENTS)
        server.MOCK_SEGMENTS["department"] = ["11 Rooms Division", "22 Food & Beverage", "22 Food & Beverage : Kitchen",
                                              "33 Spa", "55 A&G", "77 Sales and Marketing", "88 Repair and Maintenance"]
        try:
            ledger = new_ledger()
            chat = server.chat_module.ChatService(ledger)
            wiz = lambda t: chat.design({"op": "wizard", "args": {"template": t, "month": "Oct 2026"},  # noqa: E731
                                         "context": {"layouts": [], "sheets": []}})
            hotel = wiz("hotel")
            self.assertTrue(hotel["ok"], hotel.get("error"))
            layout = [a for a in hotel["actions"] if a["sheetName"].startswith("TQ Layout")][0]
            sources = [r[3] for r in layout["rows"] if len(r) > 3 and isinstance(r[3], str)]
            self.assertIn("type:Income,OthIncome | Department: 11 Rooms Division*", sources)
            self.assertIn("type:Expense,COGS | Department: 55 A&G*", sources)
            report = [a for a in hotel["actions"] if a["sheetName"] == "Hotel P&L"][0]
            flat = json.dumps(report["rows"])
            self.assertIn("Department: 22 Food & Beverage*", flat)   # the line's filter reaches TQ.BALANCE
            self.assertIn("Gross Operating Profit", flat)
            self.assertIn("RevPAR", hotel["reply"])                   # no room statistics in the sample: said so
            parsed = server.chat_module.layouts.parse(layout["sheetName"], layout["rows"])
            rooms = [l for l in parsed["lines"] if l["label"] == "Rooms Revenue"][0]
            self.assertEqual((rooms["source"], rooms["filter"]), ("type:Income,OthIncome", "Department: 11 Rooms Division*"))
            for t, title in (("restaurant", "Restaurant P&L"), ("forecast", "Rolling Forecast"), ("rolling12", "Last 12 Months")):
                out = wiz(t)
                self.assertTrue(out["ok"], (t, out.get("error")))
                self.assertTrue(any(a["sheetName"] == title for a in out["actions"]), t)
            fc = [a for a in wiz("forecast")["actions"] if a["sheetName"].startswith("TQ Layout")][0]
            self.assertIn("budget(Budget - A)", json.dumps(fc["rows"]))
        finally:
            server.MOCK_SEGMENTS.clear()
            server.MOCK_SEGMENTS.update(saved)

    def test_comment_needed_marking(self):
        made = self.chat.design({"op": "wizard", "args": {"template": "monthytd", "month": "Oct 2026", "comments": True},
                                 "context": {"layouts": [], "sheets": []}})
        layout = [a for a in made["actions"] if a["sheetName"].startswith("TQ Layout")][0]
        rows = [r if r[:1] != ["Comment needed over"] else ["Comment needed over", "10%"] for r in layout["rows"]]
        context = {"layouts": [{"name": layout["sheetName"], "row": 0, "rows": rows}], "sheets": []}
        out = self.chat.design({"op": "build", "args": {"layout": layout["sheetName"]}, "context": context})
        report = [a for a in out["actions"] if not a["sheetName"].startswith("TQ Layout")][0]
        marks = [c for c in report.get("conditional", []) if c.get("fill") == "#FFF4CE"]
        self.assertEqual(len(marks), 1)
        self.assertIn(">=0.1", marks[0]["formula"])

    def test_company_library_round_trip(self):
        """Save a layout to the library, then use it in another workbook: same lines, new names, report built."""
        made = self.chat.design({"op": "new", "args": {"template": "summary"}, "context": {"layouts": []}})
        sheet = [a for a in made["actions"] if a["sheetName"].startswith("TQ Layout")][0]
        context = {"layouts": [{"name": sheet["sheetName"], "row": 0, "rows": sheet["rows"]}], "sheets": []}
        out = self.chat.design({"op": "libraryExport", "args": {"layout": sheet["sheetName"]}, "context": context})
        self.assertTrue(out["ok"])
        backend = server.MockBackend()
        saved = backend.handle({"action": "librarySave", "name": "Group monthly pack", "design": out["design"]})
        self.assertEqual(saved["saved"], "group-monthly-pack")
        self.assertEqual([i["name"] for i in backend.handle({"action": "libraryList"})["items"]], ["Group monthly pack"])
        item = backend.handle({"action": "libraryLoad", "id": "group-monthly-pack"})["item"]
        other = {"layouts": [{"name": sheet["sheetName"], "row": 0, "rows": sheet["rows"]}],
                 "sheets": [out["title"][:31]]}
        used = self.chat.design({"op": "libraryImport", "args": {"design": item["design"]}, "context": other})
        self.assertTrue(used["ok"], used.get("error"))
        names = [a["sheetName"] for a in used["actions"]]
        self.assertNotIn(sheet["sheetName"], names)   # a free layout name: the workbook already has the original
        self.assertTrue(any(n.startswith("TQ Layout") for n in names))
        self.assertIn("company library", used["reply"])
        self.assertEqual(backend.handle({"action": "libraryDelete", "id": "group-monthly-pack"})["items"], [])
        bad = self.chat.design({"op": "libraryImport", "args": {"design": {"x": 1}}, "context": other})
        self.assertFalse(bad["ok"])

    def total(self, account, start, end, **segments):
        return self.ledger.balances([dict(account=account, fromPeriod=start, toPeriod=end, **segments)])[0]["value"]

    def test_custom_report_by_month(self):
        out = self.ask("custom P&L for 2026 by month")
        layout, action = out["actions"]
        rows = action["rows"]
        self.assertEqual((layout["sheetName"], layout["target"], layout["activate"]), ("TQ Layout", "sheet", False))
        self.assertEqual(len(server.chat_module.layouts.parse("TQ Layout", layout["rows"])["columns"]), 13)
        self.assertEqual((action["sheetName"], action["target"]), ("Restaurant P&L", "sheet"))
        self.assertEqual(rows[7][:2], ["Year", 2026])
        self.assertEqual(rows[9][:4], ["", "From period", '="Jan "&$B$8', '="Feb "&$B$8'])
        self.assertEqual(rows[11][:3] + [rows[11][14]], ["Accounts", "Name", '="Jan "&$B$8', "Total"])
        self.assertIn("2 profit and loss account(s)", out["reply"])
        self.assertIn("6100 Rent, 6200 Marketing", out["reply"])
        self.assertEqual(action["dropdowns"][1], {"r": 3, "c": 1, "free": True, "formula": "=$Q$2:$Q$4"})
        self.assertEqual(rows[1][16], "Parent Company")
        self.assertTrue(action["groups"])
        book = Workbook(rows, self.ledger)
        income = self.total("type:Income", "Mar 2026", "Mar 2026")
        self.assertAlmostEqual(book.value("E%d" % book.find("Net Sales")), income, places=2)
        self.assertAlmostEqual(book.value("E%d" % book.find("Food Cost")), -self.total("5000", "Mar 2026", "Mar 2026"),
                               places=2)
        year_income = self.total("type:Income", "Jan 2026", "Dec 2026")
        gross = year_income - self.total("5000", "Jan 2026", "Dec 2026")
        self.assertAlmostEqual(book.value("O%d" % book.find("Gross Profit")), gross, places=2)
        self.assertAlmostEqual(book.value("O%d" % book.find("Gross Margin")), gross / year_income, places=6)
        daily = book.find("Average Sales Daily")
        self.assertAlmostEqual(book.value("C%d" % daily), self.total("type:Income", "Jan 2026", "Jan 2026") / 31, places=6)
        self.assertAlmostEqual(book.value("O%d" % daily), year_income / 365, places=6)
        rows[7][1] = 2025  # the Year cell moves every column
        self.assertAlmostEqual(book.value("E%d" % book.find("Net Sales")),
                               self.total("type:Income", "Mar 2025", "Mar 2025"), places=2)

    def test_follow_up_comparison_and_split(self):
        first = self.ask("custom report for 2026 by quarter")
        self.assertEqual((first["spec"]["report"], first["spec"]["layoutName"]), ("layout", "TQ Layout"))
        second = self.ask("compare with last year", rows=first["actions"][0]["rows"], last=first["spec"])
        rows = second["actions"][1]["rows"]
        self.assertEqual(rows[11][2:6], ['="Q1 "&$B$8', '="Q1 "&($B$8-1)', '="Change Q1 "&$B$8', '="Change % Q1 "&$B$8'])
        book = Workbook(rows, self.ledger)
        row = book.find("Net Sales")
        was, now = (self.total("type:Income", "Jan %d" % y, "Mar %d" % y) for y in (2025, 2026))
        self.assertAlmostEqual(book.value("E%d" % row), now - was, places=2)
        self.assertAlmostEqual(book.value("F%d" % row), (now - was) / was, places=6)
        margin = book.find("Gross Margin")
        self.assertEqual(rows[margin - 1][5], "")  # no "change %" of a percentage
        self.assertAlmostEqual(book.value("E%d" % margin), book.value("C%d" % margin) - book.value("D%d" % margin))
        split = self.ask("custom report Jan to Jun 2026 by department")["actions"][1]["rows"]
        self.assertEqual(split[11][1:5], ["Department", "Sales", "Operations", "Finance"])
        self.assertEqual(split[12][2:7], ["Sales", "Operations", "Finance", "No department", "Total"])
        book = Workbook(split, self.ledger)
        self.assertAlmostEqual(book.value("C%d" % book.find("Net Sales")),
                               self.total("type:Income", "Jan 2026", "Jun 2026", department="Sales"), places=2)
        self.assertAlmostEqual(book.value("G%d" % book.find("Average Sales Daily")),
                               self.total("type:Income", "Jan 2026", "Jun 2026") / 181, places=6)

    def test_chat_edits_change_the_same_sheets(self):
        out = self.ask("custom P&L for 2026 by quarter")
        out = self.ask("add column Q1 2025", rows=out["actions"][0]["rows"])
        self.assertIn('Added column "Q1 2025"', out["reply"])
        self.assertEqual(out["actions"][1]["sheetName"], "Restaurant P&L")
        out = self.ask("add column Growth = [Q1 2026]/[Q1 2025]-1 percent", rows=out["actions"][0]["rows"])
        rows = out["actions"][1]["rows"]
        self.assertEqual(trim(rows[11][2:]), ['="Q1 "&$B$8', '="Q2 "&$B$8', '="Q3 "&$B$8', '="Q4 "&$B$8', "Total",
                                        '="Q1 "&($B$8-1)', "Growth"])
        book = Workbook(rows, self.ledger)
        was, now = (self.total("type:Income", "Jan %d" % y, "Mar %d" % y) for y in (2025, 2026))
        self.assertAlmostEqual(book.value("I%d" % book.find("Net Sales")), now / was - 1, places=6)
        out = self.ask("add column Jan to Jun 2026 for department Sales called Sales H1", rows=out["actions"][0]["rows"])
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        self.assertAlmostEqual(book.value("J%d" % book.find("Net Sales")),
                               self.total("type:Income", "Jan 2026", "Jun 2026", department="Sales"), places=2)
        out = self.ask("add line Rent accounts 6100, 6200 negative in section staff cost", rows=out["actions"][0]["rows"])
        self.assertIn('Added line "Rent" in section "Staff Cost"', out["reply"])
        self.assertIn("Every profit and loss account", out["reply"])
        out = self.ask("add subtotal Total Staff Cost in section Staff Cost", rows=out["actions"][0]["rows"])
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        staff = sum(self.total(n, "Jan 2026", "Mar 2026") for n in ("6000", "6100", "6200"))
        self.assertAlmostEqual(book.value("C%d" % book.find("Total Staff Cost")), -staff, places=2)
        out = self.ask("add line Net Profit = [Gross Profit]+[Total Staff Cost] bold", rows=out["actions"][0]["rows"])
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        self.assertAlmostEqual(book.value("G%d" % book.find("Net Profit")),
                               self.total("calc:NetIncome", "Jan 2026", "Dec 2026"), places=2)
        layout = out["actions"][0]["rows"]
        self.assertIn("There is no column", self.ask("remove column Budget", rows=layout)["reply"])
        self.assertIn("uses [Net Sales]", self.ask("remove line Net Sales", rows=layout)["reply"])
        out = self.ask("remove column Growth", rows=layout)
        self.assertEqual(len(out["design"]["columns"]), 7)
        self.assertIn("lines by 7 columns", self.ask("update report", rows=out["actions"][0]["rows"])["reply"])

    def test_one_sheet_per_location(self):
        rows = self.ask("custom P&L for 2026 by quarter")["actions"][0]["rows"]
        out = self.ask("one sheet per location", rows=rows)
        names = [a["sheetName"] for a in out["actions"]]
        self.assertEqual(names, ["TQ Layout", "Restaurant P&L", "Restaurant P&L - Summary", "Doha - Restaurant P&L",
                                 "Riyadh - Restaurant P&L"])
        self.assertEqual([a["activate"] for a in out["actions"]], [False, True, False, False, False])
        self.assertIn("one sheet for each location with a balance (2): Doha, Riyadh", out["reply"])
        # The summary: one row per location, the first key line in every column, the others at the total.
        summary = out["actions"][2]["rows"]
        self.assertEqual(summary[3][:2] + summary[3][6:], ["Location", "='Restaurant P&L'!C12", "Total Cost of Sales",
                                                           "Gross Profit"])
        self.assertEqual([r[0] for r in summary[4:]], ['=HYPERLINK("#\'Restaurant P&L\'!A1","All")',
                                                       '=HYPERLINK("#\'Doha - Restaurant P&L\'!A1","Doha")',
                                                       '=HYPERLINK("#\'Riyadh - Restaurant P&L\'!A1","Riyadh")'])
        book = Workbook(summary, self.ledger, {a["sheetName"]: a["rows"] for a in out["actions"]})
        income = lambda start, end, **seg: self.total("type:Income", start, end, **seg)  # noqa: E731
        self.assertAlmostEqual(book.value("B6"), income("Jan 2026", "Mar 2026", location="Doha"), places=2)
        self.assertAlmostEqual(book.value("F5"), income("Jan 2026", "Dec 2026"), places=2)
        self.assertAlmostEqual(book.value("H7"), income("Jan 2026", "Dec 2026", location="Riyadh")
                               - self.total("5000", "Jan 2026", "Dec 2026", location="Riyadh"), places=2)
        doha = out["actions"][3]["rows"]
        self.assertEqual(doha[6][:2], ["Location", "Doha"])
        book = Workbook(doha, self.ledger)
        self.assertAlmostEqual(book.value("C%d" % book.find("Net Sales")),
                               self.total("type:Income", "Jan 2026", "Mar 2026", location="Doha"), places=2)
        self.assertEqual(out["design"]["pages"], "location")
        layout = out["actions"][0]["rows"]
        self.assertEqual(len(self.ask("update report", rows=layout)["actions"]), 5)   # it is kept in the layout
        # Picking it under Sheets in the panel builds the sheets at once, without Update report.
        plain = self.ask("one sheet only", rows=layout)["actions"][0]["rows"]
        out = self.chat.design({"op": "setPages", "args": {"pages": "location"}, "context": self.context(plain)})
        self.assertTrue(out["ok"], out)
        self.assertEqual([a["sheetName"] for a in out["actions"]], names)
        self.assertIn("one sheet for each location with a balance (2)", out["reply"])
        out = self.chat.design({"op": "setPages", "args": {"pages": "Departments"}, "context": self.context(layout)})
        self.assertEqual(out["design"]["pages"], "department")
        self.assertGreater(len(out["actions"]), 2)
        self.assertTrue(all(a["sheetName"].endswith("Restaurant P&L") or a["sheetName"] in ("TQ Layout",
                        "Restaurant P&L - Summary") for a in out["actions"]))
        off = self.chat.design({"op": "setPages", "args": {"pages": ""}, "context": self.context(layout)})
        self.assertEqual([a["sheetName"] for a in off["actions"]], ["TQ Layout", "Restaurant P&L"])
        self.assertEqual(len(self.ask("one sheet only", rows=layout)["actions"]), 2)

    def test_relative_periods(self):
        led = self.ledger
        now = led.periods()[led.current_index()]["name"]
        self.assertEqual(led.period_name("CM"), now)
        self.assertEqual(led.shift_period("Oct 2026", -1), "Sep 2026")
        self.assertEqual(led.shift_period("Oct 2026", "LY"), "Oct 2025")
        self.assertEqual(led.shift_period("Oct 2026", "FYS"), "Jan 2026")
        self.assertEqual(led.shift_period("Oct 2026", "FYE"), "Dec 2026")
        self.assertEqual(led.shift_period("Oct 2026", "FYS-12"), "Jan 2025")
        self.assertEqual(led.shift_period("Oct 2026", "FYS-1"), "Dec 2025")
        self.assertEqual(led.shift_period("Nov 2026", "QS"), "Oct 2026")
        self.assertEqual(led.shift_period("Nov 2026", "QE"), "Dec 2026")
        self.assertEqual(led.shift_period("Mar 2026", "12"), "Mar 2027") if any(
            p["name"] == "Mar 2027" for p in led.periods()) else None
        same = led.balances([{"account": "type:Income", "fromPeriod": "FYS", "toPeriod": "CM"},
                             {"account": "type:Income", "fromPeriod": led.shift_period("CM", "FYS"), "toPeriod": now}])
        self.assertEqual(same[0], same[1])
        self.assertIn("not found", led.balances([{"account": "4000", "fromPeriod": "XYZ"}])[0]["error"])
        out = server.chat_module.ChatService(led).ledger.handle({"action": "period", "base": "Oct 2026", "shift": -12})
        self.assertEqual(out, {"ok": True, "name": "Oct 2025"})

    def test_month_and_ytd_report_rolls_forward(self):
        out = self.ask("monthly P&L for Oct 2026")
        self.assertTrue(out["ok"], out)
        layout = out["actions"][0]["rows"]
        self.assertIn(["Report month", "Oct 2026"], [r[:2] for r in layout])
        report = next(a for a in out["actions"] if a["sheetName"] == "Monthly P&L")["rows"]
        self.assertEqual(report[7][:2], ["Month", "Oct 2026"])
        self.assertEqual(report[9][2:4], ['=TQ.PERIOD($B$8,"CM")', '=TQ.PERIOD($B$8,"LY")'])
        self.assertEqual(report[9][6:8], ['=TQ.PERIOD($B$8,"FYS")', '=TQ.PERIOD($B$8,"FYS-12")'])
        self.assertIn("the month in B8", out["reply"])
        action = next(a for a in out["actions"] if a["sheetName"] == "Monthly P&L")
        self.assertIn((7, 1), [(d["r"], d["c"]) for d in action["dropdowns"]])     # B8 picks a month
        income = lambda a, b: self.total("type:Income", a, b)  # noqa: E731
        book = Workbook(report, self.ledger)
        row = book.find("Gross Profit")
        sales = lambda col: book.value("%s%d" % (col, row))  # noqa: E731
        gp = lambda a, b: income(a, b) - self.total("type:COGS", a, b)  # noqa: E731
        self.assertAlmostEqual(sales("C"), gp("Oct 2026", "Oct 2026"), places=2)
        self.assertAlmostEqual(sales("D"), gp("Oct 2025", "Oct 2025"), places=2)
        self.assertAlmostEqual(sales("G"), gp("Jan 2026", "Oct 2026"), places=2)
        self.assertAlmostEqual(sales("H"), gp("Jan 2025", "Oct 2025"), places=2)
        report[7][1] = "Mar 2026"                    # roll forward: change B8 only
        book = Workbook(report, self.ledger)
        self.assertAlmostEqual(book.value("G%d" % row), gp("Jan 2026", "Mar 2026"), places=2)
        self.assertAlmostEqual(book.value("H%d" % row), gp("Jan 2025", "Mar 2025"), places=2)
        # The layout keeps it: rebuilding gives the same month; the balance sheet template rolls too.
        again = self.chat.design({"op": "build", "args": {}, "context": self.context(layout)})
        self.assertEqual(next(a for a in again["actions"] if a["sheetName"] == "Monthly P&L")["rows"][7][1], "Oct 2026")
        bs = self.chat.design({"op": "new", "args": {"template": "balance"}, "context": {"layouts": []}})
        sheet = next(a for a in bs["actions"] if a["sheetName"] == "Balance Sheet")["rows"]
        self.assertEqual(sheet[7][0], "Month")
        check = Workbook(sheet, self.ledger)
        now = self.ledger.periods()[self.ledger.current_index()]["name"]
        at = check.find("Bank and Cash")
        for col, period in (("C", now), ("D", self.ledger.shift_period(now, "LY")),
                            ("E", self.ledger.shift_period(now, "FYS-1"))):
            want = self.ledger.balances([{"account": "type:Bank", "fromPeriod": period}])[0]["value"]
            self.assertAlmostEqual(check.value("%s%d" % (col, at)), want, places=2)

    def test_custom_highlight_rules(self):
        parse = server.chat_module.layouts.parse_rule
        self.assertEqual(parse("6100 > 5000 red")["op"], ">")
        rule = parse("all between -100 and 100 yellow")
        self.assertEqual((rule["v1"], rule["v2"], rule["colour"]), (-100.0, 100.0, "yellow"))
        self.assertTrue(parse("Net Profit in Change % below -10% amber")["percent"])
        with self.assertRaises(server.chat_module.reports.ReportError):
            parse("make it pretty")
        layout = self.ask("custom P&L for 2026 by quarter")["actions"][0]["rows"]
        out = self.ask("highlight 6000 > 5000 red", rows=layout)
        self.assertIn("Highlight rule added: 6000 > 5000 red", out["reply"])
        report = out["actions"][1]
        book_rows = report["rows"]
        rent = next(n for n, r in enumerate(book_rows) if r[0] == "6000")
        rules = report["conditional"]
        self.assertEqual({(c["r0"], c["r1"]) for c in rules}, {(rent, rent)})
        self.assertEqual(rules[0]["formula"], "=AND(ISNUMBER(C%d),C%d>5000.0)" % (rent + 1, rent + 1))
        self.assertEqual((rules[0]["font"], rules[0]["fill"]), ("#B3261E", "#FCE8E6"))
        out = self.ask("highlight Gross Profit in Total < 0 yellow", rows=out["actions"][0]["rows"])
        report = out["actions"][1]
        total = [j for j, c in enumerate(out["design"]["columns"]) if c["label"] == "Total"][0]
        net = [c for c in report["conditional"] if "<0.0" in c["formula"]][0]
        self.assertEqual((net["c0"], net["c1"]), (2 + total, 2 + total))
        self.assertEqual(out["design"]["rules"], ["6000 > 5000 red", "Gross Profit in Total < 0 yellow"])
        parsed = server.chat_module.layouts.parse("TQ Layout", out["actions"][0]["rows"])
        self.assertEqual(parsed["rules"], out["design"]["rules"])
        self.assertIn("no line or account", self.ask("highlight Nothing Here > 1 red", rows=out["actions"][0]["rows"])["reply"])
        off = self.ask("no highlight rules", rows=out["actions"][0]["rows"])
        self.assertEqual(off["design"]["rules"], [])
        out = self.chat.design({"op": "addRule", "args": {"rule": "all between -100 and 100 grey"},
                                "context": self.context(off["actions"][0]["rows"])})
        self.assertTrue(len(out["actions"][1]["conditional"]) >= 1)

    def test_report_wizard(self):
        wiz = lambda **args: self.chat.design({"op": "wizard", "args": args, "context": {"layouts": [], "sheets": []}})  # noqa: E731
        out = wiz(template="monthytd", month="Sep 2026", filters={"location": "Doha"}, pages="location",
                  comments=True, highlight="10", dropdowns=["Brand"])
        self.assertTrue(out["ok"], out)
        names = [a["sheetName"] for a in out["actions"]]
        self.assertEqual(names[:2], ["TQ Layout", "Monthly P&L"])
        report = out["actions"][1]
        self.assertEqual(report["rows"][7][:2], ["Month", "Sep 2026"])
        self.assertEqual(report["rows"][6][:2], ["Location", "Doha"])
        self.assertIn("conditional", report)
        self.assertIn("keep", report)
        self.assertEqual(out["design"]["dropdowns"], ["Brand"])
        out = wiz(template="summary", year=2025, shape="quarter", compare="year", title="Board P&L")
        cols = [c["label"] for c in out["design"]["columns"]]
        self.assertEqual(cols[:2], ["Q1 2025", "Q1 2024"]) if "Q1 2024" in cols else self.assertIn("Q1 2025", cols)
        self.assertEqual(out["actions"][1]["sheetName"], "Board P&L")
        out = wiz(template="management", year=2026, shape="total", compare="budget")
        self.assertTrue(any(c["data"] == "budget" for c in out["design"]["columns"]))
        out = wiz(template="balance", month="Dec 2025")
        self.assertEqual(out["actions"][1]["rows"][7][1], "Dec 2025")

    def test_variance_highlights_and_comments(self):
        out = self.ask("monthly P&L for Oct 2026")
        layout = out["actions"][0]["rows"]
        out = self.ask("highlight changes over 10%", rows=layout)
        report = next(a for a in out["actions"] if a["sheetName"] == "Monthly P&L")
        self.assertEqual([c["c0"] for c in report["conditional"]], [5, 9])          # Change % and Change % YTD
        self.assertEqual(report["conditional"][0]["limit"], 0.1)
        self.assertTrue(report["conditional"][0]["cell"].startswith("F"))
        self.assertIn(["Highlight changes over", "10%"], [r[:2] for r in out["actions"][0]["rows"]])
        out = self.ask("add comments column", rows=out["actions"][0]["rows"])
        report = next(a for a in out["actions"] if a["sheetName"] == "Monthly P&L")
        header = report["keep"]["header"]
        self.assertEqual(report["rows"][header][10], "Comments")
        self.assertEqual(report["keep"], {"header": header, "title": "Comments", "key": 1})
        parsed = server.chat_module.layouts.parse("TQ Layout", out["actions"][0]["rows"])
        self.assertEqual((parsed["highlight"], parsed["comments"]), (10.0, True))
        off = self.ask("no highlights", rows=out["actions"][0]["rows"])
        self.assertNotIn("conditional", next(a for a in off["actions"] if a["sheetName"] == "Monthly P&L"))
        self.assertEqual(server.chat_module.layouts.highlight_value("0.15"), 15.0)
        with self.assertRaises(server.chat_module.reports.ReportError):
            server.chat_module.layouts.highlight_value("lots")

    def test_sheets_per_location_when_this_year_has_none(self):
        """The P&L summary template covers this year. When no 2026 posting has a location, the sheets are made
        for the locations with a balance in other years (the report follows the year in B8)."""
        made = self.chat.design({"op": "new", "args": {"template": "summary"}, "context": {"layouts": []}})
        layout = made["actions"][0]["rows"]
        real = self.ledger.balances

        def no_2026_locations(requests):
            out = real(requests)
            return [{"value": 0.0} if r.get("location") and r["fromPeriod"].endswith("2026") else o
                    for r, o in zip(requests, out)]

        self.ledger.balances = no_2026_locations
        try:
            out = self.chat.design({"op": "setPages", "args": {"pages": "location"}, "context": self.context(layout)})
        finally:
            self.ledger.balances = real
        self.assertTrue(out["ok"], out)
        names = [a["sheetName"] for a in out["actions"]]
        self.assertEqual(names[:3], ["TQ Layout", "P&L Summary", "P&L Summary - Summary"])
        self.assertIn("Doha - P&L Summary", names)
        self.assertIn("change the year in B8", out["reply"])
        self.assertIn("The report now has one sheet per location", out["reply"])

    def test_budget_columns_and_account_detail(self):
        out = self.ask("custom P&L for 2026 by quarter vs budget")
        rows = out["actions"][1]["rows"]
        self.assertEqual(rows[11][2:6], ['="Q1 "&$B$8', '="Budget Q1 "&$B$8', '="Var Q1 "&$B$8', '="Var % Q1 "&$B$8'])
        book = Workbook(rows, self.ledger)
        row = book.find("Net Sales")
        actual = self.total("type:Income", "Jan 2026", "Mar 2026")
        budget = self.total("budget:type:Income", "Jan 2026", "Mar 2026")
        self.assertNotEqual(actual, budget)
        self.assertAlmostEqual(book.value("D%d" % row), budget, places=2)
        self.assertAlmostEqual(book.value("E%d" % row), actual - budget, places=2)
        self.assertAlmostEqual(book.value("F%d" % row), (actual - budget) / budget, places=6)
        self.assertEqual(out["actions"][1]["dropdowns"][0], {"r": 7, "c": 1, "formula": "2025,2026"})  # the year picker
        out = self.ask("add column budget 2026", rows=out["actions"][0]["rows"])
        self.assertEqual(out["design"]["columns"][-1]["data"], "budget")
        self.assertIn('Added column "Budget 2026"', out["reply"])
        # Show accounts: the accounts behind a line that covers several, each on its own row.
        wide = [r[:] for r in LAYOUT]
        wide[5][1:4] = ["All revenue", "accounts", "4000-4999"]
        del wide[6]
        out = self.ask("show GL codes", rows=self.ask("custom P&L 2026 by quarter", rows=wide)["actions"][0]["rows"])
        rows = out["actions"][1]["rows"]
        at = Workbook(rows, self.ledger).find("All revenue")
        self.assertEqual([r[:2] for r in rows[at:at + 2]], [["4000", "Product Revenue"], ["4100", "Service Revenue"]])
        book = Workbook(rows, self.ledger)
        self.assertAlmostEqual(book.value("C%d" % (at + 2)), self.total("4100", "Jan 2026", "Mar 2026"), places=2)
        self.assertAlmostEqual(book.value("C%d" % book.find("Net Sales")), self.total("type:Income", "Jan 2026", "Mar 2026"),
                               places=2)   # the subtotal adds the line once, not its accounts again
        self.assertTrue(out["design"]["detail"])
        self.assertFalse(self.ask("hide account detail", rows=out["actions"][0]["rows"])["design"]["detail"])

    def test_templates_build_reports(self):
        listed = self.chat.design({"op": "templates"})["templates"]
        self.assertEqual([t["id"] for t in listed], ["monthytd", "management", "summary", "balance", "cashflow", "hotel", "restaurant", "forecast", "rolling12"])
        for template in ("monthytd", "summary", "balance", "cashflow"):
            out = self.chat.design({"op": "new", "args": {"template": template}, "context": {"sheets": []}})
            self.assertTrue(out["ok"], out)
            self.assertEqual(out["actions"][0]["sheetName"], "TQ Layout")
            book = Workbook(out["actions"][1]["rows"], self.ledger)
            if template == "summary":   # the summary's net profit is the ledger's net income
                at = book.find("Net Profit")
                want = self.ledger.balances([{"account": "calc:NetIncome", "fromPeriod": "Jan 2026"}])[0]["value"]
                self.assertAlmostEqual(book.value("C%d" % at), want, places=2)
            if template == "cashflow":  # a movement is the closing balance less the opening one
                at = book.find("Movement in Bank Accounts")
                ends = self.ledger.balances([{"account": "type:Bank", "fromPeriod": name} for name in ("Jan 2026", "Dec 2025")])
                self.assertAlmostEqual(book.value("C%d" % at), ends[0]["value"] - ends[1]["value"], places=2)
        self.assertIn("no template", self.chat.design({"op": "new", "args": {"template": "nope"}, "context": {}})["error"])

    def test_layout_from_a_format_on_a_sheet(self):
        rows = [["My Company P&L", "", ""], ["Name", "Jan", "Feb"], ["Revenue", "", ""],
                ["Product Revenue", 10, 20], ["Service Revenue", 5, 6], ["Total Revenue", 15, 26],
                ["Cost of Sales", 3, 4], ["Gross Profit", 12, 22], ["GP %", 0.8, 0.8], ["Staff", "", ""],
                ["Salaries and Wages", 1, 1], ["Rent", 1, 1], ["Marketing", 1, 1], ["Bonus pool", 1, 1],
                ["Total Expenses", 3, 3], ["Net Profit", 9, 19]]
        out = self.chat.design({"op": "new", "args": {"rows": rows, "sheet": "My format"},
                                "context": {"sheets": ["My format"]}})
        self.assertTrue(out["ok"], out)
        lines = {l["label"]: l for l in out["design"]["lines"]}
        self.assertEqual(lines["Product Revenue"]["source"], "4000")
        self.assertEqual((lines["Rent"]["source"], lines["Rent"]["sign"], lines["Rent"]["section"]), ("6100", -1, "Staff"))
        self.assertEqual(lines["Total Revenue"]["kind"], "subtotal")
        self.assertEqual(lines["Net Profit"]["source"], "=[Total Revenue]+[Cost of Sales]+[Total Expenses]")
        self.assertNotIn("Revenue", lines)          # a heading, not a line
        self.assertIn("Bonus pool", out["reply"])   # said, not guessed
        self.assertNotEqual(out["actions"][1]["sheetName"], "My format")  # the user's sheet is not overwritten
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        want = self.ledger.balances([{"account": "calc:NetIncome", "fromPeriod": "Jan 2026"}])[0]["value"]
        self.assertAlmostEqual(book.value("C%d" % book.find("Net Profit")), want, places=2)
        coded = self.chat.design({"op": "new", "args": {"sheet": "Coded", "rows": [
            ["Code", "Line"], ["4000-4100", "Sales"], ["5000", "Food cost"], ["", "Gross Profit"]]}, "context": {}})
        self.assertEqual([l["source"] for l in coded["design"]["lines"]], ["4000-4100", "5000", "=[Sales]+[Food cost]"])
        empty = self.chat.design({"op": "new", "args": {"sheet": "X", "rows": [[1, 2], [3, 4]]}, "context": {}})
        self.assertIn("no column of row names", empty["error"])

    def test_designer_requests(self):
        rows = self.ask("custom P&L for 2026 by quarter")["actions"][0]["rows"]
        design = lambda op, rows, **args: self.chat.design(  # noqa: E731
            {"op": op, "args": args, "context": self.context(rows)})
        out = design("addColumn", rows, label="Doha H1", **{"from": "Jan 2026", "to": "jun 2026",
                                                             "filters": {"location": "doha"}})
        self.assertEqual(out["design"]["columns"][-1],
                         {"label": "Doha H1", "type": "period", "from": "Jan 2026", "to": "Jun 2026",
                          "filters": {"location": "Doha"}, "formula": "", "format": "number", "data": ""})
        self.assertEqual([a["sheetName"] for a in out["actions"]], ["TQ Layout"])
        out = design("addLine", out["actions"][0]["rows"], label="Rent", kind="accounts", source="6100",
                     negative=True, section="Staff Cost")
        self.assertEqual(out["design"]["lines"][-1]["sign"], -1)
        out = design("moveLine", out["actions"][0]["rows"], label="Rent", by=-1)
        self.assertEqual([l["label"] for l in out["design"]["lines"]][-2:], ["Rent", "Salaries"])
        self.assertIn('"77" not found', design("addLine", rows, label="Bad", source="77")["error"])
        self.assertIn("was not found", design("addColumn", rows, label="X", **{"from": "Jan 2026",
                                                                               "filters": {"class": "Nope"}})["error"])
        self.assertEqual(design("get", rows)["design"]["title"], "Restaurant P&L")
        self.assertEqual(design("new", rows, name="Hotels")["actions"][0]["sheetName"], "TQ Layout Hotels")
        self.assertEqual(design("new", rows)["actions"][0]["sheetName"], "TQ Layout 2")
        built = design("build", out["actions"][0]["rows"])
        self.assertEqual([a["sheetName"] for a in built["actions"]], ["TQ Layout", "Restaurant P&L"])
        # With the report open, a request means the layout that report was built from.
        other = {"name": "TQ Layout Old", "row": 0, "rows": [r[:] for r in LAYOUT[:1]] + [["Report title", "Old"]] + LAYOUT[2:]}
        context = {"activeSheet": "Restaurant P&L", "layouts": [other, {"name": "TQ Layout", "row": 0, "rows": rows}]}
        self.assertEqual(self.chat.design({"op": "get", "context": context})["design"]["name"], "TQ Layout")

    def test_layout_mistakes_are_explained(self):
        first = self.ask("custom P&L 2026", rows=None)  # no layout yet: one is made, and the report with it
        self.assertIn("had no layout", first["reply"])
        self.assertEqual(first["actions"][0]["sheetName"], "TQ Layout")
        self.assertGreater(len(first["actions"]), 1)
        self.assertIn("no layout sheet", self.ask("show account detail", rows=None)["reply"])
        bad = LAYOUT + [["", "Oops", "formula", "=[Nett Sales]*2", "", ""]]
        self.assertIn("row 14: the formula uses [Nett Sales]", self.ask("custom P&L 2026", rows=bad)["reply"])
        bad = LAYOUT + [["", "Oops", "accounts", "4000, 77", "", ""]]
        self.assertIn('row 14: Account "77" not found', self.ask("custom P&L 2026", rows=bad)["reply"])
        self.assertIn("no header row", self.ask("custom P&L 2026", rows=[["hello"]])["reply"])
        self.assertIn("Say which period", self.ask("add column please")["reply"])

    def test_sample_layout_builds_a_report(self):
        sample = self.ask("create layout", rows=None)["actions"][0]
        self.assertEqual(sample["sheetName"], "TQ Layout")
        out = self.ask("update report", rows=sample["rows"])   # the sample's own columns: this year by month
        self.assertIn("Every profit and loss account", out["reply"])
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        self.assertAlmostEqual(book.value("N%d" % book.find("Net Profit")),
                               self.total("calc:NetIncome", "Jan 2026", "Nov 2026"), places=2)
        out = self.ask("custom P&L 2026 by quarter", rows=sample["rows"])
        book = Workbook(out["actions"][1]["rows"], self.ledger)
        self.assertAlmostEqual(book.value("G%d" % book.find("Net Profit")),
                               self.total("calc:NetIncome", "Jan 2026", "Dec 2026"), places=2)


class QueryTest(unittest.TestCase):
    """Data queries: saved searches and SuiteQL, each filling its own sheet."""

    def setUp(self):
        self.chat = server.chat_module.ChatService(new_ledger())
        self.queries = server.chat_module.queries

    def ask(self, op, sheet=None, **args):
        context = {"queries": {"row": 0, "rows": sheet}} if sheet else {}
        return self.chat.queries.handle({"op": op, "args": args, "context": context})

    def test_saved_search_onto_its_own_sheet(self):
        self.assertEqual([s["id"] for s in self.ask("searches", text="invoice")["searches"]], ["customsearch_open_invoices"])
        out = self.ask("save", name="Open invoices", type="search", source="customsearch_open_invoices")
        listing, table = out["actions"]
        self.assertEqual((listing["sheetName"], listing["activate"]), ("TQ Queries", False))
        self.assertEqual((table["sheetName"], table["target"], table["activate"]), ("Open invoices", "sheet", True))
        self.assertEqual(table["rows"][0], ["Id", "Tranid", "Type", "Trandate", "Entity", "Foreigntotal"])
        self.assertEqual(table["rows"][1], [9001, "INV-1001", "CustInvc", "2026-01-14", 101, 12500.0])
        self.assertEqual(table["rows"][3][1], "'=CM-7")  # never becomes a formula in Excel
        self.assertIn('3 row(s) on sheet "Open invoices"', out["reply"])
        self.assertEqual(out["state"]["queries"][0]["rows"], 3)
        # The listing sheet is all it needs to run again.
        taken = self.chat.queries.handle({"op": "save", "context": {"sheets": ["Sheet1", "Budget"]},
                                          "args": {"name": "budget", "source": "SELECT id FROM item"}})
        self.assertIn('already has a sheet called "budget"', taken["error"])
        again = self.ask("run", sheet=listing["rows"])
        self.assertEqual([a["sheetName"] for a in again["actions"]], ["TQ Queries", "Open invoices"])

    def test_sql_with_parameters_and_limits(self):
        sql = "SELECT id, entityid FROM customer WHERE datecreated >= TO_DATE({from}, 'YYYY-MM-DD')"
        out = self.ask("save", name="Customers", source=sql, max="2")
        self.assertIn("{from}, which has no value", out["reply"])
        sheet = self.ask("setParam", sheet=out["actions"][0]["rows"], name="From", value="2025-01-01")["actions"][0]["rows"]
        out = self.ask("run", sheet=sheet, name="customers")
        self.assertIn("2 row(s)", out["reply"])
        self.assertIn("cut at the limit of 2", out["reply"])
        self.assertEqual(out["actions"][1]["rows"][1][1], "C-0001")
        self.assertIn("must start with SELECT", self.ask("save", name="Bad", type="sql", source="DELETE FROM x")["error"])
        self.assertIn("cannot be the output sheet", self.ask("save", name="TQ Queries", source="SELECT 1 FROM item")["error"])
        self.assertIn("was not found", self.ask("save", name="Nope", source="SELECT id FROM nothing")["reply"])
        out = self.ask("remove", sheet=sheet, name="Customers")
        self.assertEqual(out["state"]["queries"], [])
        self.assertEqual(self.ask("columns", table="item")["columns"], ["id", "itemid", "displayname", "itemtype"])
        self.assertEqual(self.queries._value("00417"), "'00417")  # a code with leading zeros stays text

    def test_query_built_by_clicking(self):
        sql = self.ask("compose", spec={
            "table": "transaction",
            "fields": [{"field": "tranid"}, {"field": "entity", "show": "name"}, {"field": "customer.companyname"},
                       {"field": "foreigntotal", "total": "sum", "label": "Total amount"}],
            "joins": [{"table": "customer", "left": "transaction.entity", "right": "id"}],
            "filters": [{"field": "type", "op": "=", "value": "CustInvc"},
                        {"field": "trandate", "op": ">=", "value": "{from}"},
                        {"field": "customer.companyname", "op": "contains", "value": "O'Neil"}],
            "sort": [{"field": "trandate", "desc": True}]})["sql"]
        self.assertEqual(sql.split("\n"), [
            "SELECT transaction.tranid AS tranid, BUILTIN.DF(transaction.entity) AS entity_name, "
            "customer.companyname AS companyname, SUM(transaction.foreigntotal) AS total_amount",
            "FROM transaction",
            "LEFT JOIN customer ON customer.id = transaction.entity",
            "WHERE transaction.type = 'CustInvc' AND transaction.trandate >= {from} "
            "AND LOWER(customer.companyname) LIKE '%' || LOWER('O''Neil') || '%'",
            "GROUP BY transaction.tranid, BUILTIN.DF(transaction.entity), customer.companyname",
            "ORDER BY transaction.trandate DESC"])
        self.assertIn("not a field name", self.ask("compose", spec={"table": "item", "fields": [{"field": "id; DROP"}]})["error"])
        self.assertTrue(any(c["table"] == "customer" for c in self.ask("catalog")["catalog"]))

    def test_chat_commands(self):
        say = lambda text, sheet=None: self.chat.handle(  # noqa: E731
            {"message": text, "context": {"queries": {"row": 0, "rows": sheet}} if sheet else {}})
        self.assertIn("Customer List  (customsearch_customer_list, Customer)", say("list saved searches")["reply"])
        out = say("load search customer list")
        self.assertEqual([a["sheetName"] for a in out["actions"]], ["TQ Queries", "Customer List"])
        self.assertIn("No saved search matched", say("load search payroll")["reply"])
        self.assertIn("no data queries yet", say("refresh data")["reply"])
        self.assertIn('Customer List: 3 row(s)', say("refresh data", out["actions"][0]["rows"])["reply"])


class ChatTest(unittest.TestCase):
    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def reply(self, message):
        return self.chat.handle({"message": message})

    def balance(self, account, start, end=""):
        return self.ledger.balances([{"account": account, "fromPeriod": start, "toPeriod": end}])[0]["value"]

    def test_quick_answers(self):
        self.assertIn("Rent (6100), Mar 2026: {:,.2f}".format(self.balance("6100", "Mar 2026")),
                      self.reply("What was rent in March 2026?")["reply"])
        # 2000 is an account here, not a year.
        self.assertIn("Accounts Payable (2000), as at Jan 2026", self.reply("balance of 2000 for jan 2026")["reply"])
        out = self.reply("net income this year")
        self.assertIn("{:,.2f}".format(self.balance("calc:NetIncome", "Jan 2026", "Nov 2026")), out["reply"])
        self.assertEqual(out["actions"], [])
        out = self.reply("total revenue q1 2026 here")
        self.assertEqual(out["actions"][0]["target"], "selection")
        self.assertIn('TQ.BALANCE("type:Income","Jan 2026","Mar 2026"', out["actions"][0]["rows"][0][2])

    def test_defaults_are_said(self):
        out = self.reply("p&l")
        self.assertIn("No period was named", out["reply"])
        self.assertEqual((out["spec"]["from"], out["spec"]["to"], out["spec"]["layout"]),
                         ("Jan 2026", "Nov 2026", "month"))

    def test_lists_and_transactions(self):
        self.assertEqual(len(self.reply("list accounts")["actions"][0]["rows"]), len(server.MOCK_ACCOUNTS) + 1)
        out = self.reply("list periods 2026 here")
        self.assertEqual((out["actions"][0]["target"], len(out["actions"][0]["rows"])), ("selection", 13))
        self.assertIn("4 transaction line(s)", self.reply("transactions for 6100 in Jan 2026")["reply"])
        self.assertIn("Name one account", self.reply("transactions for expenses in Jan 2026")["reply"])

    def test_things_it_does_not_understand(self):
        self.assertIn("did not understand", self.reply("what is the weather")["reply"])
        self.assertIn("Which report", self.reply("by quarter")["reply"])
        self.assertIn("no accounting period", self.reply("p&l for mar 2031")["reply"])
        self.assertIn("Reports", self.reply("help")["reply"])
        self.assertFalse(self.reply("   ")["ok"])


class FilterTest(unittest.TestCase):
    """Any filter in a report, several values per segment, and consolidation."""

    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def value(self, **request):
        out = self.ledger.balances([dict({"account": "4000", "fromPeriod": "Jan 2026", "toPeriod": "Mar 2026"}, **request)])[0]
        return out.get("value", out.get("error"))

    def test_several_values_and_consolidation(self):
        each = [self.value(subsidiary=name) for name in ("Parent Company", "Qatar", "Saudi Arabia")]
        self.assertAlmostEqual(self.value(subsidiary="Parent Company*"), sum(each), places=2)
        self.assertAlmostEqual(self.value(subsidiary="Parent Company (consolidated)"), sum(each), places=2)
        self.assertAlmostEqual(self.value(subsidiary="Qatar, Saudi Arabia"), each[1] + each[2], places=2)
        self.assertAlmostEqual(self.value(subsidiary="not Qatar"), each[0] + each[2], places=2)
        self.assertIn("not found", self.value(subsidiary="Narnia"))
        self.assertNotEqual(self.value(filters="Brand: Premium"), self.value())
        self.assertIn('Filter field "Colour" not found', self.value(filters="Colour: Red"))

    def test_a_netsuite_hiccup_is_not_taken_for_old_scripts(self):
        calls = []

        def flaky(payload):
            if payload.get("action") == "setup":
                calls.append(1)
                if len(calls) == 1:
                    return {"ok": False, "error": "NetSuite is busy"}
            return server.MockBackend().handle(payload)
        ledger = server.ledger_module.Ledger(flaky, today=TODAY)
        first = ledger.balances([{"account": "4000", "fromPeriod": "Jan 2026", "filters": "Brand: Premium"}])[0]
        self.assertIn("busy", first["error"])
        second = ledger.balances([{"account": "4000", "fromPeriod": "Jan 2026", "filters": "Brand: Premium"}])[0]
        self.assertIn("value", second)

    def test_old_netsuite_scripts_refuse_new_filters(self):
        def old(payload):
            if payload.get("action") == "setup":
                return {"ok": False, "error": 'Unknown action "setup"'}
            return server.MockBackend().handle(payload)
        ledger = server.ledger_module.Ledger(old, today=TODAY)
        out = ledger.balances([{"account": "4000", "fromPeriod": "Jan 2026", "department": "Sales, Finance"},
                               {"account": "4000", "fromPeriod": "Jan 2026", "department": "Sales"},
                               {"account": "4000", "fromPeriod": "Jan 2026", "filters": "Brand: Premium"}])
        self.assertIn("newer NetSuite scripts", out[0]["error"])
        self.assertIn("value", out[1])
        self.assertIn("newer NetSuite scripts", out[2]["error"])
        drill = ledger.handle({"action": "transactions", "request": {"account": "4000", "fromPeriod": "Jan 2026",
                                                                    "filters": "Brand: Premium"}})
        self.assertIn("newer NetSuite scripts", drill["error"])
        fields = ledger.handle({"action": "filterFields"})["fields"]
        self.assertEqual([f["id"] for f in fields], ["subsidiary", "department", "class", "location"])

    def test_reports_carry_the_other_filters_cell(self):
        out = self.chat.handle({"message": "P&L for 2026 by quarter filter Brand: Premium", "context": {}})
        rows = out["actions"][0]["rows"]
        self.assertEqual(rows[2][:2], ["Other filters", "Brand: Premium"])
        self.assertTrue(all(cell.endswith(",$B$3)") for row in rows for cell in row
                            if isinstance(cell, str) and cell.startswith("=TQ.BALANCE(")))
        book = Workbook(rows, self.ledger)
        row = book.find("Net Income")
        expected = self.ledger.balances([{"account": "calc:NetIncome", "fromPeriod": "Jan 2026", "toPeriod": "Mar 2026",
                                          "filters": "Brand: Premium"}])[0]["value"]
        self.assertAlmostEqual(book.value("C%d" % row), expected, places=2)
        self.assertEqual(out["spec"]["filters"], {"filters": "Brand: Premium"})
        for message, wanted in (("P&L filter Brand: Premium for 2026 by quarter", "Brand: Premium"),
                                ("P&L for 2026 filter Brand: Premium by quarter", "Brand: Premium"),
                                ("P&L for 2026 by quarter filter Brand: Premium, Basic; Name: Al Noor",
                                 "Brand: Premium, Basic; Name: Al Noor")):
            spec = self.chat.handle({"message": message, "context": {}})["spec"]
            self.assertEqual(spec["filters"]["filters"], wanted, message)
            self.assertEqual(spec["from"], "Jan 2026", message)
        spec = self.chat.handle({"message": "P&L for 2026 filter Location: not Doha summary", "context": {}})["spec"]
        self.assertEqual(spec["filters"], {"location": "not Doha"})        # its own cell, B7
        self.assertEqual(spec["from"], "Jan 2026")

    def test_consolidated_balance_sheet_balances(self):
        out = self.chat.handle({"message": "balance sheet as at Oct 2026 for subsidiary Parent Company consolidated",
                                "context": {}})
        rows = out["actions"][0]["rows"]
        self.assertEqual(rows[3][:2], ["Subsidiary", "Parent Company*"])
        book = Workbook(rows, self.ledger)
        self.assertTrue(book.find("Currency translation (consolidation)"))
        check = book.find("Check: assets less liabilities and equity (should be 0)")
        self.assertEqual(book.value("C%d" % check), 0)
        one = self.chat.handle({"message": "balance sheet as at Oct 2026 for subsidiary Qatar", "context": {}})["actions"][0]["rows"]
        book = Workbook(one, self.ledger)
        self.assertEqual(book.value("C%d" % book.find("Currency translation (consolidation)")), 0)  # one subsidiary: own currency
        tb = self.chat.handle({"message": "trial balance Oct 2026 for subsidiary Parent Company consolidated",
                               "context": {}})["actions"][0]["rows"]
        book = Workbook(tb, self.ledger)
        check = next(n + 1 for n, r in enumerate(tb) if str(r[1]).startswith("Check"))
        self.assertEqual(book.value("D%d" % check), 0)

    def test_layouts_keep_other_filters(self):
        out = self.chat.handle({"message": "custom P&L for 2026 by quarter", "context": {}})
        design = out["design"]
        sheet = server.chat_module.layouts.emit(dict(design, filters={"filters": "Brand: Premium", "department": "Sales, Finance"},
                                                 lines=[server.chat_module.layouts.new_line(l["label"], l["kind"], l["source"], l["sign"], l["format"], l["section"])
                                                        for l in design["lines"]],
                                                 columns=[server.chat_module.layouts.new_column(c["label"], c["type"], c["from"], c["to"],
                                                                                           dict(c["filters"], filters="Name: Al Noor") if n == 0 else c["filters"],
                                                                                           c["formula"], c["format"], data=c["data"])
                                                          for n, c in enumerate(design["columns"])]))
        parsed = server.chat_module.layouts.parse("TQ Layout", sheet["rows"])
        self.assertEqual(parsed["filters"], {"filters": "Brand: Premium", "department": "Sales, Finance"})
        self.assertEqual(parsed["columns"][0]["filters"], {"filters": "Name: Al Noor"})
        action, info = server.chat_module.layouts.build_report(self.ledger, parsed)
        rows = action["rows"]
        self.assertEqual(rows[2][:2], ["Other filters", "Brand: Premium"])
        self.assertEqual(rows[4][:2], ["Department", "Sales, Finance"])
        self.assertTrue(any(r[1] == "Other filters" and r[2] == "Name: Al Noor" for r in rows))


class InstallTest(unittest.TestCase):
    """The licence, the setup page and this PC's own certificate."""

    def setUp(self):
        self.key = server.tqcrypto.generate_key(1024)
        self.saved = sys.modules.get("licence_key")
        fake = type(sys)("licence_key")
        fake.PUBLIC_KEY = {"n": self.key["n"], "e": self.key["e"]}
        sys.modules["licence_key"] = fake
        self.folder = os.path.join(os.path.dirname(__file__), "_setup_test")
        os.makedirs(self.folder, exist_ok=True)

    def tearDown(self):
        if self.saved is not None:
            sys.modules["licence_key"] = self.saved
        else:
            sys.modules.pop("licence_key", None)
        for name in os.listdir(self.folder):
            os.remove(os.path.join(self.folder, name))
        os.rmdir(self.folder)

    def licence(self, **change):
        made = server.licence_module.make(self.key, "World of Hospitality", "1234567_sb1", 5, "2027-12-31", today=TODAY)
        return json.dumps(dict(made, **change))

    def test_seats_per_pc(self):
        answers = []

        def netsuite(payload):
            answers.append(payload["action"])
            if payload["action"] == "seat":
                return {"ok": True, "allowed": payload["machine"] != "pc6", "used": 5, "users": 5,
                        "machines": [{"label": "PC1", "user": "Ali", "last": "2026-11-01"}]}
            return {"ok": True, "rows": []}
        seats = server.updates_module.Seats(netsuite, {"id": "pc1", "label": "PC1"}, 5)
        self.assertTrue(seats({"action": "accounts"})["ok"])
        seats({"action": "periods"})                                  # checked once, not on every call
        self.assertEqual(answers, ["seat", "accounts", "periods"])
        self.assertIn("5 of 5 in use", seats.state["message"])
        full = server.updates_module.Seats(netsuite, {"id": "pc6", "label": "PC6"}, 5)
        out = full({"action": "accounts"})
        self.assertIn("All 5 seats of the licence are in use: PC1 (Ali, last 2026-11-01)", out["error"])
        self.assertTrue(full({"action": "ping"})["ok"])               # the connection test still works
        old = server.updates_module.Seats(lambda p: {"ok": False, "error": 'Unknown action "seat"'}
                                          if p["action"] == "seat" else {"ok": True}, {"id": "x"}, 5)
        self.assertTrue(old({"action": "accounts"})["ok"])             # older NetSuite scripts: not counted
        per_user = server.updates_module.Seats(netsuite, {"id": "pc1", "label": "PC1"}, 5, by_user=True)
        sent = []
        per_user.backend = lambda p: sent.append(p) or {"ok": True, "allowed": True, "used": 1, "users": 5}
        per_user({"action": "accounts"})
        self.assertTrue(sent[0]["byUser"])                            # signed in: the seat is the person's
        self.assertTrue(full({"action": "seatList"})["ok"])           # the list still opens when the seats are full
        pc = server.updates_module.machine(self.folder)
        self.assertEqual(server.updates_module.machine(self.folder)["id"], pc["id"])   # kept

    def test_signed_updates(self):
        u = server.updates_module
        good = u.sign_manifest(self.key, "9.9.0", "https://example.com/TQ.exe", "a" * 64, "New things")
        u.verify(good)
        with self.assertRaises(ValueError):
            u.verify(dict(good, url="https://evil.example.com/x.exe"))   # the signature covers the address
        with self.assertRaises(ValueError):
            u.verify(dict(good, signature=base64.b64encode(b"x" * 128).decode()))
        saved = u._fetch
        try:
            u._fetch = lambda url, limit: json.dumps(good).encode() if url.endswith("update.json") else b"setup bytes"
            found = u.check("https://example.com/update.json", "1.0.0")
            self.assertTrue(found["available"])
            self.assertEqual(u.check("https://example.com/update.json", "9.9.0")["available"], False)
            with self.assertRaises(ValueError):          # the download must match the signed sha256
                u.install("https://example.com/update.json", "1.0.0", run=False)
            good = u.sign_manifest(self.key, "9.9.0", "https://example.com/TQ.exe",
                                   hashlib.sha256(b"setup bytes").hexdigest())
            out = u.install("https://example.com/update.json", "1.0.0", run=False)
            self.assertTrue(os.path.exists(out["path"]))
            os.remove(out["path"])
        finally:
            u._fetch = saved

    def test_weekly_update_watch(self):
        u = server.updates_module
        good = u.sign_manifest(self.key, "9.9.0", "https://example.com/TQ.exe", "a" * 64, "New things")
        saved = u._fetch
        try:
            u._fetch = lambda url, limit: json.dumps(good).encode()
            watch = u.Watcher(lambda: "https://example.com/update.json", "1.0.0", lambda: True)
            self.assertEqual((watch.check()["available"], watch.state["version"]), (True, "9.9.0"))
            self.assertNotIn("manifest", watch.state)
            off = u.Watcher(lambda: "https://example.com/update.json", "1.0.0", lambda: False)
            self.assertFalse(off.check()["available"])                # weekly check turned off by the admin
            u._fetch = lambda url, limit: b"not json"
            self.assertIn("error", watch.check())
        finally:
            u._fetch = saved

    def test_licence_is_checked(self):
        check = server.licence_module.check
        good = json.loads(self.licence())
        self.assertTrue(check(good, "1234567-SB1", TODAY)["ok"])
        self.assertIn("not 7654321", check(good, "7654321", TODAY)["message"])
        self.assertIn("ended on 2027-12-31", check(good, "1234567_SB1", datetime.date(2028, 1, 1))["message"])
        self.assertIn("signature does not match", check(dict(good, users=50), "1234567_SB1", TODAY)["message"])
        self.assertIn("No licence", check(None, "1234567_SB1", TODAY)["message"])

    def test_only_the_account_id_is_needed(self):
        """Sign in with Techno Q's integration and a RESTlet address made from the account: nothing else to type."""
        cfg = server.check_config({"accountId": "1234567_sb1"})
        self.assertEqual(cfg["accountId"], "1234567_SB1")
        self.assertEqual(cfg["clientId"], server.DEFAULT_CLIENT_ID)
        self.assertEqual(cfg["restletUrl"], "https://1234567-sb1.restlets.api.netsuite.com/app/site/hosting/restlet.nl"
                                            "?script=customscript_tq_excel_restlet&deploy=customdeploy_tq_excel_restlet")
        own = server.check_config({"accountId": "1", "clientId": "mine", "restletUrl": "https://1.restlets.api.netsuite.com/x"})
        self.assertEqual((own["clientId"], own["restletUrl"]), ("mine", "https://1.restlets.api.netsuite.com/x"))
        tba = server.check_config(dict(CFG))   # four keys and no client ID: token-based authentication stays
        self.assertFalse(tba.get("clientId"))

    def test_licence_comes_from_netsuite(self):
        """A new PC without a licence file gets the one an administrator stored in NetSuite."""
        cfg = server.check_config(dict(CFG, accountId="1234567_SB1"))
        stored = {"licence": self.licence(account="1234567_SB1")}
        calls = []
        saved = server.call_netsuite

        def netsuite(cfg, payload):
            calls.append(payload["action"])
            if payload["action"] == "licenceGet":
                return {"ok": True, "licence": stored["licence"]}
            return {"ok": True, "allowed": True, "used": 1, "users": 5}
        server.call_netsuite = netsuite
        try:
            handler = type("H", (server.Handler,), {"config_path": os.path.join(self.folder, "config.json"),
                                                   "licence_path": os.path.join(self.folder, "licence.json")})
            server.configure(handler, cfg, TODAY, licence_text="")
            self.assertFalse(handler.licence["ok"])
            self.assertTrue(server.fetch_licence(handler, force=True))
            self.assertTrue(handler.licence["ok"])
            self.assertTrue(os.path.exists(handler.licence_path))
            calls.clear()
            self.assertFalse(server.fetch_licence(handler, force=True))   # already licensed: NetSuite is not asked again
            self.assertNotIn("licenceSave", calls)                        # NetSuite already has it
            stored["licence"] = self.licence(account="7654321")            # a licence for another account is refused
            os.remove(handler.licence_path)
            server.configure(handler, cfg, TODAY, licence_text="")
            self.assertFalse(server.fetch_licence(handler, force=True))
            self.assertFalse(handler.licence["ok"])
        finally:
            server.call_netsuite = saved

    def test_a_licensed_pc_stores_its_licence_in_netsuite_once(self):
        """A PC signed in before 1.5 never went through the sign-in that shares the licence: it stores it once."""
        cfg = server.check_config(dict(CFG, accountId="1234567_SB1"))
        stored = {"licence": ""}
        calls = []
        saved = server.call_netsuite

        def netsuite(cfg, payload):
            calls.append(payload["action"])
            if payload["action"] == "licenceGet":
                return {"ok": True, "licence": stored["licence"]}
            if payload["action"] == "licenceSave":
                stored["licence"] = payload["licence"]
            return {"ok": True, "allowed": True, "used": 1, "users": 5}
        server.call_netsuite = netsuite
        try:
            handler = type("H", (server.Handler,), {"config_path": os.path.join(self.folder, "config.json"),
                                                   "licence_path": os.path.join(self.folder, "licence.json")})
            server.configure(handler, cfg, TODAY, licence_text=self.licence(account="1234567_SB1"))
            with open(handler.licence_path, "w", encoding="utf-8") as f:
                f.write(self.licence(account="1234567_SB1"))
            self.assertTrue(handler.licence["ok"])
            server.fetch_licence(handler)
            self.assertIn("licenceSave", calls)
            self.assertTrue(stored["licence"])
            calls.clear()
            server.fetch_licence(handler)                                   # not again before the 6-hour check
            self.assertEqual(calls, [])
        finally:
            server.call_netsuite = saved

    def test_a_renewed_licence_in_netsuite_reaches_a_licensed_pc(self):
        """The administrator pastes a renewal (more users): a PC that still has a valid old licence takes it."""
        cfg = server.check_config(dict(CFG, accountId="1234567_SB1"))
        old = self.licence()
        newer = json.dumps(server.licence_module.make(self.key, "World of Hospitality", "1234567_sb1", 10, "2028-12-31",
                                                      today=TODAY + datetime.timedelta(days=1)))
        stored = {"licence": old}
        saved = server.call_netsuite

        def netsuite(cfg, payload):
            if payload["action"] == "licenceGet":
                return {"ok": True, "licence": stored["licence"]}
            return {"ok": True, "allowed": True, "used": 1, "users": 5}
        server.call_netsuite = netsuite
        try:
            handler = type("H", (server.Handler,), {"config_path": os.path.join(self.folder, "config.json"),
                                                   "licence_path": os.path.join(self.folder, "licence.json")})
            with open(handler.licence_path, "w", encoding="utf-8") as f:
                f.write(old)
            server.configure(handler, cfg, TODAY, licence_text=old)
            self.assertFalse(server.fetch_licence(handler, force=True))      # same licence: nothing changes
            stored["licence"] = newer
            self.assertFalse(server.fetch_licence(handler))                  # not before the 6-hour check
            self.assertTrue(server.fetch_licence(handler, force=True))
            self.assertEqual(handler.licence["users"], 10)
            with open(handler.licence_path, encoding="utf-8") as f:
                self.assertEqual(json.load(f)["users"], 10)
            stored["licence"] = old                                           # an older one never replaces it
            self.assertFalse(server.fetch_licence(handler, force=True))
            self.assertEqual(handler.licence["users"], 10)
        finally:
            server.call_netsuite = saved

    def test_netsuite_waits_for_a_licence(self):
        cfg = server.check_config(dict(CFG))
        calls = []
        saved = server.call_netsuite
        server.call_netsuite = lambda cfg, payload: calls.append(payload["action"]) or {"ok": True, "account": "X", "user": "U"}
        try:
            handler = type("H", (server.Handler,), {})
            server.configure(handler, cfg, TODAY, licence_text="")
            self.assertFalse(handler.licence["ok"])
            self.assertIn("Licence: ", handler.backend({"action": "accounts"})["error"])
            self.assertTrue(handler.backend({"action": "ping"})["ok"])
            self.assertEqual(calls, ["ping"])
            server.configure(handler, cfg, TODAY, licence_text=self.licence(account="1234567_SB1"))
            self.assertTrue(handler.licence["ok"])
            handler.backend({"action": "accounts"})
            # this PC's seat is checked first ("setup" is the fiscal year look-up, in the background)
            self.assertEqual([c for c in calls if c != "setup"], ["ping", "seat", "accounts"])
        finally:
            server.call_netsuite = saved

    def test_setup_page_saves_without_showing_secrets(self):
        saved = server.call_netsuite
        server.call_netsuite = lambda cfg, payload: {"ok": True, "account": cfg["accountId"], "user": "Finance"}
        try:
            handler = type("H", (server.Handler,), {"config_path": os.path.join(self.folder, "config.json"),
                                                   "licence_path": os.path.join(self.folder, "licence.json")})
            server.configure(handler, None, TODAY)
            self.assertEqual(server.save_setup(handler, dict(CFG, restletUrl="https://evil.example.com/x"))["ok"], False)
            out = server.save_setup(handler, dict(CFG, licence=self.licence(), aiKey="sk-test"))
            self.assertTrue(out["ok"] and out["connected"] and out["licence"]["ok"], out)
            self.assertEqual(handler.mode, "netsuite")
            view = server.setup_view(handler)
            self.assertTrue(view["hasSecrets"] and view["aiOn"])
            self.assertNotIn("cs", json.dumps(view).replace("consumer", ""))   # no secret is sent back
            wrong = server.save_setup(handler, dict(CFG, licence=self.licence(account="7654321")))
            self.assertIn("Nothing was saved", wrong["error"])
            self.assertTrue(server.setup_view(handler)["licence"]["ok"])   # the working licence stays
            again = server.save_setup(handler, dict(CFG, consumerSecret="", tokenSecret="", aiOff=True))
            self.assertTrue(again["ok"])
            with open(handler.config_path, encoding="utf-8") as f:
                stored = json.load(f)
            self.assertEqual((stored["consumerSecret"], stored["tokenSecret"], stored["ai"]["apiKey"]), ("cs", "ts", ""))
        finally:
            server.call_netsuite = saved

    def test_certificate_for_this_pc(self):
        cert, key = os.path.join(self.folder, "c.crt"), os.path.join(self.folder, "c.key")
        thumbprint = server.make_certificate(cert, key)
        self.assertEqual(len(thumbprint), 40)
        context = server.ssl.SSLContext(server.ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(cert, key)
        names = server.ssl._ssl._test_decode_cert(cert)["subjectAltName"]
        self.assertEqual(names, (("DNS", "localhost"), ("IP Address", "127.0.0.1")))


class BudgetUploadTest(unittest.TestCase):
    """A budget sheet made from NetSuite's fiscal year, checked (nothing written), then uploaded."""

    def setUp(self):
        self.ledger = new_ledger()

    def test_sheet_check_and_upload(self):
        made = server.budget_module.handle(self.ledger, {"op": "sheet", "year": "FY2026", "category": "Budget - A",
                                                         "fill": True})
        self.assertTrue(made["ok"], made)
        action = made["actions"][0]
        rows = action["rows"]
        self.assertEqual(action["sheetName"], "TQ Budget FY2026")
        self.assertEqual(rows[1][:2], ["Budget category", "Budget - A"])
        self.assertEqual(rows[9][:8], ["Account", "Name", "Subsidiary", "Department", "Class", "Location",
                                       "Jan 2026", "Feb 2026"])
        first = rows[10]
        last_year = self.ledger.balances([{"account": first[0], "fromPeriod": "Jan 2025"}])[0]["value"]
        self.assertAlmostEqual(first[6], last_year, places=2)            # filled with last year's January
        self.assertTrue(first[18].startswith("=SUM(G11:R11)"))
        self.assertIn("filled with last year's actuals", made["reply"])
        # Excel gives back values: the Total column is a number there, and the user changes a few cells.
        sheet = [list(r) for r in rows]
        sheet[10][18] = 0
        sheet[10][6] = 1234.5
        sheet[11][4] = "Food"                 # this row's own class
        checked = server.budget_module.handle(self.ledger, {"op": "check", "rows": sheet})
        self.assertTrue(checked["ok"], checked)
        self.assertIn("would be done", checked["reply"])
        self.assertIn("create", checked["reply"])
        self.assertEqual(checked["problems"], 0)
        done = server.budget_module.handle(self.ledger, {"op": "upload", "rows": sheet})
        self.assertIn("created", done["reply"])
        again = server.budget_module.handle(self.ledger, {"op": "check", "rows": sheet})
        self.assertIn("update", again["reply"])                          # the same lines now exist
        sheet[12][0] = "9999"
        bad = server.budget_module.handle(self.ledger, {"op": "check", "rows": sheet})
        self.assertEqual(bad["problems"], 1)
        self.assertIn("Row 13 (9999)", bad["reply"])
        sheet[12][0], sheet[12][7] = first[0], "abc"
        self.assertIn("is not a number", server.budget_module.handle(self.ledger, {"op": "check", "rows": sheet})["error"])
        self.assertIn("not a budget sheet", server.budget_module.handle(self.ledger, {"op": "check", "rows": [["x"]]})["error"])

    def test_unknown_request_and_wrong_months(self):
        self.assertFalse(server.budget_module.handle(self.ledger, {"op": "nope"})["ok"])
        rows = server.budget_module.handle(self.ledger, {"op": "sheet", "year": "FY2025"})["actions"][0]["rows"]
        rows[9][6] = "Feb 2025"
        self.assertIn("do not match", server.budget_module.handle(self.ledger, {"op": "check", "rows": rows})["error"])
        self.assertIn("not found", server.budget_module.handle(self.ledger, {"op": "sheet", "year": "FY1999"})["error"])


class DashboardTest(unittest.TestCase):
    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def value(self, rows, formula):
        """Works out one TQ.BALANCE inside a dashboard formula, with B3:B8 and the period cells read from rows."""
        cell = lambda ref: Workbook(rows, self.ledger).value(ref)  # noqa: E731
        inner = re.search(r"TQ\.BALANCE\(([^()]*)\)", formula).group(1).split(",")
        args = [a.strip('"') if a.startswith('"') else cell(a) for a in inner]
        return self.ledger.balances([dict(zip(Workbook.KEYS, args))])[0]["value"]

    def test_three_dashboards(self):
        out = self.chat.handle({"message": "dashboard for Oct 2026", "context": {"sheets": []}})
        action = out["actions"][0]
        rows = action["rows"]
        self.assertEqual(action["sheetName"], "Executive dashboard")
        self.assertEqual(rows[7][:2], ["Month", "Oct 2026"])
        self.assertEqual([c["type"] for c in action["charts"]], ["Line", "BarClustered", "Pie", "Area"])
        self.assertEqual([c["topLeft"] for c in action["charts"]], ["A14", "E14", "A32", "E32"])
        self.assertIn((7, 1), [(d["r"], d["c"]) for d in action["dropdowns"]])         # B8 picks a month
        self.assertTrue(any(f.get("merge") for f in action["formats"]))                # the big tile numbers
        self.assertIn("4 charts", out["reply"])
        # The Sales tile reads the year to date, which is the ledger's sales from Jan to Oct 2026.
        ytd = "".join(re.search(r"\$([A-Z]+)\$(\d+)", rows[10][0]).groups())
        r, c = Workbook.position(ytd)
        self.assertAlmostEqual(self.value(rows, rows[r][c]),
                               self.ledger.balances([{"account": "type:Income", "fromPeriod": "Jan 2026",
                                                      "toPeriod": "Oct 2026"}])[0]["value"], places=2)
        # The monthly series stops at the chosen month (a gap after it) and last year runs all year.
        series = next(ch for ch in action["charts"] if ch["type"] == "Line")["range"]
        first = int(re.match(r"[A-Z]+(\d+)", series).group(1))
        self.assertTrue(rows[first][10].startswith("=IF(1<=$O$"))
        self.assertEqual(Workbook(rows, self.ledger).value("J%d" % (first + 10)), "Oct 2026")
        sales = self.chat.handle({"message": "sales dashboard for Oct 2026 for location Doha", "context": {"sheets": []}})
        self.assertEqual(sales["actions"][0]["rows"][6][:2], ["Location", "Doha"])
        cash = self.chat.handle({"message": "cash dashboard", "context": {"sheets": ["Cash dashboard"]}})
        self.assertEqual(cash["actions"][0]["sheetName"], "Cash dashboard 2")     # never over an existing sheet
        self.assertIn("ColumnClustered", [c["type"] for c in cash["actions"][0]["charts"]])
        wiz = self.chat.design({"op": "dashboard", "args": {"kind": "sales", "month": "Sep 2026"}, "context": {"sheets": []}})
        self.assertEqual(wiz["actions"][0]["rows"][7][1], "Sep 2026")

    def test_finance_dashboards(self):
        for message, sheet, tiles in (
                ("finance dashboard for Oct 2026", "Profit and loss dashboard",
                 ["Revenue, year to date", "Gross margin, year to date", "Operating expenses, year to date",
                  "Net margin, year to date"]),
                ("financial position dashboard for Oct 2026", "Financial position",
                 ["Working capital at month end", "Current ratio", "Debt to equity", "Debtor days"]),
                ("budget vs actual dashboard for Oct 2026", "Budget vs actual dashboard",
                 ["Revenue, year to date", "Costs, year to date", "Net income, year to date", "Revenue, this month"])):
            out = self.chat.handle({"message": message, "context": {"sheets": []}})
            action = out["actions"][0]
            self.assertEqual(action["sheetName"], sheet, message)
            self.assertEqual([action["rows"][9][c] for c in (0, 2, 4, 6)], tiles)
            self.assertEqual(len(action["charts"]), 4)
        # costs and debtor days: up is worse, so the colours swap
        pl = self.chat.handle({"message": "p&l dashboard for Oct 2026", "context": {"sheets": []}})["actions"][0]
        opex = [c for c in pl["conditional"] if c["c0"] == 4]
        self.assertEqual([c["font"] for c in opex if c["formula"].endswith("<0)")], ["#1A7F4B"])
        # the gross margin tile works out from the ledger
        rows = pl["rows"]
        rev = self.ledger.balances([{"account": "type:Income", "fromPeriod": "Jan 2026", "toPeriod": "Oct 2026"}])[0]["value"]
        cogs = self.ledger.balances([{"account": "type:COGS", "fromPeriod": "Jan 2026", "toPeriod": "Oct 2026"}])[0]["value"]
        formula = rows[10][2]
        refs = re.findall(r"\$([A-Z]+)\$(\d+)", formula)
        values = {}
        for col, row in set(refs):
            r, c = Workbook.position(col + row)
            values[col + row] = DashboardTest.value(self, rows, rows[r][c])
        self.assertAlmostEqual(eval(re.sub(r"\$([A-Z]+)\$(\d+)", lambda m: repr(values[m.group(1) + m.group(2)]),
                                          formula[len('=IFERROR('):-len(',"")')])), (rev - cogs) / rev, places=6)


class SignInTest(unittest.TestCase):
    """Sign in with NetSuite: OAuth 2.0 with PKCE, tokens kept per Windows user, renewed when they end."""

    def setUp(self):
        self.folder = os.path.join(os.path.dirname(__file__), "_signin_test")
        os.makedirs(self.folder, exist_ok=True)
        self.calls = []
        self.answers = []
        self.cfg = {"accountId": "1234567_SB1", "clientId": "client-abc",
                    "restletUrl": "https://1234567-sb1.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=1&deploy=1"}

    def tearDown(self):
        for name in os.listdir(self.folder):
            os.remove(os.path.join(self.folder, name))
        os.rmdir(self.folder)

    def post(self, url, fields):
        self.calls.append((url, dict(fields)))
        return self.answers.pop(0)

    def test_sign_in_refresh_and_sign_out(self):
        s = server.signin_module
        session = s.Session(self.cfg, self.folder, post=self.post)
        self.assertFalse(session.info()["signedIn"])
        with self.assertRaises(s.SignInNeeded):
            session.token()
        url = session.start()
        self.assertTrue(url.startswith("https://1234567-sb1.app.netsuite.com/app/login/oauth2/authorize.nl?"))
        query = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(url).query))
        self.assertEqual((query["client_id"], query["scope"], query["code_challenge_method"], query["redirect_uri"]),
                         ("client-abc", "restlets", "S256", "https://localhost:3000/oauth/callback"))
        with self.assertRaises(ValueError):                    # a code that this PC did not ask for
            session.finish("code-1", "someone-elses-state")
        self.answers.append((200, json.dumps({"access_token": "A1", "refresh_token": "R1", "expires_in": 3600})))
        session.finish("code-1", query["state"])
        token_call = self.calls[-1]
        self.assertTrue(token_call[0].endswith(".suitetalk.api.netsuite.com/services/rest/auth/oauth2/v1/token"))
        verifier = token_call[1]["code_verifier"]
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
        self.assertEqual(challenge, query["code_challenge"])   # PKCE: the code came back to the PC that asked
        self.assertNotIn("R1", open(os.path.join(self.folder, "session.json"), encoding="utf-8").read())  # encrypted
        session.remember_user("Zubair", "Administrator")
        self.assertEqual(session.info()["user"], "Zubair")
        self.assertEqual(session.token(), "A1")
        # an hour later: renewed with the refresh token, which NetSuite replaces (rotation)
        session.expires = 0
        self.answers.append((200, json.dumps({"access_token": "A2", "refresh_token": "R2", "expires_in": 3600})))
        self.assertEqual(session.token(), "A2")
        self.assertEqual(self.calls[-1][1], {"grant_type": "refresh_token", "refresh_token": "R1", "client_id": "client-abc"})
        # a refused refresh: the person signs in again
        session.expires = 0
        self.answers.append((400, json.dumps({"error": "invalid_grant"})))
        with self.assertRaises(s.SignInNeeded):
            session.token()
        self.assertFalse(session.info()["signedIn"])

    def test_connector_uses_the_bearer_token(self):
        session = server.signin_module.Session(self.cfg, self.folder, post=self.post)
        cfg = dict(self.cfg, _session=session)
        out = server.call_netsuite(cfg, {"action": "ping"})
        self.assertTrue(out.get("signIn"))
        self.assertIn("Sign in with NetSuite", out["error"])
        session.access, session.expires = "A9", time.time() + 3600
        self.assertEqual(server.auth_header(cfg), "Bearer A9")
        self.assertEqual(server.check_config(dict(self.cfg))["accountId"], "1234567_SB1")   # no token keys needed


class NetSuitePackageTest(unittest.TestCase):
    """The SuiteCloud project that installs the connector's scripts, deployments and role in NetSuite."""

    def test_package_is_complete(self):
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))
        import package_netsuite
        package_netsuite.copy()
        self.assertEqual(package_netsuite.check(), [])


class JournalUploadTest(unittest.TestCase):
    def setUp(self):
        self.ledger = new_ledger()

    def test_sheet_check_and_upload(self):
        made = server.journal_module.handle(self.ledger, {"op": "sheet"})
        rows = [list(r) for r in made["actions"][0]["rows"]]
        self.assertEqual(rows[4][:5], ["Entry", "Date", "Account", "Debit", "Credit"])
        rows[5][:5] = ["1", 46344, "6100", 500, ""]          # an Excel date serial: 18 Nov 2026
        rows[6][:5] = ["1", 46344, "1000", "", 500]
        rows.append(["2", "2026-11-19", "6200", 75.5, ""] + [""] * 7)
        rows.append(["2", "2026-11-19", "1000", "", 70] + [""] * 7)
        out = server.journal_module.handle(self.ledger, {"op": "check", "rows": rows})
        self.assertEqual(out["problems"], 1)
        self.assertIn("Entry 2 (row 8): Debits 75.5 and credits 70.0 differ", out["reply"])
        rows[-1][4] = 75.5
        out = server.journal_module.handle(self.ledger, {"op": "check", "rows": rows})
        self.assertEqual(out["problems"], 0, out)
        self.assertIn("Checked 2 entries: 2 ready", out["reply"])
        entries = server.journal_module.parse(rows)
        self.assertEqual(entries[0]["date"], "2026-11-18")
        done = server.journal_module.handle(self.ledger, {"op": "upload", "rows": rows})
        self.assertIn("Created 2 journal entries: JE1 (entry 1), JE2 (entry 2)", done["reply"])
        rows[5][3] = "lots"
        self.assertIn("is not a number", server.journal_module.handle(self.ledger, {"op": "check", "rows": rows})["error"])


class DropdownFilterTest(unittest.TestCase):
    """Filters picked from dropdowns in Excel: B4:B7 and a row per extra field at E4:F7."""

    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def ask(self, message, last=None, context=None):
        return self.chat.handle({"message": message, "context": dict(context or {}, last=last)})

    def test_standard_report_gets_dropdown_rows(self):
        first = self.ask("P&L for 2026 by quarter filter Brand: Premium; Name: Al Noor")
        out = self.ask("add dropdown filter Brand, Name", first["spec"])
        action = out["actions"][0]
        rows = action["rows"]
        self.assertEqual(out["spec"]["dropdowns"], ["Brand", "Name"])
        self.assertEqual(rows[2][4:6], ["More filters", "Pick (or type A, B / not A)"])
        self.assertEqual((rows[3][4], rows[3][5], rows[4][4], rows[4][5]), ("Brand", "Premium", "Name", "Al Noor"))
        self.assertEqual(rows[2][1], '=MID(""&IF($F$4="","","; "&$E$4&": "&$F$4)'
                                     '&IF($F$5="","","; "&$E$5&": "&$F$5),3,2000)')
        drops = {(d["r"], d["c"]): d for d in action["dropdowns"]}
        self.assertTrue(all((r, 1) in drops for r in (3, 4, 5, 6)))      # B4:B7
        self.assertIn((3, 5), drops)                                     # Brand has a list
        self.assertNotIn((4, 5), drops)                                  # Name is typed: its list is too long
        self.assertTrue(drops[(3, 5)]["free"])
        list_col = int([c for c in range(len(rows[0])) if rows[0][c] == "Brand list"][0])
        self.assertEqual([rows[1][list_col], rows[2][list_col]], ["Premium", "Basic"])
        self.assertFalse(any(b[0] <= 2 <= b[2] and b[1] <= 1 <= b[3] for b in action["textRanges"]))  # B3 is a formula
        # With the formula worked out, every number follows the dropdowns.
        rows[2][1] = "Brand: Premium; Name: Al Noor"
        book = Workbook(rows, self.ledger)
        expected = self.ledger.balances([{"account": "calc:NetIncome", "fromPeriod": "Jan 2026", "toPeriod": "Mar 2026",
                                          "filters": "Brand: Premium; Name: Al Noor"}])[0]["value"]
        self.assertAlmostEqual(book.value("C%d" % book.find("Net Income")), expected, places=2)
        again = self.ask("by month", out["spec"])                        # kept on a follow-up
        self.assertEqual(again["actions"][0]["rows"][3][4], "Brand")
        off = self.ask("no dropdown filters", again["spec"])
        self.assertEqual(off["spec"]["dropdowns"], [])
        self.assertEqual(off["actions"][0]["rows"][2][1], "Brand: Premium; Name: Al Noor")

    def test_segment_named_in_other_filters_replaces_the_old_one(self):
        first = self.ask("P&L for 2026 filter location Doha")
        out = self.ask("filter Location: not Doha; Brand: Premium", first["spec"])
        self.assertEqual(out["spec"]["filters"], {"location": "not Doha", "filters": "Brand: Premium"})
        rows = out["actions"][0]["rows"]
        self.assertEqual((rows[2][1], rows[6][1]), ("Brand: Premium", "not Doha"))   # B3 and B7
        again = self.ask("filter location Riyadh", out["spec"])
        self.assertEqual(again["spec"]["filters"], {"location": "Riyadh", "filters": "Brand: Premium"})
        old = dict(first["spec"], filters={"filters": "Location: Doha; Brand: Premium"})   # saved by an older version
        self.assertEqual(self.ask("filter location Riyadh", old)["spec"]["filters"],
                         {"location": "Riyadh", "filters": "Brand: Premium"})

    def test_unknown_or_segment_fields(self):
        first = self.ask("P&L for 2026")
        self.assertIn("not a field", self.ask("add dropdown filter Colour", first["spec"])["reply"])
        same = self.ask("add dropdown filter department", first["spec"])
        self.assertEqual(same["spec"]["dropdowns"], [])

    def test_layout_report_keeps_its_dropdowns(self):
        built = self.ask("custom P&L for 2026 by quarter")
        layout = [a for a in built["actions"] if a["sheetName"].startswith("TQ Layout")][0]
        context = {"layouts": [{"name": layout["sheetName"], "row": 0, "rows": layout["rows"]}],
                   "activeSheet": "Management P&L"}
        out = self.ask("add dropdown filter Brand", built["spec"], context)
        self.assertIn("dropdown filter for Brand", out["reply"])
        sheet = [a for a in out["actions"] if a["sheetName"].startswith("TQ Layout")][0]
        self.assertIn(["Dropdown filters", "Brand"], [r[:2] for r in sheet["rows"]])
        report = [a for a in out["actions"] if a["sheetName"] == "Management P&L"][0]
        self.assertEqual(report["rows"][3][4], "Brand")
        parsed = server.chat_module.layouts.parse(sheet["sheetName"], sheet["rows"])
        self.assertEqual(parsed["dropdowns"], ["Brand"])
        context["layouts"][0]["rows"] = sheet["rows"]
        rebuilt = self.ask("update report", None, context)
        self.assertEqual([a for a in rebuilt["actions"] if a["sheetName"] == "Management P&L"][0]["rows"][3][4], "Brand")
        off = self.ask("no dropdown filters", built["spec"], context)
        self.assertIn("removed", off["reply"])


class SmartChatTest(unittest.TestCase):
    """The built-in reader's tolerance, and the optional AI add-on with a stand-in for the AI service."""

    def setUp(self):
        self.ledger = new_ledger()
        self.chat = server.chat_module.ChatService(self.ledger)

    def ask(self, message, context=None, **more):
        return self.chat.handle(dict({"message": message, "context": context or {}}, **more))

    def test_polite_wording_and_spelling_slips(self):
        out = self.ask("can you please show me the ballance sheet for Oct 2026")
        self.assertIn('I read that as "show me the balance sheet for Oct 2026"', out["reply"])
        self.assertEqual(out["actions"][0]["sheetName"], "BS Oct 2026")
        self.assertIn("by quarter", self.ask("P&L for 2026 by quater")["reply"])
        plain = self.ask("I want to see the profit and loss for 2026")
        self.assertEqual(plain["actions"][0]["sheetName"], "P&L 2026")
        self.assertNotIn("I read that as", plain["reply"])
        # A name the user chose is not "corrected", and neither is a word from the chart of accounts.
        self.assertEqual(self.chat._spelling("add line Marketting accounts 6200"), "add line Marketting accounts 6200")
        self.assertEqual(self.chat._spelling('add colum "Budjet" for 2026'), 'add column "Budjet" for 2026')
        self.assertEqual(self.chat._spelling("what was salaries in 2026"), "what was salaries in 2026")

    def test_several_requests_in_one_message(self):
        out = self.ask("custom P&L for 2026 by quarter then show account detail; remove column Total")
        self.assertEqual([a["sheetName"] for a in out["actions"]], ["TQ Layout", "Management P&L"])
        self.assertTrue(out["design"]["detail"])
        self.assertNotIn("Total", [c["label"] for c in out["design"]["columns"]])
        self.assertIn("remove column Total: Removed column", out["reply"])
        stopped = self.ask("custom P&L for 2026 then remove column Nope then show account detail")
        self.assertIn("I stopped there", stopped["reply"])
        self.assertFalse(stopped["design"]["detail"])

    def test_not_understood_gives_suggestions(self):
        out = self.ask("ballance shet by quartr 2026 pls")
        self.assertNotIn("missed", out)
        near = self.ask("remove the colmn")
        self.assertIn("Did you mean", near["reply"])
        self.assertIn("remove column Total", near["reply"])

    def test_ai_add_on_is_optional(self):
        picture = {"mediaType": "image/png", "data": "aGk="}
        self.assertIn("switched off", self.ask("build this", image=picture)["reply"])
        self.assertIn("switched off", self.ask("ai: do something clever")["reply"])
        self.assertIn("not a picture", self.ask("x", image={"mediaType": "text/html", "data": "aGk="})["reply"])

    def test_ai_add_on_builds_a_format_and_runs_commands(self):
        sent = {}

        def service(url, headers, body):
            sent.update(url=url, headers=headers, body=json.loads(body.decode()))
            return 200, json.dumps({"stop_reason": "tool_use", "content": [
                {"type": "text", "text": "ok"},
                {"type": "tool_use", "id": "t1", "name": "respond", "input": {
                    "reply": "I rebuilt the format from your picture.",
                    "format": {"title": "Store P&L", "rows": [
                        {"label": "Revenue"}, {"label": "Food Sales", "accounts": "4000"},
                        {"label": "Service", "accounts": "4100"}, {"label": "Total Revenue"},
                        {"label": "Cost of Sales", "accounts": "5000"}, {"label": "Gross Profit"},
                        {"label": "Mystery", "accounts": "9999"}]},
                    "commands": ["custom P&L for 2026 by quarter"]}}]})

        self.chat.assistant = server.assistant_module.Assistant({"apiKey": "k-test"}, "grammar", transport=service)
        out = self.ask("make this", image={"mediaType": "image/png", "data": "aGk="})
        self.assertTrue(out["ai"])
        self.assertEqual([(l["section"], l["label"], l["source"]) for l in out["design"]["lines"]], [
            ("Revenue", "Food Sales", "4000"), ("Revenue", "Service", "4100"), ("Revenue", "Total Revenue", ""),
            ("Revenue", "Cost of Sales", "5000"), ("Gross Profit", "Gross Profit", "=[Total Revenue]+[Cost of Sales]")])
        self.assertIn("Mystery", out["reply"])                       # an unknown account is said, not used
        self.assertEqual(len(out["design"]["columns"]), 5)           # the command after the format ran
        self.assertEqual([a["sheetName"] for a in out["actions"]], ["TQ Layout Store P&L", "Store P&L"])
        self.assertEqual(sent["url"], "https://api.anthropic.com/v1/messages")
        self.assertEqual(sent["headers"]["x-api-key"], "k-test")
        self.assertEqual([b["type"] for b in sent["body"]["messages"][0]["content"]], ["image", "text"])
        details = sent["body"]["system"][1]["text"]
        self.assertIn("4000 | Product Revenue | Income", details)
        self.assertNotRegex(details, r"\d{2,},\d{3}")                # names only: no balances leave the PC
        # Only what the built-in reader misses goes to the AI.
        sent.clear()
        self.ask("P&L for 2026")
        self.assertEqual(sent, {})
        self.ask("please do something about coffee")
        self.assertIn("coffee", sent["body"]["messages"][0]["content"][0]["text"])

    def test_ai_add_on_errors_are_plain(self):
        make = lambda status, text: server.assistant_module.Assistant(  # noqa: E731
            {"apiKey": "k"}, "g", transport=lambda *a: (status, text))
        self.chat.assistant = make(401, '{"error": {"message": "invalid x-api-key"}}')
        self.assertIn("did not accept the API key", self.ask("ai: hello")["reply"])
        self.chat.assistant = make(200, json.dumps({"stop_reason": "refusal", "content": []}))
        self.assertIn("declined", self.ask("ai: hello")["reply"])
        self.chat.assistant = make(200, json.dumps({"stop_reason": "end_turn", "content": [{"type": "text", "text": "Just words."}]}))
        self.assertEqual(self.ask("ai: hello")["reply"], "Just words.")
        self.chat.assistant = make(200, json.dumps({"content": [{"type": "tool_use", "name": "respond", "input": {
            "reply": "Done.", "commands": ["delete everything now"]}}]}))
        self.assertIn("I did not understand", self.ask("ai: hello")["reply"])   # a made-up command does nothing


if __name__ == "__main__":
    unittest.main()
