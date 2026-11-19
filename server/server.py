r"""TQ Excel Connector - local server.

Serves the Excel add-in files and forwards the add-in's requests to the
NetSuite RESTlet, signing them with token-based authentication. The NetSuite
secrets stay in server/config.json and never reach Excel.

With no config.json (or with --mock) it answers from built-in sample data, so
the add-in can be demonstrated without a NetSuite account.

Connecting to NetSuite needs a licence from Techno Q (licence.json, checked by
licence.py). The setup page (/setup.html) writes config.json and the licence.

Standard library only:  python server/server.py
Installed (TQConnector.exe): the add-in files are next to the exe in addin\, and
config.json, licence.json and the certificate are in data\.
"""

import argparse
import base64
import functools
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import ssl
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import assistant as assistant_module  # noqa: E402
import budget as budget_module  # noqa: E402
import care  # noqa: E402
import signin as signin_module  # noqa: E402
import journal as journal_module  # noqa: E402
import updates as updates_module  # noqa: E402
import chat as chat_module  # noqa: E402
import ledger as ledger_module  # noqa: E402
import licence as licence_module  # noqa: E402
import tqcrypto  # noqa: E402

VERSION = "1.4.0"
FROZEN = bool(getattr(sys, "frozen", False))   # running as the installed TQConnector.exe
if FROZEN:
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
    ADDIN_DIR = os.path.join(ROOT, "addin")
    DATA_DIR = os.path.join(ROOT, "data")
    CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
    LICENCE_PATH = os.path.join(DATA_DIR, "licence.json")
    CERT_PATH = os.path.join(DATA_DIR, "certs", "localhost.crt")
    KEY_PATH = os.path.join(DATA_DIR, "certs", "localhost.key")
else:
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ADDIN_DIR = os.path.join(ROOT, "addin")
    DATA_DIR = os.path.join(ROOT, "server")
    CONFIG_PATH = os.path.join(ROOT, "server", "config.json")
    LICENCE_PATH = os.path.join(ROOT, "server", "licence.json")
    CERT_PATH = os.path.join(ROOT, "certs", "localhost.crt")
    KEY_PATH = os.path.join(ROOT, "certs", "localhost.key")
SECRET_KEYS = ("consumerSecret", "tokenSecret")

MAX_BODY = 8 * 1024 * 1024   # a chat message may carry a screenshot
RESTLET_HOST_SUFFIX = ".restlets.api.netsuite.com"
CONFIG_KEYS = ("accountId", "restletUrl", "consumerKey", "consumerSecret", "tokenId", "tokenSecret")
# NetSuite allows an account only a few requests at the same moment and refuses the rest with
# SSS_REQUEST_LIMIT_EXCEEDED. So requests go to NetSuite one at a time, and a refused one is tried again.
NETSUITE_QUEUE = threading.Lock()
RETRY_DELAYS = (1, 2, 4, 8)  # seconds before each further try
BUSY_MARKS = ("SSS_REQUEST_LIMIT_EXCEEDED", "Request Limit Exceeded")


# ---------------------------------------------------------------- NetSuite --

def pct(value):
    """Percent-encoding as OAuth 1.0 requires (RFC 3986 unreserved only)."""
    return urllib.parse.quote(str(value), safe="-._~")


def oauth_header(method, url, cfg, nonce=None, timestamp=None):
    """Builds the OAuth 1.0 (HMAC-SHA256) Authorization header for NetSuite."""
    parts = urllib.parse.urlsplit(url)
    base_url = "%s://%s%s" % (parts.scheme.lower(), parts.netloc.lower(), parts.path)
    oauth = {
        "oauth_consumer_key": cfg["consumerKey"],
        "oauth_token": cfg["tokenId"],
        "oauth_signature_method": "HMAC-SHA256",
        "oauth_timestamp": str(timestamp or int(time.time())),
        "oauth_nonce": nonce or secrets.token_hex(16),
        "oauth_version": "1.0",
    }
    params = urllib.parse.parse_qsl(parts.query, keep_blank_values=True) + list(oauth.items())
    encoded = sorted((pct(k), pct(v)) for k, v in params)
    param_string = "&".join("%s=%s" % kv for kv in encoded)
    base_string = "&".join([method.upper(), pct(base_url), pct(param_string)])
    key = "%s&%s" % (pct(cfg["consumerSecret"]), pct(cfg["tokenSecret"]))
    digest = hmac.new(key.encode(), base_string.encode(), hashlib.sha256).digest()
    oauth["oauth_signature"] = base64.b64encode(digest).decode()
    fields = ['realm="%s"' % cfg["accountId"]]
    fields += ['%s="%s"' % (k, pct(v)) for k, v in oauth.items()]
    return "OAuth " + ", ".join(fields)


def load_config(path=CONFIG_PATH):
    """Returns the NetSuite config, or None when there is no config file."""
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as handle:
        cfg = json.load(handle)
    return check_config(cfg)


def check_config(cfg):
    """Raises ValueError for a config that cannot work; returns it tidied."""
    if not isinstance(cfg, dict):
        raise ValueError("config.json must hold one JSON object")
    # Either "Sign in with NetSuite" (OAuth 2.0: a client ID) or token-based authentication (four keys)
    oauth2 = bool(str(cfg.get("clientId", "")).strip())
    needed = ("accountId", "restletUrl") if oauth2 else CONFIG_KEYS
    missing = [k for k in needed if not str(cfg.get(k, "")).strip()]
    if missing:
        raise ValueError("config.json is missing: " + ", ".join(missing))
    parts = urllib.parse.urlsplit(cfg["restletUrl"])
    host = (parts.hostname or "").lower()
    if parts.scheme != "https" or not host.endswith(RESTLET_HOST_SUFFIX):
        raise ValueError("restletUrl must be an https://<account>%s address" % RESTLET_HOST_SUFFIX)
    cfg["accountId"] = cfg["accountId"].strip().upper().replace("-", "_")
    return cfg


def post_to_netsuite(cfg, body):
    """One signed POST to the RESTlet. Returns (HTTP status, response text)."""
    request = urllib.request.Request(
        cfg["restletUrl"],
        data=body,
        method="POST",
        headers={
            "Authorization": auth_header(cfg),
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            return response.status, response.read().decode()
    except urllib.error.HTTPError as err:
        return err.code, err.read().decode(errors="replace")


def auth_header(cfg, renew=False):
    """Bearer token of the signed-in person (OAuth 2.0), else a token-based signature with a fresh nonce."""
    session = cfg.get("_session")
    if session is not None:
        return "Bearer " + session.token(renew)
    return oauth_header("POST", cfg["restletUrl"], cfg)


def call_netsuite(cfg, payload):
    body = json.dumps(payload).encode()
    try:
        with NETSUITE_QUEUE:
            for delay in RETRY_DELAYS + (None,):
                status, text = post_to_netsuite(cfg, body)
                busy = status == 429 or (status >= 400 and any(mark in text for mark in BUSY_MARKS))
                if not busy or delay is None:
                    break
                time.sleep(delay)
        if busy:
            return {"ok": False, "error": "NetSuite is busy: its limit on requests at the same moment was reached, "
                                          "even after %d tries. Wait a few seconds and try again."
                                          % (len(RETRY_DELAYS) + 1)}
        if status == 401 and cfg.get("_session") is not None:   # the access token was refused: renew once
            cfg["_session"].token(renew=True)
            status, text = post_to_netsuite(cfg, body)
        if status >= 400:
            return {"ok": False, "error": "NetSuite returned HTTP %s: %s" % (status, text[:400])}
        return json.loads(text)
    except signin_module.SignInNeeded as err:   # not signed in, or the sign-in has ended
        return {"ok": False, "error": str(err), "signIn": True}
    except (urllib.error.URLError, TimeoutError, ValueError) as err:
        return {"ok": False, "error": "Could not reach NetSuite: %s" % err}


# -------------------------------------------------------------- sample data --

PL_TYPES = {"Income", "COGS", "Expense", "OthIncome", "OthExpense"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
MOCK_ACCOUNTS = [
    ("1000", "Cash at Bank", "Bank"),
    ("1100", "Accounts Receivable", "AcctRec"),
    ("1200", "Inventory", "OthCurrAsset"),
    ("1500", "Fixed Assets", "FixedAsset"),
    ("2000", "Accounts Payable", "AcctPay"),
    ("2100", "Accrued Liabilities", "OthCurrLiab"),
    ("3000", "Share Capital", "Equity"),
    ("4000", "Product Revenue", "Income"),
    ("4100", "Service Revenue", "Income"),
    ("5000", "Cost of Goods Sold", "COGS"),
    ("6000", "Salaries and Wages", "Expense"),
    ("6100", "Rent", "Expense"),
    ("6200", "Marketing", "Expense"),
]
MOCK_PERIODS = [
    {
        "name": "%s %d" % (MONTHS[m], year),
        "start": "%d-%02d-01" % (year, m + 1),
        "end": "%d-%02d-%02d" % (year, m + 1, MONTH_DAYS[m]),
    }
    for year in (2025, 2026)
    for m in range(12)
]
MOCK_SEGMENTS = {
    "subsidiary": ["Parent Company", "Parent Company : Qatar", "Parent Company : Saudi Arabia"],
    "department": ["Sales", "Operations", "Finance"],
    "class": ["Hardware", "Services"],
    "location": ["Doha", "Riyadh"],
    "cseg1": ["Premium", "Basic"],     # a custom segment, "Brand"
}
MOCK_FIELDS = [{"id": k, "label": k.title(), "kind": "segment"} for k in ("subsidiary", "department", "class", "location")] \
    + [{"id": "name", "label": "Name", "kind": "name"}, {"id": "item", "label": "Item", "kind": "item"},
       {"id": "cseg1", "label": "Brand", "kind": "custom"}]


MOCK_TABLES = {  # sample rows for the Data tab; a sample query returns its table whole
    "customer": (["id", "entityid", "companyname", "email", "datecreated"], [
        [101, "C-0001", "Al Noor Trading", "accounts@alnoor.example", "2024-03-11"],
        [102, "C-0002", "Gulf Hospitality Group", "finance@gulfhg.example", "2024-06-02"],
        [103, "00417", "Doha Catering Co", "ap@dohacatering.example", "2025-01-19"],
    ]),
    "item": (["id", "itemid", "displayname", "itemtype"], [
        [201, "ESP-250", "Espresso Beans 250g", "InvtPart"],
        [202, "SRV-INST", "Installation Service", "Service"],
    ]),
    "transaction": (["id", "tranid", "type", "trandate", "entity", "foreigntotal"], [
        [9001, "INV-1001", "CustInvc", "2026-01-14", 101, "12500.00"],
        [9002, "INV-1002", "CustInvc", "2026-02-03", 102, "8300.50"],
        [9003, "=CM-7", "CustCred", "2026-02-20", 101, "-450.00"],
    ]),
}
MOCK_SEARCHES = [
    {"id": "customsearch_open_invoices", "title": "Open Invoices", "type": "Transaction", "table": "transaction"},
    {"id": "customsearch_customer_list", "title": "Customer List", "type": "Customer", "table": "customer"},
]


def _text(value):
    return "" if value is None else str(value).strip()


class MockBackend:
    """Answers the same requests as the RESTlet, from made-up numbers."""

    def __init__(self):
        self.accounts = {number.lower(): (number, name, kind) for number, name, kind in MOCK_ACCOUNTS}
        self.period_index = {p["name"].lower(): i for i, p in enumerate(MOCK_PERIODS)}

    def _monthly(self, account, index, segment_key):
        seed = hashlib.sha256(("%s|%d|%s" % (account, index, segment_key)).encode()).hexdigest()
        amount = 5000 + int(seed[:8], 16) % 45000
        return round(amount * (0.35 if segment_key else 1.0), 2)

    def _resolve(self, request):
        request = request or {}
        account_text = _text(request.get("account"))
        if not account_text:
            raise ValueError("Account is required")
        account = self.accounts.get(account_text.lower())
        if not account:
            raise ValueError('Account "%s" not found' % account_text)
        from_text = _text(request.get("fromPeriod"))
        if not from_text:
            raise ValueError("Period is required")
        if from_text.lower() not in self.period_index:
            raise ValueError('Period "%s" not found' % from_text)
        start = self.period_index[from_text.lower()]
        to_text = _text(request.get("toPeriod"))
        if to_text and to_text.lower() not in self.period_index:
            raise ValueError('To period "%s" not found' % to_text)
        end = self.period_index[to_text.lower()] if to_text else start
        if end < start:
            raise ValueError("To period is before from period")
        if account[2] not in PL_TYPES:
            start = 0  # balance sheet accounts are cumulative
        self._check_filters(request)
        segment_key = "|".join(
            _text(request.get(k)).lower() for k in ("department", "class", "location", "filters")
        ).strip("|")
        return account[0], range(start, end + 1), segment_key

    @staticmethod
    def _subsidiaries(text):
        """Sample subsidiaries a subsidiary filter covers: "Parent Company*" is all three (consolidated
        at a rate of 1), "Qatar, Saudi Arabia" two, "not Qatar" the other two. Blank = all, as one figure."""
        names = MOCK_SEGMENTS["subsidiary"]
        short = {n.split(" : ")[-1].lower(): n for n in names}
        sub = _text(text)
        if not sub:
            return [""]
        exclude = bool(re.match(r"(not|except|excluding)\s+", sub, re.I))
        sub = re.sub(r"^(not|except|excluding)\s+", "", sub, flags=re.I)
        out = []
        for item in sub.split(","):
            under = bool(re.search(r"(\*|\(consolidated\)|and below)\s*$", item.strip(), re.I))
            item = re.sub(r"\s*(\*|\(consolidated\)|and below)\s*$", "", item.strip(), flags=re.I).lower()
            name = short.get(item) or next((n for n in names if n.lower() == item), None)
            if not name:
                raise ValueError('subsidiary "%s" not found' % item)
            out += [n for n in names if n == name or (under and n.startswith(name + " : "))]
        return [n for n in names if (n in out) != exclude]

    @staticmethod
    def _check_filters(request):
        for part in [p for p in _text(request.get("filters")).split(";") if p.strip()]:
            field = part.split(":")[0].strip().lower() if ":" in part else ""
            if not field:
                raise ValueError('Filter "%s" needs a field and values, e.g. Brand: Premium' % part.strip())
            if field not in {f["id"] for f in MOCK_FIELDS} | {f["label"].lower() for f in MOCK_FIELDS} | {"customer", "vendor"}:
                raise ValueError('Filter field "%s" not found. Use a segment, Name or Item.' % part.split(":")[0].strip())

    def _balance(self, request):
        return round(sum(self._one(dict(request, subsidiary=sub)) for sub in self._subsidiaries(request.get("subsidiary"))), 2)

    def _one(self, request):
        account, indexes, segment_key = self._resolve(request)
        segment_key = "|".join(x for x in (_text(request.get("subsidiary")).lower(), segment_key) if x)
        if _text((request or {}).get("budget")):  # a sample budget: a round figure near the actuals
            start = self.period_index[_text(request.get("fromPeriod")).lower()]
            indexes = range(start, indexes[-1] + 1)  # a budget is for its periods, never cumulative
            return round(sum(round(self._monthly(account, i, segment_key), -3) for i in indexes), 2)
        return round(sum(self._monthly(account, i, segment_key) for i in indexes), 2)

    def _transactions(self, request):
        account, indexes, segment_key = self._resolve(request)
        rows = []
        for i in indexes:
            total = self._monthly(account, i, segment_key)
            parts = [round(total * share, 2) for share in (0.4, 0.3, 0.2)]
            parts.append(round(total - sum(parts), 2))
            for n, amount in enumerate(parts):
                rows.append({
                    "id": 100000 + i * 10 + n,
                    "date": MOCK_PERIODS[i]["start"][:8] + "%02d" % (5 + n * 6),
                    "number": "JE%05d" % (i * 10 + n + 1),
                    "type": "Journal",
                    "entity": "",
                    "memo": "Sample entry %d" % (n + 1),
                    "amount": amount,
                })
        return {"truncated": len(rows) > 1000, "rows": rows[:1000]}

    def _library(self, action, payload):
        """The company report library, kept in memory for sample data (in NetSuite: tq_library.json)."""
        items = self.__dict__.setdefault("library", {})
        rows = lambda: {"items": sorted([{k: v for k, v in dict(it, id=i, mine=True).items() if k != "design"}  # noqa: E731
                                         for i, it in items.items()], key=lambda r: r["name"]), "canManage": True}
        if action == "librarySave":
            name = _text(payload.get("name"))[:60]
            if not name or not isinstance(payload.get("design"), dict):
                raise ValueError("Give the layout a name")
            key = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "layout"
            replaced = key in items
            items[key] = {"name": name, "description": _text(payload.get("description"))[:200], "by": "Sample User",
                          "at": "2026-11-19", "design": payload["design"]}
            return dict(rows(), saved=key, replaced=replaced)
        key = _text(payload.get("id"))
        if action == "libraryLoad":
            if key not in items:
                raise ValueError("That layout is no longer in the library")
            return {"item": dict(items[key], id=key)}
        if action == "libraryDelete":
            items.pop(key, None)
        return rows()

    def handle(self, payload):
        try:
            action = payload.get("action")
            if action == "ping":
                return {"ok": True, "version": 2, "account": "SAMPLE", "user": "Sample data", "role": "sample"}
            if action == "setup":
                return {"ok": True, "version": 2, "oneWorld": True, "multiCurrency": True}
            if action == "filterFields":
                return {"ok": True, "fields": MOCK_FIELDS}
            if action == "balances":
                requests = payload.get("requests")
                if not isinstance(requests, list):
                    raise ValueError("requests must be a list")
                results = []
                for request in requests:
                    try:
                        results.append({"value": self._balance(request)})
                    except ValueError as err:
                        results.append({"error": str(err)})
                return {"ok": True, "results": results}
            if action == "transactions":
                return dict(ok=True, **self._transactions(payload.get("request")))
            if action == "accounts":
                return {"ok": True, "accounts": [
                    {"number": n, "name": name, "type": kind} for n, name, kind in MOCK_ACCOUNTS
                ]}
            if action == "periods":
                return {"ok": True, "periods": MOCK_PERIODS}
            if action == "segments":
                kind = _text(payload.get("kind")).lower()
                kind = "cseg1" if kind == "brand" else kind
                if kind not in MOCK_SEGMENTS:
                    raise ValueError('Unknown segment "%s"' % kind)
                return {"ok": True, "segments": [
                    {"id": i + 1, "name": name} for i, name in enumerate(MOCK_SEGMENTS[kind])
                ]}
            if action in ("searchList", "searchRun", "sql", "tableColumns"):
                return dict(ok=True, **self._data(action, payload))
            if action == "seatList":
                return {"ok": True, "canRelease": True, "days": 30, "seats": [
                    {"key": "u1", "label": "PC-1", "user": "Sample User", "last": "2026-11-18", "mine": True}]}
            if action == "seatRelease":
                return {"ok": True, "canRelease": True, "days": 30, "seats": []}
            if action in ("libraryList", "librarySave", "libraryLoad", "libraryDelete"):
                return dict(ok=True, **self._library(action, payload))
            if action == "budgetInfo":
                years = sorted({p["name"][-4:] for p in MOCK_PERIODS})
                return {"ok": True, "canWrite": True, "recordOk": True, "categories": ["Budget - A", "Budget - B"],
                        "years": [{"id": n + 1, "name": "FY" + y, "months": [p["name"] for p in MOCK_PERIODS
                                                                            if p["name"].endswith(y)]}
                                  for n, y in enumerate(years)]}
            if action == "branding":   # the sample company: the add-in's own icon stands in for a logo
                with open(os.path.join(ADDIN_DIR, "assets", "icon-80.png"), "rb") as f:
                    logo = base64.b64encode(f.read()).decode()
                return {"ok": True, "company": "Sample Company", "subsidiary": _text(payload.get("subsidiary")),
                        "logo": {"data": logo, "type": "PNGIMAGE", "name": "logo.png"}}
            if action == "journalSave":
                results = []
                for entry in payload.get("entries") or []:
                    lines = entry.get("lines") or []
                    unknown = [l.get("account") for l in lines if _text(l.get("account")).lower() not in self.accounts]
                    debit = round(sum(float(l.get("debit") or 0) for l in lines), 2)
                    credit = round(sum(float(l.get("credit") or 0) for l in lines), 2)
                    if unknown:
                        results.append({"ok": False, "error": 'line: account "%s" not found' % _text(unknown[0])})
                    elif debit != credit:
                        results.append({"ok": False, "error": "Debits %s and credits %s differ" % (debit, credit)})
                    elif payload.get("dryRun") is False:
                        n = self.__dict__.setdefault("journals", 0) + 1
                        self.journals = n
                        results.append({"ok": True, "action": "created", "id": 1000 + n, "number": "JE%d" % n, "total": debit})
                    else:
                        results.append({"ok": True, "action": "create", "total": debit})
                return {"ok": True, "results": results, "done": len(results), "canWrite": True}
            if action == "budgetSave":
                saved = self.__dict__.setdefault("budgets", {})
                results = []
                for line in payload.get("lines") or []:
                    if _text(line.get("account")).lower() not in self.accounts:
                        results.append({"ok": False, "error": 'Account "%s" not found' % _text(line.get("account"))})
                        continue
                    key = (_text(line.get("account")), _text(payload.get("year")), _text(payload.get("category")),
                           tuple(_text(line.get(k)) for k in ("subsidiary", "department", "class", "location")))
                    known = key in saved
                    if payload.get("dryRun") is False:
                        saved[key] = list(line.get("amounts") or [])
                    results.append({"ok": True, "action": ("update" if known else "create") if payload.get("dryRun") is not False
                                    else ("updated" if known else "created")})
                return {"ok": True, "results": results, "done": len(results)}
            raise ValueError('Unknown action "%s"' % _text(action))
        except ValueError as err:
            return {"ok": False, "error": str(err)}

    @staticmethod
    def _data(action, payload):
        """Stands in for the data queries in tq_excel_restlet.js: a query returns every row of the sample table it names."""
        try:
            limit = max(1, int(payload.get("maxRows") or 5000))
        except (TypeError, ValueError):
            limit = 5000
        if action == "searchList":
            words = _text(payload.get("text")).lower().split()
            return {"searches": [{k: s[k] for k in ("id", "title", "type")} for s in MOCK_SEARCHES
                                 if all(w in (s["id"] + " " + s["title"]).lower() for w in words)]}
        if action == "searchRun":
            found = next((s for s in MOCK_SEARCHES if s["id"] == _text(payload.get("id")).lower()), None)
            if not found:
                raise ValueError('Saved search "%s" could not be opened' % _text(payload.get("id")))
            table = found["table"]
        else:
            text = _text(payload.get("query") if action == "sql" else "FROM " + _text(payload.get("table")))
            if action == "sql" and not re.match(r"(select|with)\b", text, re.I):
                raise ValueError("A query must start with SELECT")
            if action == "sql" and text.count("?") != len(payload.get("params") or []):
                raise ValueError("The query has %d ? placeholders but %d values were given"
                                 % (text.count("?"), len(payload.get("params") or [])))
            match = re.search(r"\bfrom\s+(\w+)", text, re.I)
            table = match.group(1).lower() if match else ""
            if table not in MOCK_TABLES:
                raise ValueError("Search error occurred: Record '%s' was not found." % table)
        columns, rows = MOCK_TABLES[table]
        if action == "tableColumns":
            return {"table": table, "columns": columns, "empty": False}
        return {"columns": columns, "rows": [list(r) for r in rows[:limit]], "truncated": len(rows) > limit,
                "total": len(rows)}


# -------------------------------------------------------------- HTTP server --

class Handler(SimpleHTTPRequestHandler):
    backend = None  # callable(payload) -> dict
    mode = "sample"
    chat = None  # ChatService
    licence = None  # licence check result in NetSuite mode
    account = ""
    today = None
    config_path = CONFIG_PATH
    licence_path = LICENCE_PATH

    def end_headers(self):
        if self.path.startswith("/assets/"):
            # Icons: Excel's ribbon saves them to disk before it draws them, and shows a placeholder for an icon
            # it may not keep. Their addresses carry a version tag (?v=...), so a new icon gets a new address.
            self.send_header("Cache-Control", "public, max-age=604800")
        else:
            # Excel caches add-in files aggressively; always serve the latest.
            self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _page(self, status, title, text, ok=True):
        body = ("<!DOCTYPE html><html><head><meta charset='utf-8'><title>%s</title></head>"
                "<body style=\"font:15px 'Segoe UI',sans-serif;margin:48px;color:#1f2933\">"
                "<h2 style='color:%s'>%s</h2><p>%s</p></body></html>"
                % (html.escape(title), "#1a7f4b" if ok else "#b3261e", html.escape(title), html.escape(text))).encode()
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _logo(self):
        """The company logo: the one uploaded on the setup page, else the one in NetSuite (Company Information)."""
        found = custom_logo(type(self))
        if not found and type(self).mode == "netsuite":
            brand = type(self).backend({"action": "branding", "subsidiary": ""})
            logo = brand.get("logo") if isinstance(brand, dict) else None
            if logo and logo.get("data"):
                try:
                    found = (base64.b64decode(logo["data"]), "image/jpeg" if "JPG" in str(logo.get("type")).upper() else "image/png")
                except ValueError:
                    found = None
        if not found:
            return self._json(404, {"ok": False, "error": "No logo"})
        self.send_response(200)
        self.send_header("Content-Type", found[1])
        self.send_header("Content-Length", str(len(found[0])))
        self.end_headers()
        self.wfile.write(found[0])

    def _oauth_start(self):
        """Sends the browser to NetSuite's sign-in page."""
        session = getattr(type(self), "session", None)
        if session is None:
            return self._page(400, "Sign-in is not set up", "Enter the NetSuite account and the client ID of the "
                              "TQ Excel Connector integration on the setup page first.", ok=False)
        self.send_response(302)
        self.send_header("Location", session.start())
        self.end_headers()

    def _oauth_callback(self):
        """NetSuite sends the browser back here with a one-time code."""
        session = getattr(type(self), "session", None)
        query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
        if session is None:
            return self._page(400, "Sign-in is not set up", "Open the setup page first.", ok=False)
        if query.get("error"):
            return self._page(400, "Not signed in", "NetSuite said: %s" % (query.get("error_description") or
                                                                         query["error"])[0], ok=False)
        try:
            session.finish((query.get("code") or [""])[0], (query.get("state") or [""])[0])
        except (ValueError, OSError, urllib.error.URLError) as err:
            return self._page(400, "Not signed in", str(err), ok=False)
        who = type(self).backend({"action": "ping"})
        session.remember_user(who.get("user"), who.get("role"))
        if getattr(type(self), "seats", None) is not None:
            type(self).seats.checked = 0   # a new person: their own seat is checked on the next request
        return self._page(200, "Signed in to NetSuite", "Signed in as %s. You can close this tab and go back to Excel."
                          % (who.get("user") or "you"))

    def _json(self, status, payload):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.urlsplit(self.path).path
        if path == "/api/status":
            return self._json(200, {"ok": True, "mode": self.mode, "chat": type(self).chat.available(),
                                    "ai": bool(type(self).chat.assistant), "version": VERSION,
                                    "licence": type(self).licence, "configured": self.mode == "netsuite",
                                    "seats": getattr(type(self), "seats", None) and type(self).seats.state,
                                    "update": getattr(type(self), "watcher", None) and type(self).watcher.state,
                                    "signIn": type(self).session.info() if getattr(type(self), "session", None)
                                    else {"mode": "tba" if self.mode == "netsuite" else "sample"}})
        if path == "/logo":
            return self._logo()
        if path == "/oauth/start":
            return self._oauth_start()
        if path == "/oauth/callback":
            return self._oauth_callback()
        if path == "/api/update":
            try:
                found = updates_module.check(_saved_config(type(self).config_path).get("updateUrl", ""), VERSION)
                found.pop("manifest", None)
                return self._json(200, dict(found, ok=True))
            except ValueError as err:
                return self._json(200, {"ok": False, "error": str(err)})
        if path == "/api/setup":
            return self._json(200, setup_view(type(self)))
        if path == "/api/prefs":
            return self._json(200, type(self).prefs.handle({"op": "get"}))
        if path.startswith("/api"):
            return self._json(404, {"ok": False, "error": "Not found"})
        return super().do_GET()

    def do_POST(self):
        path = urllib.parse.urlsplit(self.path).path
        if path not in ("/api", "/api/chat", "/api/design", "/api/query", "/api/setup", "/api/budget", "/api/journal", "/api/update",
                        "/api/signout", "/api/logo", "/api/prefs", "/api/diagnostics", "/api/export", "/api/import"):
            return self._json(404, {"ok": False, "error": "Not found"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = -1
        if length < 0 or length > MAX_BODY:
            self.close_connection = True  # the body is not read, so the connection cannot be reused
            return self._json(413, {"ok": False, "error": "Request too large"})
        body = self.rfile.read(length)  # read before any refusal, or the sender sees a broken connection
        # Requiring JSON (exactly, so "text/plain; x=application/json" does not pass) stops other web pages
        # from posting here without the browser asking first; a request a browser marks as coming from
        # another site is refused as well.
        media = self.headers.get("Content-Type", "").split(";")[0].strip().lower()
        if media != "application/json":
            return self._json(415, {"ok": False, "error": "Content-Type must be application/json"})
        if self.headers.get("Sec-Fetch-Site", "same-origin").lower() not in ("same-origin", "none"):
            return self._json(403, {"ok": False, "error": "Requests from other sites are not accepted"})
        try:
            payload = json.loads(body.decode())
        except ValueError:
            return self._json(400, {"ok": False, "error": "Body is not valid JSON"})
        if not isinstance(payload, dict):
            return self._json(400, {"ok": False, "error": "Body must be a JSON object"})
        if path == "/api/setup":
            return self._json(200, save_setup(type(self), payload))
        if path == "/api/logo":
            return self._json(200, save_logo(type(self), payload))
        if path == "/api/prefs":
            return self._json(200, type(self).prefs.handle(payload))
        if path == "/api/diagnostics":
            return self._json(200, save_diagnostics(type(self)))
        if path == "/api/export":
            return self._json(200, export_setup(type(self), payload))
        if path == "/api/import":
            return self._json(200, import_setup(type(self), payload))
        if path == "/api/chat":
            return self._json(200, type(self).chat.handle(payload))
        if path == "/api/design":
            return self._json(200, type(self).chat.design(payload))
        if path == "/api/query":
            return self._json(200, type(self).chat.queries.handle(payload))
        if path == "/api/budget":
            return self._json(200, budget_module.handle(type(self).chat.ledger, payload))
        if path == "/api/signout":
            session = getattr(type(self), "session", None)
            if session is not None:
                session.sign_out()
            return self._json(200, {"ok": True, "signIn": session.info() if session else {"mode": "tba"}})
        if path == "/api/update":
            try:
                return self._json(200, updates_module.install(
                    _saved_config(type(self).config_path).get("updateUrl", ""), VERSION))
            except (ValueError, OSError) as err:
                return self._json(200, {"ok": False, "error": str(err)})
        if path == "/api/journal":
            return self._json(200, journal_module.handle(type(self).chat.ledger, payload))
        if payload.get("action") in ("budgetSave", "journalSave"):   # writing budgets goes through /api/budget and its checks only
            return self._json(400, {"ok": False, "error": "Use /api/budget"})
        answer = type(self).backend(payload)
        if payload.get("action") == "branding" and isinstance(answer, dict):   # the uploaded logo wins
            found = custom_logo(type(self))
            if found:
                answer = dict(answer, ok=True, logo={"data": base64.b64encode(found[0]).decode(),
                                                     "type": "PNGIMAGE" if found[1] == "image/png" else "JPGIMAGE",
                                                     "name": "logo"})
        return self._json(200, answer)


def configure(handler, cfg, today=None, licence_text=None):
    """Points the handler at NetSuite (cfg) or at the sample data (cfg=None). Also used after setup.

    licence_text: for tests; otherwise the licence file is read. Without a valid licence for the
    account, NetSuite is not asked for anything but the connection test.
    """
    handler.session = None
    handler.prefs = care.Prefs(os.path.dirname(handler.licence_path))
    if cfg and str(cfg.get("clientId", "")).strip():   # "Sign in with NetSuite": each person's own token
        handler.session = signin_module.Session(cfg, os.path.dirname(handler.licence_path))
        cfg["_session"] = handler.session
    raw = functools.partial(call_netsuite, cfg) if cfg else MockBackend().handle
    handler.mode = "netsuite" if cfg else "sample"
    handler.account = cfg["accountId"] if cfg else ""
    handler.today = today
    handler.licence = None
    if cfg:
        try:
            found = licence_module.parse(licence_text) if licence_text is not None \
                else licence_module.read(handler.licence_path)
        except (ValueError, OSError) as err:
            found, handler.licence = None, {"ok": False, "message": str(err)}
        handler.licence = handler.licence or licence_module.check(found, cfg["accountId"], today)
        if not handler.licence["ok"]:
            netsuite, message = raw, handler.licence["message"]

            def raw(payload):  # noqa: F811 - only the connection test passes without a licence
                if payload.get("action") == "ping":
                    return netsuite(payload)
                return {"ok": False, "error": "Licence: " + message}
    handler.seats = None
    if cfg and handler.licence and handler.licence.get("ok"):   # this PC must hold one of the licence's seats
        handler.seats = updates_module.Seats(raw, updates_module.machine(os.path.dirname(handler.licence_path)),
                                             handler.licence.get("users"), by_user=handler.session is not None)
        raw = handler.seats
    try:
        fiscal_start = int((cfg or {}).get("fiscalYearStartMonth") or 1)
    except (TypeError, ValueError):
        fiscal_start = 1
    # The ledger adds totals (type:..., calc:...) on top of NetSuite's single-account balances.
    ledger = ledger_module.Ledger(raw, fiscal_start_month=fiscal_start, today=today)
    if cfg:   # drill-down rows link to the record: https://65987-sb1.app.netsuite.com/app/accounting/...
        ledger.record_url = "https://%s.app.netsuite.com/app/accounting/transactions/transaction.nl?id=" % \
            cfg["accountId"].lower().replace("_", "-")
    handler.backend = staticmethod(ledger.handle)
    handler.chat = chat_module.ChatService(ledger)
    ai = (cfg or {}).get("ai") if isinstance((cfg or {}).get("ai"), dict) else {}
    if str(ai.get("apiKey") or "").strip():  # the optional AI add-on, with the customer's own key
        handler.chat.assistant = assistant_module.Assistant(ai, chat_module.HELP_TEXT)


def _saved_config(path):
    """config.json as a dict; {} when it is missing or not a JSON object."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        return {}
    return cfg if isinstance(cfg, dict) else {}


def custom_logo(handler):
    """(bytes, content type) of the logo uploaded on the setup page, or None."""
    folder = os.path.dirname(handler.licence_path)
    for name, kind in (("logo.png", "image/png"), ("logo.jpg", "image/jpeg")):
        path = os.path.join(folder, name)
        if os.path.exists(path):
            with open(path, "rb") as f:
                return f.read(), kind
    return None


def save_logo(handler, payload):
    """Setup page: {op: save, data: base64 of a PNG or JPEG (up to 1 MB)} or {op: remove}."""
    folder = os.path.dirname(handler.licence_path)
    for name in ("logo.png", "logo.jpg"):
        if os.path.exists(os.path.join(folder, name)):
            os.remove(os.path.join(folder, name))
    if payload.get("op") == "remove":
        return {"ok": True, "logo": "netsuite"}
    try:
        data = base64.b64decode(str(payload.get("data") or ""), validate=True)
    except ValueError:
        return {"ok": False, "error": "That file could not be read."}
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        name = "logo.png"
    elif data[:3] == b"\xff\xd8\xff":
        name = "logo.jpg"
    else:
        return {"ok": False, "error": "Use a PNG or JPEG picture."}
    if len(data) > 1024 * 1024:
        return {"ok": False, "error": "The picture is over 1 MB; use a smaller one."}
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, name), "wb") as f:
        f.write(data)
    return {"ok": True, "logo": "custom"}


def setup_view(handler):
    """What the setup page shows: the settings without their secrets."""
    cfg = _saved_config(handler.config_path)
    ai = cfg.get("ai") if isinstance(cfg.get("ai"), dict) else {}
    return {"ok": True, "version": VERSION, "mode": handler.mode,
            "accountId": cfg.get("accountId", ""), "restletUrl": cfg.get("restletUrl", ""),
            "consumerKey": cfg.get("consumerKey", ""), "tokenId": cfg.get("tokenId", ""),
            "fiscalYearStartMonth": cfg.get("fiscalYearStartMonth", 1), "updateUrl": cfg.get("updateUrl", ""),
            "clientId": cfg.get("clientId", ""), "authMode": "oauth2" if cfg.get("clientId") else "tba",
            "signIn": handler.session.info() if getattr(handler, "session", None) else None,
            "redirectUri": signin_module.REDIRECT_URI,
            "logo": "custom" if custom_logo(handler) else "netsuite",
            "hasSecrets": all(str(cfg.get(k) or "").strip() for k in SECRET_KEYS),
            "aiOn": bool(str(ai.get("apiKey") or "").strip()), "licence": handler.licence,
            "dataFolder": os.path.dirname(handler.config_path),
            "certExpires": str(care.cert_expires(CERT_PATH) or "") if os.path.exists(CERT_PATH) else "",
            "certDaysLeft": care.cert_days_left(CERT_PATH) if os.path.exists(CERT_PATH) else None}


def _status(handler):
    return {"mode": handler.mode, "version": VERSION, "licence": handler.licence,
            "seats": getattr(handler, "seats", None) and handler.seats.state,
            "update": getattr(handler, "watcher", None) and handler.watcher.state,
            "signIn": handler.session.info() if getattr(handler, "session", None) else None}


def save_diagnostics(handler):
    """Setup page > Save diagnostics: a zip for Techno Q support in Documents\\TQ Connector (no secrets)."""
    try:
        path = care.diagnostics(os.path.dirname(handler.config_path), VERSION, _status(handler),
                                _saved_config(handler.config_path), CERT_PATH, open_folder=FROZEN)
    except OSError as err:
        return {"ok": False, "error": "Could not write the diagnostics file: %s" % err}
    return {"ok": True, "path": path}


def export_setup(handler, payload):
    """Move to another PC > Export: settings (no secrets), licence, logo, filter sets, plus the layouts and
    queries the panel sends from the workbook, in one file in Documents\\TQ Connector."""
    licence_text = ""
    try:
        with open(handler.licence_path, encoding="utf-8") as f:
            licence_text = f.read()
    except OSError:
        pass
    try:
        path, summary = care.export_bundle(_saved_config(handler.config_path), licence_text, custom_logo(handler),
                                           handler.prefs.read(), payload.get("layouts"), payload.get("queries"),
                                           VERSION, payload.get("name"))
    except OSError as err:
        return {"ok": False, "error": "Could not write the setup file: %s" % err}
    if FROZEN:
        care.show_in_folder(path)
    return {"ok": True, "path": path, "summary": summary}


def import_setup(handler, payload):
    """Move to another PC > Import: {bundle, use: {settings, licence, logo, filterSets}} -> what was applied.
    The layouts and queries are written into the workbook by the panel itself."""
    try:
        bundle = care.read_bundle(payload.get("bundle"))
    except ValueError as err:
        return {"ok": False, "error": str(err)}
    use = payload.get("use") if isinstance(payload.get("use"), dict) else {}
    done, notes = [], []
    if use.get("filterSets") and bundle.get("filterSets"):
        prefs = handler.prefs.read()
        have = {s.get("name", "").lower(): s for s in prefs.get("filterSets") or [] if isinstance(s, dict)}
        for item in bundle["filterSets"]:
            if isinstance(item, dict) and item.get("name"):
                have[str(item["name"]).lower()] = {"name": str(item["name"])[:60], "cells": item.get("cells") or {}}
        prefs["filterSets"] = sorted(have.values(), key=lambda x: x["name"].lower())
        handler.prefs.write(prefs)
        done.append("%d filter set(s)" % len(bundle["filterSets"]))
    if use.get("logo") and isinstance(bundle.get("logo"), dict):
        answer = save_logo(handler, {"op": "save", "data": bundle["logo"].get("data", "")})
        (done if answer.get("ok") else notes).append("the logo" if answer.get("ok") else "Logo: " + answer["error"])
    if (use.get("settings") or use.get("licence")) and (bundle.get("settings") or bundle.get("licence")):
        current = _saved_config(handler.config_path)
        settings = bundle["settings"] if use.get("settings") else {k: current.get(k, "") for k in care.MOVED_SETTINGS}
        body = dict(settings, licence=bundle.get("licence", "") if use.get("licence") else "")
        body["authMode"] = "oauth2" if settings.get("clientId") else "tba"
        answer = save_setup(handler, body)
        if answer.get("ok"):
            done.append("the NetSuite connection" if use.get("settings") else "the licence")
            if use.get("licence") and bundle.get("licence") and use.get("settings"):
                done.append("the licence")
            if answer.get("signIn"):
                notes.append("Sign in with NetSuite on this PC to finish (the panel shows the link).")
            elif not answer.get("connected"):
                notes.append("NetSuite did not answer yet: " + (answer.get("error") or "check the setup page."))
        else:
            notes.append(answer.get("error") or "The settings could not be saved.")
    return {"ok": True, "done": done, "notes": notes}


def save_setup(handler, payload):
    """Writes config.json (and the licence, when one is pasted) from the setup page, then reconnects.
    A secret left blank keeps the one already saved, so the page never has to show secrets."""
    old = _saved_config(handler.config_path)
    cfg = dict(old)
    for key in CONFIG_KEYS:
        value = str(payload.get(key) or "").strip()
        if value or key not in SECRET_KEYS:
            cfg[key] = value
    # "Sign in with NetSuite": the integration's client ID (public, no secret). Empty = access tokens.
    cfg["clientId"] = str(payload.get("clientId") or "").strip() if payload.get("authMode") == "oauth2" else ""
    try:
        month = int(payload.get("fiscalYearStartMonth") or 1)
    except (TypeError, ValueError):
        month = 1
    cfg["fiscalYearStartMonth"] = month if 1 <= month <= 12 else 1
    update_url = str(payload.get("updateUrl") or "").strip()
    if update_url and not update_url.lower().startswith("https://"):
        return {"ok": False, "error": "The update address must start with https://"}
    cfg["updateUrl"] = update_url
    ai = dict(old.get("ai") or {}) if isinstance(old.get("ai"), dict) else {}
    if payload.get("aiOff"):
        ai["apiKey"] = ""
    elif str(payload.get("aiKey") or "").strip():
        ai["apiKey"] = str(payload["aiKey"]).strip()
    ai.setdefault("model", assistant_module.DEFAULT_MODEL)
    cfg["ai"] = ai
    try:
        checked = check_config(dict(cfg))
        text = str(payload.get("licence") or "").strip()
        found = licence_module.parse(text) if text else None
    except ValueError as err:
        return {"ok": False, "error": str(err)}
    if found:  # a licence that does not fit is refused, so it cannot replace one that works
        verdict = licence_module.check(found, checked["accountId"], handler.today)
        if not verdict["ok"]:
            return {"ok": False, "error": "Nothing was saved. " + verdict["message"]}
    os.makedirs(os.path.dirname(handler.config_path), exist_ok=True)
    with open(handler.config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    if found:
        with open(handler.licence_path, "w", encoding="utf-8") as f:
            json.dump(found, f, indent=1)
    configure(handler, checked, handler.today)
    if handler.session is not None and not handler.session.info()["signedIn"]:
        return {"ok": True, "connected": False, "signIn": True, "licence": handler.licence,
                "error": "Saved. Now sign in with NetSuite (the button below, or Sign in in the panel)."}
    ping = handler.backend({"action": "ping"})
    return {"ok": True, "connected": bool(ping.get("ok")), "error": ping.get("error", ""),
            "user": ping.get("user", ""), "licence": handler.licence, "signIn": bool(ping.get("signIn"))}


def make_certificate(cert_path=CERT_PATH, key_path=KEY_PATH):
    """A new certificate for https://localhost on this PC. Returns its thumbprint."""
    cert, key, thumbprint = tqcrypto.make_certificate()
    os.makedirs(os.path.dirname(cert_path), exist_ok=True)
    with open(key_path, "w", encoding="ascii") as f:
        f.write(key)
    with open(cert_path, "w", encoding="ascii") as f:
        f.write(cert)
    write_thumbprint(cert_path)
    return thumbprint


def write_thumbprint(cert_path=CERT_PATH):
    """localhost.crt.sha1: the certificate's Windows thumbprint, so the uninstaller removes exactly this one
    certificate from the user's trusted list (and never another "localhost" certificate)."""
    try:
        with open(cert_path, "r", encoding="ascii") as f:
            pem = f.read()
        der = base64.b64decode("".join(line for line in pem.splitlines() if line and not line.startswith("-----")))
        with open(cert_path + ".sha1", "w", encoding="ascii") as f:
            f.write(hashlib.sha1(der).hexdigest().upper())
    except (OSError, ValueError):
        pass


def make_server(port, cfg=None, use_tls=True, today=None, licence_text=None):
    """Builds the server. cfg=None means sample data. today and licence_text are for tests."""
    handler = type("BoundHandler", (Handler,), {})
    configure(handler, cfg, today, licence_text)
    if use_tls:   # the real connector (not the tests): look for updates a minute after starting, then weekly
        handler.watcher = updates_module.Watcher(
            lambda: _saved_config(handler.config_path).get("updateUrl", ""), VERSION,
            lambda: _saved_config(handler.config_path).get("autoUpdateCheck", True) is not False).start()
    httpd = ThreadingHTTPServer(("127.0.0.1", port), functools.partial(handler, directory=ADDIN_DIR))
    if use_tls:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(CERT_PATH, KEY_PATH)
        httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    return httpd


def main():
    parser = argparse.ArgumentParser(description="TQ Excel Connector local server")
    parser.add_argument("--port", type=int, default=3000)
    parser.add_argument("--mock", action="store_true", help="use sample data even if config.json exists")
    parser.add_argument("--http", action="store_true", help="plain HTTP (browser preview only; Excel needs HTTPS)")
    parser.add_argument("--make-cert", action="store_true", help="make this PC's localhost certificate and stop")
    parser.add_argument("--watch", action="store_true", help="keep the connector running: start it again if it stops")
    args = parser.parse_args()

    if sys.stdout is None or FROZEN:  # the installed exe has no console: keep a log instead
        os.makedirs(DATA_DIR, exist_ok=True)
        log = open(os.path.join(DATA_DIR, "server.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = log

    if args.make_cert:
        print("Certificate made, thumbprint %s" % make_certificate())
        return
    if args.watch:   # the start-up shortcut runs this: the connector itself is a child that is restarted if it stops
        care.watch(sys.argv[1:], FROZEN, os.path.abspath(__file__), args.port)
        return

    try:
        cfg = None if args.mock else load_config()
    except (ValueError, OSError) as err:
        print("Config problem: %s. Starting with sample data; fix it on the setup page." % err)
        cfg = None

    if os.path.exists(CERT_PATH) and not os.path.exists(CERT_PATH + ".sha1"):
        write_thumbprint()   # a certificate made by an older version
    if not args.http and not (os.path.exists(CERT_PATH) and os.path.exists(KEY_PATH)):
        if not FROZEN:
            sys.exit("No HTTPS certificate found. Run scripts\\setup-excel.ps1 first (see README).")
        make_certificate()
    if FROZEN and not args.http and os.path.exists(CERT_PATH):
        note = care.renew_certificate(CERT_PATH, KEY_PATH, make_certificate)
        if note:
            print(note)

    try:
        httpd = make_server(args.port, cfg, use_tls=not args.http)
    except OSError as err:
        sys.exit("Could not start on port %d (is the connector already running?): %s" % (args.port, err))
    scheme = "http" if args.http else "https"
    source = "NetSuite account %s" % cfg["accountId"] if cfg else "sample data"
    print("TQ Excel Connector running at %s://localhost:%d  (%s)" % (scheme, args.port, source))
    sys.stdout.flush()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
