"""TQ Excel Connector - ledger helper.

Sits between the add-in and NetSuite (or the sample data). It adds totals that
one formula can ask for, without any change in NetSuite:

    type:Income              every account of a NetSuite account type
    type:Bank,AcctRec        several types added together
    calc:NetIncome           income less costs between fromPeriod and toPeriod
    calc:NetIncomeYTD        the same, from the start of the fiscal year
    calc:RetainedEarnings    net income of all fiscal years before the period's
    4000, 4010, 4100-4199    a list of accounts and account ranges added together
    4000-4999, type:COGS     lists can mix numbers, ranges and types
    calc:Days                days from the start of fromPeriod to the end of toPeriod
    budget:4000-4999         the budget for those accounts and periods, not the actual postings
    budget(Plan A):type:Income   the same, for one budget category
    move:1100, 1200          the change in balance sheet accounts over the periods (for a cash flow)
    calc:Translation         currency translation difference of a consolidated balance sheet (0 otherwise)

Filters: each segment may hold several values ("Sales, Bar"), "not Sales" or "Food*" (with
everything under it), and a request may carry "filters": "Brand: Premium; Name: Al Noor".
They are worked out in NetSuite (tq_excel_core.js version 2 or later).

For drill-down, 'breakdown' lists the accounts behind a total with each one's amount.

It also keeps short-lived copies of the account, period and segment lists and
answers period questions (fiscal year, quarter) for the reports and the chat.
"""

import datetime
import fnmatch
import re
import threading
import time

CACHE_SECONDS = 120
NETSUITE_BATCH = 2000  # MAX_BATCH in tq_excel_core.js
SEGMENT_KINDS = ("subsidiary", "department", "class", "location")
FILTER_KINDS = SEGMENT_KINDS + ("filters",)   # "filters": any other filter, e.g. "Brand: Premium; Name: Al Noor"
CONSOLIDATED_RE = re.compile(r"(\*|\(consolidated\)|\band below)\s*$", re.I)
NEW_FILTER_RE = re.compile(r",|\*\s*$|^\s*(?:not|except|excluding)\s|\(consolidated\)|\band below\s*$", re.I)
CREDIT_TYPES = {"AcctPay", "CredCard", "OthCurrLiab", "LongTermLiab", "DeferRevenue", "Equity", "Income", "OthIncome"}
PL_SIGNS = {"Income": 1, "OthIncome": 1, "COGS": -1, "Expense": -1, "OthExpense": -1}
ACCOUNT_TYPES = (
    "Bank", "AcctRec", "UnbilledRec", "OthCurrAsset", "DeferExpense", "FixedAsset", "OthAsset",
    "AcctPay", "CredCard", "OthCurrLiab", "DeferRevenue", "LongTermLiab", "Equity",
    "Income", "COGS", "Expense", "OthIncome", "OthExpense", "Stat",
)
STAT_TYPES = {"Stat"}   # statistical accounts: units such as covers or headcount, per period
_TYPE_BY_LOWER = {t.lower(): t for t in ACCOUNT_TYPES}
EXCLUDING_RE = re.compile(r"\s+(excluding|except)\s+", re.I)   # "type:Expense excluding 8100-8199"
PREFIX_RE = re.compile(r"^(budget|move)(?:\(([^)]*)\))?:\s*(.+)$", re.I | re.S)
PERIOD_EXPR_RE = re.compile(r"(CM|LY|FYS|FYE|QS|QE)?([+-]\d{1,3})?", re.I)
RANGE_RE = re.compile(r"^([\w.]+)\s*(?:\.\.|-)\s*([\w.]+)$")
CALCS = {"netincome": "calc:NetIncome", "netincomeytd": "calc:NetIncomeYTD",
         "retainedearnings": "calc:RetainedEarnings", "translation": "calc:Translation"}


class LedgerError(Exception):
    pass


def text(value):
    return "" if value is None else str(value).strip()


def is_total(account):
    lowered = text(account).lower()
    return lowered.startswith("type:") or lowered.startswith("calc:")


def _date(value):
    try:
        return datetime.date.fromisoformat(text(value)[:10])
    except ValueError:
        return None


class Ledger:
    def __init__(self, backend, fiscal_start_month=1, today=None):
        self.backend = backend
        self.fiscal_start_month = fiscal_start_month if 1 <= fiscal_start_month <= 12 else 1
        self._today = today
        self._cache = {}
        self._lock = threading.Lock()
        self.record_url = ""   # set by server.py: a transaction's id is added to open it in NetSuite

    # ----------------------------------------------------------- lookups --

    def today(self):
        return self._today or datetime.date.today()

    def call(self, payload):
        result = self.backend(payload)
        if not isinstance(result, dict) or not result.get("ok"):
            raise LedgerError((result or {}).get("error") if isinstance(result, dict) else "Request failed")
        return result

    def _cached(self, key, load):
        with self._lock:
            hit = self._cache.get(key)
            if hit and time.time() - hit[0] < CACHE_SECONDS:
                return hit[1]
        value = load()
        with self._lock:
            self._cache[key] = (time.time(), value)
        return value

    def accounts(self):
        return self._cached("accounts", lambda: [
            {"number": text(a.get("number")), "name": text(a.get("name")), "type": text(a.get("type"))}
            for a in self.handle({"action": "accounts"}).get("accounts", []) if text(a.get("number"))
        ])

    def periods(self):
        """Periods in date order, each with its fiscal year and position in that year."""
        def load():
            out = []
            for p in self.call({"action": "periods"}).get("periods", []):
                start, end = _date(p.get("start")), _date(p.get("end"))
                if text(p.get("name")) and start and end:
                    out.append({"name": text(p["name"]), "start": start, "end": end})
            out.sort(key=lambda p: (p["start"], p["end"]))
            first = self.fiscal_start_month
            for p in out:
                month, year = p["start"].month, p["start"].year
                p["fy"] = year if first == 1 else (year + 1 if month >= first else year)
                p["fm"] = (month - first) % 12  # 0 = first month of the fiscal year
            return out
        return self._cached("periods", load)

    def setup(self):
        """What the NetSuite scripts can do: version, OneWorld, more than one currency."""
        def load():
            try:
                found = self.call({"action": "setup"})
            except LedgerError as err:
                if "unknown action" not in str(err).lower():
                    raise  # a time-out, a licence or permission problem: say that, and ask again next time
                return {"version": 1, "oneWorld": False, "multiCurrency": False}  # scripts from before filters
            return {"version": int(found.get("version") or 1), "oneWorld": bool(found.get("oneWorld")),
                    "multiCurrency": bool(found.get("multiCurrency"))}
        return self._cached("setup", load)

    def consolidating(self, subsidiary):
        """True when amounts of several subsidiaries are translated into one currency."""
        # As tq_excel_core.js decides: one subsidiary on its own keeps its currency; none, several, "not A"
        # or "Parent*" are translated when the subsidiaries use more than one currency.
        sub = text(subsidiary)
        if not self.setup()["multiCurrency"]:
            return False
        return not sub or bool(CONSOLIDATED_RE.search(sub) or NEW_FILTER_RE.search(sub))

    def _needs_new_scripts(self, request):
        return bool(text(request.get("filters"))) or any(
            NEW_FILTER_RE.search(text(request.get(k))) for k in SEGMENT_KINDS)

    def _needs_v4(self, request):
        """Book: and Currency: in the other filters need tq_excel_core.js version 4."""
        return bool(re.search(r"(^|;)\s*(?:accounting\s+)?(?:book|(?:reporting\s+)?currency)\s*:", text(request.get("filters")), re.I))

    def filter_fields(self):
        """Fields a report can be filtered on: [{id, label, kind}] (segments, Name, Item, custom segments)."""
        return self._cached("fields", lambda: [f for f in self.handle({"action": "filterFields"}).get("fields") or []
                                               if text(f.get("id")) and text(f.get("label"))])

    def segments(self, kind):
        return self._cached("seg:" + kind, lambda: [
            text(s.get("name")) for s in self.call({"action": "segments", "kind": kind}).get("segments", [])
            if text(s.get("name"))
        ])

    def _numbers(self):
        return self._cached("numbers", lambda: {a["number"].lower(): a["number"] for a in self.accounts()})

    def is_total(self, account):
        """True when the account argument is a total or a list, not one account number."""
        lowered = text(account).lower()
        if lowered.startswith(("type:", "calc:", "budget:", "budget(", "move:")):
            return True
        if not ("," in lowered or "-" in lowered or ".." in lowered):
            return False
        return lowered not in self._numbers()  # an account number may itself contain a dash

    def members(self, token):
        """Account numbers named by '4000, 4100-4199, type:COGS'; 'type:Expense excluding 8100-8199' leaves the
        accounts after "excluding" out (so a new account still lands in the first part). Raises LedgerError."""
        numbers, accounts = self._numbers(), self.accounts()
        split = EXCLUDING_RE.split(text(token), maxsplit=1)
        if len(split) == 3:   # [head, the word, tail]
            head, word, tail = split
            if not head.strip():
                raise LedgerError('"%s": put the accounts before "%s"' % (text(token), word))
            drop = set(self.members(tail)) if tail.strip() else set()
            return [n for n in self.members(head) if n not in drop]
        if text(token).lower() in numbers:
            return [numbers[text(token).lower()]]
        found, in_types = [], False
        for part in text(token).split(","):
            part = part.strip()
            lowered = part.lower()
            if not part:
                continue
            if lowered.startswith("calc:"):
                raise LedgerError('"%s" cannot be part of a list of accounts' % part)
            if lowered.startswith("type:"):
                in_types, lowered = True, lowered[5:].strip()
            elif lowered in numbers:
                in_types = False
                found.append(numbers[lowered])
                continue
            elif not (in_types and lowered in _TYPE_BY_LOWER):  # "type:Bank,AcctRec" carries on
                in_types = False
                match = RANGE_RE.match(part)
                if not match:
                    raise LedgerError('Account "%s" not found' % part)
                low, high = match.group(1).lower(), match.group(2).lower()
                if low.isdigit() and high.isdigit():
                    low, high = sorted((int(low), int(high)))
                    found += [a["number"] for a in accounts
                              if a["number"].isdigit() and low <= int(a["number"]) <= high]
                else:
                    low, high = sorted((low, high))
                    found += [a["number"] for a in accounts if low <= a["number"].lower() <= high]
                continue
            if lowered not in _TYPE_BY_LOWER:
                raise LedgerError('Unknown account type "%s"' % lowered)
            found += [a["number"] for a in accounts if a["type"] == _TYPE_BY_LOWER[lowered]]
        return list(dict.fromkeys(found))

    def days(self, from_period, to_period=""):
        periods = self.periods()
        first = self.period_index(from_period)
        last = self.period_index(to_period) if text(to_period) else first
        first, last = min(first, last), max(first, last)
        return (periods[last]["end"] - periods[first]["start"]).days + 1

    def period_index(self, name):
        wanted = text(name).lower()
        for i, p in enumerate(self.periods()):
            if p["name"].lower() == wanted:
                return i
        raise LedgerError('Period "%s" not found' % text(name))

    def month_index(self, year, month):
        """The ordinary period that starts in this calendar month (longest one, to skip adjustment periods)."""
        best = None
        for i, p in enumerate(self.periods()):
            if p["start"].year == year and p["start"].month == month:
                if best is None or (p["end"] - p["start"]) > (self.periods()[best]["end"] - self.periods()[best]["start"]):
                    best = i
        if best is None:
            raise LedgerError("There is no accounting period for %d-%02d in NetSuite" % (year, month))
        return best

    def year_range(self, fiscal_year):
        hits = [i for i, p in enumerate(self.periods()) if p["fy"] == fiscal_year]
        if not hits:
            raise LedgerError("There are no accounting periods for %d in NetSuite" % fiscal_year)
        return hits[0], hits[-1]

    def current_index(self):
        """The period containing today, else the latest one that has started."""
        periods, today = self.periods(), self.today()
        if not periods:
            raise LedgerError("NetSuite returned no accounting periods")
        best = 0
        for i, p in enumerate(periods):
            if p["start"] <= today:
                best = i
        for i, p in enumerate(periods):
            if p["start"] <= today <= p["end"] and (p["end"] - p["start"]).days > 20:
                return i
        return best

    def _step(self, index, months):
        """The ordinary month period `months` calendar months after the one that holds period `index`."""
        start = self.periods()[index]["start"]
        n = start.year * 12 + start.month - 1 + months
        return self.month_index(n // 12, n % 12 + 1)

    def resolve_period(self, expr, base=None):
        """A period name, or a period relative to today (or to `base`, an index), -> index. Raises LedgerError.

        "CM" this month, "CM-1" last month, "LY" the same month last year, "FYS"/"FYE" the first/last month of the
        fiscal year, "QS"/"QE" of the fiscal quarter; any of them may end in +n/-n months ("FYS-12" is the
        start of last year, "FYS-1" last year's end), and a bare "-1" / "+2" moves from the base."""
        raw = text(expr)
        if base is None:
            try:
                return self.period_index(raw)
            except LedgerError:
                pass
        match = PERIOD_EXPR_RE.fullmatch(raw.replace(" ", ""))
        if not match or not (match.group(1) or match.group(2)) or (base is None and not match.group(1)):
            raise LedgerError('Period "%s" not found. Use a period name such as Jan 2026, or CM, CM-1, FYS, FYE, '
                              'QS, QE, LY.' % raw)
        anchor, shift = (match.group(1) or "").upper(), int(match.group(2) or 0)
        at = self.current_index() if base is None else base
        periods = self.periods()
        fm = periods[at]["fm"]
        if anchor == "LY":
            shift -= 12
        elif anchor == "FYS":
            shift -= fm
        elif anchor == "FYE":
            shift += 11 - fm
        elif anchor == "QS":
            shift -= fm % 3
        elif anchor == "QE":
            shift += 2 - fm % 3
        return self._step(at, shift) if shift or anchor else at

    def period_name(self, expr, base=None):
        return self.periods()[self.resolve_period(expr, base)]["name"]

    def shift_period(self, base, how=""):
        """TQ.PERIOD(base, how): base is a period (or CM...), how a number of months or FYS/FYE/QS/QE/LY[+-n]."""
        at = self.resolve_period(base)
        how = text(how)
        if re.fullmatch(r"[+-]?\d+(\.0+)?", how):
            how = "%+d" % int(float(how))
        return self.periods()[self.resolve_period(how, at)]["name"] if how else self.periods()[at]["name"]

    def _named(self, request):
        """The request with relative periods (CM-1, FYS...) turned into period names."""
        out = dict(request)
        for key in ("fromPeriod", "toPeriod"):
            value = text(out.get(key))
            if value:
                out[key] = self.period_name(value)
        return out

    def fy_start(self, index):
        periods = self.periods()
        fiscal_year = periods[index]["fy"]
        while index > 0 and periods[index - 1]["fy"] == fiscal_year:
            index -= 1
        return index

    # ------------------------------------------------------------ totals --

    def split_prefix(self, account):
        """'budget(Plan A):4000' -> ('budget', 'plan a', '4000'); a plain account -> ('', '', account)."""
        token = text(account)
        match = PREFIX_RE.match(token)
        if not match or token.lower() in self._numbers():
            return "", "", token
        return match.group(1).lower(), text(match.group(2)).lower(), match.group(3).strip()

    def _movement(self, parts):
        """For balance sheet accounts: the closing balance less the balance before the first period."""
        periods, types = self.periods(), {a["number"]: a["type"] for a in self.accounts()}
        out = []
        for sign, part in parts:
            out.append((sign, part))
            if types.get(part["account"]) in PL_SIGNS:
                continue  # a P&L account's amount is already the movement
            first = self.period_index(part["fromPeriod"])
            if first > 0:
                out.append((-sign, dict(part, fromPeriod=periods[first - 1]["name"], toPeriod="")))
        return out

    def _expand(self, request):
        """One total request -> [(sign, plain request)]. Raises LedgerError."""
        prefix, option, token = self.split_prefix(request.get("account"))
        kind, _, rest = token.partition(":")
        base = {k: text(request.get(k)) for k in ("fromPeriod", "toPeriod") + FILTER_KINDS}
        if not base["fromPeriod"]:
            raise LedgerError("Period is required")
        if prefix == "budget":
            base["budget"] = option or "*"
        if prefix == "move":
            base["toPeriod"] = base["toPeriod"] or base["fromPeriod"]
        if kind.lower() != "calc":
            if kind.lower() == "type" and not rest.strip():
                raise LedgerError('Unknown account type ""')
            parts = [(1, dict(base, account=number)) for number in self.members(token)]
            return self._movement(parts) if prefix == "move" else parts
        if prefix == "move":
            raise LedgerError("move: works on accounts, not on %s" % token)
        calc = CALCS.get(rest.strip().lower())
        if not calc:
            raise LedgerError('Unknown total "%s"' % token)
        if calc == "calc:Translation":
            return self._translation(base)
        if calc != "calc:NetIncome":
            periods = self.periods()
            at = self.period_index(base["toPeriod"] or base["fromPeriod"])
            first = self.fy_start(at)
            if calc == "calc:NetIncomeYTD":
                base.update(fromPeriod=periods[first]["name"], toPeriod=periods[at]["name"])
            elif first == 0:
                return []  # nothing before the first fiscal year
            else:
                base.update(fromPeriod=periods[0]["name"], toPeriod=periods[first - 1]["name"])
        return [(PL_SIGNS[a["type"]], dict(base, account=a["number"]))
                for a in self.accounts() if a["type"] in PL_SIGNS]

    def _translation(self, base):
        """What makes a consolidated balance sheet balance (NetSuite's cumulative translation adjustment):
        assets less liabilities and equity less all net income to date. Nothing when not consolidating."""
        currency = re.search(r"(^|;)\s*(?:reporting\s+)?currency\s*:", text(base.get("filters")), re.I)
        if not (self.consolidating(base.get("subsidiary")) or currency):
            return []
        periods = self.periods()
        at = self.period_index(base["toPeriod"] or base["fromPeriod"])
        closing = dict(base, fromPeriod=periods[at]["name"], toPeriod="")
        earned = dict(base, fromPeriod=periods[0]["name"], toPeriod=periods[at]["name"])
        parts = []
        for a in self.accounts():
            if a["type"] in PL_SIGNS:
                parts.append((-PL_SIGNS[a["type"]], dict(earned, account=a["number"])))
            elif a["type"] in _TYPE_BY_LOWER.values():
                parts.append((-1 if a["type"] in CREDIT_TYPES else 1, dict(closing, account=a["number"])))
        return parts

    def balances(self, requests):
        """Like the NetSuite 'balances' action, with totals expanded and any batch size."""
        results = [None] * len(requests)
        flat, owners = [], []  # owners[n] = (request position, sign)
        for i, request in enumerate(requests):
            request = request if isinstance(request, dict) else {}
            try:
                request = self._named(request)
                if self._needs_new_scripts(request) and self.setup()["version"] < 2:
                    raise LedgerError("This filter needs the newer NetSuite scripts (tq_excel_core.js version 2). "
                                      "Ask your administrator to update them.")
                if self._needs_v4(request) and self.setup()["version"] < 4:
                    raise LedgerError("Book and Currency need the newer NetSuite scripts (tq_excel_core.js version 4). "
                                      "Ask your administrator to update them.")
                if not self.is_total(request.get("account")):
                    flat.append(request)
                    owners.append((i, None))
                    continue
                if text(request.get("account")).lower() == "calc:days":
                    results[i] = {"value": float(self.days(request.get("fromPeriod"), request.get("toPeriod")))}
                    continue
                parts = self._expand(request)
            except LedgerError as err:
                results[i] = {"error": str(err)}
                continue
            results[i] = {"value": 0.0}
            for sign, part in parts:
                flat.append(part)
                owners.append((i, sign))
        for start in range(0, len(flat), NETSUITE_BATCH):
            chunk = flat[start:start + NETSUITE_BATCH]
            answers = self.call({"action": "balances", "requests": chunk}).get("results", [])
            for n, _ in enumerate(chunk):
                i, sign = owners[start + n]
                answer = answers[n] if n < len(answers) and isinstance(answers[n], dict) else {"error": "No result returned"}
                if sign is None:
                    results[i] = answer
                elif "error" in results[i]:
                    pass
                elif isinstance(answer.get("value"), (int, float)):
                    results[i]["value"] += sign * answer["value"]
                else:
                    results[i] = {"error": answer.get("error") or "No value"}
        for result in results:
            if "value" in result:
                result["value"] = round(float(result["value"]), 2)
        return results

    # ------------------------------------------------------------- entry --

    def breakdown(self, request):
        """The accounts behind one cell, each with its amount: what a drill-down on a total shows first."""
        prefix, option, token = self.split_prefix(request.get("account"))
        if token.lower().startswith("calc:"):
            if token.lower() == "calc:days":
                raise LedgerError("That cell counts days; there is nothing behind it to drill into.")
            parts = self._expand(request)  # net income: every P&L account, already signed
        else:
            base = {k: text(request.get(k)) for k in ("fromPeriod", "toPeriod") + FILTER_KINDS}
            lead = "%s%s:" % (prefix, "(%s)" % option if option else "") if prefix else ""
            parts = [(1, dict(base, account=lead + number)) for number in self.members(token)]
        names = {a["number"]: a for a in self.accounts()}
        results = self.balances([part for _, part in parts])
        errors = [r["error"] for r in results if "error" in r]
        if errors:
            raise LedgerError(errors[0])
        rows = []
        for (sign, part), result in zip(parts, results):
            number = self.split_prefix(part["account"])[2]
            if abs(result["value"]) >= 0.005:
                rows.append({"account": number, "name": names.get(number, {}).get("name", ""),
                             "type": names.get(number, {}).get("type", ""),
                             "amount": round(sign * result["value"], 2),
                             "request": dict(part, account=number) if prefix != "budget" else None})
        return {"ok": True, "rows": rows, "total": round(sum(r["amount"] for r in rows), 2),
                "kind": prefix or "actual"}

    LIST_KINDS = {"accounts": "accounts", "account": "accounts", "periods": "periods", "period": "periods",
                  "subsidiaries": "subsidiary", "subsidiary": "subsidiary", "departments": "department",
                  "department": "department", "classes": "class", "class": "class", "locations": "location",
                  "location": "location"}
    LIST_MAX = 2000

    def list_rows(self, payload):
        """=TQ.LIST: the rows of a list that grows by itself (it is worked out again on every refresh).

        payload: kind (accounts, departments, classes, locations, subsidiaries, periods), filter (accounts:
        "4000-4999", "type:Expense", "61*"; others: a name or "Food*"), fromPeriod/toPeriod (periods: the range;
        with hideZero: the periods whose balance decides), hideZero, subsidiary, show (accounts: both, number,
        name). Returns rows: accounts [number, name] (or one of them), the others [name]."""
        kind = self.LIST_KINDS.get(text(payload.get("kind")).lower())
        if not kind:
            raise LedgerError('TQ.LIST: say which list, one of accounts, departments, classes, locations, '
                              'subsidiaries or periods (not "%s")' % text(payload.get("kind")))
        wanted = text(payload.get("filter"))
        first, last = text(payload.get("fromPeriod")), text(payload.get("toPeriod"))
        hide = str(payload.get("hideZero")).lower() in ("true", "1", "yes", "y")

        def like(value):   # a name filter: "Food*" from the start, "*bar" anywhere, plain words anywhere
            if not wanted:
                return True
            if "*" in wanted or "?" in wanted:
                return fnmatch.fnmatchcase(value.lower(), wanted.lower())
            return wanted.lower() in value.lower()

        if kind == "periods":
            periods = self.periods()
            lo = self.period_index(first) if first else 0
            hi = self.period_index(last) if last else len(periods) - 1
            lo, hi = min(lo, hi), max(lo, hi)
            names = [p["name"] for p in periods[lo:hi + 1]
                     if 20 < (p["end"] - p["start"]).days < 40 and like(p["name"])]   # months only
            return [[n] for n in names[:self.LIST_MAX]]
        if hide and not first:
            raise LedgerError("TQ.LIST: to leave out the ones with no balance, give the from period too")
        period = {"fromPeriod": first, "toPeriod": last or first}
        extra = {"subsidiary": text(payload.get("subsidiary"))} if text(payload.get("subsidiary")) else {}
        if kind == "accounts":
            accounts = [a for a in self.accounts() if a["type"] not in STAT_TYPES]
            if wanted and "*" not in wanted and "?" not in wanted:
                keep = set(self.members(wanted))
                found = [a for a in accounts if a["number"] in keep]
            else:
                found = [a for a in accounts if like(a["number"]) or like(a["name"])]
            if hide and found:
                values = self.balances([dict(period, account=a["number"], **extra) for a in found])
                errors = [v["error"] for v in values if "error" in v]
                if errors:
                    raise LedgerError(errors[0])
                found = [a for a, v in zip(found, values) if abs(v.get("value") or 0) >= 0.005]
            show = text(payload.get("show")).lower()
            cut = (lambda a: [a["number"]]) if show == "number" else (lambda a: [a["name"]]) if show == "name" \
                else (lambda a: [a["number"], a["name"]])
            return [cut(a) for a in found[:self.LIST_MAX]]
        names = [n for n in self.segments(kind) if like(n)]
        if hide and names:   # the ones with a profit and loss balance in the periods
            values = self.balances([dict(period, account="type:Income,OthIncome,COGS,Expense,OthExpense",
                                         **dict(extra, **{kind: n})) for n in names])
            names = [n for n, v in zip(names, values) if "error" not in v and abs(v.get("value") or 0) >= 0.005]
        return [[n] for n in names[:self.LIST_MAX]]

    def handle(self, payload):
        """What POST /api calls: totals handled here, the rest passed through."""
        action = payload.get("action")
        try:
            if action == "period":
                return {"ok": True, "name": self.shift_period(payload.get("base") or "CM", payload.get("shift"))}
            if action == "branding":   # the logo is the same for every sheet of a pack: asked once
                sub = text(payload.get("subsidiary"))
                return self._cached("brand:" + sub.lower(), lambda: self.call({"action": "branding", "subsidiary": sub}))
            if action in ("transactions", "breakdown") and isinstance(payload.get("request"), dict):
                payload = dict(payload, request=self._named(payload["request"]))
            if action in ("transactions", "breakdown") and isinstance(payload.get("request"), dict) \
                    and self._needs_new_scripts(payload["request"]) and self.setup()["version"] < 2:
                raise LedgerError("This filter needs the newer NetSuite scripts (tq_excel_core.js version 2). "
                                  "Ask your administrator to update them.")
            if action == "list":
                return {"ok": True, "rows": self.list_rows(payload)}
            if action == "balances" and isinstance(payload.get("requests"), list):
                return {"ok": True, "results": self.balances(payload["requests"])}
            if action == "breakdown" and isinstance(payload.get("request"), dict):
                return self.breakdown(payload["request"])
            if action == "transactions" and isinstance(payload.get("request"), dict) \
                    and self.is_total(payload["request"].get("account")):
                return {"ok": False, "error": "Drill-down works on a single account, not on a total."}
            if action == "setup":
                return dict({"ok": True}, **self.setup())
            if action == "filterFields" and self.setup()["version"] < 2:
                return {"ok": True, "fields": [{"id": k, "label": k.title(), "kind": "segment"} for k in SEGMENT_KINDS]}
            if action in ("accounts", "periods", "filterFields") or action == "segments":
                # The panel's pickers ask for these lists often; answer from the short-lived copy.
                key = "raw:%s:%s" % (action, payload.get("kind") or "")
                return self._cached(key, lambda: self.call({"action": action, "kind": payload.get("kind")}
                                                           if action == "segments" else {"action": action}))

        except LedgerError as err:
            return {"ok": False, "error": str(err)}
        answer = self.backend(payload)
        if action == "transactions" and self.record_url and isinstance(answer, dict):
            for row in answer.get("rows") or []:   # each line opens its transaction in NetSuite
                if isinstance(row, dict) and row.get("id"):
                    row["url"] = "%s%d" % (self.record_url, int(row["id"]))
        return answer
