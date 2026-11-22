"""TQ Excel Connector - chat for the side panel.

A built-in command reader, not an AI service: it recognises report requests
written in plain words ("P&L by quarter for 2026 vs last year, summary"),
builds them with reports.py and answers quick balance questions. Nothing is
sent anywhere except the usual lookups in NetSuite, and nothing is billed.

The panel posts {message, context}; context.last is the previous report's spec,
so a short follow-up ("by department", "now for 2025") changes that report.
The answer is {ok, reply, actions, spec}; the panel applies the actions in Excel.
"""

import difflib
import json
import re

import assistant as assistant_module
import dashboard
import layouts
import templates
import queries
import reports
from assistant import AssistantError
from ledger import FILTER_KINDS, LedgerError, PL_SIGNS, SEGMENT_KINDS

MAX_MESSAGE_CHARS = 2000
MAX_ANSWER_ACCOUNTS = 12
MAX_LIST_ROWS = 20000
MAX_LAYOUT_SHEETS = 10
MAX_LAYOUT_ROWS = 1000
MAX_LAYOUT_COLS = 24

MONTH_PAT = (r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|"
             r"sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)")
MONTH_NUM = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
ORDINALS = {"1st": 1, "first": 1, "2nd": 2, "second": 2, "3rd": 3, "third": 3, "4th": 4, "fourth": 4}
YEAR_PAT = r"(?:20\d{2}|\d{2})"
TOKEN_RE = re.compile(
    r"\b(?P<my>%(m)s)\.?[\s\-/,']*(?P<myy>%(y)s)\b"
    r"|\b(?:q|quarter\s*)(?P<qn>[1-4])(?:[\s\-/,']*(?:of\s+)?(?:fy\s*)?(?P<qy>%(y)s))?\b"
    r"|\b(?P<qo>1st|first|2nd|second|3rd|third|4th|fourth)\s+quarter(?:\s+(?:of\s+)?(?:fy\s*)?(?P<qoy>%(y)s))?\b"
    r"|\bh(?P<hn>[12])(?:[\s\-/,']*(?:fy\s*)?(?P<hy>%(y)s))?\b"
    r"|\bfy\s*(?P<fy2>\d{2})\b"
    r"|\b(?:fy\s*|year\s+)?(?P<y>20\d{2})\b"
    r"|\b(?P<rel>(?:this|current|last|previous|prior)\s+(?:financial\s+|fiscal\s+)?(?:year|month|quarter))\b"
    r"|\b(?P<m>%(m)s)\b(?![\s\-/,']*\d)" % {"m": MONTH_PAT, "y": YEAR_PAT})
RANGE_LINK = re.compile(r"\s*(?:to|through|thru|till|until|up\s*to|-|–|—)\s*")
COMPARE_PRIOR = re.compile(
    r"\b(?:vs\.?|versus|against|compared?\s+(?:to|with|against)|comparison\s+(?:to|with)|and|with)\s+(?:the\s+)?"
    r"(?:last|previous|prior|preceding)\s+(?P<unit>year|period|month|quarter)s?\b"
    r"|\b(?P<yoy>yoy|year[\s-]on[\s-]year|year[\s-]over[\s-]year|prior[\s-]year\s+comparison)\b")
COMPARE_WORD = re.compile(r"\b(?:compare[ds]?|comparison|comparative|vs\.?|versus|against|variance)\b")
NO_COMPARE = re.compile(r"\b(?:no|without|remove|drop)\s+(?:the\s+)?(?:comparison|compare|comparative|variance)\b")
LAST_N = re.compile(r"\b(?:last|past|trailing|previous|rolling)\s+(?P<n>\d{1,2}|twelve|six|three)\s+months?\b|\b(?P<ttm>ttm)\b")
YTD = re.compile(r"\b(?:ytd|year[\s-]to[\s-]date)\b")
LAYOUTS = [
    ("month", re.compile(r"\b(?:by|per|each)\s+month\b|\bmonth(?:ly|[\s-]?wise|[\s-]by[\s-]month)\b")),
    ("quarter", re.compile(r"\b(?:by|per|each)\s+quarter\b|\bquarter(?:ly|[\s-]?wise)\b")),
    ("total", re.compile(r"\b(?:one|single)\s+column\b|\btotal\s+only\b|\bas\s+a\s+total\b")),
]
SUMMARY = re.compile(r"\b(?:summary|summari[sz]ed?|condensed|high[\s-]level|totals\s+only|collapsed)\b")
DETAIL = re.compile(r"\b(?:detail(?:ed|s)?|by\s+account|account[\s-]level|expanded)\b")
ZERO = re.compile(r"\b(?:all\s+accounts|(?:include|including|with|show)\s+zero\w*(?:\s+balances?)?|zero\s+balances?)\b")
DROPDOWN_ADD = re.compile(r"\b(?:add\s+(?:a\s+)?)?drop[\s-]?downs?(?:\s+filters?)?\s+(?:for\s+|on\s+|of\s+)?"
                          r"(?P<f>[a-z][^;]*)$", re.I)
DROPDOWN_OFF = re.compile(r"\b(?:no|remove|without)\s+(?:the\s+)?drop[\s-]?downs?(?:\s+filters?)?\b", re.I)
OTHER_FILTER = re.compile(r"\b(?:filter(?:ed)?|where)\s+(?:by\s+|on\s+)?(?P<f>[a-z][\w &|./-]{0,40}:\s*\S.*)$", re.I)
CONSOLIDATED = re.compile(r"\bconsolidated\b", re.I)
NO_FILTER = re.compile(r"\b(?:no|clear|remove|without)\s+filters?\b|\ball\s+(?:departments|classes|locations|subsidiaries)\b")
HERE = re.compile(r"\b(?:here|(?:at|in|into|to)\s+(?:the\s+)?(?:selected|current|active|this)\s+cell|"
                  r"(?:selected|current|active)\s+cell|(?:at|in)\s+(?:the\s+)?selection|at\s+(?:the\s+)?cursor)\b")
TO_SHEET = re.compile(r"\b(?:(?:on|in|into|to)\s+(?:a\s+|the\s+)?(?:new\s+)?(?:work)?sheet|in\s+excel|write\s+it|put\s+it|insert\s+it)\b")
KIND_WORDS = {
    "subsidiary": r"(?:subsidiar(?:y|ies)|entit(?:y|ies)|compan(?:y|ies))",
    "department": r"(?:departments?|depts?)",
    "class": r"(?:class(?:es)?)",
    "location": r"(?:locations?|branch(?:es)?)",
}
DATA_REFRESH = re.compile(r"^\W*(?:refresh|update|reload|rerun|run)\s+(?:all\s+)?(?:the\s+|my\s+)?"
                          r"(?:data|quer(?:y|ies)|searches|lists)\W*$", re.I)
SEARCH_LIST = re.compile(r"^\W*(?:list|show|find)\s+(?:me\s+)?(?:the\s+|my\s+|all\s+)?saved\s+search(?:es)?\b[\s:]*(.*)$", re.I)
SEARCH_LOAD = re.compile(r"^\W*(?:load|run|get|import|add|open)\s+(?:the\s+|a\s+|my\s+)?(?:saved\s+)?search\b[\s:]*(.+)$", re.I)
EDIT_COLUMN_ADD = re.compile(r"^\W*(?:please\s+)?add\s+(?:a\s+|another\s+|one\s+more\s+)?(?:new\s+)?column\b[\s:]*(.*)$", re.I)
EDIT_COLUMN_DEL = re.compile(r"^\W*(?:please\s+)?(?:remove|delete|drop)\s+(?:the\s+)?column\b[\s:]*(.+)$", re.I)
EDIT_LINE_ADD = re.compile(r"^\W*(?:please\s+)?add\s+(?:a\s+|another\s+)?(?:new\s+)?(line|row|subtotal|sub-total|total)\b[\s:]*(.+)$", re.I)
EDIT_LINE_DEL = re.compile(r"^\W*(?:please\s+)?(?:remove|delete|drop)\s+(?:the\s+)?(?:line|row|subtotal)\b[\s:]*(.+)$", re.I)
EDIT_BUILD = re.compile(r"^\W*(?:please\s+)?(?:update|rebuild|refresh|build|run|redo)\s+(?:the\s+|my\s+)?(?:custom\s+)?(?:report|layout)\W*$", re.I)
_KIND = "(%s)" % "|".join(p[3:-1] for p in KIND_WORDS.values())  # any segment word, as one group
PAGES_ON = re.compile(r"\b(?:one|a|separate|own|different)\s+(?:sheet|tab)s?\s+(?:per|for\s+each|for\s+every|by|for)\s+%s\b"
                      r"|\b(?:each|every)\s+%s\s+(?:on|in|into|to)\s+(?:its|a|their)\s+(?:own\s+|separate\s+)?(?:sheet|tab)\b"
                      r"|\b(?:sheets|tabs)\s+by\s+%s\b|\b%s[\s-]?wise\s+(?:sheets|tabs)\b" % ((_KIND,) * 4), re.I)
PAGES_OFF = re.compile(r"\b(?:no|remove|without|stop)\s+(?:the\s+)?(?:separate\s+|extra\s+)?(?:sheets|tabs)\b"
                       r"|\b(?:one|single)\s+sheet\s+only\b", re.I)
BUDGET_COMPARE = re.compile(r"\b(?:vs\.?|versus|against|compared?\s+(?:to|with)|with|and)\s+(?:the\s+)?budgets?\b"
                            r"|\bbudgets?\s+(?:comparison|variance)\b|\bvariance\s+(?:to|against|vs\.?)\s+budgets?\b")
_DETAIL = r"(?:(?:the\s+)?(?:gl\s+codes?|gl\s+accounts?|account\s+details?|accounts?\s+(?:in|on|under|behind)\s+.+))"
ACCOUNTS_ON = re.compile(r"^\W*(?:(?:show|list|add|include|with)\s+%s|expand\s+(?:the\s+)?accounts?)\W*$" % _DETAIL, re.I)
ACCOUNTS_OFF = re.compile(r"^\W*(?:(?:hide|remove|no|without)\s+%s|collapse\s+(?:the\s+)?accounts?)\W*$" % _DETAIL, re.I)
LAYOUT_NEW = re.compile(r"\b(?:create|new|make|sample|start|set\s*up|add|blank)\s+(?:me\s+)?(?:a\s+|an\s+|the\s+|my\s+)?"
                        r"(?:(?:new|sample|custom|report|blank)\s+)*(?:layout|template)\b")
RULE_ADD = re.compile(r"^\W*(?:add\s+(?:a\s+)?)?highlight(?:\s+rule)?:?\s+(.+\s(?:red|yellow|amber|orange|green|blue|grey|gray))"
                      r"\W*$", re.I)
RULES_OFF = re.compile(r"\b(?:no|remove|clear|delete)\s+(?:the\s+|all\s+)?highlight\s+rules?\b", re.I)
HIGHLIGHT_ON = re.compile(r"\bhighlight\w*\s+(?:changes?\s+|variances?\s+)?(?:over|above|of|at|beyond|>=?)\s*"
                          r"(\d+(?:\.\d+)?\s*%?)", re.I)
HIGHLIGHT_OFF = re.compile(r"\b(?:no|remove|without|stop|turn\s+off)\s+(?:the\s+)?(?:variance\s+)?highlight", re.I)
COMMENTS_ON = re.compile(r"\b(?:add|with|show)\s+(?:a\s+)?comments?(?:\s+column)?\b", re.I)
COMMENTS_OFF = re.compile(r"\b(?:no|remove|without|hide)\s+(?:the\s+)?comments?(?:\s+column)?\b", re.I)
DASHBOARD = re.compile(r"\b(?:(executive|sales|cash|management|kpi|finance|financial|profit\s+and\s+loss|p\s*&\s*l|pnl|position|balance\s+sheet|ratios?|budget(?:\s+vs\.?\s+actual)?)\s+)?dash\s*-?board\b", re.I)
MONTH_YTD = re.compile(r"\b(?:month\s*(?:and|&|\+|/)\s*(?:ytd|year[\s-]to[\s-]date)|mtd\s*(?:and|&|/|\+)\s*ytd|"
                       r"roll(?:ing|s)?[\s-]?forward\w*|monthly\s+(?:management\s+)?(?:pack|report|p\s*&\s*l|pnl))(?![\w&])")
INTENTS = [
    ("layout", re.compile(r"\b(?:custom|layouts?|templates?)\b")),
    ("pl", re.compile(r"\b(?:p\s*&\s*l|p\s+and\s+l|pnl|p/l|profit\s*(?:and|&|n)?\s*loss|income\s+statement|"
                      r"statement\s+of\s+(?:comprehensive\s+)?income)\b|(?<![\w&])p&l(?![\w&])")),
    ("bs", re.compile(r"\b(?:balance\s*sheet|bs|statement\s+of\s+financial\s+position|financial\s+position)\b")),
    ("tb", re.compile(r"\b(?:trial\s+balance|tb)\b")),
    ("txn", re.compile(r"\b(?:transactions?|entries|journal\s+lines|postings|drill(?:\s*down)?|gl\s+detail)\b")),
]
LISTS = [
    ("accounts", re.compile(r"\b(?:chart\s+of\s+accounts|coa|accounts)\b")),
    ("periods", re.compile(r"\b(?:accounting\s+)?periods\b")),
    ("subsidiary", re.compile(r"\bsubsidiaries\b")),
    ("department", re.compile(r"\bdepartments\b")),
    ("class", re.compile(r"\bclasses\b")),
    ("location", re.compile(r"\blocations\b")),
]
LIST_VERB = re.compile(r"\b(?:list|show|give|get|insert|display|all|what\s+are)\b")
HELP = re.compile(r"^\W*(?:help|hi|hello|hey|examples?|commands?|what\s+can\s+(?:you|i)\s+(?:do|ask|say).*)\W*$")
METRICS = [  # (pattern, label, [(sign, key)], is profit and loss)
    (r"net\s+(?:income|profit|loss|result)|bottom\s+line", "Net income", [(1, "calc:NetIncome")], True),
    (r"gross\s+(?:profit|margin)", "Gross profit", [(1, "type:Income"), (-1, "type:COGS")], True),
    (r"other\s+income", "Other income", [(1, "type:OthIncome")], True),
    (r"other\s+expenses?", "Other expenses", [(1, "type:OthExpense")], True),
    (r"cost\s+of\s+(?:sales|goods(?:\s+sold)?)|cogs", "Cost of sales", [(1, "type:COGS")], True),
    (r"revenues?|income|sales|turnover", "Income", [(1, "type:Income")], True),
    (r"expenses?|opex|overheads?", "Expenses", [(1, "type:Expense")], True),
    (r"profit", "Net income", [(1, "calc:NetIncome")], True),
    (r"cash|bank\s+balances?", "Bank accounts", [(1, "type:Bank")], False),
    (r"receivables?", "Accounts receivable", [(1, "type:AcctRec")], False),
    (r"payables?", "Accounts payable", [(1, "type:AcctPay")], False),
]
STOPWORDS = set("""a an the of for in on at to from by with and or vs versus as is was were are be been it its this
that these those what whats which how much many show me give tell get need want can could would you i we my our us
please pls kindly balance balances total totals value amount amounts number figure account accounts acct now same
make do did does also add change switch instead again report sheet new create build generate prepare run let lets
see about just only too between through period periods month months year years quarter date today current latest
excel cell write put insert workbook up till until end ending ended there their not no yes ok okay thanks thank
out into all each per so far did had have has will if then than more less list display""".split())

HELP_TEXT = """I build NetSuite reports in this workbook from plain requests. I am a built-in command reader, \
so keep to wording like this:

Reports (each goes on a new sheet as live formulas)
- P&L for 2026 by month
- income statement Q1 2026 vs last year
- P&L Jan to Jun 2026 by department, summary
- balance sheet as at Oct 2026 vs Dec 2025
- balance sheet by quarter for 2026
- trial balance Oct 2026

Your own report format (Reports tab, Report layout, does the same by clicking)
- dashboard for Oct 2026      tiles and charts (executive); also sales, cash, finance (profit and loss),
                              financial position (ratios) and budget dashboard
- monthly P&L for Oct 2026    month and year to date against last year; pick another month in B8 to roll forward
- highlight changes over 10%  change % cells green at or above, red at or below; no highlights
- add comments column         a Comments column kept when the report is rebuilt; no comments column
- highlight 6100 > 5000 red   your own rules: an account, range or line, (in a column), > < between, a colour;
                              e.g. highlight Net Profit < 0 red, highlight all between -100 and 100 yellow; no highlight rules
- create layout            adds a sample layout sheet "TQ Layout"
- custom P&L for 2026 by month          sets the columns and builds the report
- add column Q1 2025
- add column Jan to Jun 2026 for department Sales called Sales H1
- add column Change = [2026]-[2025]
- remove column Total
- add line Rent accounts 6100, 6200-6299 negative in section Staff Cost
- add line Margin = [Gross Profit]/[Net Sales] percent
- add subtotal Total Staff Cost in section Staff Cost
- remove line Rent
- custom P&L for 2026 by quarter vs budget     actual, budget, variance and variance % for each period
- add column budget 2026
- show account detail / hide account detail   lists the GL accounts behind each line
- one sheet per location   the report for all, a summary, then one sheet for each location (or department, class, subsidiary)
- one sheet only
- update report            rebuilds it on the same sheets

Then change it with a short follow-up
- by quarter / by month / summary / detailed
- compare with last year / no comparison
- for department Sales / no filters
- filter Brand: Premium (any field: custom segments, Name, Item; several: Location: A, B; leave out: Location: not A)
- for subsidiary Parent consolidated (the parent and its subsidiaries in its currency)
- add dropdown filter Brand, Name (a row with a dropdown on the report for each; no dropdown filters)
- now for 2025

Other NetSuite data (Reports tab, Data tables, builds queries by clicking)
- list saved searches invoice
- load search Open Invoices     brings it onto its own sheet
- refresh data                  runs every data query in the workbook again

Quick answers and lists
- what was rent in Mar 2026?
- net income this year
- transactions for 6100 in Jan 2026
- list accounts / list periods / list departments"""

EXAMPLES = [part.strip() for line in HELP_TEXT.split("\n") if line.startswith("- ")
            for part in re.split(r"\s{2,}", line[2:])[0].split(" / ") if part.strip()]
AI_PREFIX = re.compile(r"^\W*(?:ai|ask\s+(?:the\s+)?ai)\s*[:,]\s*", re.I)
MAX_STEPS = 8
SPLIT = re.compile(r"\s*(?:\n|;(?!\s*[a-z][\w &|./-]{0,40}:)|,?\s+(?:and\s+)?then\s+|,?\s+and\s+also\s+|,?\s+after\s+that\s+)\s*", re.I)
POLITE = re.compile(r"^\W*(?:(?:hi|hello|hey)\b[,!\s]+)?(?:(?:please|pls|kindly)\s+)?(?:(?:can|could|would|will)\s+you\s+"
                    r"(?:please\s+|kindly\s+)?|i\s+(?:want|need|would\s+like|wish)(?:\s+you)?\s+to\s+(?:see\s+|have\s+)?|"
                    r"i\s+(?:want|need|would\s+like)\s+|let'?s\s+|(?:please|pls|kindly)\s+)", re.I)
COMMAND_WORD_LIST = """column columns line lines subtotal remove delete budget budgets balance sheet sheets trial
profit income statement quarter quarterly quarters month monthly months yearly department departments location
locations subsidiary subsidiaries class classes custom layout report update rebuild account accounts detail details
transactions summary detailed compare compared versus variance receivable receivables payable payables expenses
expense revenue sales saved search searches refresh create january february march april august september october
november december period periods section negative percent called total against without filters filter collapse
expand actual trailing rolling previous""".split()
COMMAND_WORDS = set(COMMAND_WORD_LIST)


def _text(value):
    return "" if value is None else str(value).strip()


def _year(digits):
    year = int(digits)
    return year if year >= 100 else 2000 + year


def _money(value):
    return "{:,.2f}".format(value)


class ChatError(Exception):
    pass


FILTER_STOP = re.compile(r"\s(?:by|vs\.?|versus|compared?(?:\s+with)?|summary|detailed|then|and\s+also|"
                         r"all\s+accounts|for\s+(?:19|20)\d\d|for\s+(?:this|last)\s+year)\b.*$", re.I)


def _filter_value(msg, match):
    """The text of "filter Brand: Premium ..." in the user's own capitals, without the words other parts
    of the message already used (periods, layouts) and stopping before words that begin another part."""
    start, end = match.start("f"), match.end("f")
    same = len(msg.original) == len(msg.text)
    chars = []
    for n in range(start, end):
        used = msg.text[n] == " " and (not same or msg.original[n] != " ")  # blanked by an earlier take()
        chars.append("\0" if used else (msg.original[n] if same else msg.text[n]))
    value = "".join(chars).split("\0")[0]
    value = FILTER_STOP.sub("", " " + value).strip()
    value = re.sub(r"(?:\s+(?:for|in|of|at|from|to|with|and|on))+\s*$", "", value, flags=re.I)
    return re.sub(r"\s+", " ", value).rstrip(" .,;")


def arg_text(args, key):
    value = args.get(key)
    return "" if value is None or isinstance(value, bool) else str(value).strip()[:200]


def _segment_of(field):
    """"Location" -> "location": the segment a "Field:" part names, else None."""
    word = field.strip().lower()
    return next((k for k in SEGMENT_KINDS if re.fullmatch(KIND_WORDS[k], word)), None)


def _route_other(text):
    """'Location: not Doha; Brand: Premium' -> ({"location": "not Doha"}, "Brand: Premium").

    A part that names subsidiary, department, class or location goes to that segment's own cell (B4:B7),
    so a later "filter location Riyadh" replaces it instead of both applying at once."""
    segments, rest = {}, []
    for part in [p.strip() for p in text.split(";") if p.strip()]:
        field, _, value = part.partition(":")
        kind = _segment_of(field) if value.strip() else None
        if kind:
            segments[kind] = value.strip()
        else:
            rest.append(part)
    return segments, "; ".join(rest)


def _merge_filters(old, new):
    """The filters after a follow-up: new values replace old ones; an empty dict clears them all."""
    if not new:
        return {}
    out = dict(old, **new)
    if "filters" not in new and out.get("filters"):
        # "Location: Doha" kept from earlier gives way to a new location filter
        kept = [p.strip() for p in out["filters"].split(";")
                if p.strip() and _segment_of(p.partition(":")[0]) not in new]
        out["filters"] = "; ".join(kept)
    return {k: v for k, v in out.items() if v}


def _filter_words(filters):
    """{"department": "Sales", "filters": "Brand: Premium"} -> 'department Sales, Brand: Premium'."""
    return ", ".join(v if k == "filters" else "%s %s" % (k, v) for k, v in filters.items() if v)


class _Message:
    """The message being read; recognised phrases are blanked out as they are used."""

    def __init__(self, raw):
        self.text = " " + re.sub(r"\s+", " ", raw.lower()) + " "
        self.original = " " + re.sub(r"\s+", " ", raw) + " "   # same positions, own capitals

    def take(self, pattern):
        match = pattern.search(self.text)
        if match:
            self.blank(match.start(), match.end())
        return match

    def blank(self, start, end):
        self.text = self.text[:start] + " " * (end - start) + self.text[end:]

    def words(self):
        return [w for w in re.findall(r"[a-z0-9][a-z0-9.\-]*", self.text) if w.strip(".-")]


class ChatService:
    """Answers POST /api/chat."""

    def __init__(self, ledger):
        self.ledger = ledger
        self.queries = queries.QueryService(ledger)  # also answers POST /api/query for the Data tab

    assistant = None  # the optional AI add-on; set by the server when config.json has an "ai" section

    def available(self):
        return True

    # ----------------------------------------------------------- periods --

    def _mention_range(self, item, default_year):
        ledger, periods = self.ledger, self.ledger.periods()
        kind = item[0]
        if kind == "month":
            index = ledger.month_index(item[2] or default_year, item[1])
            return index, index
        if kind == "year":
            return ledger.year_range(item[1])
        current = ledger.current_index()
        if kind in ("quarter", "half"):
            size = 3 if kind == "quarter" else 6
            year = item[2] or periods[current]["fy"]
            hits = [i for i, p in enumerate(periods) if p["fy"] == year and p["fm"] // size == item[1] - 1]
            if not hits:
                raise LedgerError("There are no accounting periods for %s%d %d in NetSuite"
                                  % ("Q" if size == 3 else "H", item[1], year))
            return hits[0], hits[-1]
        previous = item[1].split()[0] in ("last", "previous", "prior")
        unit = item[1].split()[-1]
        if unit == "month":
            index = current - 1 if previous else current
            if index < 0:
                raise LedgerError("There is no accounting period before %s" % periods[0]["name"])
            return index, index
        if unit == "year":
            if not previous:
                return ledger.fy_start(current), current
            return ledger.year_range(periods[current]["fy"] - 1)
        key = (periods[current]["fy"], periods[current]["fm"] // 3)
        hits = [i for i, p in enumerate(periods) if (p["fy"], p["fm"] // 3) == key]
        if not previous:
            return hits[0], current
        if hits[0] == 0:
            raise LedgerError("There is no accounting period before %s" % periods[0]["name"])
        key = (periods[hits[0] - 1]["fy"], periods[hits[0] - 1]["fm"] // 3)
        hits = [i for i, p in enumerate(periods) if (p["fy"], p["fm"] // 3) == key]
        return hits[0], hits[-1]

    def _ranges(self, msg):
        """Period ranges named in the message, in order. Consumes them."""
        periods = self.ledger.periods()
        known_years = {p["fy"] for p in periods} | {p["start"].year for p in periods}
        found = []
        for match in TOKEN_RE.finditer(msg.text):
            g = match.groupdict()
            if g["my"]:
                item = ("month", MONTH_NUM[g["my"][:3]], _year(g["myy"]))
            elif g["qn"]:
                item = ("quarter", int(g["qn"]), _year(g["qy"]) if g["qy"] else None)
            elif g["qo"]:
                item = ("quarter", ORDINALS[g["qo"]], _year(g["qoy"]) if g["qoy"] else None)
            elif g["hn"]:
                item = ("half", int(g["hn"]), _year(g["hy"]) if g["hy"] else None)
            elif g["fy2"]:
                item = ("year", 2000 + int(g["fy2"]))
            elif g["y"]:
                if int(g["y"]) not in known_years:
                    continue  # more likely an account number
                item = ("year", int(g["y"]))
            elif g["rel"]:
                item = ("rel", re.sub(r"\s+(?:financial|fiscal)", "", g["rel"]))
            else:
                before, after = msg.text[:match.start()], msg.text[match.end():]
                linked = re.search(r"(?:\bto|\bfrom|\bthrough|\bbetween|\band|-)\s*$", before) or RANGE_LINK.match(after)
                if g["m"] == "may" and not linked:
                    continue  # "may" the verb
                item = ("month", MONTH_NUM[g["m"][:3]], None)
            found.append((match.start(), match.end(), item))

        def year_of(item):
            return item[2] if item[0] in ("month", "quarter", "half") else item[1] if item[0] == "year" else None

        resolved = []
        for n, (start, end, item) in enumerate(found):
            years = [year_of(i) for _, _, i in found[n + 1:]] + [year_of(i) for _, _, i in reversed(found[:n])]
            default = next((y for y in years if y), self.ledger.today().year)
            resolved.append((start, end, self._mention_range(item, default)))
        ranges, n = [], 0
        while n < len(resolved):
            start, end, (first, last) = resolved[n]
            between = re.search(r"\bbetween\s*$", msg.text[:start])
            while n + 1 < len(resolved):
                gap = msg.text[end:resolved[n + 1][0]]
                if not (RANGE_LINK.fullmatch(gap) or (between and re.fullmatch(r"\s*and\s*", gap))):
                    break
                n += 1
                end, last = resolved[n][1], resolved[n][2][1]
            ranges.append((min(first, last), max(first, last)))
            n += 1
        for start, end, _ in found:
            msg.blank(start, end)
        return ranges

    # ------------------------------------------------------------ reading --

    def _segments(self, kind):
        try:
            return self.ledger.segments(kind)
        except LedgerError:
            return []

    def _modifiers(self, msg):
        """Everything in the message that shapes a report or a balance. Consumes it."""
        mods = {}
        if msg.take(NO_COMPARE):
            mods["compare"] = None
        else:
            match = msg.take(COMPARE_PRIOR)
            if match:
                mods["compare"] = "period" if match.group("unit") in ("period", "month", "quarter") else "year"
        match = msg.take(LAST_N)
        if match:
            count = 12 if match.group("ttm") else {"twelve": 12, "six": 6, "three": 3}.get(match.group("n")) \
                or int(match.group("n"))
            current = self.ledger.current_index()
            mods["range"] = (max(0, current - max(count, 1) + 1), current)
        ytd = bool(msg.take(YTD))
        ranges = self._ranges(msg)
        if ranges:
            mods["range"] = ranges[0]
            if len(ranges) > 1 and "compare" not in mods:
                mods["compare"] = ranges[1]
        if ytd:
            anchor = mods["range"][1] if "range" in mods else self.ledger.current_index()
            mods["range"] = (self.ledger.fy_start(anchor), anchor)
            mods.setdefault("layout", "total")
        if "compare" not in mods and msg.take(COMPARE_WORD):
            mods["compare"] = "year"
        else:
            msg.take(COMPARE_WORD)
        for name, pattern in LAYOUTS:
            if msg.take(pattern):
                mods["layout"] = name
        if msg.take(NO_FILTER):
            mods["filters"] = {}
        if msg.take(DROPDOWN_OFF):
            mods["dropdowns"] = []
        drop = DROPDOWN_ADD.search(msg.text)
        if drop:
            names = [n.strip() for n in FILTER_STOP.sub("", " " + drop.group("f")).split(",") if n.strip()]
            mods["dropdowns"] = ("add", names)
            msg.blank(drop.start(), drop.end())
        other = OTHER_FILTER.search(msg.text)
        if other:
            segments, rest = _route_other(_filter_value(msg, other))
            mods.setdefault("filters", {}).update(segments)
            if rest:
                mods["filters"]["filters"] = rest
            msg.blank(other.start(), other.end())
        for kind in SEGMENT_KINDS:
            words = KIND_WORDS[kind]
            if msg.take(re.compile(r"\b(?:by|per|across|for\s+each|each|split\s+by)\s+(?:all\s+)?%s\b|\b%s[\s-]?wise\b"
                                   % (words, words))):
                mods["split"] = kind
                continue
            for name in sorted(self._segments(kind), key=len, reverse=True):
                options = {name.lower(), name.split(" : ")[-1].strip().lower()}
                pattern = "|".join(re.escape(o) for o in sorted(options, key=len, reverse=True) if len(o) >= 2)
                if not pattern:
                    continue
                hit = msg.take(re.compile(r"\b%s\s*[:=]?\s*(?:is\s+|=\s*)?(?:%s)(?!\w)|(?<!\w)(?:%s)\s+%s\b"
                                          % (words, pattern, pattern, words)))
                if not hit and len(name) >= 3:
                    hit = msg.take(re.compile(r"\b(?:for|in|of|at)\s+(?:the\s+)?(?:%s)(?!\w)" % pattern))
                if hit:
                    mods.setdefault("filters", {})[kind] = name
                    break
        if msg.take(CONSOLIDATED) and (mods.get("filters") or {}).get("subsidiary"):
            mods["filters"]["subsidiary"] += "*"   # the parent and everything under it, in the parent's currency
        if msg.take(re.compile(r"\b(?:no|without|remove)\s+split\b")):
            mods["split"] = None
        if msg.take(SUMMARY):
            mods["detail"] = False
        elif msg.take(DETAIL):
            mods["detail"] = True
        if msg.take(ZERO):
            mods["zero"] = True
        return mods

    def _shift(self, bounds, mode):
        """The comparison range for 'year' (same months a year earlier) or 'period' (the periods just before)."""
        first, last = bounds
        periods = self.ledger.periods()
        if mode == "period":
            size = last - first + 1
            if first - size < 0:
                raise LedgerError("There are no earlier periods in NetSuite to compare with")
            return first - size, first - 1
        try:
            return tuple(self.ledger.month_index(periods[i]["start"].year - 1, periods[i]["start"].month)
                         for i in (first, last))
        except LedgerError:
            raise LedgerError("There are no periods a year earlier in NetSuite to compare with")

    def _spec(self, report, mods, last):
        """Combines defaults, the previous report (for a follow-up) and this message."""
        ledger, periods = self.ledger, self.ledger.periods()
        current = ledger.current_index()
        flows = report in ("pl", "layout")  # covers a range of periods, not a position at one date
        spec = {"report": report, "layout": None, "split": None, "detail": True, "zero": False,
                "filters": {}, "compareMode": None, "dropdowns": []}
        notes = []
        if last:
            spec.update({k: last[k] for k in ("layout", "split", "detail", "zero", "compareMode", "layoutName")
                         if k in last})
            spec["filters"] = dict(last.get("filters") or {})
            spec["dropdowns"] = list(last.get("dropdowns") or [])
            spec["range"] = (ledger.period_index(last["from"]), ledger.period_index(last["to"]))
            if last.get("compare"):
                spec["compare"] = tuple(ledger.period_index(name) for name in last["compare"])
            if "layout" in mods or "split" in mods:
                spec["layout"] = None
        if "range" in mods:
            spec["range"] = mods["range"]
            if spec.get("compare") and "compare" not in mods and spec["compareMode"] in ("year", "period"):
                spec["compare"] = self._shift(mods["range"], spec["compareMode"])
        elif "range" not in spec:
            spec["range"] = (ledger.fy_start(current), current) if flows else (current, current)
            notes.append("No period was named, so I used %s." % (
                "this fiscal year to date" if flows else periods[current]["name"]))
        for key in ("layout", "split", "detail", "zero"):
            if key in mods:
                spec[key] = mods[key]
        if "layout" in mods and "split" not in mods:
            spec["split"] = None  # "by month" replaces an earlier "by department"
        if "filters" in mods:
            spec["filters"] = _merge_filters(spec["filters"], mods["filters"])
        if "dropdowns" in mods:
            if mods["dropdowns"] and mods["dropdowns"][0] == "add":
                spec["dropdowns"] = [f["label"] for f in reports.dropdown_fields(
                    self.ledger, spec["dropdowns"] + mods["dropdowns"][1])]
            else:
                spec["dropdowns"] = []
        first, last_index = spec["range"]
        if not flows and first <= current < last_index:
            last_index = current  # a position cannot be later than today
        if "compare" in mods:
            if mods["compare"] is None:
                spec["compare"], spec["compareMode"] = None, None
            elif isinstance(mods["compare"], str):
                spec["compare"], spec["compareMode"] = self._shift((first, last_index), mods["compare"]), mods["compare"]
            else:
                spec["compare"], spec["compareMode"] = mods["compare"], "custom"
        if spec["split"]:
            spec["filters"].pop(spec["split"], None)
            if spec.get("compare"):
                notes.append("A split by %s cannot also show a comparison, so the comparison is left out."
                             % spec["split"])
                spec["compare"], spec["compareMode"] = None, None
        if report == "tb":
            spec.update(layout="total", split=None, compare=None, compareMode=None, detail=True)
        if not spec["layout"]:
            spec["layout"] = "month" if (flows and first != last_index and not spec.get("compare")
                                         and not spec["split"]) else "total"
        spec.update(start=first, end=last_index, compare=spec.get("compare"))
        spec.pop("range")
        return spec, notes

    def _wire(self, spec):
        """The spec as sent to the panel and back: period names, not positions."""
        periods = self.ledger.periods()
        return {
            "report": spec["report"], "from": periods[spec["start"]]["name"], "to": periods[spec["end"]]["name"],
            "compare": [periods[i]["name"] for i in spec["compare"]] if spec.get("compare") else None,
            "compareMode": spec.get("compareMode"), "layout": spec["layout"], "split": spec["split"],
            "detail": spec["detail"], "zero": spec["zero"], "filters": spec["filters"],
            "layoutName": spec.get("layoutName"), "dropdowns": list(spec.get("dropdowns") or []),
        }

    @staticmethod
    def _last(context):
        last = context.get("last")
        if not isinstance(last, dict) or last.get("report") not in ("pl", "bs", "tb", "layout"):
            return None
        if not (_text(last.get("from")) and _text(last.get("to"))):
            return None
        clean = {"report": last["report"], "from": _text(last["from"]), "to": _text(last["to"])}
        if _text(last.get("layoutName")):
            clean["layoutName"] = _text(last["layoutName"])[:31]
        if isinstance(last.get("compare"), list) and len(last["compare"]) == 2:
            clean["compare"] = [_text(n) for n in last["compare"]]
        if last.get("layout") in ("total", "month", "quarter"):
            clean["layout"] = last["layout"]
        if last.get("split") in SEGMENT_KINDS:
            clean["split"] = last["split"]
        if last.get("compareMode") in ("year", "period", "custom"):
            clean["compareMode"] = last["compareMode"]
        clean["detail"] = last.get("detail") is not False
        if isinstance(last.get("dropdowns"), list):
            clean["dropdowns"] = [_text(n)[:60] for n in last["dropdowns"] if _text(n)][:reports.MAX_DROPDOWNS]
        clean["zero"] = last.get("zero") is True
        if isinstance(last.get("filters"), dict):
            clean["filters"] = {k: _text(v)[:200] for k, v in last["filters"].items()
                                if k in FILTER_KINDS and _text(v)}
        return clean

    # ----------------------------------------------------------- answers --

    @staticmethod
    def _layout_sheets(context):
        """The layout sheets the panel found in the workbook: [{name, row, rows}]."""
        found = []
        sheets = context.get("layouts")
        for sheet in (sheets if isinstance(sheets, list) else [])[:MAX_LAYOUT_SHEETS]:
            if not isinstance(sheet, dict) or not isinstance(sheet.get("rows"), list) or not _text(sheet.get("name")):
                continue
            rows = [row[:MAX_LAYOUT_COLS] if isinstance(row, list) else [] for row in sheet["rows"][:MAX_LAYOUT_ROWS]]
            first = sheet.get("row")
            found.append({"name": _text(sheet["name"])[:31], "rows": rows,
                          "row": first if isinstance(first, int) and first >= 0 else 0})
        return found

    def _on_layout_report(self, context):
        """True when the request is about a layout report: the last report was one, or the open sheet is one."""
        last = self._last(context)
        if last:
            return last["report"] == "layout"
        active = _text(context.get("activeSheet")).lower()
        for sheet in self._layout_sheets(context):
            try:
                design = layouts.parse(sheet["name"], sheet["rows"], sheet["row"])
            except reports.ReportError:
                continue
            if active in (sheet["name"].lower(), layouts.report_sheet(design).lower()):
                return True
        return False

    def _pick_layout(self, context, msg=None, wanted=None):
        """The design a request means: a layout it names, the previous one, the open sheet's, else the first."""
        sheets = self._layout_sheets(context)
        if not sheets:
            raise ChatError('This workbook has no layout sheet yet. Say "create layout" (or use New in the Designer '
                            'tab) and I will add a sample one called "%s" for you to edit.' % layouts.SHEET_PREFIX)
        parse = lambda sheet: layouts.parse(sheet["name"], sheet["rows"], sheet["row"])  # noqa: E731
        if msg is not None:
            for sheet in sorted(sheets, key=lambda s: -len(s["name"])):
                suffix = sheet["name"][len(layouts.SHEET_PREFIX):].strip(" -_").lower()
                if suffix and msg.take(re.compile(r"(?<!\w)%s(?!\w)" % re.escape(suffix))):
                    return parse(sheet), len(sheets)
        active = _text(context.get("activeSheet")).lower()
        for name in (_text(wanted).lower(), active):
            hit = next((s for s in sheets if name and s["name"].lower() == name), None)
            if hit:
                return parse(hit), len(sheets)
        for sheet in sheets:  # the open sheet may be a report: use the layout it was built from
            try:
                design = parse(sheet)
            except reports.ReportError:
                continue
            if active and layouts.report_sheet(design).lower() == active:
                return design, len(sheets)
        return parse(sheets[0]), len(sheets)

    def _publish(self, design, lead, sheet_count=1, notes=(), spec=None):
        """Writes the layout sheet back and rebuilds its report, both in place."""
        if not design["columns"]:
            design["columns"] = layouts.default_columns(self.ledger)
        built, info = layouts.build_pack(self.ledger, design)
        misses = list(dict.fromkeys(design.pop("misses", [])))
        if misses:
            notes = list(notes) + ["Highlight rule %s matches no line on the report: a line is coloured when all its "
                                   "accounts are in the rule (or with Show accounts on, the account's own row)."
                                   % ", ".join('"%s"' % m for m in misses)]
        if design.get("pages") and not info["pages"] and lead.startswith("The report now has one sheet per"):
            lead = ""   # the line below says why no sheet was made
        lines = [lead] if lead else []
        lines.append('%s is on sheet "%s": %d lines by %d columns, from layout sheet "%s".'
                     % (info["title"], info["sheet"], info["lines"], info["columns"], design["name"]))
        if info["pages"]:
            shown = ", ".join(info["pages"][:10]) + (" and %d more" % (len(info["pages"]) - 10)
                                                     if len(info["pages"]) > 10 else "")
            lines.append("After it there is one sheet for each %s with a balance (%d): %s."
                         % (design["pages"], len(info["pages"]), shown))
            if design.get("pagesWide"):
                lines.append("No %s has a balance in the report's own periods, so these are the ones with a balance "
                             "in other periods: their sheets show zero until you change the year in B8 (or the "
                             "periods) to one with postings." % design["pages"])
            if info["summary"]:
                lines.append('Sheet "%s" puts them side by side, one row each.' % info["summary"])
        elif design.get("pages"):
            lines.append("No sheet per %s was made: no %s has a balance on these accounts in any period. Check that "
                         "postings in NetSuite carry a %s." % ((design["pages"],) * 3))
        if sheet_count > 1:
            lines.append("This workbook has %d layout sheets. To use another, open it or add the end of its name "
                         "to the request." % sheet_count)
        unmapped = info["unmapped"]
        if unmapped:
            shown = ", ".join("%s %s" % (a["number"], a["name"]) for a in unmapped[:8])
            more = " and %d more" % (len(unmapped) - 8) if len(unmapped) > 8 else ""
            lines.append("%d profit and loss account(s) with a balance in these periods are not in the layout: %s%s."
                         % (len(unmapped), shown, more))
        elif unmapped is not None:
            lines.append("Every profit and loss account with a balance in these periods is in the layout.")
        if info["empty"]:
            lines.append("These lines match no account: %s." % ", ".join(info["empty"][:8]))
        lines += list(notes)
        if design["filters"]:
            lines.append("Filtered to " + _filter_words(design["filters"]) + ".")
        lines.append("The numbers are live formulas: change %sthe filters in B3:B7, then use Refresh all. "
                     "To change the report itself, say e.g. add column Q1 2025, remove column Total, "
                     "add line Rent accounts 6100, or use Report layout in the Reports tab."
                     % ("the year in B8 or " if info["year"] else
                        "the month in B8 (every column rolls forward with it) or " if info.get("month") else ""))
        periods = self.ledger.periods()
        spans = [c["at"] for c in design["columns"] if c.get("at")]
        wire = self._wire(spec) if spec else {
            "report": "layout", "compare": None, "compareMode": None, "layout": "total", "split": None,
            "detail": True, "zero": False, "filters": design["filters"],
            "from": periods[min(s[0] for s in spans)]["name"] if spans else "",
            "to": periods[max(s[1] for s in spans)]["name"] if spans else ""}
        wire["layoutName"] = design["name"]
        return {"ok": True, "reply": "\n".join(lines), "spec": wire, "design": layouts.public(design),
                "actions": [layouts.emit(design, activate=False)] + built}

    def _layout_report(self, mods, last, design, sheet_count, lead=""):
        """A report from a layout. Naming periods or a column shape replaces the layout's columns."""
        spec, notes = None, []
        budget = mods.pop("budget", False)
        if budget:
            mods["compare"] = None  # the comparison is with the budget, not with other periods
        if budget or any(key in mods for key in ("range", "layout", "split", "compare")) or not design["columns"]:
            spec, notes = self._spec("layout", mods, last)
            design["columns"] = layouts.columns_from_spec(self.ledger, spec)
            if spec["split"]:
                design["filters"].pop(spec["split"], None)
            if budget:
                design["columns"] = layouts.with_budget(design["columns"])
        if "filters" in mods:
            design["filters"] = _merge_filters(design["filters"], mods["filters"])
        return self._publish(design, lead, sheet_count, notes, spec)

    def _sample_design(self, context, name=""):
        taken = {s["name"].lower() for s in self._layout_sheets(context)}
        name = reports.sheet_name(name)
        if name and not name.lower().startswith(layouts.SHEET_PREFIX.lower()):
            name = reports.sheet_name("%s %s" % (layouts.SHEET_PREFIX, name))
        base = name or layouts.SHEET_PREFIX
        name, n = base, 2
        while name.lower() in taken:
            name, n = reports.sheet_name("%s %d" % (base[:28], n)), n + 1
        return layouts.sample(self.ledger, name)

    def _library_import(self, context, args):
        """A layout from the company library: its own layout sheet (a free name) and its report, built now."""
        design = args.get("design")
        if not isinstance(design, dict) or not isinstance(design.get("lines"), list):
            raise ChatError("That library entry is not a layout.")
        design = json.loads(json.dumps(design))
        taken = {s["name"].lower() for s in self._layout_sheets(context)}
        design["name"] = templates.layout_name(_text(args.get("name"))[:40] or _text(design.get("title"))[:40], taken)
        design["title"] = _text(design.get("title"))[:60] or "Report"
        sheets = [_text(s).lower() for s in context.get("sheets") or []]
        if reports.sheet_name(design["title"]).lower() in sheets:
            base, n = design["title"], 2   # never write the report over a sheet that is already there
            while reports.sheet_name("%s %d" % (base, n)).lower() in sheets:
                n += 1
            design["title"] = "%s %d" % (base, n)
        layouts.check(self.ledger, design)
        return self._publish(design, 'From the company library: layout sheet "%s".' % design["name"], 1)

    def _layout_new(self, context, args):
        """Designer: a new layout from a ready-made template, or from a format typed on a sheet."""
        taken = {s["name"].lower() for s in self._layout_sheets(context)}
        wanted = _text(args.get("name"))[:60]
        rows = args.get("rows")
        if isinstance(rows, list):
            source = _text(args.get("sheet"))[:31]
            name = templates.layout_name(wanted or source, taken)
            design, notes = templates.from_sheet(self.ledger, name, wanted or source or "Report", rows)
            lead = 'Layout sheet "%s" was made from the format on sheet "%s".' % (name, source)
        else:
            template = _text(args.get("template")).lower() or "management"
            design = templates.build(self.ledger, template, templates.layout_name(wanted, taken))
            if wanted:
                design["title"] = wanted
            notes = design.pop("notes", [])
            lead = 'Layout sheet "%s" was made from the template "%s".' % (design["name"], next(
                (t["name"] for t in templates.TEMPLATES if t["id"] == template), template))
        if _text(args.get("month")):
            design["month"] = _text(args.get("month"))[:40]
        if isinstance(args.get("filters"), dict):
            design["filters"] = {k: _text(v)[:200] for k, v in args["filters"].items() if k in FILTER_KINDS and _text(v)}
        if any(reports.sheet_name(design["title"]).lower() == _text(s).lower() for s in context.get("sheets") or []):
            base, n = design["title"], 2  # never write the report over a sheet that is already there
            while any(reports.sheet_name("%s %d" % (base, n)).lower() == _text(s).lower()
                      for s in context.get("sheets") or []):
                n += 1
            design["title"] = "%s %d" % (base, n)
        return self._publish(design, lead, 1, notes)

    def _dashboard(self, context, kind, month="", filters=None, title=""):
        try:
            action, info = dashboard.build(self.ledger, kind, month, filters or {}, title,
                                           [_text(n) for n in context.get("sheets") or []])
        except (LedgerError, reports.ReportError) as err:
            return {"ok": True, "actions": [], "reply": "The dashboard could not be built: %s" % err}
        return {"ok": True, "actions": [action], "reply": (
            'Dashboard "%s" for %s: %d charts and the key figures as tiles. Pick another month in B8 (or change the '
            "filters in B3:B7) and use Refresh this sheet; everything moves. The numbers come from the data block on "
            "the right of the sheet." % (info["sheet"], info["month"], info["charts"]))}

    WIZARD_ROLLING = ("monthytd", "balance", "hotel", "usali", "restaurant", "forecast", "rolling12")   # follow a month cell

    def _wizard(self, context, args):
        """The report wizard: every report it makes is a layout, so it can be changed and rebuilt in place.

        args: template, title, month (rolling reports), year (or from + to, period names) + shape (month/quarter/total) + compare
        (none/year/budget) for the others, filters {}, dropdowns [], pages, detail, comments, highlight."""
        template = _text(args.get("template")).lower() or "monthytd"
        taken = {s["name"].lower() for s in self._layout_sheets(context)}
        design = templates.build(self.ledger, template, templates.layout_name(_text(args.get("title"))[:60], taken))
        if _text(args.get("title")):
            design["title"] = _text(args.get("title"))[:60]
        periods = self.ledger.periods()
        if template in self.WIZARD_ROLLING:
            if _text(args.get("month")):
                design["month"] = _text(args.get("month"))
        elif template != "cashflow" or _text(args.get("year")) or _text(args.get("from")):
            if _text(args.get("from")) and _text(args.get("to")):   # a range of months: from one period to another
                first = self.ledger.period_index(_text(args.get("from")))
                last = self.ledger.period_index(_text(args.get("to")))
                first, last = min(first, last), max(first, last)
            else:
                year = int(args.get("year") or periods[self.ledger.current_index()]["fy"])
                first, last = self.ledger.year_range(year)
                last = min(last, max(first, self.ledger.current_index())) \
                    if year == periods[self.ledger.current_index()]["fy"] else last
            shape = _text(args.get("shape")).lower()
            shape = shape if shape in ("month", "quarter", "total") else "month"
            compare_mode = _text(args.get("compare")).lower()
            compare = None
            if compare_mode == "year":
                try:
                    compare = (self.ledger._step(first, -12), self.ledger._step(last, -12))
                except LedgerError:
                    compare = None
            design["columns"] = layouts.columns_from_spec(self.ledger, {
                "start": first, "end": last, "range": (first, last), "layout": shape, "split": None,
                "compare": compare, "compareMode": "year" if compare else None, "filters": {}})
            if compare_mode == "budget":
                design["columns"] = layouts.with_budget(design["columns"])
        filters = args.get("filters") if isinstance(args.get("filters"), dict) else {}
        design["filters"] = {k: _text(v)[:200] for k, v in filters.items() if k in FILTER_KINDS and _text(v)}
        if isinstance(args.get("dropdowns"), list):
            design["dropdowns"] = [f["label"] for f in reports.dropdown_fields(self.ledger, args["dropdowns"])]
        design["pages"] = layouts.page_kind(arg_text(args, "pages"))
        design["detail"] = bool(args.get("detail")) if "detail" in args else bool(design.get("detail"))
        design["comments"] = bool(args.get("comments"))
        design["highlight"] = layouts.highlight_value(arg_text(args, "highlight"))
        design["commentAt"] = layouts.highlight_value(arg_text(args, "commentAt"))
        if design["commentAt"]:
            design["comments"] = True
        sheets = [_text(s).lower() for s in context.get("sheets") or []]
        if reports.sheet_name(design["title"]).lower() in sheets:   # never write over a sheet that is there
            base, n = design["title"], 2
            while reports.sheet_name("%s %d" % (base, n)).lower() in sheets:
                n += 1
            design["title"] = "%s %d" % (base, n)
        notes = design.pop("notes", [])
        return self._publish(design, 'Report "%s" is built. Its layout is on sheet "%s": change it there, in Report '
                                     "layout, or in the chat, and it is rebuilt in place." % (design["title"], design["name"]),
                             notes=notes)

    def _layout_sample(self, context, name=""):
        design = self._sample_design(context, name)
        name = design["name"]
        return {"ok": True, "actions": [layouts.emit(design)], "design": layouts.public(design),
                "reply": 'I added layout sheet "%s" with %d lines and %d columns, made from the profit and loss '
                         "accounts that have a balance.\n"
                         "Change it to your format in the Reports tab, on the sheet itself, or here in the chat: "
                         "add column Q1 2025, add line Rent accounts 6100-6199, remove line Rent.\n"
                         "Then say: update report" % (name, len(design["lines"]), len(design["columns"]))}

    def _range_name(self, first, last):
        """A short name for a run of periods: "Jan 2026", "Q1 2026", "2026", else "Jan 2026 to May 2026"."""
        periods = self.ledger.periods()
        quarter = lambda i: (periods[i]["fy"], periods[i]["fm"] // 3)  # noqa: E731
        whole = (first != last and quarter(first) == quarter(last)
                 and (first == 0 or quarter(first - 1) != quarter(first))
                 and (last + 1 == len(periods) or quarter(last + 1) != quarter(last)))
        if whole:
            return "Q%d %d" % (quarter(first)[1] + 1, quarter(first)[0])
        return reports.range_label(self.ledger, first, last)

    # ------------------------------------------------------------ editing --

    def _data(self, raw, context):
        """Chat requests about data queries: saved searches and refreshing. None if it is not one."""
        text = re.sub(r"\s+", " ", raw).strip()
        sheet = context.get("queries") if isinstance(context.get("queries"), dict) else {}
        try:
            if DATA_REFRESH.match(text):
                state = queries.parse(sheet.get("rows"), sheet.get("row") if isinstance(sheet.get("row"), int) else 0)
                actions, lines = self.queries.run_all(state)
                return {"ok": True, "reply": "\n".join(lines), "actions": actions}
            match = SEARCH_LIST.match(text)
            if match:
                found = self.queries.searches(match.group(1).strip(' "\''))
                if not found:
                    return {"ok": True, "actions": [], "reply": "No saved search matched."}
                lines = ["%s  (%s, %s)" % (s["title"], s["id"], s["type"]) for s in found[:20]]
                more = "\n… and %d more. Add a word to narrow it down." % (len(found) - 20) if len(found) > 20 else ""
                return {"ok": True, "actions": [], "reply": "\n".join(lines) + more
                        + "\nTo bring one into Excel, say: load search <name or ID>"}
            match = SEARCH_LOAD.match(text)
            if match:
                wanted = match.group(1).strip(' "\'')
                found = self.queries.searches(wanted)
                exact = [s for s in found if wanted.lower() in (s["id"].lower(), s["title"].lower())]
                if len(exact or found) != 1:
                    if not found:
                        raise queries.QueryError('No saved search matched "%s". Say: list saved searches' % wanted)
                    raise queries.QueryError("%d saved searches match. Say which one:\n%s" % (
                        len(found), "\n".join("%s  (%s)" % (s["title"], s["id"]) for s in found[:15])))
                hit = (exact or found)[0]
                return self.queries.handle({"op": "save", "context": context,
                                            "args": {"name": hit["title"], "type": "search", "source": hit["id"]}})
        except (queries.QueryError, LedgerError) as err:
            return {"ok": True, "reply": str(err), "actions": []}
        return None

    def _edit(self, raw, context):
        """Chat requests that change a layout: add/remove a column or a line, or rebuild. None if it is not one."""
        text = re.sub(r"\s+", " ", raw).strip()
        match = PAGES_ON.search(text)
        if match or PAGES_OFF.search(text):
            design, count = self._pick_layout(context)
            word = next(g for g in match.groups() if g) if match else ""
            design["pages"] = next((k for k in SEGMENT_KINDS if re.fullmatch(KIND_WORDS[k], word.lower())), "")
            return self._publish(design, "The report now has one sheet per %s as well as the summary."
                                 % design["pages"] if design["pages"] else
                                 "The report is one sheet again. Sheets made earlier for each value are still in the "
                                 "workbook; delete the ones you do not need.", count)
        if EDIT_BUILD.match(text):
            design, count = self._pick_layout(context)
            return self._publish(design, "", count)
        drop, off = DROPDOWN_ADD.search(text), DROPDOWN_OFF.search(text)
        if (drop or off) and self._on_layout_report(context):
            design, count = self._pick_layout(context)
            if off:
                design["dropdowns"] = []
                return self._publish(design, "The dropdown filters are removed from the report.", count)
            names = [n.strip() for n in FILTER_STOP.sub("", " " + drop.group("f")).split(",") if n.strip()]
            fields = reports.dropdown_fields(self.ledger, list(design.get("dropdowns") or []) + names)
            design["dropdowns"] = [f["label"] for f in fields]
            if not fields:
                return {"ok": True, "actions": [], "reply": "Subsidiary, department, class and location already have "
                        "their own cells with a dropdown (B4:B7). A dropdown row is for other fields, such as a custom "
                        "segment, Name or Item."}
            return self._publish(design, "The report now has a dropdown filter for %s (cells E4:F%d). Pick a value "
                                 "there and use Refresh all." % (", ".join(design["dropdowns"]), 3 + len(fields)), count)
        rule = RULE_ADD.match(text)
        if rule or RULES_OFF.search(text):
            design, count = self._pick_layout(context)
            if rule:
                found = layouts.parse_rule(rule.group(1))
                design["rules"] = [r for r in design.get("rules") or [] if r.lower() != found["text"].lower()]
                design["rules"].append(found["text"])
                return self._publish(design, 'Highlight rule added: %s. The layout has %d rule(s); "no highlight '
                                     'rules" removes them.' % (found["text"], len(design["rules"])), count)
            design["rules"] = []
            return self._publish(design, "The highlight rules are removed.", count)
        match = HIGHLIGHT_ON.search(text)
        if match or HIGHLIGHT_OFF.search(text):
            design, count = self._pick_layout(context)
            design["highlight"] = layouts.highlight_value(match.group(1)) if match else 0.0
            return self._publish(design, "Change %% cells at or above %g%% are now green, at or below -%g%% red."
                                 % (design["highlight"], design["highlight"]) if match else
                                 "The variance highlights are off.", count)
        if COMMENTS_ON.search(text) or COMMENTS_OFF.search(text):
            design, count = self._pick_layout(context)
            design["comments"] = bool(COMMENTS_ON.search(text))
            return self._publish(design, "The report has a Comments column at the right; what you type there stays "
                                 "when the report is rebuilt." if design["comments"] else
                                 "The Comments column is removed (its notes are gone with it).", count)
        if ACCOUNTS_ON.match(text) or ACCOUNTS_OFF.match(text):
            design, count = self._pick_layout(context)
            design["detail"] = bool(ACCOUNTS_ON.match(text))
            return self._publish(design, "Each line now lists the accounts behind it that have a balance."
                                 if design["detail"] else "The accounts under each line are hidden again.", count)
        match = EDIT_COLUMN_ADD.match(text)
        if match:
            design, count = self._pick_layout(context)
            rest, fmt = match.group(1).strip(), ""
            flag = re.search(r"[\s,]+(?:as\s+|in\s+)?(percent|%)\s*$", rest, re.I)
            if flag:
                rest, fmt = rest[:flag.start()], "percent"
            if "=" in rest and "[" in rest:
                label, formula = rest.split("=", 1)
                column = layouts.new_column(label.strip(' "\':'), "", "", "", None, formula, fmt)
            else:
                named = re.search(r"\s+(?:called|named|labell?ed|as)\s+\"?([^\"]+?)\"?\s*$", rest, re.I)
                label = named.group(1).strip() if named else ""
                rest = rest[:named.start()] if named else rest
                budget = re.search(r"\bbudget(?:ed|s)?\b", rest, re.I)
                if budget:
                    rest = rest[:budget.start()] + " " + rest[budget.end():]
                mods = self._modifiers(_Message(rest))
                if "range" not in mods:
                    raise ChatError("Say which period the column covers, e.g. add column Q1 2025, add column "
                                    "Jan to Jun 2026 for department Sales, or a formula: add column Change = [A]-[B].")
                periods, (first, last) = self.ledger.periods(), mods["range"]
                filters = mods.get("filters") or {}
                label = label or " ".join((["Budget"] if budget else []) + [self._range_name(first, last)]
                                          + list(filters.values()))
                column = layouts.new_column(label, "period", periods[first]["name"], periods[last]["name"], filters,
                                            data="budget" if budget else "")
            layouts.add_column(design, column)
            return self._publish(design, 'Added column "%s".' % column["label"], count)
        match = EDIT_COLUMN_DEL.match(text)
        if match:
            design, count = self._pick_layout(context)
            gone = layouts.remove(design, "column", match.group(1))
            return self._publish(design, 'Removed column "%s".' % gone["label"], count)
        match = EDIT_LINE_ADD.match(text)
        if match:
            design, count = self._pick_layout(context)
            what, rest = match.group(1).lower(), match.group(2).strip()
            section, sign, words = "", "", []
            placed = re.search(r"\s+(?:in|under|to)\s+(?:the\s+)?section\s+\"?(.+?)\"?\s*$", rest, re.I)
            if placed:
                rest, section = rest[:placed.start()], placed.group(1).strip()
            while True:
                flag = re.search(r"[\s,]+(?:as\s+|in\s+|shown\s+as\s+)?(negative|percent|bold|integer)\s*$", rest, re.I)
                if not flag:
                    break
                rest = rest[:flag.start()]
                if flag.group(1).lower() == "negative":
                    sign = "-1"
                else:
                    words.append(flag.group(1).lower())
            if what not in ("line", "row"):
                line = layouts.new_line(rest.strip(' "\''), "subtotal", "", "", " ".join(words), section)
            elif "=" in rest and not re.search(r"\baccounts?\s*[:=]", rest, re.I):
                label, formula = rest.split("=", 1)
                line = layouts.new_line(label.strip(' "\':'), "formula", formula, "", " ".join(words), section)
            else:
                parts = re.match(r"(.+?)\s+(?:with\s+|from\s+|using\s+|for\s+)?accounts?\s*[:=]?\s*(.+)$", rest, re.I)
                if not parts:
                    raise ChatError("Say what the line shows, e.g. add line Rent accounts 6100, 6200-6299 negative, "
                                    "add line Margin = [Gross Profit]/[Net Sales] percent, "
                                    "or add subtotal Total Staff Cost in section Staff Cost.")
                line = layouts.new_line(parts.group(1).strip(' "\':'), "accounts", parts.group(2), sign,
                                        " ".join(words), section)
            layouts.add_line(design, line)
            return self._publish(design, 'Added line "%s"%s.' % (
                line["label"], ' in section "%s"' % line["section"] if line["section"] else ""), count)
        match = EDIT_LINE_DEL.match(text)
        if match:
            design, count = self._pick_layout(context)
            gone = layouts.remove(design, "line", match.group(1))
            return self._publish(design, 'Removed line "%s".' % gone["label"], count)
        return None

    def design(self, payload):
        """POST /api/design, used by the Designer tab: {op, args, context} -> {ok, design, actions, reply}."""
        context = payload.get("context") if isinstance(payload.get("context"), dict) else {}
        args = payload.get("args") if isinstance(payload.get("args"), dict) else {}
        op = _text(payload.get("op"))
        arg = lambda key: _text(args.get(key))[:2000]  # noqa: E731
        try:
            if op == "templates":
                return {"ok": True, "templates": templates.TEMPLATES, "actions": [], "reply": ""}
            if op == "new":
                return self._layout_new(context, args)
            if op == "wizard":
                return self._wizard(context, args)
            if op == "dashboard":
                filters = args.get("filters") if isinstance(args.get("filters"), dict) else {}
                return self._dashboard(context, arg("kind"), arg("month"),
                                       {k: _text(v)[:200] for k, v in filters.items() if k in FILTER_KINDS}, arg("title"))
            if op == "libraryImport":
                return self._library_import(context, args)
            design, count = self._pick_layout(context, wanted=arg("layout"))
            if op == "libraryExport":   # the layout as it is, for the company library
                return {"ok": True, "design": json.loads(json.dumps(design, default=str)), "title": design["title"],
                        "name": design["name"], "actions": [], "reply": ""}
            if op == "get":
                return {"ok": True, "design": layouts.public(design), "actions": [], "reply": ""}
            if op == "build":
                return self._publish(design, "", count)
            if op == "addLine":
                layouts.add_line(design, layouts.new_line(
                    arg("label"), arg("kind"), arg("source"), "-1" if args.get("negative") else "",
                    arg("format"), arg("section")))
            elif op == "addColumn":
                filters = args.get("filters") if isinstance(args.get("filters"), dict) else {}
                layouts.add_column(design, layouts.new_column(
                    arg("label"), arg("type"), arg("from"), arg("to"), filters, arg("formula"), arg("format"),
                    data="budget" if args.get("budget") else ""))
            elif op in ("removeLine", "removeColumn"):
                layouts.remove(design, "line" if op == "removeLine" else "column", arg("label"))
            elif op in ("moveLine", "moveColumn"):
                layouts.move(design, "line" if op == "moveLine" else "column", arg("label"),
                             -1 if args.get("by") == -1 else 1)
            elif op == "setTitle":
                if not arg("title"):
                    raise ChatError("Type a report title first.")
                design["title"] = arg("title")[:60]
            elif op == "setPages":
                # Built straight away: picking "Location" is expected to make the location sheets.
                design["pages"] = layouts.page_kind(arg("pages"))
                return self._publish(design, "The report now has one sheet per %s as well as the summary."
                                     % design["pages"] if design["pages"] else
                                     "The report is one sheet again. Sheets made earlier for each value are still "
                                     "in the workbook; delete the ones you do not need.", count)
            elif op == "setDetail":
                design["detail"] = bool(args.get("detail"))
                return self._publish(design, "Each line now lists the accounts behind it that have a balance."
                                     if design["detail"] else "The accounts under each line are hidden again.", count)
            elif op == "setHighlight":
                design["highlight"] = layouts.highlight_value(arg("highlight"))
                return self._publish(design, "Variance highlights: %s." % (
                    "at or above %g%% green, at or below -%g%% red" % (design["highlight"], design["highlight"])
                    if design["highlight"] else "off"), count)
            elif op in ("addRule", "removeRule"):
                rules = list(design.get("rules") or [])
                if op == "addRule":
                    found = layouts.parse_rule(arg("rule"))
                    rules = [r for r in rules if r.lower() != found["text"].lower()] + [found["text"]]
                    lead = "Highlight rule added: %s." % found["text"]
                else:
                    rules = [r for r in rules if r.lower() != arg("rule").lower()]
                    lead = "Highlight rule removed."
                design["rules"] = rules
                return self._publish(design, lead, count)
            elif op == "setCommentAt":
                design["commentAt"] = layouts.highlight_value(arg("commentAt"))
                if design["commentAt"]:
                    design["comments"] = True   # asking for comments needs the Comments column
                return self._publish(design, "Comments are asked for when a change is %s." % (
                    "at least %g%% up or down" % design["commentAt"] if design["commentAt"] else "any size (marking off)"), count)
            elif op == "setComments":
                design["comments"] = bool(args.get("comments"))
                return self._publish(design, "Comments column: %s." % ("on" if design["comments"] else "off"), count)
            elif op == "setDropdowns":
                given = args.get("dropdowns") if isinstance(args.get("dropdowns"), list) else []
                design["dropdowns"] = [f["label"] for f in reports.dropdown_fields(self.ledger, given)]
            elif op == "setFilters":
                given = args.get("filters") if isinstance(args.get("filters"), dict) else {}
                design["filters"] = {k: _text(v)[:200] for k, v in given.items() if k in FILTER_KINDS and _text(v)}
            else:
                return {"ok": False, "error": "Unknown designer request"}
            layouts.check(self.ledger, design)  # a mistake is refused before it reaches the sheet
            return {"ok": True, "design": layouts.public(design), "actions": [layouts.emit(design, activate=False)],
                    "reply": "Layout updated. Use Update report to see it."}
        except (LedgerError, ChatError, reports.ReportError) as err:
            return {"ok": False, "error": str(err)}

    def _report(self, report, mods, last):
        spec, notes = self._spec(report, mods, last)
        action, count = reports.build(self.ledger, spec)
        lines = ["%s: %s. It is on a new sheet." % (reports.TITLES[report], reports.describe(self.ledger, spec))]
        if spec["detail"]:
            lines.append("%d account(s) %s." % (count, "listed" if spec["zero"] else "with a balance are listed"))
        lines += notes
        if spec["filters"]:
            lines.append("Filtered to " + _filter_words(spec["filters"]) + ".")
        if report == "tb":
            lines.append("The numbers are live formulas. Edit the period in B9:B10 or the filters in B3:B7, "
                         "then use Refresh all.")
        else:
            lines.append("The numbers are live formulas. Edit the periods in the header or the filters in B3:B7, "
                         'then use Refresh all. Or tell me what to change, e.g. "by quarter", '
                         '"compare with last year", "summary".')
        return {"ok": True, "reply": "\n".join(lines), "actions": [action], "spec": self._wire(spec)}

    def _targets(self, msg):
        """Accounts or totals the rest of the message names: [(label, [(sign, account)], is P&L)]."""
        accounts = self.ledger.accounts()
        by_number = {a["number"].lower(): a for a in accounts}

        def one(account):
            return ("%s (%s)" % (account["name"], account["number"]), [(1, account["number"])],
                    account["type"] in PL_SIGNS)

        hits = [by_number[w] for w in msg.words() if w in by_number]
        if hits:
            return [one(a) for a in hits]
        flat = re.sub(r"\s+", " ", msg.text)
        named = [a for a in accounts
                 if any(len(n) >= 4 and re.search(r"(?<!\w)%s(?!\w)" % re.escape(n), flat)
                        for n in {a["name"].lower(), a["name"].split(" : ")[-1].strip().lower()})]
        if named:
            longest = max(len(a["name"]) for a in named)
            return [one(a) for a in named if len(a["name"]) == longest]
        for pattern, label, parts, is_pl in METRICS:
            if re.search(r"\b(?:%s)\b" % pattern, flat):
                return [(label, parts, is_pl)]
        words = [w for w in msg.words() if w not in STOPWORDS and len(w) >= 3 and not w.isdigit()]
        if not words:
            return []

        def has(name, word):
            stem = word[:-1] if word.endswith("s") and len(word) > 4 else word
            return re.search(r"\b%s(?:s|es)?\b" % re.escape(stem), name)

        found = [a for a in accounts if all(has(a["name"].lower(), w) for w in words)]
        if len(found) > MAX_ANSWER_ACCOUNTS:
            raise ChatError('"%s" matches %d accounts. Give the account number or more of the name.'
                            % (" ".join(words), len(found)))
        return [one(a) for a in found]

    def _period_args(self, mods, is_pl, notes):
        periods, current = self.ledger.periods(), self.ledger.current_index()
        if "range" in mods:
            first, last = mods["range"]
        else:
            first, last = (self.ledger.fy_start(current), current) if is_pl else (current, current)
            if not notes:
                notes.append("No period was named, so this is %s."
                             % ("this fiscal year to date" if is_pl else "as at " + periods[current]["name"]))
        if is_pl:
            return periods[first]["name"], periods[last]["name"], reports.range_label(self.ledger, first, last)
        if "range" in mods and first <= current < last:
            last = current
        return periods[last]["name"], "", "as at " + periods[last]["name"]

    def _answer(self, targets, mods, write_here):
        filters = mods.get("filters") or {}
        requests, plan, notes, rows = [], [], [], []
        for label, parts, is_pl in targets:
            start, end, when = self._period_args(mods, is_pl, notes)
            plan.append((label, when, len(requests), len(parts)))
            for sign, account in parts:
                requests.append(dict({"account": account, "fromPeriod": start, "toPeriod": end}, **filters))
            args = '"%s","%s",%s' % (start, end, ",".join('"%s"' % filters.get(k, "").replace('"', '""')
                                                         for k in FILTER_KINDS))
            rows.append([label, when, "=" + "".join(
                '%sTQ.BALANCE("%s",%s)' % ("+" if sign > 0 else "-", account, args) for sign, account in parts
            ).lstrip("+")])
        results = self.ledger.balances(requests)
        lines = []
        for (label, when, at, size), (_, parts, _) in zip(plan, targets):
            errors = [r["error"] for r in results[at:at + size] if "error" in r]
            if errors:
                raise ChatError(errors[0])
            total = sum(sign * results[at + n]["value"] for n, (sign, _) in enumerate(parts))
            lines.append("%s, %s: %s" % (label, when, _money(total)))
        if filters:
            lines.append("Filtered to " + _filter_words(filters) + ".")
        actions = []
        if write_here:
            actions.append({"type": "write", "target": "selection", "sheetName": "", "startCell": "", "rows": rows,
                            "textRanges": [[0, 0, len(rows) - 1, 1]],
                            "formats": [{"r0": 0, "c0": 2, "r1": len(rows) - 1, "c1": 2,
                                         "numberFormat": reports.NUMBER_FORMAT}]})
        return {"ok": True, "reply": "\n".join(lines + notes), "actions": actions}

    def _transactions(self, targets, mods):
        real = [t for t in targets if len(t[1]) == 1 and not t[1][0][1].lower().startswith(("type:", "calc:"))]
        if len(real) != 1 or len(targets) != 1:
            raise ChatError("Name one account for the transactions, e.g. transactions for 6100 in Jan 2026.")
        label, parts, is_pl = real[0]
        notes = []
        start, end, when = self._period_args(mods, is_pl, notes)
        request = dict({"account": parts[0][1], "fromPeriod": start, "toPeriod": end}, **(mods.get("filters") or {}))
        data = self.ledger.call({"action": "transactions", "request": request})
        found = data.get("rows", [])
        rows = [["Transactions: %s, %s" % (label, when)], [],
                ["Date", "Number", "Type", "Name", "Memo", "Amount"]]
        base = getattr(self.ledger, "record_url", "")

        def number(r):   # a link that opens the transaction in NetSuite
            if base and r.get("id"):
                return '=HYPERLINK("%s%d","%s")' % (base, int(r["id"]), (_text(r.get("number")) or "open").replace('"', '""'))
            return "'" + _text(r.get("number")) if _text(r.get("number")) else ""
        rows += [[_text(r.get("date")), number(r), _text(r.get("type")), _text(r.get("entity")),
                  _text(r.get("memo")), r.get("amount") or 0] for r in found]
        rows.append(["", "", "", "", "Total", "=SUM(F4:F%d)" % (len(found) + 3)] if found
                    else ["", "", "", "", "Total", 0])
        rows = [r + [""] * (6 - len(r)) for r in rows]
        end_row = len(rows) - 1
        action = {
            "type": "write", "target": "new_sheet", "startCell": "", "rows": rows,
            "sheetName": reports.sheet_name("Txn %s %s" % (parts[0][1], end or start)),
            "textRanges": [[3, 0, end_row, 0], [3, 2, end_row, 4]] if found else [],
            "formats": [{"r0": 0, "c0": 0, "r1": 0, "c1": 5, "bold": True, "size": 12},
                        {"r0": 2, "c0": 0, "r1": 2, "c1": 5, "bold": True, "bottom": "thin"},
                        {"r0": 3, "c0": 5, "r1": end_row, "c1": 5, "numberFormat": reports.NUMBER_FORMAT},
                        {"r0": end_row, "c0": 0, "r1": end_row, "c1": 5, "bold": True, "top": "thin"}],
            "freeze": {"rows": 3, "cols": 0}, "autofit": True,
        }
        total = sum(float(r.get("amount") or 0) for r in found)
        lines = ["%d transaction line(s) for %s, %s, totalling %s. They are on a new sheet."
                 % (len(found), label, when, _money(total))]
        if data.get("truncated"):
            lines.append("NetSuite cut the list at 1,000 lines, so it is incomplete. Ask for a shorter period.")
        return {"ok": True, "reply": "\n".join(lines + notes), "actions": [action]}

    def _list(self, what, msg, mods, here):
        if what == "accounts":
            words = [w for w in msg.words() if w not in STOPWORDS and w not in ("chart", "coa")]
            found = [a for a in self.ledger.accounts()
                     if all(w in ("%s %s %s" % (a["number"], a["name"], a["type"])).lower() for w in words)]
            rows = [["Account", "Name", "Type"]] + [[a["number"], a["name"], a["type"]] for a in found]
            title, text_cols = "Accounts", [0]
        elif what == "periods":
            periods = self.ledger.periods()
            first, last = mods.get("range", (0, len(periods) - 1))
            rows = [["Period", "Start", "End"]] + [
                [p["name"], p["start"].isoformat(), p["end"].isoformat()] for p in periods[first:last + 1]]
            title, text_cols = "Periods", [0, 1, 2]
        else:
            rows = [[reports.KIND_LABEL[what]]] + [[name] for name in self.ledger.segments(what)]
            title, text_cols = reports.KIND_LABEL[what] + " list", [0]
        rows = rows[:MAX_LIST_ROWS]
        count = len(rows) - 1
        if not count:
            return {"ok": True, "reply": "Nothing matched, so nothing was written.", "actions": []}
        action = {
            "type": "write", "target": "selection" if here else "new_sheet", "sheetName": title, "startCell": "",
            "rows": rows, "textRanges": [[1, c, count, c] for c in text_cols],
            "formats": [{"r0": 0, "c0": 0, "r1": 0, "c1": len(rows[0]) - 1, "bold": True}], "autofit": True,
        }
        return {"ok": True, "actions": [action],
                "reply": "%d %s written %s." % (count, "row" if count == 1 else "rows",
                                                "at the selected cell" if here else "to a new sheet")}

    # -------------------------------------------------------------- entry --

    def handle(self, payload):
        message = _text(payload.get("message"))
        context = payload.get("context") if isinstance(payload.get("context"), dict) else {}
        try:
            image = assistant_module.clean_image(payload.get("image"))
        except AssistantError as err:
            return {"ok": True, "reply": str(err), "actions": []}
        if not message and not image:
            return {"ok": False, "error": "Type a message first"}
        if len(message) > MAX_MESSAGE_CHARS:
            return {"ok": False, "error": "Message is too long"}
        forced = AI_PREFIX.match(message)
        if forced:
            message = message[forced.end():].strip()
        try:
            if image or forced:
                return self._ai(message, image, context)
            out = self._run(message, context)
            if out.pop("missed", False) and self.assistant:
                return self._ai(message, None, context)
            return out
        except (LedgerError, ChatError, reports.ReportError) as err:
            return {"ok": True, "reply": str(err), "actions": []}

    # ------------------------------------------- reading a message more kindly --

    def _run(self, message, context):
        parts = [p.strip(" .") for p in SPLIT.split(message) if p and p.strip(" .")][:MAX_STEPS]
        return self._run_many(parts, context) if len(parts) > 1 else self._one(message, context)

    def _one(self, text, context):
        """One request: politeness dropped, spelling slips in command words put right, and a hint if it is not understood."""
        for _ in range(2):
            text = POLITE.sub("", text, count=1).strip() or text
        fixed = self._spelling(text)
        if fixed != text:
            out = self._respond(fixed, context)  # an error here is about the corrected reading: let it show
            if not out.get("missed"):
                out["reply"] = 'I read that as "%s".\n%s' % (fixed, out["reply"])
                return out
        out = self._respond(text, context)
        if out.get("missed"):
            out["reply"] = self._suggest(fixed)
        return out

    def _known_words(self):
        """Words that are names in this NetSuite account, so they are never "corrected"."""
        words = set()
        try:
            names = [a["name"] for a in self.ledger.accounts()]
            for kind in SEGMENT_KINDS:
                names += layouts._segments(self.ledger, kind)
        except (LedgerError, reports.ReportError):
            names = []
        for name in names:
            words.update(w.lower() for w in re.findall(r"[A-Za-z]+", name))
        return words

    def _spelling(self, text):
        known = self._known_words()
        limit = 3 if re.match(r"^\W*(?:add|remove|delete|drop|load|run|open|get|import)\b", text, re.I) else 0
        seen = [0]  # after "add line ..." comes a name the user chose: only the command words are looked at

        def fix(match):
            seen[0] += 1
            word = match.group(0)
            low = word.lower()
            if (limit and seen[0] > limit) or len(word) < 5 or low in COMMAND_WORDS or low in STOPWORDS or low in known:
                return word
            hit = difflib.get_close_matches(low, COMMAND_WORD_LIST, 1, 0.8)
            return hit[0] if hit and hit[0][0] == low[0] else word

        parts = re.split(r'("[^"]*"|\[[^\]]*\])', text)  # quoted names and [references] are left alone
        return "".join(part if n % 2 else re.sub(r"[A-Za-z]+", fix, part) for n, part in enumerate(parts))

    @staticmethod
    def _suggest(text):
        words = set(re.findall(r"[a-z&]{3,}", text.lower()))
        scored = []
        for example in EXAMPLES:
            theirs = set(re.findall(r"[a-z&]{3,}", example.lower()))
            overlap = len(words & theirs) / float(len(words | theirs) or 1)
            score = difflib.SequenceMatcher(None, text.lower(), example.lower()).ratio() + overlap
            scored.append((score, example))
        best = [example for score, example in sorted(scored, reverse=True)[:3] if score >= 0.55]
        if best:
            return ("I did not understand that. Did you mean one of these?\n%s\nType help for the full list."
                    % "\n".join("- " + example for example in best))
        return ("I did not understand that. I read report requests such as:\n"
                "- P&L for 2026 by month\n- balance sheet as at Oct 2026 vs Dec 2025\n"
                "- what was rent in Mar 2026?\nType help for the full list.")

    def _run_many(self, commands, context, replies=None, actions=None):
        """Several requests in one message, each working on what the one before left in the workbook."""
        context = dict(context)
        replies, actions = list(replies or []), list(actions or [])
        result = {"ok": True}
        for step, command in enumerate(commands):
            try:
                out = self._one(command, context)
            except (LedgerError, ChatError, reports.ReportError) as err:
                out = {"reply": str(err), "actions": [], "missed": True}
            stopped = out.pop("missed", False) or not out.get("ok", True)
            text = out.get("reply") or out.get("error") or ""
            last = step == len(commands) - 1 or stopped
            brief = text.split("\n")
            brief = " ".join(brief[:2]) if brief[0].startswith("I read that as") else brief[0]
            replies.append("%s: %s" % (command, text if last else brief))
            if stopped:
                if step < len(commands) - 1:
                    replies.append("I stopped there; the steps after it were not done.")
                break
            for action in out.get("actions", []):
                same = [n for n, a in enumerate(actions) if action.get("target") == "sheet" == a.get("target")
                        and a.get("sheetName") == action.get("sheetName")]
                if same:
                    actions[same[0]] = action     # a later step rewrote the same sheet
                else:
                    actions.append(action)
                self._after(context, action)
            for key in ("spec", "design"):
                if out.get(key):
                    result[key] = out[key]
            if out.get("spec"):
                context["last"] = out["spec"]
        result.update(reply="\n".join(replies), actions=actions)
        return result

    @staticmethod
    def _after(context, action):
        """Keeps the picture of the workbook in step with an action that has not been written yet."""
        name = _text(action.get("sheetName"))
        if not name or action.get("target") != "sheet":
            return
        context["sheets"] = list(context.get("sheets") or []) + [name]
        if name.lower().startswith(layouts.SHEET_PREFIX.lower()):
            others = [s for s in context.get("layouts") or [] if isinstance(s, dict) and s.get("name") != name]
            context["layouts"] = [{"name": name, "row": 0, "rows": action.get("rows", [])}] + others
            context["activeSheet"] = name

    # ------------------------------------------------------------ AI add-on --

    def _ai_details(self, context):
        """What the AI add-on is told about this NetSuite account and workbook. Names only, no figures."""
        ledger, lines = self.ledger, []
        periods = ledger.periods()
        if periods:
            lines.append("Periods run from %s to %s. The current period is %s." % (
                periods[0]["name"], periods[-1]["name"], periods[ledger.current_index()]["name"]))
        for kind in SEGMENT_KINDS:
            names = layouts._segments(ledger, kind)
            if names:
                lines.append("%s: %s" % (reports.KIND_LABEL[kind], "; ".join(names[:150])))
        for sheet in self._layout_sheets(context):
            try:
                design = layouts.parse(sheet["name"], sheet["rows"], sheet["row"])
            except reports.ReportError:
                continue
            lines.append('Layout sheet "%s" (report "%s"). Lines: %s. Columns: %s.' % (
                design["name"], design["title"],
                "; ".join("%s [%s]" % (l["label"], l["source"] or l["kind"]) for l in design["lines"][:120]),
                "; ".join(c["label"] for c in design["columns"][:60])))
        if not self._layout_sheets(context):
            lines.append("The workbook has no layout sheet yet.")
        if context.get("sheets"):
            lines.append("Sheets in the workbook: %s" % "; ".join(_text(s) for s in context["sheets"][:60]))
        lines.append("Chart of accounts (number | name | type):")
        lines += ["%s | %s | %s" % (a["number"], a["name"], a["type"]) for a in ledger.accounts()[:3000]]
        return "\n".join(lines)

    def _free_title(self, title, context):
        """A report title whose sheet does not exist yet, so nothing of the user's is written over."""
        taken = {_text(s).lower() for s in context.get("sheets") or []}
        out, n = title, 2
        while reports.sheet_name(out).lower() in taken:
            out, n = "%s %d" % (title, n), n + 1
        return out

    def _ai(self, message, image, context):
        if not self.assistant:
            return {"ok": True, "actions": [], "reply":
                    ("Reading a picture needs the AI add-on, which is switched off. " if image else
                     "The AI add-on is switched off. ") +
                    "It is optional and uses the customer's own AI key (see \"AI add-on\" in the README). "
                    "Without it, type the format on a sheet and use Designer > Start from > My own format, "
                    "or type help for the commands I understand."}
        try:
            plan = self.assistant.plan(message, image, self._ai_details(context))
        except AssistantError as err:
            return {"ok": True, "reply": str(err), "actions": [], "ai": True}
        replies, actions, context = [plan["reply"]] if plan["reply"] else [], [], dict(context)
        result = {}
        if plan["format"]:
            taken = {s["name"].lower() for s in self._layout_sheets(context)}
            title = self._free_title(plan["format"]["title"], context)
            name = templates.layout_name(plan["format"]["title"], taken)
            try:
                design, notes = templates.from_sheet(
                    self.ledger, name, title, [[row["accounts"], row["label"]] for row in plan["format"]["rows"]],
                    columns=(1, 0))
                built = self._publish(design, 'Layout sheet "%s" was made from that format.' % name, 1, notes)
            except (LedgerError, ChatError, reports.ReportError) as err:
                return {"ok": True, "ai": True, "actions": [],
                        "reply": "\n".join(replies + ["The format could not be built: %s" % err])}
            replies.append(built["reply"])
            for action in built["actions"]:
                actions.append(action)
                self._after(context, action)
            result = {"spec": built.get("spec"), "design": built.get("design")}
            context["last"] = built.get("spec")
        if plan["commands"]:
            ran = self._run_many(plan["commands"], context, replies, actions)
            result.update({k: v for k, v in ran.items() if k in ("spec", "design")})
            replies, actions = [ran["reply"]], ran["actions"]
        return dict({k: v for k, v in result.items() if v}, ok=True, ai=True,
                    reply="\n".join(replies) or "The AI add-on gave no answer.", actions=actions)

    def _respond(self, message, context):
        msg = _Message(message)
        if HELP.match(msg.text.strip()):
            return {"ok": True, "reply": HELP_TEXT, "actions": []}
        answered = self._data(message, context) or self._edit(message, context)
        if answered:
            if not answered.get("ok"):  # the Data tab's own error form
                return {"ok": True, "reply": answered.get("error", "That did not work."), "actions": []}
            return answered
        if msg.take(LAYOUT_NEW):
            return self._layout_sample(context)
        match = DASHBOARD.search(msg.text)
        if match:
            msg.blank(match.start(), match.end())
            mods = self._modifiers(msg)
            month = self.ledger.periods()[mods["range"][1]]["name"] if mods.get("range") else ""
            word = (match.group(1) or "executive").lower()
            kind = ("pl" if re.match(r"finance|financial|profit|p\s*&|pnl", word) else
                    "position" if re.match(r"position|balance|ratio", word) else
                    "budget" if word.startswith("budget") else word)
            return self._dashboard(context, kind, month, mods.get("filters") or {})
        if msg.take(MONTH_YTD):   # the monthly pack: month and year to date, rolling forward from B8
            msg.take(INTENTS[1][1])
            mods = self._modifiers(msg)
            month = self.ledger.periods()[mods["range"][1]]["name"] if mods.get("range") else ""
            return self._layout_new(context, {"template": "monthytd", "month": month,
                                              "filters": mods.get("filters") or {}})
        budget = bool(msg.take(BUDGET_COMPARE))
        intent = next((name for name, pattern in INTENTS if msg.take(pattern)), None)
        last = self._last(context)
        if intent == "layout":
            msg.take(INTENTS[1][1])  # "custom P&L": the P&L part says nothing more
            lead = ""
            if self._layout_sheets(context):
                design, count = self._pick_layout(context, msg)
            else:  # a workbook with no layout yet: start one from the accounts that have a balance
                design, count = self._sample_design(context), 1
                lead = ('This workbook had no layout, so I added layout sheet "%s", made from the profit and loss '
                        "accounts that have a balance. Change it to your own format at any time." % design["name"])
            return self._layout_report(dict(self._modifiers(msg), budget=budget), None, design, count, lead)
        here = bool(msg.take(HERE))
        to_sheet = bool(msg.take(TO_SHEET))
        if intent is None and LIST_VERB.search(msg.text):
            for what, pattern in LISTS:
                if pattern.search(msg.text):
                    msg.take(LIST_VERB)
                    msg.take(pattern)
                    return self._list(what, msg, self._modifiers(msg) if what == "periods" else {}, here)
        mods = self._modifiers(msg)
        if budget and last and last["report"] == "layout":
            design, count = self._pick_layout(context, msg, last.get("layoutName"))
            return self._layout_report(dict(mods, budget=True), last, design, count)
        if budget:
            return {"ok": True, "actions": [],
                    "reply": "A budget comparison is built from a layout. Say: custom P&L for 2026 by quarter vs budget"}
        if intent in ("pl", "bs", "tb"):
            return self._report(intent, mods, None)
        targets = self._targets(msg)
        if intent == "txn":
            return self._transactions(targets, mods)
        if targets:
            return self._answer(targets, mods, here or to_sheet)
        if mods and last and last["report"] == "layout":
            design, count = self._pick_layout(context, msg, last.get("layoutName"))
            return self._layout_report(mods, last, design, count)
        if mods and last:
            return self._report(last["report"], mods, last)
        if mods:
            return {"ok": True, "actions": [],
                    "reply": 'Which report should that be? For example "P&L" or "balance sheet" with the period.'}
        return {"ok": True, "actions": [], "missed": True, "reply": "I did not understand that."}
