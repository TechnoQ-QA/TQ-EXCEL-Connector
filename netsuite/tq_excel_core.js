/**
 * TQ Excel Connector - query logic (a library, not an entry point)
 *
 * Answers requests from the Excel add-in. Called by the elevated Suitelet
 * (tq_excel_elevated.js) so balances are complete; who may ask is decided
 * by assertReportAccess below.
 *
 * POST body: { action: 'ping' | 'balances' | 'transactions' | 'accounts' |
 *                      'periods' | 'segments', ... }
 *
 * Sign convention: balances are returned in their natural sign, so income,
 * liabilities and equity come back positive (as on NetSuite's own reports).
 *
 * A balance request may carry budget: '*' (any budget category) or a budget
 * category name; it is then answered from NetSuite's budgets, not postings.
 *
 * An account with no account number is known as "ID" + its internal id
 * (ID123), so every account can be asked for and the balance sheet balances.
 *
 * Filters (version 2). Each segment argument takes one value or several:
 *   "Sales"  "Sales, Bar"  "not Sales"  "Food*" (Food and everything under it)
 * and a request may carry filters: "Brand: Premium, Basic; Name: Al Noor"
 * for custom segments (by label or cseg id), Name (customer, vendor...) and Item.
 *
 * Consolidation. In a OneWorld account whose subsidiaries use more than one
 * currency, amounts of several subsidiaries are translated with NetSuite's
 * consolidated exchange rates (BUILTIN.CONSOLIDATE): to the named parent for
 * "Parent*" or "Parent (consolidated)", else to the top subsidiary. Income and
 * expenses use each posting period's rate, balance sheet accounts the rate of
 * the report period, as on NetSuite's own consolidated reports.
 *
 * @NApiVersion 2.1
 */
define(['N/query', 'N/runtime', 'N/search', 'N/record', 'N/file', 'N/config'], (query, runtime, search, record, file, config) => {
  const PL_TYPES = new Set(['Income', 'COGS', 'Expense', 'OthIncome', 'OthExpense']);
  const CREDIT_TYPES = new Set([
    'AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue',
    'Equity', 'Income', 'OthIncome'
  ]);
  const STAT_TYPES = new Set(['Stat']);
  const SEGMENTS = {
    subsidiary: { table: 'subsidiary', feature: 'SUBSIDIARIES', column: 'tl.subsidiary' },
    department: { table: 'department', feature: 'DEPARTMENTS', column: 'tl.department' },
    class: { table: 'classification', feature: 'CLASSES', column: 'tl.class' },
    location: { table: 'location', feature: 'LOCATIONS', column: 'tl.location' }
  };
  const VERSION = 8;   // 3 budgets; 4 statistical accounts, Book/Currency, journals, seats; 5 branding; 6 seats per user; 7 report library; 8 licence in NetSuite
  const MAX_BATCH = 2000;
  const EXTRA_FIELDS = {
    name: { column: 'COALESCE(tl.entity, t.entity)', label: 'Name' },
    item: { column: 'tl.item', label: 'Item' }
  };
  const FIELD_WORDS = { customer: 'name', vendor: 'name', entity: 'name', employee: 'name', items: 'item' };
  const MAX_DRILL_ROWS = 1000;
  const ADMINISTRATOR_ROLE = 3;
  const REPORT_PERMISSIONS = ['REPO_FINANCIALS', 'REPO_TRIALBALANCE', 'REPO_GL'];

  const clean = (v) => (v === null || v === undefined ? '' : String(v).trim());
  const round2 = (n) => Math.round(Number(n || 0) * 100) / 100;

  const feature = (id) => {
    try {
      return runtime.isFeatureInEffect({ feature: id });
    } catch (e) {
      return false;
    }
  };

  const rows = (sql, params) =>
    query.runSuiteQL({ query: sql, params: params || [] }).asMappedResults();

  const allRows = (sql, params) => {
    const out = [];
    const paged = query.runSuiteQLPaged({ query: sql, params: params || [], pageSize: 1000 });
    paged.pageRanges.forEach((range) => {
      paged.fetch({ index: range.index }).data.asMappedResults().forEach((r) => out.push(r));
    });
    return out;
  };

  /** Per-request lookups, loaded only when first needed. */
  const createContext = () => {
    const cache = {};
    const lazy = (key, load) => () => (key in cache ? cache[key] : (cache[key] = load()));

    const periodList = lazy('periods', () =>
      rows(
        "SELECT id, periodname, TO_CHAR(startdate, 'YYYY-MM-DD') AS startdate, " +
          "TO_CHAR(enddate, 'YYYY-MM-DD') AS enddate FROM accountingperiod " +
          "WHERE isquarter = 'F' AND isyear = 'F' ORDER BY startdate, id"
      )
    );

    const periodMap = lazy('periodMap', () => {
      const map = new Map();
      periodList().forEach((p) => map.set(clean(p.periodname).toLowerCase(), p));
      return map;
    });

    const accountList = lazy('accounts', () =>
      allRows('SELECT id, acctnumber, fullname, accttype FROM account ORDER BY id').map((a) =>
        Object.assign(a, { acctnumber: clean(a.acctnumber) || 'ID' + a.id })
      )
    );

    const accountMap = lazy('accountMap', () => {
      const map = new Map();
      accountList().forEach((a) => map.set(clean(a.acctnumber).toLowerCase(), a));
      return map;
    });

    const segmentList = (kind) =>
      lazy('seg:' + kind, () => {
        const table = SEGMENTS[kind].table;
        const tries = [
          'id, name, fullname, parent' + (kind === 'subsidiary' ? ', currency' : ''),
          'id, name, fullname, parent',
          'id, name, fullname',
          'id, name'
        ];
        for (let n = 0; n < tries.length; n += 1) {
          try {
            return allRows('SELECT ' + tries[n] + ' FROM ' + table + ' ORDER BY id');
          } catch (e) {
            if (n === tries.length - 1) throw e;
          }
        }
        return [];
      })();

    /** Custom segments: [{ column: 'cseg8', label: 'Brand' }]. */
    const customSegments = lazy('cseg', () => {
      try {
        return rows("SELECT scriptid, name FROM customrecordtype WHERE scriptid LIKE 'CUSTOMRECORD_CSEG%'")
          .map((r) => ({ column: clean(r.scriptid).toLowerCase().replace(/^customrecord_/, ''), label: clean(r.name) }))
          .filter((s) => /^cseg\w+$/.test(s.column));
      } catch (e) {
        return [];
      }
    });

    const primaryBook = lazy('book', () => {
      if (!feature('MULTIBOOK')) return null;
      const r = rows("SELECT id FROM accountingbook WHERE isprimary = 'T'");
      return r.length ? Number(r[0].id) : null;
    });

    /**
     * Every value set up for a custom segment (active ones), read from the segment's own record type, so a
     * dropdown also offers values nothing has been posted to yet. Falls back to the values on postings.
     */
    const customValues = (column) =>
      lazy('cv:' + column, () => {
        const types = ['customrecord_' + column, column];
        for (let n = 0; n < types.length; n += 1) {
          try {
            const out = [];
            const paged = search.create({
              type: types[n],
              filters: [['isinactive', 'is', 'F']],
              columns: [search.createColumn({ name: 'name', sort: search.Sort.ASC })]
            }).runPaged({ pageSize: 1000 });
            paged.pageRanges.slice(0, 5).forEach((range) => {
              paged.fetch({ index: range.index }).data.forEach((r) => {
                out.push({ id: Number(r.id), name: clean(r.getValue('name')) });
              });
            });
            return out;
          } catch (e) {
            // not this record type name: try the next way
          }
        }
        return rows('SELECT DISTINCT ' + 'tl.' + column + ' AS id, BUILTIN.DF(tl.' + column +
          ') AS name FROM transactionline tl WHERE tl.' + column + ' IS NOT NULL')
          .map((r) => ({ id: Number(r.id), name: clean(r.name) }))
          .sort((a, b) => a.name.localeCompare(b.name));
      })();

    /** The same filter text is looked up once per request, not once per account. */
    const memo = (key, load) => {
      if (!(key in cache)) {
        try {
          cache[key] = { value: load() };
        } catch (e) {
          cache[key] = { error: e };  // a filter that is not found is reported for every cell, but looked up once
        }
      }
      if (cache[key].error) throw cache[key].error;
      return cache[key].value;
    };

    return { periodList, periodMap, accountList, accountMap, segmentList, customSegments, customValues, primaryBook, memo };
  };

  const resolvePeriod = (ctx, name, label) => {
    const p = ctx.periodMap().get(clean(name).toLowerCase());
    if (!p) throw new Error(label + ' "' + clean(name) + '" not found');
    return p;
  };

  /**
   * "not A, B*" -> { exclude: true, items: [{ text: 'A' }, { text: 'B', under: true }] }.
   * isName(text) says whether the whole text is one name, so a name with a comma still works.
   */
  const UNDER_RE = /\s*(\*|\(consolidated\)|\band below)$/i;
  const parseList = (value, isName) => {
    let t = clean(value);
    if (isName(t)) return { exclude: false, items: [{ text: t, under: false }] };  // a name such as "Not Applicable"
    let exclude = false;
    const not = /^(not|except|excluding)\s+/i.exec(t);
    if (not) {
      exclude = true;
      t = t.slice(not[0].length).trim();
    }
    const one = (s) => ({ text: clean(s).replace(UNDER_RE, '').trim(), under: UNDER_RE.test(clean(s)) });
    const whole = one(t);
    const items = isName(whole.text) ? [whole] : t.split(',').map(one).filter((i) => i.text);
    return { exclude, items };
  };

  const listHit = (list, text) => {
    const wanted = clean(text).toLowerCase();
    if (/^\d+$/.test(wanted)) return list.find((s) => String(s.id) === wanted) || { id: Number(wanted) };
    return (
      list.find((s) => clean(s.fullname).toLowerCase() === wanted || clean(s.name).toLowerCase() === wanted) ||
      list.find((s) => clean(s.fullname).toLowerCase().split(' : ').pop() === wanted)
    );
  };

  const hasParent = (s) => s.parent !== null && s.parent !== undefined && s.parent !== '';

  const descendants = (list, id) => {
    const out = [Number(id)];
    for (let n = 0; n < out.length; n += 1) {
      list.forEach((s) => {
        if (hasParent(s) && Number(s.parent) === out[n] && out.indexOf(Number(s.id)) === -1) out.push(Number(s.id));
      });
    }
    return out;
  };

  /** One segment argument -> { column, field, ids, exclude, items } or null. */
  const resolveSegment = (ctx, kind, value) => {
    const text = clean(value);
    if (!text) return null;
    if (!feature(SEGMENTS[kind].feature)) {
      throw new Error(kind + ' is not enabled in this NetSuite account');
    }
    const list = ctx.segmentList(kind);
    const parsed = parseList(text, (t) => !!listHit(list, t));
    const ids = [];
    parsed.items.forEach((item) => {
      const hit = listHit(list, item.text);
      if (!hit) throw new Error(kind + ' "' + item.text + '" not found');
      (item.under ? descendants(list, hit.id) : [Number(hit.id)]).forEach((id) => {
        if (ids.indexOf(id) === -1) ids.push(id);
      });
    });
    if (!ids.length) return null;
    return { column: SEGMENTS[kind].column, field: kind, ids, exclude: parsed.exclude, items: parsed.items };
  };

  /** Values of a custom segment, Name or Item -> internal ids. known: the segment's own values, when read. */
  const resolveValues = (column, kind, values, known) => {
    const parsed = parseList(values, () => false);
    const ids = [];
    parsed.items.forEach((item) => {
      let found;
      const wanted = item.text.toLowerCase();
      if (/^\d+$/.test(item.text)) {
        found = [{ id: item.text }];
      } else if (kind === 'name') {
        found = rows('SELECT id FROM entity WHERE LOWER(entitytitle) = ? OR LOWER(entityid) = ? OR LOWER(altname) = ?',
          [wanted, wanted, wanted]);
      } else if (kind === 'item') {
        found = rows('SELECT id FROM item WHERE LOWER(itemid) = ? OR LOWER(displayname) = ?', [wanted, wanted]);
      } else {
        found = (known || []).filter((v) => v.name.toLowerCase() === wanted);
        if (!found.length) found = (known || []).filter((v) => v.name.toLowerCase().split(' : ').pop() === wanted);
        if (!found.length) {
          found = rows('SELECT DISTINCT ' + column + ' AS id FROM transactionline tl WHERE LOWER(BUILTIN.DF(' +
            column + ')) = ?', [wanted]);
        }
      }
      if (!found.length) throw new Error('"' + item.text + '" not found');
      found.forEach((r) => {
        if (ids.indexOf(Number(r.id)) === -1) ids.push(Number(r.id));
      });
    });
    return { ids, exclude: parsed.exclude };
  };

  /** "Brand: Premium, Basic; Name: Al Noor; Department: Sales" -> conditions. */
  const resolveFilters = (ctx, value) => {
    const out = [];
    clean(value).split(';').map(clean).filter(Boolean).forEach((part) => {
      const at = part.indexOf(':');
      if (at < 1) throw new Error('Filter "' + part + '" needs a field and values, e.g. Brand: Premium');
      const word = clean(part.slice(0, at)).toLowerCase();
      const values = part.slice(at + 1);
      const key = FIELD_WORDS[word] || word;
      if (SEGMENTS[key]) {
        const seg = resolveSegment(ctx, key, values);
        if (seg) out.push(seg);
        return;
      }
      if (EXTRA_FIELDS[key]) {
        const r = resolveValues(EXTRA_FIELDS[key].column, key, values);
        if (r.ids.length) out.push({ column: EXTRA_FIELDS[key].column, field: key, ids: r.ids, exclude: r.exclude });
        return;
      }
      const seg = ctx.customSegments().find((s) => s.column === word || s.label.toLowerCase() === word);
      if (!seg) throw new Error('Filter field "' + clean(part.slice(0, at)) + '" not found. Use a segment, Name or Item.');
      const r = resolveValues('tl.' + seg.column, 'custom', values, ctx.customValues(seg.column));
      if (r.ids.length) out.push({ column: 'tl.' + seg.column, field: seg.column, ids: r.ids, exclude: r.exclude });
    });
    return out;
  };

  const oneWorld = () => feature('SUBSIDIARIES');

  /** How many currencies the subsidiaries use (all of them, or those in ids). */
  const currencies = (ctx, ids) => {
    if (!oneWorld()) return 1;
    const set = {};
    ctx.segmentList('subsidiary').forEach((s) => {
      if (!ids || ids.indexOf(Number(s.id)) !== -1) set[clean(s.currency)] = true;
    });
    return Object.keys(set).length;
  };

  /**
   * The subsidiary to translate to, or null when the amounts are all in one currency.
   * "Parent*" -> that parent; several subsidiaries or none -> the top subsidiary.
   */
  const consolidationTarget = (ctx, sub) => {
    if (!oneWorld()) return null;
    const list = ctx.segmentList('subsidiary');
    let chosen = null;
    if (sub && !sub.exclude) chosen = sub.ids;
    else if (sub) chosen = list.map((s) => Number(s.id)).filter((id) => sub.ids.indexOf(id) === -1);
    if (chosen && chosen.length === 1) return null;
    if (currencies(ctx, chosen) < 2) return null;
    if (sub && !sub.exclude && sub.items.length === 1 && sub.items[0].under) {
      return Number(listHit(list, sub.items[0].text).id);
    }
    const top = list.filter((s) => !hasParent(s));
    if (!top.length) throw new Error('Cannot find the top subsidiary to consolidate to');
    return Number(top[0].id);
  };

  /** "Book: Secondary; Currency: USD; Brand: Premium" -> { special: {book, currency}, rest: 'Brand: Premium' }. */
  const splitSpecial = (value) => {
    const special = {};
    const rest = [];
    clean(value).split(';').map(clean).filter(Boolean).forEach((part) => {
      const at = part.indexOf(':');
      const word = at > 0 ? clean(part.slice(0, at)).toLowerCase() : '';
      if (word === 'book' || word === 'accounting book') special.book = clean(part.slice(at + 1));
      else if (word === 'currency' || word === 'reporting currency') special.currency = clean(part.slice(at + 1));
      else rest.push(part);
    });
    return { special, rest: rest.join('; ') };
  };

  /** An accounting book by name (Multi-Book Accounting). */
  const resolveBook = (ctx, name) =>
    ctx.memo('book:' + name.toLowerCase(), () => {
      if (!feature('MULTIBOOK')) throw new Error('Multi-Book Accounting is not enabled, so there is only one book');
      const hit = rows('SELECT id FROM accountingbook WHERE LOWER(name) = ?', [name.toLowerCase()])[0];
      if (!hit) throw new Error('Accounting book "' + name + '" not found');
      return Number(hit.id);
    });

  /**
   * Reporting currency: NetSuite translates to a subsidiary's currency, so the amounts are consolidated to the
   * highest subsidiary that uses that currency (an elimination or reporting subsidiary set up for it, typically).
   */
  const currencyTarget = (ctx, code) =>
    ctx.memo('cur:' + code.toLowerCase(), () => {
      if (!oneWorld()) throw new Error('A reporting currency needs NetSuite OneWorld');
      const cur = rows('SELECT id FROM currency WHERE LOWER(symbol) = ? OR LOWER(name) = ?',
        [code.toLowerCase(), code.toLowerCase()])[0];
      if (!cur) throw new Error('Currency "' + code + '" not found');
      const list = ctx.segmentList('subsidiary');
      const using = list.filter((s) => Number(s.currency) === Number(cur.id));
      if (!using.length) throw new Error('No subsidiary uses ' + code + ', so NetSuite has no rates to translate to it');
      const depth = (s) => {
        let n = 0;
        let at = s;
        while (at && hasParent(at) && n < 20) {
          at = list.find((p) => Number(p.id) === Number(at.parent));
          n += 1;
        }
        return n;
      };
      using.sort((a, b) => depth(a) - depth(b));
      return Number(using[0].id);
    });

  /** Turns one cell's arguments into ids and dates. Throws a readable error. */
  const resolveRequest = (ctx, r) => {
    const acctText = clean(r.account);
    if (!acctText) throw new Error('Account is required');
    const acct = ctx.accountMap().get(acctText.toLowerCase());
    if (!acct) throw new Error('Account "' + acctText + '" not found');

    if (!clean(r.fromPeriod)) throw new Error('Period is required');
    const from = resolvePeriod(ctx, r.fromPeriod, 'Period');
    const to = clean(r.toPeriod) ? resolvePeriod(ctx, r.toPeriod, 'To period') : from;
    if (to.enddate < from.startdate) throw new Error('To period is before from period');

    const stat = STAT_TYPES.has(acct.accttype);   // statistical: units (covers, headcount) per period, never translated
    const isPL = PL_TYPES.has(acct.accttype) || stat;
    const budget = clean(r.budget).toLowerCase();
    const sp = splitSpecial(r.filters);
    const seg = (kind) => ctx.memo('r:' + kind + ':' + clean(r[kind]).toLowerCase(), () => resolveSegment(ctx, kind, r[kind]));
    const others = ctx.memo('f:' + sp.rest.toLowerCase(), () => resolveFilters(ctx, sp.rest));
    const conds = [seg('subsidiary'), seg('department'), seg('class'), seg('location')].filter(Boolean).concat(others);
    // "Subsidiary: X" in the other filters counts as the subsidiary when the argument is empty
    const sub = seg('subsidiary') || others.filter((c) => c.field === 'subsidiary')[0] || null;
    if (budget && conds.some((c) => !SEGMENTS[c.field])) {
      throw new Error('A budget can be filtered by subsidiary, department, class and location only');
    }
    if (budget && sp.special.currency) throw new Error('A budget is shown in its own currency; leave Currency out');
    let target = null;
    if (!budget && !stat) {
      target = sp.special.currency ? currencyTarget(ctx, sp.special.currency)
        : ctx.memo('t:' + clean(r.subsidiary).toLowerCase() + '|' + clean(r.filters).toLowerCase(), () => consolidationTarget(ctx, sub));
    }
    return {
      accountId: Number(acct.id),
      accountType: acct.accttype,
      isPL,
      stat,
      budget,
      // Balance sheet accounts are cumulative, so the start date is ignored (a budget is always for its periods).
      fromStart: isPL || budget ? from.startdate : '',
      toEnd: to.enddate,
      toPeriodId: Number(to.id),
      conds,
      target,
      book: sp.special.book ? resolveBook(ctx, sp.special.book) : ctx.primaryBook()
    };
  };

  const condKey = (c) => (c.exclude ? '!' : '') + c.column + '=' + c.ids.join(',');

  const groupKey = (q) =>
    [q.fromStart, q.toEnd, q.book, q.budget, q.target, q.target === null ? '' : q.toPeriodId, q.stat ? 'stat' : '']
      .concat(q.conds.map(condKey)).join('|');

  /** The amount in the subsidiary's currency, or translated with the consolidated exchange rates. */
  const amountExpr = (q) =>
    q.target === null
      ? 'tal.amount'
      : "BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', " + q.target + ', ' +
        (q.isPL ? 't.postingperiod' : q.toPeriodId) + ", 'DEFAULT')";

  /** FROM / WHERE shared by the balance and drill-down queries. */
  const buildFilter = (q, accountIds) => {
    const params = [q.toEnd];
    const where = [
      'tal.account IN (' + accountIds.join(',') + ')',
      "ap.enddate <= TO_DATE(?, 'YYYY-MM-DD')"
    ];
    if (!q.stat) where.unshift("tal.posting = 'T'");   // statistical journals do not post to the general ledger
    if (q.fromStart) {
      where.push("ap.startdate >= TO_DATE(?, 'YYYY-MM-DD')");
      params.push(q.fromStart);
    }
    if (q.book !== null) where.push('tal.accountingbook = ' + q.book);

    let needLine = false;
    q.conds.forEach((c) => {
      needLine = true;
      const list = c.ids.map(Number).join(',');
      where.push(c.exclude ? '(' + c.column + ' IS NULL OR ' + c.column + ' NOT IN (' + list + '))'
        : c.column + ' IN (' + list + ')');
    });

    const from =
      'FROM transactionaccountingline tal ' +
      'INNER JOIN transaction t ON t.id = tal.transaction ' +
      'INNER JOIN accountingperiod ap ON ap.id = t.postingperiod ' +
      (needLine
        ? 'INNER JOIN transactionline tl ON tl.transaction = tal.transaction ' +
          'AND tl.id = tal.transactionline '
        : '');

    return { from, where: 'WHERE ' + where.join(' AND '), params };
  };

  /** Budget amounts per account for the request's periods and segments. Budgets are stored in their natural sign. */
  const budgetTotals = (q, accountIds) => {
    const params = [q.toEnd, q.fromStart];
    const where = [
      'b.account IN (' + accountIds.join(',') + ')',
      "ap.isquarter = 'F'",
      "ap.isyear = 'F'",
      "ap.enddate <= TO_DATE(?, 'YYYY-MM-DD')",
      "ap.startdate >= TO_DATE(?, 'YYYY-MM-DD')"
    ];
    if (q.book !== null) where.push('b.accountingbook = ' + q.book);
    q.conds.forEach((c) => {
      const column = 'b.' + c.field;
      const list = c.ids.map(Number).join(',');
      where.push(c.exclude ? '(' + column + ' IS NULL OR ' + column + ' NOT IN (' + list + '))'
        : column + ' IN (' + list + ')');
    });
    if (q.budget !== '*') {
      where.push('LOWER(BUILTIN.DF(b.category)) = ?');
      params.push(q.budget);
    }
    return rows(
      'SELECT b.account AS account, SUM(bm.amount) AS amount FROM budgets b ' +
        'INNER JOIN budgetsmachine bm ON bm.budget = b.id ' +
        'INNER JOIN accountingperiod ap ON ap.id = bm.period ' +
        'WHERE ' + where.join(' AND ') + ' GROUP BY b.account',
      params
    );
  };

  const signed = (accountType, amount) =>
    round2(CREDIT_TYPES.has(accountType) ? -Number(amount || 0) : Number(amount || 0));

  /**
   * Many cells in, few queries out: cells that share the same dates and
   * segment filters are answered by one query grouped by account.
   */
  const balances = (ctx, requests) => {
    if (!Array.isArray(requests)) throw new Error('requests must be a list');
    if (requests.length > MAX_BATCH) throw new Error('Too many requests in one batch');

    const results = new Array(requests.length);
    const groups = new Map();

    requests.forEach((r, i) => {
      try {
        const q = resolveRequest(ctx, r || {});
        const key = groupKey(q);
        if (!groups.has(key)) groups.set(key, { q, cells: [] });
        groups.get(key).cells.push({ i, q });
      } catch (e) {
        results[i] = { error: e.message };
      }
    });

    groups.forEach((group) => {
      const ids = Array.from(new Set(group.cells.map((c) => c.q.accountId)));
      const totals = new Map();
      for (let start = 0; start < ids.length; start += 500) {
        const part = ids.slice(start, start + 500);
        let found;
        if (group.q.budget) {
          found = budgetTotals(group.q, part);
        } else {
          const f = buildFilter(group.q, part);
          found = rows(
            'SELECT tal.account AS account, SUM(' + amountExpr(group.q) + ') AS amount ' +
              f.from + f.where + ' GROUP BY tal.account',
            f.params
          );
        }
        found.forEach((row) => totals.set(Number(row.account), Number(row.amount || 0)));
      }
      group.cells.forEach((c) => {
        const amount = totals.get(c.q.accountId) || 0;
        results[c.i] = { value: group.q.budget ? round2(amount) : signed(c.q.accountType, amount) };
      });
    });

    return results;
  };

  const transactions = (ctx, request) => {
    const q = resolveRequest(ctx, request || {});
    const f = buildFilter(q, [q.accountId]);
    const found = rows(
      "SELECT t.id AS id, TO_CHAR(t.trandate, 'YYYY-MM-DD') AS trandate, " +
        't.tranid AS tranid, BUILTIN.DF(t.type) AS type, ' +
        'BUILTIN.DF(t.entity) AS entity, t.memo AS memo, ' + amountExpr(q) + ' AS amount ' +
        f.from + f.where + ' ORDER BY t.trandate, t.id',
      f.params
    );
    return {
      truncated: found.length > MAX_DRILL_ROWS,
      rows: found.slice(0, MAX_DRILL_ROWS).map((r) => ({
        id: Number(r.id),
        date: r.trandate,
        number: r.tranid || '',
        type: r.type || '',
        entity: r.entity || '',
        memo: r.memo || '',
        amount: signed(q.accountType, r.amount)
      }))
    };
  };

  /**
   * Balances must not depend on which transaction types the caller's own role
   * can view, so the work runs in a Suitelet deployed as Administrator. This
   * check replaces those permissions: the caller must hold a role that is
   * allowed to run financial reports.
   */
  const assertReportAccess = (user) => {
    let found;
    try {
      found = rows(
        'SELECT COUNT(*) AS n FROM employeerolesforsearch er ' +
          'LEFT JOIN rolepermissions rp ON rp.role = er.role ' +
          "AND rp.permkey IN ('" + REPORT_PERMISSIONS.join("','") + "') " +
          'WHERE er.entity = ? AND (er.role = ' + ADMINISTRATOR_ROLE + ' OR rp.permkey IS NOT NULL)',
        [Number(user.id)]
      );
    } catch (e) {
      // Role tables are unreadable unless the Suitelet really runs as Administrator.
      throw new Error('Cannot check report access. Deploy the elevated Suitelet with Execute As Role = Administrator.');
    }
    if (!found.length || Number(found[0].n) < 1) {
      throw new Error('None of your NetSuite roles may run financial reports');
    }
  };

  // ------------------------------------------------------------------ budgets --
  // Budget upload (version 3). The script runs as Administrator, so writing is allowed only to a caller whose
  // role is Administrator or holds one of these permissions at Edit level or higher.
  const BUDGET_PERMISSIONS = ['ADMI_BUDGET', 'LIST_BUDGET', 'TRAN_BUDGET', 'ADMI_SETUPBUDGET'];
  const BUDGET_RECORD = 'budgetimport';

  const canWriteBudgets = (user) => {
    try {   // as assertReportAccess: the caller's own roles, read with the elevated script's rights
      const found = rows(
        'SELECT COUNT(*) AS n FROM employeerolesforsearch er LEFT JOIN rolepermissions rp ON rp.role = er.role ' +
          "AND rp.permkey IN ('" + BUDGET_PERMISSIONS.join("','") + "') AND rp.permlevel >= 2 " +
          'WHERE er.entity = ? AND (er.role = ' + ADMINISTRATOR_ROLE + ' OR rp.permkey IS NOT NULL)',
        [Number(user.id)]
      );
      return found.length > 0 && Number(found[0].n) > 0;
    } catch (e) {
      return false;
    }
  };

  /** Fiscal years with their months in order, the budget categories, and whether this user may upload. */
  const budgetInfo = (ctx, user) => {
    const years = rows(
      "SELECT id, periodname, TO_CHAR(startdate, 'YYYY-MM-DD') AS startdate, TO_CHAR(enddate, 'YYYY-MM-DD') AS enddate " +
        "FROM accountingperiod WHERE isyear = 'T' ORDER BY startdate"
    ).map((y) => ({
      id: Number(y.id),
      name: clean(y.periodname),
      months: rows(
        "SELECT periodname FROM accountingperiod WHERE isyear = 'F' AND isquarter = 'F' AND isadjust = 'F' " +
          "AND startdate >= TO_DATE(?, 'YYYY-MM-DD') AND enddate <= TO_DATE(?, 'YYYY-MM-DD') ORDER BY startdate",
        [y.startdate, y.enddate]
      ).map((p) => clean(p.periodname))
    }));
    let categories = [];
    try {
      categories = rows('SELECT name FROM budgetcategory ORDER BY name').map((c) => clean(c.name));
    } catch (e) {
      categories = [];   // Multiple Budgets is off: one budget per account and year
    }
    let recordOk = true;
    let recordError = '';
    try {
      record.create({ type: BUDGET_RECORD });   // not saved: only checks that this account can script budgets
    } catch (e) {
      recordOk = false;
      recordError = e.message || String(e);
    }
    return { years, categories, canWrite: canWriteBudgets(user), recordOk, recordError };
  };

  /**
   * lines: [{ account, subsidiary, department, class, location, amounts: [month 1, month 2, ...] }]
   * Each line is one account (and segment set) for the year; amounts follow the year's months, in their natural
   * sign (income and expenses both positive). An existing budget line is updated, else one is created.
   * dryRun: says what would happen and writes nothing.
   */
  const budgetSave = (ctx, user, request) => {
    const dryRun = request.dryRun !== false;
    if (!dryRun && !canWriteBudgets(user)) throw new Error('Your NetSuite role may not change budgets');
    const year = rows("SELECT id, periodname FROM accountingperiod WHERE isyear = 'T' AND LOWER(periodname) = ?",
      [clean(request.year).toLowerCase()])[0];
    if (!year) throw new Error('Fiscal year "' + clean(request.year) + '" not found');
    let category = null;
    if (clean(request.category)) {
      const hit = rows('SELECT id FROM budgetcategory WHERE LOWER(name) = ?', [clean(request.category).toLowerCase()])[0];
      if (!hit) throw new Error('Budget category "' + clean(request.category) + '" not found');
      category = Number(hit.id);
    }
    const book = ctx.primaryBook();
    const lines = Array.isArray(request.lines) ? request.lines : [];
    const results = [];
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i] || {};
      if (!dryRun && runtime.getCurrentScript().getRemainingUsage() < 120) {
        return { results, done: i };   // the server sends the rest in the next call
      }
      try {
        const acct = ctx.accountMap().get(clean(line.account).toLowerCase());
        if (!acct) throw new Error('Account "' + clean(line.account) + '" not found');
        const segs = {};
        Object.keys(SEGMENTS).forEach((kind) => {
          const seg = resolveSegment(ctx, kind, line[kind]);
          if (seg && (seg.exclude || seg.ids.length !== 1)) throw new Error(kind + ': a budget line takes one value');
          segs[kind] = seg ? seg.ids[0] : null;
        });
        const where = ['account = ?', 'year = ?'];
        const params = [Number(acct.id), Number(year.id)];
        where.push(category === null ? 'category IS NULL' : 'category = ' + category);
        Object.keys(SEGMENTS).forEach((kind) => {
          // Without OneWorld NetSuite still stores subsidiary 1 on a budget, so a blank one is not looked for.
          if (kind === 'subsidiary' && segs[kind] === null && !oneWorld()) return;
          where.push(segs[kind] === null ? kind + ' IS NULL' : kind + ' = ' + Number(segs[kind]));
        });
        if (book !== null) where.push('accountingbook = ' + book);
        const existing = rows('SELECT id FROM budgets WHERE ' + where.join(' AND '), params)[0];
        const amounts = (Array.isArray(line.amounts) ? line.amounts : []).map((a) => (a === null || a === '' ? null : Number(a)));
        if (amounts.some((a) => a !== null && !isFinite(a))) throw new Error('An amount is not a number');
        if (dryRun) {
          results.push({ ok: true, action: existing ? 'update' : 'create', id: existing ? Number(existing.id) : null });
          continue;
        }
        const rec = existing ? record.load({ type: BUDGET_RECORD, id: Number(existing.id) }) : record.create({ type: BUDGET_RECORD });
        if (!existing) {
          rec.setValue({ fieldId: 'year', value: Number(year.id) });
          if (category !== null) rec.setValue({ fieldId: 'category', value: category });
          rec.setValue({ fieldId: 'account', value: Number(acct.id) });
          Object.keys(SEGMENTS).forEach((kind) => {
            if (segs[kind] !== null) rec.setValue({ fieldId: kind, value: segs[kind] });
          });
        }
        amounts.forEach((a, n) => {
          rec.setValue({ fieldId: 'periodamount' + (n + 1), value: a === null ? 0 : round2(a) });
        });
        const id = rec.save({ ignoreMandatoryFields: false });
        results.push({ ok: true, action: existing ? 'updated' : 'created', id: Number(id) });
      } catch (e) {
        results.push({ ok: false, error: e.message || String(e) });
      }
    }
    return { results, done: lines.length };
  };

  // ----------------------------------------------------------------- branding --
  /** Company (or subsidiary) name and logo for the printed pack (version 5). The logo comes back as base64. */
  const branding = (ctx, request) => {
    let company = '';
    let logoId = null;
    try {
      const info = config.load({ type: config.Type.COMPANY_INFORMATION });
      company = clean(info.getValue({ fieldId: 'companyname' }));
      logoId = info.getValue({ fieldId: 'formlogo' }) || info.getValue({ fieldId: 'pagelogo' }) || null;
    } catch (e) {
      // no access to Company Information: the name and logo stay empty
    }
    let subsidiary = '';
    if (oneWorld() && clean(request.subsidiary)) {
      try {
        const s = resolveSegment(ctx, 'subsidiary', request.subsidiary);
        if (s && !s.exclude && s.ids.length === 1) {
          const sub = record.load({ type: record.Type.SUBSIDIARY, id: s.ids[0] });
          subsidiary = clean(sub.getValue({ fieldId: 'legalname' })) || clean(sub.getValue({ fieldId: 'name' }));
          logoId = sub.getValue({ fieldId: 'logo' }) || logoId;
        }
      } catch (e) {
        subsidiary = '';
      }
    }
    let logo = null;
    if (logoId) {
      try {
        const f = file.load({ id: Number(logoId) });
        if (Number(f.size) < 3000000) logo = { data: f.getContents(), type: clean(f.fileType), name: clean(f.name) };
      } catch (e) {
        logo = null;
      }
    }
    return { company, subsidiary, logo };
  };

  // -------------------------------------------------------------------- seats --
  // Each PC's connector registers here (version 4). The licence allows a number of users; a PC not seen for
  // SEAT_DAYS frees its seat. The list is a small file next to this script in the File Cabinet.
  const SEAT_FILE = 'tq_seats.json';
  const SEAT_DAYS = 30;

  const seatFile = () => {
    const home = rows("SELECT id, folder FROM file WHERE name = 'tq_excel_core.js'")[0];
    if (!home) throw new Error('Cannot find tq_excel_core.js in the File Cabinet');
    const found = rows('SELECT id FROM file WHERE name = ? AND folder = ?', [SEAT_FILE, Number(home.folder)])[0];
    let data = { machines: {} };
    if (found) {
      try {
        data = JSON.parse(file.load({ id: Number(found.id) }).getContents()) || data;
      } catch (e) {
        data = { machines: {} };
      }
    }
    data.machines = data.machines || {};
    return { folder: Number(home.folder), data };
  };

  const saveSeats = (folder, data) => {
    file.create({ name: SEAT_FILE, fileType: file.Type.JSON, contents: JSON.stringify(data), folder }).save();
  };

  const activeSeats = (data) => {
    const cutoff = Date.now() - SEAT_DAYS * 86400000;
    return Object.keys(data.machines).filter((id) => Number(data.machines[id].last) >= cutoff);
  };

  const seatRows = (data, keys, user) => keys.map((id) => ({
    key: id, label: data.machines[id].label, user: data.machines[id].user,
    last: new Date(Number(data.machines[id].last)).toISOString().slice(0, 10), mine: id === 'u' + user.id
  }));

  /**
   * byUser (Sign in with NetSuite): the seat belongs to the NetSuite user, whichever PC they use. Otherwise
   * (shared access token) it belongs to the PC.
   */
  const seat = (user, request) => {
    const key = request.byUser ? 'u' + Number(user.id) : clean(request.machine).slice(0, 64);
    if (!key || key === 'u0') throw new Error('No PC or user to give the seat to');
    const users = Math.max(1, Math.floor(Number(request.users)) || 1);
    const { folder, data } = seatFile();
    const now = Date.now();
    const active = activeSeats(data);
    const known = active.indexOf(key) !== -1;
    if (!known && active.length >= users) {
      return { allowed: false, used: active.length, users, machines: seatRows(data, active, user) };
    }
    const before = data.machines[key];
    data.machines[key] = { label: clean(request.label).slice(0, 60), user: clean(user.name).slice(0, 60), last: now };
    Object.keys(data.machines).forEach((id) => {   // forget seats gone for a long time
      if (Number(data.machines[id].last) < now - 4 * SEAT_DAYS * 86400000) delete data.machines[id];
    });
    if (!before || now - Number(before.last) > 12 * 3600000 || before.label !== data.machines[key].label) {
      saveSeats(folder, data);   // written at most twice a day per seat
    }
    return { allowed: true, used: known ? active.length : active.length + 1, users };
  };

  const isAdministrator = (user) => {
    try {
      const found = rows('SELECT COUNT(*) AS n FROM employeerolesforsearch WHERE entity = ? AND role = ' + ADMINISTRATOR_ROLE,
        [Number(user.id)]);
      return found.length > 0 && Number(found[0].n) > 0;
    } catch (e) {
      return false;
    }
  };

  /** Who holds a seat, and whether this caller may release seats other than their own. */
  const seatList = (user) => {
    const { data } = seatFile();
    return { seats: seatRows(data, activeSeats(data), user), canRelease: isAdministrator(user), days: SEAT_DAYS };
  };

  /** Frees a seat at once: your own, or anyone's if you are an administrator. */
  const seatRelease = (user, request) => {
    const key = clean(request.key);
    if (key !== 'u' + Number(user.id) && !isAdministrator(user)) throw new Error('Only an administrator can release another person\'s seat');
    const { folder, data } = seatFile();
    if (!data.machines[key]) throw new Error('That seat is not in use');
    delete data.machines[key];
    saveSeats(folder, data);
    return seatList(user);
  };

  // ------------------------------------------------------------ report library --
  // Report layouts shared by everyone in the company (version 7): one small file next to this script. Anyone may
  // add a layout; the person who added it, or an administrator, may replace or remove it.
  const LIBRARY_FILE = 'tq_library.json';
  const LIBRARY_MAX = 200000;   // characters per layout

  const libraryFile = () => {
    const home = rows("SELECT id, folder FROM file WHERE name = 'tq_excel_core.js'")[0];
    if (!home) throw new Error('Cannot find tq_excel_core.js in the File Cabinet');
    const found = rows('SELECT id FROM file WHERE name = ? AND folder = ?', [LIBRARY_FILE, Number(home.folder)])[0];
    let data = { items: {} };
    if (found) {
      try {
        data = JSON.parse(file.load({ id: Number(found.id) }).getContents()) || data;
      } catch (e) {
        data = { items: {} };
      }
    }
    data.items = data.items || {};
    return { folder: Number(home.folder), data };
  };

  const saveLibrary = (folder, data) => {
    file.create({ name: LIBRARY_FILE, fileType: file.Type.JSON, contents: JSON.stringify(data), folder }).save();
  };

  const libraryRows = (data, user) => Object.keys(data.items).map((id) => {
    const it = data.items[id];
    return { id, name: it.name, description: it.description || '', by: it.by, at: it.at, mine: Number(it.byId) === Number(user.id) };
  }).sort((a, b) => a.name.localeCompare(b.name));

  const libraryList = (user) => {
    const { data } = libraryFile();
    return { items: libraryRows(data, user), canManage: isAdministrator(user) };
  };

  const librarySave = (user, request) => {
    const name = clean(request.name).slice(0, 60);
    if (!name) throw new Error('Give the layout a name');
    const design = JSON.stringify(request.design || null);
    if (!request.design || typeof request.design !== 'object') throw new Error('No layout to save');
    if (design.length > LIBRARY_MAX) throw new Error('This layout is too large for the library');
    const id = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'layout';
    const { folder, data } = libraryFile();
    const before = data.items[id];
    if (before && Number(before.byId) !== Number(user.id) && !isAdministrator(user)) {
      throw new Error('"' + before.name + '" was added by ' + before.by + '. Use another name, or ask an administrator.');
    }
    data.items[id] = { name, description: clean(request.description).slice(0, 200), design: request.design,
      by: clean(user.name).slice(0, 60), byId: Number(user.id), at: new Date().toISOString().slice(0, 10) };
    saveLibrary(folder, data);
    return Object.assign({ saved: id, replaced: !!before }, libraryList(user));
  };

  const libraryLoad = (user, request) => {
    const { data } = libraryFile();
    const it = data.items[clean(request.id)];
    if (!it) throw new Error('That layout is no longer in the library');
    return { item: { id: clean(request.id), name: it.name, description: it.description || '', by: it.by, at: it.at, design: it.design } };
  };

  const libraryDelete = (user, request) => {
    const { folder, data } = libraryFile();
    const id = clean(request.id);
    const it = data.items[id];
    if (!it) throw new Error('That layout is no longer in the library');
    if (Number(it.byId) !== Number(user.id) && !isAdministrator(user)) throw new Error('Only ' + it.by + ' or an administrator can remove it');
    delete data.items[id];
    saveLibrary(folder, data);
    return libraryList(user);
  };

  // ------------------------------------------------------------------ licence --
  // The customer's licence from Techno Q, stored once by an administrator (version 8), so every user's connector
  // gets it at its first sign-in instead of a file being copied to each PC. The connector checks its signature.
  const LICENCE_FILE = 'tq_licence.json';

  const licenceGet = () => {
    const home = rows("SELECT id, folder FROM file WHERE name = 'tq_excel_core.js'")[0];
    if (!home) return { licence: '' };
    const found = rows('SELECT id FROM file WHERE name = ? AND folder = ?', [LICENCE_FILE, Number(home.folder)])[0];
    return { licence: found ? file.load({ id: Number(found.id) }).getContents() : '' };
  };

  const licenceSave = (user, request) => {
    if (!isAdministrator(user)) throw new Error('Only an administrator can store the licence in NetSuite');
    const text = clean(request.licence);
    let parsed = null;
    try { parsed = JSON.parse(text); } catch (e) { parsed = null; }
    if (!parsed || !parsed.signature || !parsed.account) throw new Error('That is not a TQ licence');
    const home = rows("SELECT id, folder FROM file WHERE name = 'tq_excel_core.js'")[0];
    if (!home) throw new Error('Cannot find tq_excel_core.js in the File Cabinet');
    file.create({ name: LICENCE_FILE, fileType: file.Type.JSON, contents: text, folder: Number(home.folder) }).save();
    return { saved: true };
  };

  // ----------------------------------------------------------------- journals --
  // Journal upload (version 4): rows of the Excel sheet with the same Entry number become one journal entry.
  const JOURNAL_PERMISSIONS = ['TRAN_JOURNAL'];

  const canWriteJournals = (user) => {
    try {
      const found = rows(
        'SELECT COUNT(*) AS n FROM employeerolesforsearch er LEFT JOIN rolepermissions rp ON rp.role = er.role ' +
          "AND rp.permkey IN ('" + JOURNAL_PERMISSIONS.join("','") + "') AND rp.permlevel >= 2 " +
          'WHERE er.entity = ? AND (er.role = ' + ADMINISTRATOR_ROLE + ' OR rp.permkey IS NOT NULL)',
        [Number(user.id)]
      );
      return found.length > 0 && Number(found[0].n) > 0;
    } catch (e) {
      return false;
    }
  };

  /**
   * entries: [{ date: 'YYYY-MM-DD', subsidiary, memo, lines: [{ account, debit, credit, memo, name, department,
   * class, location }] }]. dryRun (the default) checks everything and writes nothing.
   */
  const journalSave = (ctx, user, request) => {
    const dryRun = request.dryRun !== false;
    if (!dryRun && !canWriteJournals(user)) throw new Error('Your NetSuite role may not make journal entries');
    const entries = Array.isArray(request.entries) ? request.entries : [];
    const results = [];
    for (let i = 0; i < entries.length; i += 1) {
      const entry = entries[i] || {};
      if (!dryRun && runtime.getCurrentScript().getRemainingUsage() < 150) return { results, done: i };
      try {
        const date = clean(entry.date);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date "' + date + '" is not a date');
        const holding = (columns) => rows('SELECT id, periodname, ' + columns + " FROM accountingperiod WHERE isyear = 'F' AND " +
          "isquarter = 'F' AND isadjust = 'F' AND startdate <= TO_DATE(?, 'YYYY-MM-DD') AND enddate >= TO_DATE(?, 'YYYY-MM-DD')",
        [date, date])[0];
        let period;
        try {
          period = holding('closed, alllocked');
        } catch (e) {
          period = holding('closed');   // an account without period locking
        }
        if (!period) throw new Error('No accounting period holds ' + date);
        if (period.closed === 'T' || period.alllocked === 'T') throw new Error('Period ' + period.periodname + ' is closed or locked');
        let sub = null;
        if (oneWorld()) {
          const s = resolveSegment(ctx, 'subsidiary', entry.subsidiary);
          if (!s || s.exclude || s.ids.length !== 1) throw new Error('Give the entry one subsidiary');
          sub = s.ids[0];
        }
        const lines = (Array.isArray(entry.lines) ? entry.lines : []).map((line, n) => {
          const where = 'line ' + (n + 1) + ': ';
          const acct = ctx.accountMap().get(clean(line.account).toLowerCase());
          if (!acct) throw new Error(where + 'account "' + clean(line.account) + '" not found');
          if (STAT_TYPES.has(acct.accttype)) throw new Error(where + 'a statistical account needs a statistical journal');
          const debit = round2(line.debit);
          const credit = round2(line.credit);
          if ((debit && credit) || (!debit && !credit) || debit < 0 || credit < 0) {
            throw new Error(where + 'give either a debit or a credit, above zero');
          }
          const out = { account: Number(acct.id), debit, credit, memo: clean(line.memo) };
          ['department', 'class', 'location'].forEach((kind) => {
            const seg = resolveSegment(ctx, kind, line[kind]);
            if (seg && (seg.exclude || seg.ids.length !== 1)) throw new Error(where + kind + ' takes one value');
            out[kind] = seg ? seg.ids[0] : null;
          });
          out.entity = null;
          if (clean(line.name)) {
            const r = resolveValues(EXTRA_FIELDS.name.column, 'name', clean(line.name));
            if (r.exclude || r.ids.length !== 1) throw new Error(where + 'name "' + clean(line.name) + '" must match one record');
            out.entity = r.ids[0];
          }
          return out;
        });
        if (lines.length < 2) throw new Error('An entry needs at least two lines');
        const debits = round2(lines.reduce((s, l) => s + l.debit, 0));
        const credits = round2(lines.reduce((s, l) => s + l.credit, 0));
        if (debits !== credits) throw new Error('Debits ' + debits + ' and credits ' + credits + ' differ');
        if (dryRun) {
          results.push({ ok: true, action: 'create', period: clean(period.periodname), total: debits });
          continue;
        }
        const je = record.create({ type: record.Type.JOURNAL_ENTRY });
        if (sub !== null) je.setValue({ fieldId: 'subsidiary', value: sub });
        const parts = date.split('-').map(Number);
        je.setValue({ fieldId: 'trandate', value: new Date(parts[0], parts[1] - 1, parts[2]) });
        if (clean(entry.memo)) je.setValue({ fieldId: 'memo', value: clean(entry.memo) });
        lines.forEach((l, n) => {
          const set = (fieldId, value) => je.setSublistValue({ sublistId: 'line', fieldId, line: n, value });
          set('account', l.account);
          if (l.debit) set('debit', l.debit);
          if (l.credit) set('credit', l.credit);
          if (l.memo) set('memo', l.memo);
          if (l.entity !== null) set('entity', l.entity);
          ['department', 'class', 'location'].forEach((kind) => { if (l[kind] !== null) set(kind, l[kind]); });
        });
        const id = Number(je.save({ ignoreMandatoryFields: false }));
        const tran = rows('SELECT tranid FROM transaction WHERE id = ?', [id])[0];
        results.push({ ok: true, action: 'created', id, number: tran ? clean(tran.tranid) : '', total: debits });
      } catch (e) {
        results.push({ ok: false, error: e.message || String(e) });
      }
    }
    return { results, done: entries.length, canWrite: canWriteJournals(user) };
  };

  /** Does the work. Called by the elevated Suitelet, never directly by the RESTlet. */
  const handle = (body) => {
    try {
      const request = body || {};
      const user = runtime.getCurrentUser();
      if (request.action === 'ping') {
        return {
          ok: true,
          version: VERSION,
          account: runtime.accountId,
          user: user.name,
          role: user.roleId
        };
      }
      assertReportAccess(user);
      const ctx = createContext();
      switch (request.action) {
        case 'balances':
          return { ok: true, results: balances(ctx, request.requests) };
        case 'transactions':
          return Object.assign({ ok: true }, transactions(ctx, request.request));
        case 'accounts':
          return {
            ok: true,
            accounts: ctx.accountList().map((a) => ({
              number: a.acctnumber,
              name: a.fullname,
              type: a.accttype
            }))
          };
        case 'periods':
          return {
            ok: true,
            periods: ctx.periodList().map((p) => ({
              name: p.periodname,
              start: p.startdate,
              end: p.enddate
            }))
          };
        case 'setup': {
          // the first month of the fiscal year, from the latest fiscal year period (version 8), so nobody sets it
          let fiscalStartMonth = 0;
          try {
            const year = rows("SELECT TO_CHAR(startdate, 'MM') AS m FROM accountingperiod WHERE isyear = 'T' ORDER BY startdate DESC")[0];
            fiscalStartMonth = year ? Number(year.m) : 0;
          } catch (e) {
            fiscalStartMonth = 0;
          }
          return { ok: true, version: VERSION, oneWorld: oneWorld(), multiCurrency: currencies(ctx, null) > 1, fiscalStartMonth };
        }
        case 'budgetInfo':
          return Object.assign({ ok: true }, budgetInfo(ctx, user));
        case 'budgetSave':
          return Object.assign({ ok: true }, budgetSave(ctx, user, request));
        case 'journalSave':
          return Object.assign({ ok: true }, journalSave(ctx, user, request));
        case 'seat':
          return Object.assign({ ok: true }, seat(user, request));
        case 'seatList':
          return Object.assign({ ok: true }, seatList(user));
        case 'seatRelease':
          return Object.assign({ ok: true }, seatRelease(user, request));
        case 'branding':
          return Object.assign({ ok: true }, branding(ctx, request));
        case 'licenceGet':
          return Object.assign({ ok: true }, licenceGet());
        case 'licenceSave':
          return Object.assign({ ok: true }, licenceSave(user, request));
        case 'libraryList':
          return Object.assign({ ok: true }, libraryList(user));
        case 'librarySave':
          return Object.assign({ ok: true }, librarySave(user, request));
        case 'libraryLoad':
          return Object.assign({ ok: true }, libraryLoad(user, request));
        case 'libraryDelete':
          return Object.assign({ ok: true }, libraryDelete(user, request));
        case 'filterFields':
          return {
            ok: true,
            fields: Object.keys(SEGMENTS).filter((k) => feature(SEGMENTS[k].feature))
              .map((k) => ({ id: k, label: k.charAt(0).toUpperCase() + k.slice(1), kind: 'segment' }))
              .concat([{ id: 'name', label: 'Name', kind: 'name' }, { id: 'item', label: 'Item', kind: 'item' }])
              .concat(ctx.customSegments().map((s) => ({ id: s.column, label: s.label, kind: 'custom' })))
              .concat(feature('MULTIBOOK') ? [{ id: 'book', label: 'Book', kind: 'book' }] : [])
              .concat(currencies(ctx, null) > 1 ? [{ id: 'currency', label: 'Currency', kind: 'currency' }] : [])
          };
        case 'segments': {
          const kind = clean(request.kind).toLowerCase();
          if (kind === 'book') {
            return { ok: true, segments: feature('MULTIBOOK')
              ? rows('SELECT id, name FROM accountingbook ORDER BY name').map((b) => ({ id: Number(b.id), name: clean(b.name) })) : [] };
          }
          if (kind === 'currency') {
            return { ok: true, segments: rows('SELECT id, symbol FROM currency ORDER BY symbol')
              .map((c) => ({ id: Number(c.id), name: clean(c.symbol) })) };
          }
          const custom = ctx.customSegments().find((s) => s.column === kind || s.label.toLowerCase() === kind);
          if (custom && !SEGMENTS[kind]) {
            return { ok: true, segments: ctx.customValues(custom.column) };
          }
          if (!SEGMENTS[kind]) throw new Error('Unknown segment "' + kind + '"');
          if (!feature(SEGMENTS[kind].feature)) return { ok: true, segments: [] };
          return {
            ok: true,
            segments: ctx.segmentList(kind).map((s) => ({
              id: Number(s.id),
              name: s.fullname || s.name
            }))
          };
        }
        default:
          throw new Error('Unknown action "' + clean(request.action) + '"');
      }
    } catch (e) {
      return { ok: false, error: e.message || String(e) };
    }
  };

  return { handle };
});
