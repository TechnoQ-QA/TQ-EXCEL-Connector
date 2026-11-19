/*
 * TQ Excel Connector - custom formulas.
 *
 *   =TQ.BALANCE(account, fromPeriod, [toPeriod], [subsidiary], [department], [class], [location], [filters])
 *
 * A segment argument takes one value or several ("Sales, Bar"), "not Sales" or "Food*" (with
 * everything under it; for a subsidiary, consolidated in its currency). filters takes any other
 * filter: "Brand: Premium; Name: Al Noor".
 *   =TQ.ACCOUNTNAME(account)
 *
 * A workbook can hold thousands of these. Cells that calculate together are
 * collected for a moment and sent to the server as one request.
 */
(function () {
  'use strict';

  var API = '/api';
  var BATCH_DELAY_MS = 50;
  var BATCH_SIZE = 500;

  var cache = new Map(); // request key -> number
  var pending = new Map(); // request key -> { request, waiters: [{resolve, reject}] }
  var registry = new Map(); // cell address -> request, used by drill-down
  var timer = null;
  var accountsPromise = null;

  /* What the panel shows while numbers come in: how many were asked for, answered and failed (with their
   * cells), and which sheets just got fresh numbers (for each report's "refreshed at" line). */
  var stats = { asked: 0, done: 0, failed: [], inFlight: 0, started: 0 };
  var listeners = [];
  function changed(what) { listeners.forEach(function (fn) { try { fn(what, stats); } catch (e) { /* a closed panel */ } }); }
  function sheetOf(address) { var a = normalizeAddress(address); var i = a.lastIndexOf('!'); return i > 0 ? a.slice(0, i) : ''; }

  function text(value) {
    return value === null || value === undefined ? '' : String(value).trim();
  }

  function normalizeAddress(address) {
    return text(address).replace(/'/g, '').replace(/\$/g, '').toUpperCase();
  }

  function fail(message) {
    if (typeof CustomFunctions !== 'undefined' && CustomFunctions.Error) {
      return new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, message);
    }
    return new Error(message);
  }

  function callApi(payload) {
    return fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (response) {
      return response.json().then(
        function (data) {
          if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
          return data;
        },
        function () {
          throw new Error('Connector server returned HTTP ' + response.status);
        }
      );
    });
  }

  function flush() {
    timer = null;
    var entries = Array.from(pending.entries());
    pending = new Map();

    for (var start = 0; start < entries.length; start += BATCH_SIZE) {
      sendBatch(entries.slice(start, start + BATCH_SIZE));
    }
  }

  function sendBatch(entries) {
    var requests = entries.map(function (entry) { return entry[1].request; });
    stats.inFlight += 1;
    changed('sent');
    var sheets = {};
    var settle = function (w, ok, message) {
      stats.done += 1;
      if (ok) { if (w.sheet) sheets[w.sheet] = true; }
      else if (stats.failed.length < 500) stats.failed.push({ address: w.address || '', message: message });
    };
    callApi({ action: 'balances', requests: requests }).then(
      function (data) {
        entries.forEach(function (entry, i) {
          var result = (data.results || [])[i] || { error: 'No result returned' };
          var waiters = entry[1].waiters;
          if (typeof result.value === 'number') {
            cache.set(entry[0], result.value);
            waiters.forEach(function (w) { settle(w, true); w.resolve(result.value); });
          } else {
            waiters.forEach(function (w) { settle(w, false, result.error || 'No value'); w.reject(fail(result.error || 'No value')); });
          }
        });
      },
      function (error) {
        entries.forEach(function (entry) {
          entry[1].waiters.forEach(function (w) { settle(w, false, error.message); w.reject(fail(error.message)); });
        });
      }
    ).then(function () {
      stats.inFlight -= 1;
      if (!stats.inFlight) stats.finished = Date.now();
      changed({ sheets: Object.keys(sheets) });
    });
  }

  /* An optional argument that points at an empty cell can arrive as 0; treat it as not given. */
  function optional(value) {
    var t = text(value);
    return t === '0' ? '' : t;
  }

  function balance(account, fromPeriod, toPeriod, subsidiary, department, klass, location, filters, invocation) {
    var request = {
      account: text(account),
      fromPeriod: text(fromPeriod),
      toPeriod: optional(toPeriod),
      subsidiary: optional(subsidiary),
      department: optional(department),
      class: optional(klass),
      location: optional(location),
      filters: optional(filters)
    };
    if (!request.filters) delete request.filters;
    if (invocation && invocation.address) {
      registry.set(normalizeAddress(invocation.address), request);
    }

    var key = JSON.stringify(request);
    if (cache.has(key)) return Promise.resolve(cache.get(key));

    stats.asked += 1;
    if (!stats.started) stats.started = Date.now();
    var address = invocation && invocation.address ? invocation.address : '';
    return new Promise(function (resolve, reject) {
      if (!pending.has(key)) pending.set(key, { request: request, waiters: [] });
      pending.get(key).waiters.push({ resolve: resolve, reject: reject, address: address, sheet: address ? sheetOf(address) : '' });
      if (timer === null) timer = setTimeout(flush, BATCH_DELAY_MS);
    });
  }

  function loadAccounts() {
    if (!accountsPromise) {
      accountsPromise = callApi({ action: 'accounts' }).then(
        function (data) {
          var map = new Map();
          (data.accounts || []).forEach(function (a) {
            map.set(text(a.number).toLowerCase(), a);
          });
          return map;
        },
        function (error) {
          accountsPromise = null; // let the next call retry
          throw error;
        }
      );
    }
    return accountsPromise;
  }

  function accountName(account) {
    return loadAccounts().then(
      function (map) {
        var hit = map.get(text(account).toLowerCase());
        if (!hit) throw fail('Account "' + text(account) + '" not found');
        return hit.name;
      },
      function (error) {
        throw fail(error.message);
      }
    );
  }

  /* =TQ.PERIOD(period, [shift]): a period name worked out from another one, so a whole report follows
   * one month cell. shift: months (-1, 12) or FYS, FYE, QS, QE, LY, with +n/-n ("FYS-12"). */
  var periodCache = new Map();
  function period(base, shift) {
    var key = text(base) + '|' + optional(shift);
    if (!periodCache.has(key)) {
      periodCache.set(key, callApi({ action: 'period', base: text(base) || 'CM', shift: optional(shift) }).then(
        function (data) { return data.name; },
        function (error) { periodCache.delete(key); throw fail(error.message); }
      ));
    }
    return periodCache.get(key);
  }

  /* =TQ.SQL(query, [p1], [p2], [p3], [p4]): the first value a SuiteQL query returns, for data that does not post
   * to the ledger (open sales orders, purchase orders, quantities). Runs with the user's own NetSuite role.
   * Each ? in the query takes the next argument: =TQ.SQL("SELECT SUM(foreigntotal) FROM transaction WHERE
   * type = 'SalesOrd' AND trandate BETWEEN TO_DATE(?, 'YYYY-MM-DD') AND TO_DATE(?, 'YYYY-MM-DD')", "2026-01-01", "2026-12-31") */
  var sqlCache = new Map();
  function sql(query, p1, p2, p3, p4) {
    var params = [p1, p2, p3, p4].map(optional).filter(function (p, n, all) {
      return all.slice(n).some(function (q) { return q !== ''; });   // drop only the unused arguments at the end
    });
    var key = JSON.stringify([text(query), params]);
    if (!sqlCache.has(key)) {
      sqlCache.set(key, callApi({ action: 'sql', query: text(query), params: params, maxRows: 1 }).then(
        function (data) {
          var first = data.rows && data.rows.length ? data.rows[0][0] : null;
          if (first === null || first === undefined || first === '') return 0;
          return isFinite(Number(first)) ? Number(first) : String(first);
        },
        function (error) { sqlCache.delete(key); throw fail(error.message); }
      ));
    }
    return sqlCache.get(key);
  }

  function clearCache() {
    cache.clear();
    sqlCache.clear();
    periodCache.clear();
    accountsPromise = null;
  }

  window.TQ = {
    /* Refresh progress: stats() now, resetStats() before a refresh, onChange(fn) for every batch sent or back. */
    stats: function () { return { asked: stats.asked, done: stats.done, failed: stats.failed.slice(), busy: stats.inFlight > 0 || pending.size > 0 || timer !== null, started: stats.started, finished: stats.finished }; },
    resetStats: function () { stats.asked = 0; stats.done = 0; stats.failed = []; stats.started = 0; stats.finished = 0; },
    onChange: function (fn) { listeners.push(fn); },
    balance: balance,
    accountName: accountName,
    period: period,
    callApi: callApi,
    clearCache: clearCache,
    requestForAddress: function (address) {
      return registry.get(normalizeAddress(address)) || null;
    }
  };

  if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('BALANCE', balance);
    CustomFunctions.associate('ACCOUNTNAME', accountName);
    CustomFunctions.associate('PERIOD', period);
    CustomFunctions.associate('SQL', sql);
  }
})();
