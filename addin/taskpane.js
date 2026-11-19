/* TQ Excel Connector - side panel. */
(function () {
  'use strict';

  var inExcel = false;
  var lastDrill = null; // { title, rows }
  var lastAsk = null;   // { kind: 'transactions' | 'breakdown', request }: what Open large shows
  var $ = function (id) { return document.getElementById(id); };
  var hideTimer = null;

  /* The message line at the bottom. Plain news fades after a few seconds; errors, messages with a button and
   * work in progress (ending in ...) stay until they are replaced or dismissed. */
  function say(message, isError, action) {
    var el = $('message');
    clearTimeout(hideTimer);
    el.textContent = '';
    el.classList.toggle('bad', !!isError && !!message);
    if (!message) return;
    var text = document.createElement('span');
    text.className = 'grow';
    text.textContent = message;
    el.appendChild(text);
    if (action) {
      var a = document.createElement('button');
      a.className = 'btn link';
      a.textContent = action.label;
      a.addEventListener('click', action.run);
      el.appendChild(a);
    }
    var x = document.createElement('button');
    x.className = 'x';
    x.textContent = '\u2715';
    x.title = 'Dismiss';
    x.setAttribute('aria-label', 'Dismiss the message');
    x.addEventListener('click', function () { say(''); });
    el.appendChild(x);
    if (!isError && !action && !/\u2026$/.test(message)) hideTimer = setTimeout(function () { say(''); }, 6000);
  }

  /* Watches a refresh: "Refreshing... 340 of 1,284 numbers", then how many, how long, and any errors (with a
   * button that selects the cell). Excel may calculate in several waves, so it waits for a quiet moment. */
  var tracking = null;
  function track(label) {
    if (tracking) clearInterval(tracking);
    TQ.resetStats();
    var t0 = Date.now();
    var quiet = 0;
    say(label + '\u2026');
    tracking = setInterval(function () {
      var st = TQ.stats();
      if (st.busy || (!st.asked && Date.now() - t0 < 4000)) {
        quiet = 0;
        if (st.asked) say('Refreshing: ' + st.done.toLocaleString() + ' of ' + st.asked.toLocaleString() + ' numbers\u2026');
        return;
      }
      if (!quiet) { quiet = Date.now(); return; }
      if (Date.now() - quiet < 800 && Date.now() - t0 < 300000) return;
      clearInterval(tracking);
      tracking = null;
      if (!st.asked) { say('Done: no TQ numbers needed fetching from NetSuite.'); return; }
      var secs = (((st.finished || Date.now()) - (st.started || t0)) / 1000).toFixed(1);
      if (st.failed.length) {
        say('Refreshed ' + st.asked.toLocaleString() + ' numbers in ' + secs + ' s \u00b7 ' + st.failed.length + ' with an error.', true,
          { label: 'Show the first', run: function () { showFailure(st.failed, 0); } });
      } else {
        say('Refreshed ' + st.asked.toLocaleString() + ' numbers in ' + secs + ' s.');
      }
    }, 400);
  }

  /* Selects a cell whose number could not be fetched and says why; Next goes on to the following one. */
  function showFailure(list, n) {
    var f = list[n];
    if (!f) return;
    var cut = f.address.lastIndexOf('!');
    var sheet = cut > 0 ? f.address.slice(0, cut).replace(/^'|'$/g, '').split("''").join("'") : '';
    var cell = cut > 0 ? f.address.slice(cut + 1) : f.address;
    var next = n + 1 < list.length ? { label: 'Next (' + (n + 2) + ' of ' + list.length + ')', run: function () { showFailure(list, n + 1); } } : null;
    say((cell ? cell + ': ' : '') + f.message, true, next);
    if (!inExcel || !sheet || !cell) return;
    Excel.run(function (context) {
      var ws = context.workbook.worksheets.getItem(sheet);
      ws.activate();
      ws.getRange(cell).select();
      return context.sync();
    }).catch(function () {});
  }

  function needExcel() {
    if (!inExcel) say('Open this panel inside Excel to use this button.', true);
    return inExcel;
  }

  /* The line under the name in the header: the account and the user, or what is wrong. The full status (with
   * the ping) is kept in window.TQStatus for the welcome screen. */
  function showStatus() {
    var signOutItem = $('menuSignOut');
    if (signOutItem) signOutItem.hidden = true;
    window.TQStatus = fetch('/api/status').then(function (r) { return r.json(); })
      .then(function (status) {
        var s = status.signIn || {};
        if (s.mode === 'oauth2' && !s.signedIn) {   // Sign in with NetSuite: nobody signed in on this PC yet
          $('dot').className = 'dot bad';
          $('status').textContent = 'Not signed in. ';
          var link = document.createElement('a');
          link.href = '#';
          link.textContent = 'Sign in with NetSuite';
          link.addEventListener('click', function (event) { event.preventDefault(); signIn(); });
          $('status').appendChild(link);
          status.ping = null;
          return status;
        }
        return TQ.callApi({ action: 'ping' }).then(function (ping) {
          var licence = status.licence;
          var bad = status.mode === 'netsuite' && licence && !licence.ok;
          $('dot').className = bad ? 'dot bad' : 'dot ok';
          $('status').textContent = status.mode === 'netsuite'
            ? (bad ? licence.message : ping.account + ' · ' + ping.user)
            : 'Sample data (not connected to NetSuite)';
          $('status').parentNode.title = status.mode === 'netsuite' && !bad
            ? 'Connected to NetSuite account ' + ping.account + ' as ' + ping.user + (status.ai ? ' · AI add-on on' : '') : '';
          if (s.mode === 'oauth2' && s.signedIn && signOutItem) signOutItem.hidden = false;
          showUpdate(status.update);
          status.ping = ping;
          return status;
        });
      })
      .catch(function (error) {
        $('dot').className = 'dot bad';
        $('status').textContent = 'Not connected: ' + error.message;
        return { error: error.message };
      });
    return window.TQStatus;
  }

  /* A newer signed version of the connector: one line with Install (the setup restarts the connector). */
  function showUpdate(update) {
    var old = document.getElementById('updateNotice');
    if (old) old.remove();
    if (!update || !update.available) return;
    var line = document.createElement('div');
    line.id = 'updateNotice';
    line.className = 'banner';
    var text = document.createElement('div');
    text.className = 'grow';
    text.textContent = 'Version ' + update.version + ' is available' + (update.notes ? ': ' + update.notes : '.');
    line.appendChild(text);
    var b = document.createElement('button');
    b.className = 'btn';
    b.textContent = 'Install';
    b.addEventListener('click', function () {
      b.disabled = true;
      text.textContent = 'Installing ' + update.version + '... the connector restarts in a minute; then close and reopen this panel.';
      fetch('/api/update', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"op":"install"}' })
        .then(function (r) { return r.json(); })
        .then(function (data) { if (!data.ok) text.textContent = data.error; })
        .catch(function () {});
    });
    line.appendChild(b);
    $('banners').appendChild(line);
  }

  /* Opens NetSuite's sign-in page in the user's browser; the panel checks back until the sign-in is done. */
  function signIn() {
    var url = location.origin + '/oauth/start';
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(url);
      else window.open(url, '_blank');
    } catch (e) { window.open(url, '_blank'); }
    $('status').textContent = 'Finish signing in in your browser…';
    var tries = 0;
    var timer = setInterval(function () {
      tries += 1;
      fetch('/api/status').then(function (r) { return r.json(); }).then(function (st) {
        if ((st.signIn && st.signIn.signedIn) || tries > 100) { clearInterval(timer); showStatus(); }
      }).catch(function () {});
    }, 3000);
  }

  function signOut() {
    fetch('/api/signout', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
      .then(function () { TQ.clearCache(); showStatus(); });
  }

  function refreshAll() {
    if (!needExcel()) return;
    TQ.clearCache();
    track('Refreshing the workbook');
    Excel.run(function (context) {
      context.workbook.application.calculate(Excel.CalculationType.full);
      return context.sync();
    }).catch(function (e) { say(e.message, true); });
  }

  /* Only the active sheet: quicker than Refresh all once a pack has many location sheets. */
  function refreshSheet() {
    if (!needExcel()) return;
    TQ.clearCache();
    track('Refreshing this sheet');
    Excel.run(function (context) {
      context.workbook.worksheets.getActiveWorksheet().calculate(true);
      return context.sync();
    }).catch(function (e) { say(e.message, true); });
  }

  /* Only the selected cells: one changed number without waiting for the whole sheet. */
  function refreshSelection() {
    if (!needExcel()) return;
    TQ.clearCache();
    track('Refreshing the selected cells');
    Excel.run(function (context) {
      context.workbook.getSelectedRange().calculate();
      return context.sync();
    }).catch(function (e) { say(e.message, true); });
  }

  function describe(request) {
    var parts = [request.account, request.fromPeriod + (request.toPeriod ? ' to ' + request.toPeriod : '')];
    ['subsidiary', 'department', 'class', 'location', 'filters'].forEach(function (k) {
      if (request[k]) parts.push(request[k]);
    });
    return parts.join(' · ');
  }

  /* Opens a NetSuite page in the user's browser (signed in there), not inside the panel. */
  function openUrl(url) {
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (e) { /* fall back below */ }
    window.open(url, '_blank');
  }

  function renderDrill(title, data) {
    var body = $('drillBody');
    body.textContent = '';
    data.rows.forEach(function (row) {
      var tr = document.createElement('tr');
      [row.date, row.number, row.type, row.entity, row.memo].forEach(function (value, n) {
        var td = document.createElement('td');
        if (n === 1 && row.url) {   // the number opens the transaction in NetSuite
          var link = document.createElement('a');
          link.href = row.url;
          link.textContent = value || 'open';
          link.title = 'Open in NetSuite';
          link.addEventListener('click', function (event) { event.preventDefault(); openUrl(row.url); });
          td.appendChild(link);
        } else {
          td.textContent = value || '';
        }
        tr.appendChild(td);
      });
      var amount = document.createElement('td');
      amount.className = 'num';
      amount.textContent = Number(row.amount).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      tr.appendChild(amount);
      body.appendChild(tr);
    });
    $('drillTitle').textContent = 'Transactions: ' + title;
    window.TQNav.go('drill');
    say(data.rows.length + ' line(s)' + (data.truncated ? ' (list cut at 1,000)' : '') + '.');
  }

  var HEADS = {
    transactions: ['Date', 'Number', 'Type', 'Name', 'Memo', 'Amount'],
    accounts: ['Account', 'Name', 'Type', '', '', 'Amount']
  };

  function setHeads(kind) {
    var cells = document.querySelectorAll('#drillSection thead th');
    HEADS[kind].forEach(function (text, n) { cells[n].textContent = text; });
  }

  function drill(request) {
    lastAsk = { kind: 'transactions', request: request };
    say('Loading transactions\u2026');
    return TQ.callApi({ action: 'transactions', request: request }).then(function (data) {
      lastDrill = { title: describe(request), head: HEADS.transactions,
        rows: data.rows.map(function (r) {
          // On the sheet the number is a link to the transaction in NetSuite.
          var number = r.url ? '=HYPERLINK("' + r.url + '","' + String(r.number || 'open').replace(/"/g, '""') + '")' : "'" + r.number;
          return ["'" + r.date, number, r.type, r.entity, r.memo, r.amount];
        }) };
      setHeads('transactions');
      renderDrill(lastDrill.title, data);
    });
  }

  /* A cell that adds up several accounts (a list, a range, a type, net income, a budget): first the
   * accounts behind it with each one's amount; a click on an account then shows its transactions. */
  function isTotal(account) {
    return /^(type|calc|budget|move)\b|[,:]|\.\./i.test(account) || /^[\w.]+\s*-\s*[\w.]+$/.test(account);
  }

  function breakdown(request) {
    lastAsk = { kind: 'breakdown', request: request };
    say('Loading the accounts behind this cell\u2026');
    return TQ.callApi({ action: 'breakdown', request: request }).then(function (data) {
      if (data.rows.length === 1 && data.rows[0].request) return drill(data.rows[0].request);
      var body = $('drillBody');
      body.textContent = '';
      data.rows.forEach(function (row) {
        var tr = document.createElement('tr');
        [row.account, row.name, row.type, '', ''].forEach(function (value) {
          var td = document.createElement('td');
          td.textContent = value || '';
          tr.appendChild(td);
        });
        var amount = document.createElement('td');
        amount.className = 'num';
        amount.textContent = Number(row.amount).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        tr.appendChild(amount);
        if (row.request) {
          tr.className = 'clickable';
          tr.tabIndex = 0;
          tr.title = 'Show the transactions of ' + row.account;
          var open = function () { drill(row.request).catch(function (e) { say(e.message, true); }); };
          tr.addEventListener('click', open);
          tr.addEventListener('keydown', function (event) { if (event.key === 'Enter') open(); });
        }
        body.appendChild(tr);
      });
      setHeads('accounts');
      lastDrill = { title: describe(request), head: HEADS.accounts,
        rows: data.rows.map(function (r) { return ["'" + r.account, r.name, r.type, '', '', r.amount]; }) };
      $('drillTitle').textContent = (data.kind === 'budget' ? 'Budget by account: ' : 'Accounts: ') + lastDrill.title;
      window.TQNav.go('drill');
      say(data.rows.length + ' account(s) with an amount, total ' +
        Number(data.total).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
        (data.kind === 'budget' ? '. A budget has no transactions behind it.' : '. Click an account for its transactions.'));
    });
  }

  function drillSelected() {
    if (!needExcel()) return;
    Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      range.load('address');
      return context.sync().then(function () { return range.address; });
    }).then(function (address) {
      var request = TQ.requestForAddress(address);
      if (!request) {
        say('Select a cell that holds a balance (not a subtotal or a ratio). If the workbook was just opened, press Refresh all first.', true);
        return null;
      }
      return isTotal(String(request.account)) ? breakdown(request) : drill(request);
    }).catch(function (e) { say(e.message, true); });
  }

  function writeAtSelection(values) {
    return Excel.run(function (context) {
      var target = context.workbook.getSelectedRange().getCell(0, 0)
        .getResizedRange(values.length - 1, values[0].length - 1);
      target.values = values;
      return context.sync();
    });
  }

  function listValues(kind) {
    if (kind === 'accounts') {
      return TQ.callApi({ action: 'accounts' }).then(function (d) {
        return [['Account', 'Name', 'Type']].concat(d.accounts.map(function (a) {
          return ["'" + a.number, a.name, a.type];
        }));
      });
    }
    if (kind === 'periods') {
      return TQ.callApi({ action: 'periods' }).then(function (d) {
        return [['Period', 'Start', 'End']].concat(d.periods.map(function (p) {
          return ["'" + p.name, "'" + p.start, "'" + p.end];
        }));
      });
    }
    return TQ.callApi({ action: 'segments', kind: kind }).then(function (d) {
      return [['Internal ID', 'Name']].concat(d.segments.map(function (s) { return [s.id, s.name]; }));
    });
  }

  function insertList(kind) {
    if (!needExcel()) return;
    say('Loading…');
    listValues(kind)
      .then(function (values) {
        if (values.length < 2) { say('Nothing to insert.'); return null; }
        return writeAtSelection(values).then(function () { say((values.length - 1) + ' row(s) inserted.'); });
      })
      .catch(function (e) { say(e.message, true); });
  }

  /* Drill down in a large window of its own: every column, sorting, a search box and Send to a new sheet. The
   * window asks the connector itself; what it sends back (rows for a new sheet) comes here. */
  function openLarge() {
    if (!lastAsk) return;
    var url = location.origin + '/drill.html#' + encodeURIComponent(JSON.stringify(lastAsk));
    if (!(window.Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync)) {
      window.open(url, '_blank');
      return;
    }
    Office.context.ui.displayDialogAsync(url, { height: 80, width: 75, displayInIframe: false }, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) { say(result.error.message, true); return; }
      var dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = {};
        try { msg = JSON.parse(arg.message); } catch (e) { return; }
        if (msg.toSheet) toSheet(msg.toSheet.head, msg.toSheet.rows);
        if (msg.close) dialog.close();
      });
    });
  }

  function toSheet(head, rows) {
    if (!needExcel()) return;
    var values = [head].concat(rows);
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.add();
      sheet.getRange('A1').getResizedRange(values.length - 1, values[0].length - 1).formulas = values;
      sheet.getRange('A1').getResizedRange(0, values[0].length - 1).format.font.bold = true;
      sheet.getRange('A1').getResizedRange(values.length - 1, values[0].length - 1).format.autofitColumns();
      sheet.activate();
      return context.sync();
    }).then(function () { say('Sent ' + rows.length + ' row(s) to a new sheet.'); }, function (e) { say(e.message, true); });
  }

  function drillToSheet() {
    if (!needExcel() || !lastDrill) return;
    var values = [lastDrill.head].concat(lastDrill.rows);
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.add();
      sheet.getRange('A1').getResizedRange(values.length - 1, values[0].length - 1).formulas = values;
      sheet.activate();
      return context.sync();
    }).then(function () { say('Sent to a new sheet.'); }, function (e) { say(e.message, true); });
  }

  function start(info) {
    inExcel = !!(info && info.host && typeof Excel !== 'undefined');
    $('refresh').addEventListener('click', refreshAll);
    $('refreshSheet').addEventListener('click', refreshSheet);
    $('refreshSelection').addEventListener('click', refreshSelection);
    $('drill').addEventListener('click', drillSelected);
    $('drillToSheet').addEventListener('click', drillToSheet);
    if ($('drillLarge')) $('drillLarge').addEventListener('click', openLarge);
    $('drillClose').addEventListener('click', function () { say(''); });
    if ($('menuSignOut')) $('menuSignOut').addEventListener('click', signOut);
    Array.prototype.forEach.call(document.querySelectorAll('[data-list]'), function (button) {
      button.addEventListener('click', function () { insertList(button.getAttribute('data-list')); });
    });
    showStatus();
  }

  // Exposed so the panel can be checked in a plain browser, outside Excel.
  window.TQPanel = { showStatus: showStatus, say: say, track: track, toSheet: toSheet, openLarge: openLarge, drill: drill, breakdown: breakdown, refreshAll: refreshAll, refreshSheet: refreshSheet,
    refreshSelection: refreshSelection, drillSelected: drillSelected };

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(start);
  } else {
    document.addEventListener('DOMContentLoaded', function () { start(null); });
  }
})();
