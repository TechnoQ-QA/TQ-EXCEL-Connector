/* TQ Excel Connector - side panel. */
(function () {
  'use strict';

  var inExcel = false;
  var lastDrill = null; // { title, rows }
  var $ = function (id) { return document.getElementById(id); };

  function say(message, isError) {
    var el = $('message');
    el.textContent = message || '';
    el.className = isError ? 'bad' : '';
  }

  function needExcel() {
    if (!inExcel) say('Open this panel inside Excel to use this button.', true);
    return inExcel;
  }

  function showStatus() {
    fetch('/api/status')
      .then(function (r) { return r.json(); })
      .then(function (status) {
        var s = status.signIn || {};
        if (s.mode === 'oauth2' && !s.signedIn) {   // Sign in with NetSuite: nobody signed in on this PC yet
          $('dot').className = 'dot bad';
          $('status').textContent = 'Not signed in to NetSuite. ';
          var link = document.createElement('a');
          link.href = '#';
          link.textContent = 'Sign in with NetSuite';
          link.addEventListener('click', function (event) { event.preventDefault(); signIn(); });
          $('status').appendChild(link);
          return null;
        }
        return TQ.callApi({ action: 'ping' }).then(function (ping) {
          var licence = status.licence;
          var bad = status.mode === 'netsuite' && licence && !licence.ok;
          $('dot').className = bad ? 'dot bad' : 'dot ok';
          $('status').textContent = status.mode === 'netsuite'
            ? (bad ? licence.message : 'Connected to NetSuite account ' + ping.account + ' as ' + ping.user)
            : 'Sample data (not connected to NetSuite)';
          if (s.mode === 'oauth2' && s.signedIn) {
            var out = document.createElement('a');
            out.href = '#';
            out.textContent = ' - Sign out';
            out.addEventListener('click', function (event) { event.preventDefault(); signOut(); });
            $('status').appendChild(out);
          }
          if (status.ai) $('status').textContent += ' · AI add-on on';
          showUpdate(status.update);
        });
      })
      .catch(function (error) {
        $('dot').className = 'dot bad';
        $('status').textContent = 'Not connected: ' + error.message;
      });
  }

  /* A newer signed version of the connector: one line with Install (the setup restarts the connector). */
  function showUpdate(update) {
    var old = document.getElementById('updateNotice');
    if (old) old.remove();
    if (!update || !update.available) return;
    var line = document.createElement('div');
    line.id = 'updateNotice';
    line.className = 'status';
    line.style.marginTop = '6px';
    line.textContent = 'TQ Connector ' + update.version + ' is available' + (update.notes ? ': ' + update.notes : '') + ' ';
    var b = document.createElement('button');
    b.textContent = 'Install';
    b.addEventListener('click', function () {
      b.disabled = true;
      line.firstChild.textContent = 'Installing ' + update.version + '... the connector restarts in a minute; then close and reopen this panel. ';
      fetch('/api/update', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"op":"install"}' })
        .then(function (r) { return r.json(); })
        .then(function (data) { if (!data.ok) line.firstChild.textContent = data.error + ' '; })
        .catch(function () {});
    });
    line.appendChild(b);
    document.querySelector('.status').insertAdjacentElement('afterend', line);
  }

  /* Opens NetSuite's sign-in page in the user's browser; the panel checks back until the sign-in is done. */
  function signIn() {
    var url = location.origin + '/oauth/start';
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(url);
      else window.open(url, '_blank');
    } catch (e) { window.open(url, '_blank'); }
    $('status').textContent = 'Finish signing in in your browser, then come back here...';
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
    say('Refreshing…');
    Excel.run(function (context) {
      context.workbook.application.calculate(Excel.CalculationType.full);
      return context.sync();
    }).then(function () { say('Refresh requested.'); }, function (e) { say(e.message, true); });
  }

  /* Only the active sheet: quicker than Refresh all once a pack has many location sheets. */
  function refreshSheet() {
    if (!needExcel()) return;
    TQ.clearCache();
    say('Refreshing this sheet…');
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.calculate(true);
      sheet.load('name');
      return context.sync().then(function () { return sheet.name; });
    }).then(function (name) { say('Refresh requested for sheet "' + name + '".'); }, function (e) { say(e.message, true); });
  }

  /* Only the selected cells: one changed number without waiting for the whole sheet. */
  function refreshSelection() {
    if (!needExcel()) return;
    TQ.clearCache();
    say('Refreshing the selected cells…');
    Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      range.calculate();
      range.load('address');
      return context.sync().then(function () { return range.address; });
    }).then(function (address) { say('Refresh requested for ' + address.split('!').pop() + '.'); }, function (e) { say(e.message, true); });
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
    $('drillSection').hidden = false;
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
    say('Loading transactions…');
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
    say('Loading the accounts behind this cell…');
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
      $('drillSection').hidden = false;
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
    $('drillClose').addEventListener('click', function () { $('drillSection').hidden = true; say(''); });
    Array.prototype.forEach.call(document.querySelectorAll('[data-list]'), function (button) {
      button.addEventListener('click', function () { insertList(button.getAttribute('data-list')); });
    });
    showStatus();
  }

  // Exposed so the panel can be checked in a plain browser, outside Excel.
  window.TQPanel = { drill: drill, breakdown: breakdown, refreshAll: refreshAll, refreshSheet: refreshSheet,
    refreshSelection: refreshSelection, drillSelected: drillSelected };

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(start);
  } else {
    document.addEventListener('DOMContentLoaded', function () { start(null); });
  }
})();
