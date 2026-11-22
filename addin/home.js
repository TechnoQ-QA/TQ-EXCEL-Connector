/* TQ Excel Connector - Home, This report, the welcome screen, the report wizard and the ribbon.
 *   - This report follows the open sheet: on a TQ report or dashboard it shows the month (with ‹ ›), the
 *     filters as chips, and the way to its layout, send and dashboard. Other sheets go back to Home, where a
 *     layout, budget or journal sheet gets a short card of what it needs.
 *   - Home lists the reports in the workbook; a click opens that sheet.
 *   - The wizard: one question per step, then one request that builds a layout report (so every report it
 *     makes can be changed and rebuilt in place).
 *   - The ribbon tab's buttons (Office.actions), which run here because the add-in shares one runtime. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var LAYOUT_RE = /layout sheet "([^"]+)"/;
  var DASH_RE = /^Live from NetSuite/;
  var periods = null;   // [{name}] from NetSuite, loaded once
  var FOLLOW = ['home', 'report'];   // screens that change by themselves when another sheet is opened
  var currentSheet = '';             // the sheet on the This report screen

  // ------------------------------------------- when each sheet's numbers came in --
  // Kept in the workbook's own settings (saved with the file), not in a cell: SHEET NAME -> time.
  var stamps = {};
  var saveTimer = null;
  function key(name) { return String(name || '').replace(/'/g, '').toUpperCase(); }
  function loadStamps() {
    try { stamps = Office.context.document.settings.get('tqRefreshed') || {}; } catch (e) { stamps = {}; }
  }
  function stampSheets(names) {
    var now = Date.now();
    names.forEach(function (n) { stamps[key(n)] = now; });
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try {
        Office.context.document.settings.set('tqRefreshed', stamps);
        Office.context.document.settings.saveAsync(function () {});
      } catch (e) { /* outside Excel */ }
    }, 1500);
    showStamp();
  }
  function stampOf(name) {
    var t = stamps[key(name)];
    if (!t) return { text: 'Not refreshed in this workbook yet', old: true };
    var d = new Date(t);
    var today = d.toDateString() === new Date().toDateString();
    var time = d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    return today ? { text: 'Refreshed today at ' + time, short: 'refreshed ' + time, old: false }
      : { text: 'Last refreshed ' + d.toLocaleDateString([], { day: 'numeric', month: 'short' }) + ' at ' + time + ': Refresh for the latest numbers',
        short: 'refreshed ' + d.toLocaleDateString([], { day: 'numeric', month: 'short' }), old: true };
  }
  function showStamp() {
    var box = $('rpStamp');
    if (!box || !currentSheet || !inExcel()) return;
    var st = stampOf(currentSheet);
    box.textContent = st.text;
    box.className = 'stamp' + (st.old ? ' old' : '');
    box.hidden = false;
  }

  function el(tag, text, cls) {
    var n = document.createElement(tag);
    if (text) n.textContent = text;
    if (cls) n.className = cls;
    return n;
  }
  function button(text, onClick, cls) {
    var b = el('button', text, cls || 'btn');
    b.type = 'button';
    b.addEventListener('click', onClick);
    return b;
  }
  function inExcel() { return typeof Excel !== 'undefined' && window.Office && Office.context && Office.context.document; }
  function months() {
    if (!periods) {
      periods = TQ.callApi({ action: 'periods' }).then(function (d) {
        return d.periods.filter(function (p) {
          var days = (new Date(p.end) - new Date(p.start)) / 86400000;
          return days > 20 && days < 40;   // months only: no adjustment periods, quarters or years
        });
      }).catch(function (e) { periods = null; throw e; });
    }
    return periods;
  }

  /* The older section ids, as the screens that now hold them. */
  var SCREEN_OF = { secFilters: 'report', secLayout: 'layout', secPack: 'send', secBudget: 'upload', secJournal: 'upload',
    secFormula: 'formulas', secData: 'data', secCharts: 'dash' };
  function openSection(id) { window.TQNav.go(SCREEN_OF[id] || 'home'); }

  // -------------------------------------------------------- the open sheet --

  function readActive() {
    if (!inExcel()) return Promise.resolve({ name: '', a: [] });
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange('A1:B8');
      sheet.load('name');
      range.load('values');
      return context.sync().then(function () { return { name: sheet.name, a: range.values }; });
    }).catch(function () { return { name: '', a: [] }; });
  }

  function kindOf(a) {
    var cell = function (r, c) { return a[r] && a[r][c] !== undefined ? String(a[r][c]) : ''; };
    var layout = LAYOUT_RE.exec(cell(1, 0));
    if (DASH_RE.test(cell(1, 0)) && cell(2, 0) === 'Other filters') return { kind: 'dashboard', title: cell(0, 0), month: cell(7, 1) };
    if (layout || cell(2, 0) === 'Other filters') {
      return { kind: 'report', title: cell(0, 0), layout: layout && layout[1], month: cell(7, 0) === 'Month' ? cell(7, 1) : '' };
    }
    return { kind: '' };
  }

  /* Called when the panel opens and whenever another sheet is opened. */
  function showCard() {
    listReports();
    return readActive().then(function (s) {
      var k = kindOf(s.a || []);
      var follow = FOLLOW.indexOf(window.TQNav.current()) !== -1;
      contextCard(s.name);
      if (k.kind) {
        fillReport(s.name, k);
        if (follow) window.TQNav.replace('report');
      } else if (follow && window.TQNav.current() === 'report') {
        window.TQNav.replace('home');
      }
    });
  }

  /* On Home: what a layout, budget or journal sheet needs (nothing on other sheets). */
  function contextCard(name) {
    var box = $('homeCard');
    box.textContent = '';
    var add = function (title, line, buttons) {
      box.appendChild(el('h3', title));
      box.appendChild(el('div', line, 'hint'));
      var row = el('div', '', 'row wrap');
      buttons.forEach(function (b) { row.appendChild(b); });
      box.appendChild(row);
      box.hidden = false;
    };
    box.hidden = true;
    if (/^tq layout/i.test(name)) {
      add('Layout sheet', 'Change the lines and columns here or in the layout screen, then update the report.',
        [button('Update report', function () { TQDesigner.select(name).then(function () { TQDesigner.build(); }); }, 'btn primary'),
          button('Open the layout', function () { TQDesigner.select(name).then(function () { window.TQNav.go('layout'); }); })]);
    } else if (/^tq budget/i.test(name)) {
      add('Budget sheet', 'Change the numbers, then Check (writes nothing) and Upload.',
        [button('Check and upload', function () { window.TQNav.go('upload'); }, 'btn primary')]);
    } else if (/^tq journals/i.test(name)) {
      add('Journal sheet', 'Rows with the same Entry number become one journal entry.',
        [button('Check and upload', function () { window.TQNav.go('upload'); }, 'btn primary')]);
    }
  }

  var reportLayout = '';   // the layout of the report on the This report screen
  function fillReport(name, k) {
    currentSheet = name;
    showStamp();
    $('rpTitle').textContent = k.title || name;
    reportLayout = k.layout || '';
    $('rpAbout').textContent = k.kind === 'dashboard' ? 'Dashboard on sheet "' + name + '". Change the month or the filters; the tiles and charts follow.'
      : k.layout ? 'Sheet "' + name + '", from layout "' + k.layout + '". Changes there rebuild it in place.'
        : 'Sheet "' + name + '". A standard report: ask the chat for changes, or use New report for one you can edit.';
    $('rpLayoutRow').hidden = !k.layout;
    $('rpRestoreRow').hidden = !(window.TQChat && TQChat.hasBackup(name));
    $('rpMonthBox').hidden = !k.month;
    if (k.month) monthPicker(k.month);
  }

  function monthPicker(current) {
    var select = $('hMonth');
    select.textContent = '';
    var o = el('option', current);
    o.value = current;
    select.appendChild(o);
    select.value = current;
    months().then(function (list) {
      select.textContent = '';
      list.forEach(function (p) { var x = el('option', p.name); x.value = p.name; select.appendChild(x); });
      if (!list.some(function (p) { return p.name === current; })) { var y = el('option', current); y.value = current; select.appendChild(y); }
      select.value = current;
    }).catch(function () {});
  }

  function setMonth(value) {
    if (!inExcel() || !value) return;
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange('B8').values = [[value]];
      sheet.calculate(true);
      return context.sync();
    }).then(function () { TQPanel.say('Month set to ' + value + '. The numbers are being refreshed.'); },
      function (e) { TQPanel.say(e.message, true); });
  }

  function stepMonth(by) {
    var select = $('hMonth');
    var n = select.selectedIndex + by;
    if (n < 0 || n >= select.options.length) return;
    select.selectedIndex = n;
    setMonth(select.value);
  }

  var known = [];   // the reports in the workbook, for the quick search
  function openSheet(name) {
    Excel.run(function (context) {
      context.workbook.worksheets.getItem(name).activate();
      return context.sync();
    }).then(function () { window.TQNav.go('report'); showCard(); }).catch(function () {});
  }

  /* Home's list of the TQ reports and dashboards in this workbook. */
  function listReports() {
    var box = $('reportList');
    if (!inExcel()) return;
    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name,items/visibility');
      return context.sync().then(function () {
        var shown = sheets.items.filter(function (s) { return s.visibility === 'Visible'; });
        var ranges = shown.map(function (s) { var r = s.getRange('A1:B8'); r.load('values'); return r; });
        return context.sync().then(function () {
          return shown.map(function (s, n) { return { name: s.name, k: kindOf(ranges[n].values) }; })
            .filter(function (x) { return x.k.kind; });
        });
      });
    }).then(function (found) {
      known = found.map(function (x) { return { name: x.name, title: x.k.title || x.name }; });
      box.textContent = '';
      if (!found.length) { box.appendChild(el('div', 'No TQ reports yet. Start with New report.', 'empty')); return; }
      found.forEach(function (x) {
        var item = el('button', '', 'it');
        item.type = 'button';
        var grow = el('div', '', 'grow');
        grow.appendChild(el('b', x.k.title || x.name));
        grow.appendChild(el('small', (x.k.kind === 'dashboard' ? 'Dashboard' : 'Report') + ' · sheet ' + x.name + (x.k.month ? ' · ' + x.k.month : '')));
        var st = stampOf(x.name);
        if (st.short) { grow.appendChild(document.createTextNode(' ')); grow.appendChild(el('small', '· ' + st.short, st.old ? 'old' : '')); }
        item.appendChild(grow);
        item.appendChild(el('span', '›'));
        item.addEventListener('click', function () {
          Excel.run(function (context) {
            context.workbook.worksheets.getItem(x.name).activate();
            return context.sync();
          }).then(function () { window.TQNav.go('report'); showCard(); }).catch(function () {});
        });
        box.appendChild(item);
      });
    }).catch(function () {});
  }

  // ------------------------------------------------------------ the welcome --

  function welcomed() { try { return localStorage.getItem('tq.welcomed') === '1'; } catch (e) { return true; } }
  function setWelcomed() { try { localStorage.setItem('tq.welcomed', '1'); } catch (e) { /* private storage */ } }

  function welcome() {
    var box = $('welcomeChecks');
    box.textContent = '';
    var line = function (ok, title, note) {
      var r = el('div', '', 'check-row');
      r.appendChild(el('span', ok === true ? '✓' : ok === false ? '!' : '•', 'tick' + (ok === true ? '' : ok === false ? ' bad' : ' todo')));
      var t = el('div');
      t.appendChild(el('b', title));
      if (note) t.appendChild(el('small', note));
      r.appendChild(t);
      box.appendChild(r);
    };
    window.TQNav.replace('welcome');
    (window.TQStatus || Promise.resolve({})).then(function (st) {
      st = st || {};
      var sign = st.signIn || {};
      if (st.mode === 'sample') line(null, 'Sample data', 'Not connected to NetSuite yet: Setup (the gear) connects it.');
      else if (st.error) line(false, 'The connector is not answering', st.error);
      else if (sign.mode === 'oauth2' && !sign.signedIn) line(false, 'Not signed in to NetSuite', 'Use "Sign in with NetSuite" at the top of the panel.');
      else if (st.ping) line(true, 'Connected to NetSuite', 'Account ' + st.ping.account);
      if (st.ping) line(true, 'Signed in as ' + st.ping.user, (st.ping.role || '') + (st.seats && st.seats.used ? ' · seat ' + st.seats.used + ' of ' + st.seats.users : ''));
      if (st.licence) line(!!st.licence.ok, st.licence.ok ? 'Licensed to ' + st.licence.customer : 'Licence', st.licence.message);
      line(null, 'Build your first report', 'Pick a format, a month and filters.');
    });
  }

  // ----------------------------------------------------------------- the wizard --

  var KINDS = [
    { id: 'monthytd', group: 'Reports', title: 'Monthly P&L', about: 'The month and the year to date against last year. Rolls forward with one cell.', rolling: true },
    { id: 'summary', group: 'Reports', title: 'P&L summary', about: 'One line per group (sales, cost of sales, expenses...), by month or quarter.' },
    { id: 'management', group: 'Reports', title: 'P&L, account by account', about: 'Every account with a balance, grouped, with margins.' },
    { id: 'budget', group: 'Reports', title: 'Budget against actual', about: 'Actual, budget, variance and variance % per period.', template: 'management', compare: 'budget' },
    { id: 'balance', group: 'Reports', title: 'Balance sheet', about: 'As at a month, the same month last year and last year end. Rolls forward.', rolling: true },
    { id: 'cashflow', group: 'Reports', title: 'Cash flow', about: 'Indirect method, month by month, checked against the bank accounts.' },
    { id: 'own', group: 'Reports', title: 'My own format', about: 'Typed on a sheet: its rows become the report lines.' },
    { id: 'hotel', group: 'Hospitality', title: 'Hotel P&L (USALI summary)', about: 'Rooms, F&B and other operated departments, undistributed expenses, GOP and GOP %, with occupancy, ADR and RevPAR when the statistics are in NetSuite. Rolls forward.', rolling: true },
    { id: 'usali', group: 'Hospitality', title: 'USALI income statement by department', about: 'Each operated department (Rooms, F&B, Telecom, Laundry, Recreation, Rental) with revenue, payroll, other expenses and departmental profit, then undistributed expenses, GOP and net profit, with the accounts under each line. Month and YTD against budget and last year, with % of income. Rolls forward.', rolling: true },
    { id: 'restaurant', group: 'Hospitality', title: 'Restaurant P&L', about: 'Food and beverage sales, food cost %, beverage cost %, labour %, prime cost and operating profit. Rolls forward.', rolling: true },
    { id: 'forecast', group: 'Planning', title: 'Rolling forecast (outlook)', about: 'Actual so far this year plus the forecast for the months left, the full-year outlook against the budget. Rolls forward.', rolling: true },
    { id: 'rolling12', group: 'Planning', title: 'Last 12 months', about: 'The twelve months up to the report month, side by side, with the total. Rolls forward.', rolling: true },
    { id: 'op-ar', group: 'Operational reports', title: 'Receivables aging', about: 'Unpaid customer invoices by customer, in buckets by days overdue today.', query: 'Receivables aging' },
    { id: 'op-ap', group: 'Operational reports', title: 'Payables aging', about: 'Unpaid vendor bills by vendor, in buckets by days overdue today.', query: 'Payables aging' },
    { id: 'op-open', group: 'Operational reports', title: 'Open invoices', about: 'Every unpaid customer invoice with its due date and days overdue.', query: 'Open invoices' },
    { id: 'op-sales-cust', group: 'Operational reports', title: 'Sales by customer', about: 'Invoices and cash sales in a date range, totalled by customer.', query: 'Sales by customer', dates: true },
    { id: 'op-sales-item', group: 'Operational reports', title: 'Sales by item', about: 'Quantity and amount sold per item in a date range.', query: 'Sales by item', dates: true },
    { id: 'op-purch', group: 'Operational reports', title: 'Purchases by vendor', about: 'Vendor bills in a date range, totalled by vendor.', query: 'Purchases by vendor', dates: true },
    { id: 'op-stock', group: 'Operational reports', title: 'Stock on hand', about: 'Quantity on hand and available for each item and location.', query: 'Stock on hand' },
    { id: 'dash-executive', group: 'Dashboards', title: 'Dashboard: executive', about: 'Tiles for sales, gross profit, net income and cash; charts of sales by month, by location and where the money went.', rolling: true, dashboard: 'executive' },
    { id: 'dash-sales', group: 'Dashboards', title: 'Dashboard: sales', about: 'Sales this month and year to date, sales by month against last year, gross margin, sales by location.', rolling: true, dashboard: 'sales' },
    { id: 'dash-cash', group: 'Dashboards', title: 'Dashboard: cash', about: 'Cash, receivables and payables by month, cash movement and the bank accounts.', rolling: true, dashboard: 'cash' },
    { id: 'dash-pl', group: 'Dashboards', title: 'Finance: profit and loss', about: 'Revenue, gross and net margin, operating expenses; revenue, costs and net income by month; a bridge from revenue to net income.', rolling: true, dashboard: 'pl' },
    { id: 'dash-position', group: 'Dashboards', title: 'Finance: financial position', about: 'Working capital, current and quick ratio, debt to equity, debtor and creditor days; what the company owns and owes.', rolling: true, dashboard: 'position' },
    { id: 'dash-budget', group: 'Dashboards', title: 'Finance: budget against actual', about: 'Revenue, costs and net income against the budget, by month and so far this year.', rolling: true, dashboard: 'budget' }
  ];
  var STEPS = ['What', 'Periods', 'Filters', 'Sheets', 'Check'];
  var w = null;   // the answers so far

  function openWizard() {
    w = { step: 0, kind: KINDS[0], filters: {}, dropdowns: [], pages: '', detail: false, comments: false, highlight: '' };
    window.TQNav.go('wizard');
    draw();
  }
  function closeWizard() { window.TQNav.back(); }

  function field(labelText, control) {
    var box = el('div');
    box.appendChild(el('label', labelText));
    box.appendChild(control);
    return box;
  }
  function select(options, value, onChange) {
    var s = el('select', '', 'field');
    options.forEach(function (o) { var x = el('option', o[1]); x.value = o[0]; s.appendChild(x); });
    s.value = value;
    s.addEventListener('change', function () { onChange(s.value); });
    return s;
  }

  function draw() {
    var box = $('wizard');
    box.textContent = '';
    box.appendChild(el('h3', w.step === 0 ? 'New report' : w.kind.title));
    var bar = el('div', '', 'steps');
    STEPS.forEach(function (s, n) { if (!((w.kind.dashboard && n === 3) || (w.kind.query && (n === 2 || n === 3)))) bar.appendChild(el('span', '', n <= w.step ? 'on' : '')); });
    box.appendChild(bar);
    box.appendChild(el('div', 'Step: ' + STEPS[w.step], 'steps-label'));
    var body = el('div');
    box.appendChild(body);
    [stepWhat, stepPeriods, stepFilters, stepSheets, stepCheck][w.step](body);
    var row = el('div', '', 'row wrap');
    row.style.marginTop = '14px';
    if (w.step > 0) row.appendChild(button('Back', function () { w.step = w.kind.query && w.step === 4 ? 1 : w.step - ((w.kind.dashboard && w.step === 4) ? 2 : 1); draw(); }));
    if (w.step > 0 && w.step < STEPS.length - 1) {
      row.appendChild(button('Next', function () {
        w.step += 1;
        if (w.kind.query) w.step = 4;                       // a data report has no filters or extra sheets
        if (w.kind.dashboard && w.step === 3) w.step = 4;   // a dashboard has no extra sheets to choose
        draw();
      }, 'btn primary'));
    }
    else if (w.step === STEPS.length - 1) row.appendChild(button(w.kind.dashboard ? 'Create the dashboard' : 'Create the report', create, 'btn primary'));
    row.appendChild(button('Cancel', closeWizard));
    box.appendChild(row);
    var msg = el('div', '', 'status-line');
    msg.id = 'wMessage';
    box.appendChild(msg);
  }

  function stepWhat(body) {
    var list = el('div', '', 'choices');
    var group = '';
    KINDS.forEach(function (k) {
      if (k.group !== group) { group = k.group; list.appendChild(el('div', group, 'group-title')); }
      var b = button('', function () {
        w.kind = k;
        w.detail = k.id === 'usali';   // the USALI statement lists the accounts under each line, as NetSuite's does
        if (k.id === 'own') {   // made from a sheet: the layout screen does that
          $('dTemplate').value = '(sheet)';
          $('dTemplate').dispatchEvent(new Event('change'));
          window.TQNav.replace('layout');
          var fold = document.querySelector('#secLayout details.fold');
          if (fold) fold.open = true;
          return;
        }
        w.step = 1;
        draw();
      }, w.kind === k ? 'on' : '');
      b.appendChild(el('strong', k.title));
      b.appendChild(el('small', k.about));
      list.appendChild(b);
    });
    if (window.TQLibrary) TQLibrary.fillWizard(list);   // layouts the company shared, when there are any
    body.appendChild(list);
  }

  /* NetSuite did not answer (not signed in, offline): say so in the list and below, instead of "Loading…". */
  function periodsFailed(selects) {
    return function (e) {
      selects.forEach(function (sel) {
        sel.textContent = '';
        var o = el('option', 'Could not load the periods');
        o.value = '';
        sel.appendChild(o);
      });
      var m = $('wMessage');
      if (m) { m.textContent = 'NetSuite did not send the periods (' + (e && e.message || 'no answer') + '). Check you are signed in, then press Back and Next.'; m.className = 'status-line bad'; }
    };
  }

  function stepPeriods(body) {
    if (w.kind.query) {
      if (!w.kind.dates) { body.appendChild(el('p', 'Nothing to choose: it shows today\'s position. Press Next.', 'hint')); return; }
      var today = new Date();
      w.from = w.from || today.getFullYear() + '-01-01';
      w.to = w.to || today.toISOString().slice(0, 10);
      [['from', 'From date'], ['to', 'To date']].forEach(function (d) {
        var input = el('input', '', 'field');
        input.type = 'date';
        input.value = w[d[0]];
        input.addEventListener('change', function () { w[d[0]] = input.value; });
        body.appendChild(field(d[1], input));
      });
      return;
    }
    if (w.kind.rolling) {
      var s = select([['', 'Loading…']], '', function (v) { w.month = v; });
      body.appendChild(field(w.kind.id === 'balance' ? 'As at the end of' : 'Report month', s));
      body.appendChild(el('p', {
        balance: 'The balance sheet shows the position at the end of this month, the same month last year and last year end.',
        rolling12: 'The report shows the twelve months up to this month.',
        forecast: 'Actual up to this month, the forecast for the rest of the year.'
      }[w.kind.id] || (w.kind.dashboard ? 'The dashboard shows this month and the year to date up to it.'
        : 'The report shows this month and the year to date up to it, so one month is all it needs.')
        + ' To move it on later, pick another month in cell B8 of the report.', 'hint'));
      Promise.all([months(), TQ.period('CM', -1).catch(function () { return ''; })]).then(function (r) {
        s.textContent = '';
        r[0].forEach(function (p) { var x = el('option', p.name); x.value = p.name; s.appendChild(x); });
        w.month = w.month || r[1] || (r[0].length ? r[0][r[0].length - 1].name : '');
        s.value = w.month;
      }).catch(periodsFailed([s]));
      return;
    }
    // A whole fiscal year, or any run of months (from one month to another, across years if needed)
    w.span = w.span || 'year';
    body.appendChild(field('Period', select([['year', 'A whole fiscal year'], ['range', 'From one month to another']],
      w.span, function (v) { w.span = v; draw(); })));
    if (w.span === 'range') {
      var fromSel = select([['', 'Loading…']], '', function (v) { w.fromMonth = v; check(); });
      var toSel = select([['', 'Loading…']], '', function (v) { w.toMonth = v; check(); });
      var note = el('div', '', 'hint');
      var check = function () {
        var names = Array.prototype.map.call(fromSel.options, function (o) { return o.value; });
        var a = names.indexOf(w.fromMonth), b = names.indexOf(w.toMonth);
        note.textContent = a > b ? 'The From month is after the To month: they will be swapped.'
          : (b - a + 1) + ' month' + (b === a ? '' : 's') + '.';
      };
      body.appendChild(field('From month', fromSel));
      body.appendChild(field('To month', toSel));
      body.appendChild(note);
      Promise.all([months(), TQ.period('CM', -1).catch(function () { return ''; })]).then(function (r) {
        var list = r[0];
        [fromSel, toSel].forEach(function (s) {
          s.textContent = '';
          list.forEach(function (p) { var x = el('option', p.name); x.value = p.name; s.appendChild(x); });
        });
        var names = list.map(function (p) { return p.name; });
        var last = names.indexOf(r[1]) >= 0 ? r[1] : names[names.length - 1];
        w.toMonth = names.indexOf(w.toMonth) >= 0 ? w.toMonth : last;
        w.fromMonth = names.indexOf(w.fromMonth) >= 0 ? w.fromMonth : names[Math.max(0, names.indexOf(w.toMonth) - 11)];
        fromSel.value = w.fromMonth;
        toSel.value = w.toMonth;
        check();
      }).catch(periodsFailed([fromSel, toSel]));
    } else {
      var years = select([['', 'Loading…']], '', function (v) { w.year = v; });
      body.appendChild(field('Fiscal year', years));
      months().then(function (list) {
        var found = {};
        list.forEach(function (p) { var y = /(\d{4})/.exec(p.name); if (y) found[y[1]] = true; });
        var keys = Object.keys(found).sort();
        years.textContent = '';
        keys.forEach(function (y) { var x = el('option', y); x.value = y; years.appendChild(x); });
        w.year = w.year || String(new Date().getFullYear());
        if (keys.indexOf(w.year) === -1 && keys.length) w.year = keys[keys.length - 1];
        years.value = w.year;
      }).catch(periodsFailed([years]));
    }
    if (w.kind.id !== 'cashflow') {
      w.shape = w.shape || 'month';
      body.appendChild(field('Columns', select([['month', 'One per month'], ['quarter', 'One per quarter'], ['total', 'The year as one column']],
        w.shape, function (v) { w.shape = v; })));
      w.compare = w.kind.compare || w.compare || '';
      body.appendChild(field('Compare with', select([['', 'Nothing'], ['year', 'Last year'], ['budget', 'The budget']],
        w.compare, function (v) { w.compare = v; })));
    }
  }

  /* One filter box of the wizard: "All", or several values ticked in a searchable list (with or without
   * everything under them). Stored as the report's filter text: "Rooms, Food & Beverage*". */
  function multiPick(kind, values) {
    var box = el('div', '', 'mpick');
    var head = el('button', '', 'mpick-head');
    head.type = 'button';
    var panel = el('div', '', 'mpick-panel');
    panel.hidden = true;
    var search = el('input', '', 'field');
    search.type = 'search';
    search.placeholder = 'Search…';
    var list = el('div', '', 'pick tall');
    var under = el('label', '', 'check');
    var underBox = el('input');
    underBox.type = 'checkbox';
    under.appendChild(underBox);
    under.appendChild(document.createTextNode(' With everything under each one'));
    var current = String(w.filters[kind] || '');
    underBox.checked = /\*/.test(current);
    var chosen = current ? current.split(',').map(function (x) { return x.replace(/\*$/, '').trim(); }).filter(Boolean) : [];
    function store() {
      w.filters[kind] = chosen.map(function (x) { return x + (underBox.checked ? '*' : ''); }).join(', ');
      head.textContent = (chosen.length ? chosen.slice(0, 2).join(', ') + (chosen.length > 2 ? ' +' + (chosen.length - 2) : '') : 'All') + '  ▾';
    }
    function draw() {
      list.textContent = '';
      var q = search.value.trim().toLowerCase();
      values.filter(function (v) { return !q || v.toLowerCase().indexOf(q) !== -1; }).slice(0, 300).forEach(function (v) {
        var l = el('label', '', 'check');
        var c = el('input');
        c.type = 'checkbox';
        c.checked = chosen.indexOf(v) !== -1;
        c.addEventListener('change', function () {
          chosen = chosen.filter(function (x) { return x !== v; });
          if (c.checked) chosen.push(v);
          store();
        });
        l.appendChild(c);
        l.appendChild(document.createTextNode(' ' + v));
        list.appendChild(l);
      });
    }
    head.addEventListener('click', function () { panel.hidden = !panel.hidden; if (!panel.hidden) { draw(); search.focus(); } });
    search.addEventListener('input', draw);
    underBox.addEventListener('change', store);
    var clear = button('All (clear)', function () { chosen = []; store(); draw(); });
    panel.appendChild(search);
    panel.appendChild(list);
    panel.appendChild(under);
    panel.appendChild(clear);
    box.appendChild(head);
    box.appendChild(panel);
    store();
    return box;
  }

  function stepFilters(body) {
    body.appendChild(el('div', 'Leave a box on All for everything, or tick several values. You can change these on the report later.', 'hint'));
    ['subsidiary', 'department', 'class', 'location'].forEach(function (kind) {
      var holder = el('div');
      var f = field(kind.charAt(0).toUpperCase() + kind.slice(1), holder);
      holder.appendChild(el('div', 'Loading…', 'hint'));
      body.appendChild(f);
      TQ.callApi({ action: 'segments', kind: kind }).then(function (d) {
        holder.textContent = '';
        if (!d.segments.length) { f.hidden = true; return; }
        holder.appendChild(multiPick(kind, d.segments.map(function (x) { return x.name; })));
      }).catch(function () { f.hidden = true; });
    });
    var extra = el('div');
    body.appendChild(extra);
    TQ.callApi({ action: 'filterFields' }).then(function (d) {
      var fields = d.fields.filter(function (x) { return x.kind !== 'segment'; });
      if (!fields.length) return;
      extra.appendChild(el('label', 'Dropdown on the report for (up to 4)'));
      fields.forEach(function (x) {
        var l = el('label', '', 'check');
        var c = el('input');
        c.type = 'checkbox';
        c.checked = w.dropdowns.indexOf(x.label) !== -1;
        c.addEventListener('change', function () {
          if (c.checked && w.dropdowns.length >= 4) { c.checked = false; $('wMessage').textContent = 'Up to 4 dropdowns: untick one first.'; return; }
          w.dropdowns = w.dropdowns.filter(function (n) { return n !== x.label; });
          if (c.checked) w.dropdowns.push(x.label);
        });
        l.appendChild(c);
        l.appendChild(document.createTextNode(' ' + x.label));
        extra.appendChild(l);
      });
    }).catch(function () {});
  }

  function stepSheets(body) {
    body.appendChild(field('Besides the report, one sheet for each', select(
      [['', '(no extra sheets)'], ['location', 'Location'], ['department', 'Department'], ['class', 'Class'], ['subsidiary', 'Subsidiary']],
      w.pages, function (v) { w.pages = v; })));
    [['detail', 'List the GL accounts under each line'], ['comments', 'A Comments column (kept on rebuild)']].forEach(function (o) {
      var l = el('label', '', 'check');
      var c = el('input');
      c.type = 'checkbox';
      c.checked = !!w[o[0]];
      c.addEventListener('change', function () { w[o[0]] = c.checked; });
      l.appendChild(c);
      l.appendChild(document.createTextNode(' ' + o[1]));
      body.appendChild(l);
    });
    body.appendChild(field('Highlight change % over', select(
      [['', 'No highlights'], ['5', '5%'], ['10', '10%'], ['20', '20%'], ['25', '25%']], w.highlight, function (v) { w.highlight = v; })));
    body.appendChild(field('Ask for a comment when a change is at least (adds a Comments column)', select(
      [['', "Don't ask"], ['5', '5%'], ['10', '10%'], ['20', '20%']], w.commentAt || '', function (v) { w.commentAt = v; })));
  }

  function sentence() {
    var parts = [w.kind.title];
    if (w.kind.query) {
      return w.kind.title + (w.kind.dates ? ' from ' + w.from + ' to ' + w.to : '') + ', on a sheet of its own. Refresh it any time from Data table > Refresh data.';
    }
    if (w.kind.dashboard) {
      var f = Object.keys(w.filters).filter(function (k) { return w.filters[k]; }).map(function (k) { return k + ' ' + w.filters[k]; });
      return w.kind.title + ' for ' + (w.month || 'last month') + ', rolling forward from B8' + (f.length ? ', for ' + f.join(', ') : '') + '.';
    }
    if (w.kind.rolling) parts.push('for ' + (w.month || 'last month') + ', rolling forward from B8');
    else {
      parts.push(w.span === 'range' && w.fromMonth && w.toMonth ? 'from ' + w.fromMonth + ' to ' + w.toMonth : (w.year || ''));
      if (w.kind.id !== 'cashflow') parts.push({ month: 'by month', quarter: 'by quarter', total: 'as one column' }[w.shape || 'month']);
      if (w.compare === 'year') parts.push('against last year');
      if (w.compare === 'budget') parts.push('against the budget');
    }
    var filters = Object.keys(w.filters).filter(function (k) { return w.filters[k]; }).map(function (k) { return k + ' ' + w.filters[k]; });
    if (filters.length) parts.push('for ' + filters.join(', '));
    if (w.dropdowns.length) parts.push('with dropdowns for ' + w.dropdowns.join(', '));
    if (w.pages) parts.push('plus one sheet per ' + w.pages);
    if (w.detail) parts.push('listing the GL accounts');
    if (w.comments) parts.push('with a Comments column');
    if (w.highlight) parts.push('highlighting changes over ' + w.highlight + '%');
    if (w.commentAt) parts.push('asking for a comment on changes of ' + w.commentAt + '% or more');
    return parts.filter(Boolean).join(', ') + '.';
  }

  function stepCheck(body) {
    body.appendChild(el('p', sentence()));
    if (w.kind.query) return;   // a data report is named after its query
    var t = el('input', '', 'field');
    t.type = 'text';
    t.placeholder = 'Report title (optional)';
    t.value = w.title || '';
    t.addEventListener('input', function () { w.title = t.value; });
    body.appendChild(field('Title', t));
  }

  function create() {
    if (w.kind.query) {
      $('wMessage').textContent = 'Asking NetSuite…';
      return TQData.runTemplate(w.kind.query, w.kind.dates ? { from: w.from, to: w.to } : {}).then(function (data) {
        if (!data) throw new Error('NetSuite did not answer. Try again, or open Data table to see why.');
        TQPanel.say(data.reply ? String(data.reply).split('\n')[0] : w.kind.title + ' is on its own sheet.');
        window.TQNav.replace('home');
        showCard();
      }).catch(function (e) { $('wMessage').textContent = e.message; $('wMessage').className = 'status-line bad'; });
    }
    var args = {
      template: w.kind.template || w.kind.id, title: w.title || '', month: w.month || '', year: w.year || '',
      from: w.span === 'range' ? w.fromMonth || '' : '', to: w.span === 'range' ? w.toMonth || '' : '',
      shape: w.shape || 'month', compare: w.compare || '', filters: w.filters, dropdowns: w.dropdowns,
      pages: w.pages, detail: w.detail, comments: w.comments || !!w.commentAt, highlight: w.highlight, commentAt: w.commentAt || ''
    };
    if (w.building) return;   // a second click while the first report is being built
    w.building = true;
    var createButton = document.querySelector('#wizard .btn.primary');
    if (createButton) createButton.disabled = true;
    $('wMessage').textContent = w.kind.dashboard ? 'Building the dashboard…' : 'Building the report…';
    if (w.kind.dashboard) args = { kind: w.kind.dashboard, month: w.month || '', filters: w.filters, title: w.title || '' };
    window.TQChat.workbookContext().then(function (context) {
      return fetch('/api/design', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ op: w.kind.dashboard ? 'dashboard' : 'wizard', args: args, context: context }) });
    }).then(function (r) { return r.json(); }).then(function (data) {
      if (!data.ok) throw new Error(data.error || 'Request failed');
      if (data.spec) window.TQChat.setSpec(data.spec);
      return window.TQChat.applyActions(data.actions).then(function () {
        var lines = String(data.reply || '').split('\n');
        // later lines are the checks (accounts missing from the layout, lines with no account): keep them on screen
        if (lines.length > 1) TQPanel.say(lines.join(' '), false, { label: 'OK', run: function () { TQPanel.say(''); } });
        else TQPanel.say(lines[0]);
        if (w.kind.dashboard && data.actions && data.actions[0] && window.TQPack) TQPack.placeLogo(data.actions[0].sheetName);
        if (data.design && window.TQDesigner) TQDesigner.select(data.design.name);
        window.TQNav.replace('report');
        showCard();
      });
    }).catch(function (e) {
      $('wMessage').textContent = e.message; $('wMessage').className = 'status-line bad';
    }).then(function () {
      w.building = false;
      var b = document.querySelector('#wizard .btn.primary');
      if (b) b.disabled = false;
    });
  }

  // ------------------------------------------------------------------ the ribbon --

  function viaPanel(then) {
    return function (event) {
      var done = function () { if (event && event.completed) event.completed(); };
      var show = (window.Office && Office.addin && Office.addin.showAsTaskpane) ? Office.addin.showAsTaskpane() : Promise.resolve();
      Promise.resolve(show).then(function () { then(); done(); }, function () { then(); done(); });
    };
  }
  function direct(run) {
    return function (event) {
      try { run(); } finally { if (event && event.completed) event.completed(); }
    };
  }

  function openDashboard(kind) {
    openWizard();
    w.kind = KINDS.filter(function (x) { return x.dashboard === kind; })[0];
    w.step = 1;
    draw();
  }

  function start() {
    if (!$('homeCard')) return;
    startCharts();
    $('homeNew').addEventListener('click', openWizard);
    $('welcomeStart').addEventListener('click', function () { setWelcomed(); window.TQNav.replace('home'); openWizard(); });
    $('welcomeSkip').addEventListener('click', function () { setWelcomed(); window.TQNav.replace('home'); showCard(); });
    $('menuWelcome').addEventListener('click', welcome);
    $('hMonth').addEventListener('change', function () { setMonth($('hMonth').value); });
    $('rpPrev').addEventListener('click', function () { stepMonth(-1); });
    $('rpNext').addEventListener('click', function () { stepMonth(1); });
    $('rpRestore').addEventListener('click', function () {
      if ($('rpRestore').textContent !== 'Restore?') {   // a second click confirms (Office panels have no confirm())
        $('rpRestore').textContent = 'Restore?';
        TQPanel.say('Click Restore? again to put back the version from before the last rebuild of this sheet.');
        setTimeout(function () { $('rpRestore').textContent = 'Restore'; }, 5000);
        return;
      }
      $('rpRestore').textContent = 'Restore';
      TQChat.restoreSheet(currentSheet).then(function (msg) { TQPanel.say(msg); }, function (e) { TQPanel.say(e.message, true); });
    });
    $('rpEdit').addEventListener('click', function () {
      var open = function () { window.TQNav.go('layout'); };
      if (reportLayout && window.TQDesigner) TQDesigner.select(reportLayout).then(open, open); else open();
    });
    if (inExcel()) loadStamps();
    if (window.TQ && TQ.onChange) TQ.onChange(function (what) { if (what && what.sheets && what.sheets.length) stampSheets(what.sheets); });
    if (!welcomed()) welcome(); else showCard();
    if (inExcel()) {
      try {
        Excel.run(function (context) {
          context.workbook.worksheets.onActivated.add(function () { showCard(); });
          context.workbook.worksheets.onAdded.add(function () { listReports(); });
          context.workbook.worksheets.onDeleted.add(function () { listReports(); });
          return context.sync();
        }).catch(function () {});
      } catch (e) { /* an Excel without sheet events: the screens update when the panel opens */ }
    }
    if (window.Office && Office.actions && Office.actions.associate) {
      Office.actions.associate('newReport', viaPanel(openWizard));
      Office.actions.associate('showPanel', viaPanel(function () {}));
      Office.actions.associate('newDashboard', viaPanel(function () { window.TQNav.go('dash'); }));
      Office.actions.associate('openChat', viaPanel(function () { TQChat.showTab('Chat'); }));
      Office.actions.associate('refreshAll', direct(function () { TQPanel.refreshAll(); }));
      Office.actions.associate('refreshSheet', direct(function () { TQPanel.refreshSheet(); }));
      Office.actions.associate('refreshSelection', direct(function () { TQPanel.refreshSelection(); }));
      Office.actions.associate('drillDown', viaPanel(function () { TQPanel.drillSelected(); }));
      Office.actions.associate('insertFormula', viaPanel(function () { openSection('secFormula'); }));
      Office.actions.associate('copyValues', viaPanel(function () { openSection('secPack'); TQPack.copyAsValues(); }));
      Office.actions.associate('setupPdf', viaPanel(function () { openSection('secPack'); TQPack.setUpPrint(); }));
    }
  }

  /* Chart the selection: an Excel chart of the selected cells (first row = names, first column = labels), placed
   * to the right of them. Works on any sheet; on a TQ report the chart follows its live numbers. */
  function chartSelection() {
    var msg = $('cMessage');
    if (!inExcel()) { msg.textContent = 'Open the panel in Excel and select the cells to chart.'; return; }
    var type = $('cChartType').value;
    Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      range.load('address,rowCount,columnCount,top,left,width');
      return context.sync().then(function () {
        if (range.rowCount < 2 && range.columnCount < 2) throw new Error('Select a block of cells: names in the first row or column, numbers in the rest.');
        var chart = sheet.charts.add(type, range, 'Auto');
        chart.title.text = $('cChartTitle').value || 'Chart';
        chart.left = range.left + range.width + 20;
        chart.top = range.top;
        chart.width = 420;
        chart.height = 260;
        chart.legend.position = 'Bottom';
        return context.sync().then(function () { return range.address; });
      });
    }).then(function (address) { msg.textContent = 'Chart added for ' + address.split('!').pop() + '. Move or resize it like any Excel chart.'; },
      function (e) { msg.textContent = e.message; });
  }

  function startCharts() {
    if (!$('cChart')) return;
    $('cChart').addEventListener('click', chartSelection);
    ['executive', 'sales', 'cash', 'pl', 'position', 'budget'].forEach(function (k) {
      var b = $('dash-' + k);
      if (b) b.addEventListener('click', function () { openDashboard(k); });
    });
  }

  /* Straight to one kind of report (quick search): the wizard opens at its second step. */
  function openKind(id) {
    var kind = KINDS.filter(function (k) { return k.id === id; })[0];
    if (!kind) return;
    openWizard();
    if (kind.id === 'own') { stepWhat($('wizard')); return; }
    w.kind = kind;
    w.step = 1;
    draw();
  }

  window.TQHome = { openWizard: openWizard, showCard: showCard, welcome: welcome, openKind: openKind, openSheet: openSheet,
    kinds: function () { return KINDS.filter(function (k) { return k.id !== 'own'; }); },
    reports: function () { return known.slice(); } };

  if (window.Office && Office.onReady) Office.onReady(function () { start(); });
  else if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
