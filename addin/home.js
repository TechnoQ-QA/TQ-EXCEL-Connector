/* TQ Excel Connector - the top of the Reports tab and the ribbon.
 *   - a card that follows the open sheet: on a report its month picker, refresh, edit, filters and send; on a
 *     layout, budget or journal sheet what that sheet needs; elsewhere New report.
 *   - the report wizard: one question per step, then one request that builds a layout report (so every report
 *     it makes can be changed and rebuilt in place).
 *   - the ribbon tab's buttons (Office.actions), which run here because the add-in shares one runtime. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var LAYOUT_RE = /layout sheet "([^"]+)"/;
  var periods = null;   // [{name}] from NetSuite, loaded once
  var notice = '';      // the last wizard result, shown on the card until another sheet is opened

  function el(tag, text, cls) {
    var n = document.createElement(tag);
    if (text) n.textContent = text;
    if (cls) n.className = cls;
    return n;
  }
  function button(text, onClick, cls) {
    var b = el('button', text, cls);
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
  function openSection(id) {
    var s = $(id);
    if (!s) return;
    s.open = true;
    s.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ------------------------------------------------------------------ the card --

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

  function card(title, lines, buttons) {
    var box = $('homeCard');
    box.textContent = '';
    box.appendChild(el('h3', title));
    if (notice) box.appendChild(el('div', notice, 'status-line'));
    lines.forEach(function (l) { if (l) box.appendChild(typeof l === 'string' ? el('div', l, 'hint') : l); });
    var row = el('div', '', 'row');
    row.style.marginTop = '6px';
    buttons.forEach(function (b) { row.appendChild(b); });
    box.appendChild(row);
  }

  function monthPicker(current) {
    var wrap = el('div');
    var label = el('label', 'Month (rolls every column)');
    label.setAttribute('for', 'hMonth');
    var select = el('select', '', 'field');
    select.id = 'hMonth';
    var o = el('option', current);
    o.value = current;
    select.appendChild(o);
    months().then(function (list) {
      select.textContent = '';
      list.forEach(function (p) { var x = el('option', p.name); x.value = p.name; select.appendChild(x); });
      select.value = current;
    }).catch(function () {});
    select.addEventListener('change', function () {
      Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange('B8').values = [[select.value]];
        sheet.calculate(true);
        return context.sync();
      }).catch(function () {});
    });
    wrap.appendChild(label);
    wrap.appendChild(select);
    return wrap;
  }

  function showCard() {
    return readActive().then(function (s) {
      var a = s.a || [];
      var cell = function (r, c) { return a[r] && a[r][c] !== undefined ? String(a[r][c]) : ''; };
      var newReport = button('New report', openWizard, 'primary');
      var wand = el('img', '', 'ico');
      wand.src = 'assets/icon-wizard-16.png';
      wand.alt = '';
      newReport.insertBefore(wand, newReport.firstChild);
      var layout = LAYOUT_RE.exec(cell(1, 0));
      if (layout || cell(2, 0) === 'Other filters') {
        var month = cell(7, 0) === 'Month' ? cell(7, 1) : '';
        var buttons = [button('Refresh sheet', function () { TQPanel.refreshSheet(); }),
          button('Filters', function () { openSection('secFilters'); })];
        if (layout) {
          buttons.push(button('Edit layout', function () {
            TQDesigner.select(layout[1]).then(function () { openSection('secLayout'); });
          }));
        }
        buttons.push(button('Send', function () { openSection('secPack'); }), newReport);
        card(cell(0, 0) || s.name, [layout ? 'From layout "' + layout[1] + '": changes there rebuild it in place.'
          : 'A standard report: ask the chat for changes, or use New report for one you can edit.',
        month ? monthPicker(month) : null], buttons);
      } else if (/^tq layout/i.test(s.name)) {
        card('Layout sheet', ['Change lines and columns here or in Report layout, then update the report.'],
          [button('Update report', function () { TQDesigner.select(s.name).then(function () { TQDesigner.build(); }); }, 'primary'),
            button('Report layout', function () { TQDesigner.select(s.name).then(function () { openSection('secLayout'); }); })]);
      } else if (/^tq budget/i.test(s.name)) {
        card('Budget sheet', ['Change the numbers, then Check (writes nothing) and Upload.'],
          [button('Check and upload', function () { openSection('secBudget'); }, 'primary')]);
      } else if (/^tq journals/i.test(s.name)) {
        card('Journal sheet', ['Rows with the same Entry number become one journal entry.'],
          [button('Check and upload', function () { openSection('secJournal'); }, 'primary')]);
      } else {
        card('TQ Connector', ['Build a report step by step, or put a NetSuite number in this cell.'],
          [newReport, button('Insert formula', function () { openSection('secFormula'); })]);
      }
    });
  }

  // ----------------------------------------------------------------- the wizard --

  var KINDS = [
    { id: 'monthytd', title: 'Monthly P&L', about: 'The month and the year to date against last year. Rolls forward with one cell.', rolling: true },
    { id: 'summary', title: 'P&L summary', about: 'One line per group (sales, cost of sales, expenses...), by month or quarter.' },
    { id: 'management', title: 'P&L, account by account', about: 'Every account with a balance, grouped, with margins.' },
    { id: 'budget', title: 'Budget against actual', about: 'Actual, budget, variance and variance % per period.', template: 'management', compare: 'budget' },
    { id: 'balance', title: 'Balance sheet', about: 'As at a month, the same month last year and last year end. Rolls forward.', rolling: true },
    { id: 'cashflow', title: 'Cash flow', about: 'Indirect method, month by month, checked against the bank accounts.' },
    { id: 'dash-executive', title: 'Dashboard: executive', about: 'Tiles for sales, gross profit, net income and cash; charts of sales by month, by location and where the money went.', rolling: true, dashboard: 'executive' },
    { id: 'dash-sales', title: 'Dashboard: sales', about: 'Sales this month and year to date, sales by month against last year, gross margin, sales by location.', rolling: true, dashboard: 'sales' },
    { id: 'dash-cash', title: 'Dashboard: cash', about: 'Cash, receivables and payables by month, cash movement and the bank accounts.', rolling: true, dashboard: 'cash' },
    { id: 'dash-pl', title: 'Finance: profit and loss', about: 'Revenue, gross and net margin, operating expenses; revenue, costs and net income by month; a bridge from revenue to net income.', rolling: true, dashboard: 'pl' },
    { id: 'dash-position', title: 'Finance: financial position', about: 'Working capital, current and quick ratio, debt to equity, debtor and creditor days; what the company owns and owes.', rolling: true, dashboard: 'position' },
    { id: 'dash-budget', title: 'Finance: budget against actual', about: 'Revenue, costs and net income against the budget, by month and so far this year.', rolling: true, dashboard: 'budget' },
    { id: 'own', title: 'My own format', about: 'Typed on a sheet: its rows become the report lines.' }
  ];
  var STEPS = ['What', 'Periods', 'Filters', 'Sheets', 'Check'];
  var w = null;   // the answers so far

  function openWizard() {
    w = { step: 0, kind: KINDS[0], filters: {}, dropdowns: [], pages: '', detail: false, comments: false, highlight: '' };
    $('wizard').hidden = false;
    $('homeCard').hidden = true;
    draw();
    $('wizard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function closeWizard() {
    $('wizard').hidden = true;
    $('homeCard').hidden = false;
  }

  function field(labelText, control) {
    var box = el('div');
    var label = el('label', labelText);
    box.appendChild(label);
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
    box.appendChild(el('h3', 'New report'));
    box.appendChild(el('div', STEPS.map(function (s, n) { return (n === w.step ? '● ' : '○ ') + s; }).join('   '), 'steps'));
    var body = el('div');
    box.appendChild(body);
    [stepWhat, stepPeriods, stepFilters, stepSheets, stepCheck][w.step](body);
    var row = el('div', '', 'row');
    row.style.marginTop = '10px';
    if (w.step > 0) row.appendChild(button('Back', function () { w.step -= (w.kind.dashboard && w.step === 4) ? 2 : 1; draw(); }));
    if (w.step < STEPS.length - 1) {
      row.appendChild(button('Next', function () {
        w.step += 1;
        if (w.kind.dashboard && w.step === 3) w.step = 4;   // a dashboard has no extra sheets to choose
        draw();
      }, 'primary'));
    }
    else row.appendChild(button('Create the report', create, 'primary'));
    row.appendChild(button('Cancel', closeWizard));
    box.appendChild(row);
    var msg = el('div', '', 'status-line');
    msg.id = 'wMessage';
    box.appendChild(msg);
  }

  function stepWhat(body) {
    var list = el('div', '', 'choices');
    KINDS.forEach(function (k) {
      var b = button('', function () {
        w.kind = k;
        if (k.id === 'own') {   // made from a sheet: the Report layout section does that
          closeWizard();
          $('dTemplate').value = '(sheet)';
          $('dTemplate').dispatchEvent(new Event('change'));
          openSection('secLayout');
          return;
        }
        w.step = 1;
        draw();
      }, w.kind === k ? 'on' : '');
      b.appendChild(el('strong', k.title));
      b.appendChild(el('small', k.about));
      list.appendChild(b);
    });
    body.appendChild(list);
  }

  function stepPeriods(body) {
    if (w.kind.rolling) {
      var s = select([['', 'Loading…']], '', function (v) { w.month = v; });
      body.appendChild(field('Report month (change it later in B8 to roll forward)', s));
      Promise.all([months(), TQ.period('CM', -1).catch(function () { return ''; })]).then(function (r) {
        s.textContent = '';
        r[0].forEach(function (p) { var x = el('option', p.name); x.value = p.name; s.appendChild(x); });
        w.month = w.month || r[1] || (r[0].length ? r[0][r[0].length - 1].name : '');
        s.value = w.month;
      });
      return;
    }
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
    });
    if (w.kind.id !== 'cashflow') {
      w.shape = w.shape || 'month';
      body.appendChild(field('Columns', select([['month', 'One per month'], ['quarter', 'One per quarter'], ['total', 'The year as one column']],
        w.shape, function (v) { w.shape = v; })));
      w.compare = w.kind.compare || w.compare || '';
      body.appendChild(field('Compare with', select([['', 'Nothing'], ['year', 'Last year'], ['budget', 'The budget']],
        w.compare, function (v) { w.compare = v; })));
    }
  }

  function stepFilters(body) {
    body.appendChild(el('div', 'Leave a box on All for everything. You can change these on the report later (B3:B7).', 'hint'));
    ['subsidiary', 'department', 'class', 'location'].forEach(function (kind) {
      var s = select([['', 'All']], w.filters[kind] || '', function (v) { w.filters[kind] = v; });
      var f = field(kind.charAt(0).toUpperCase() + kind.slice(1), s);
      body.appendChild(f);
      TQ.callApi({ action: 'segments', kind: kind }).then(function (d) {
        d.segments.forEach(function (x) { var o = el('option', x.name); o.value = x.name; s.appendChild(o); });
        s.value = w.filters[kind] || '';
        if (!d.segments.length) f.hidden = true;
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
  }

  function sentence() {
    var parts = [w.kind.title];
    if (w.kind.dashboard) {
      var f = Object.keys(w.filters).filter(function (k) { return w.filters[k]; }).map(function (k) { return k + ' ' + w.filters[k]; });
      return w.kind.title + ' for ' + (w.month || 'last month') + ', rolling forward from B8' + (f.length ? ', for ' + f.join(', ') : '') + '.';
    }
    if (w.kind.rolling) parts.push('for ' + (w.month || 'last month') + ', rolling forward from B8');
    else {
      parts.push(w.year || '');
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
    return parts.filter(Boolean).join(', ') + '.';
  }

  function stepCheck(body) {
    body.appendChild(el('p', sentence()));
    var t = el('input', '', 'field');
    t.type = 'text';
    t.placeholder = 'Report title (optional)';
    t.value = w.title || '';
    t.addEventListener('input', function () { w.title = t.value; });
    body.appendChild(field('Title', t));
  }

  function create() {
    var args = {
      template: w.kind.template || w.kind.id, title: w.title || '', month: w.month || '', year: w.year || '',
      shape: w.shape || 'month', compare: w.compare || '', filters: w.filters, dropdowns: w.dropdowns,
      pages: w.pages, detail: w.detail, comments: w.comments, highlight: w.highlight
    };
    $('wMessage').textContent = w.kind.dashboard ? 'Building the dashboard…' : 'Building the report…';
    if (w.kind.dashboard) args = { kind: w.kind.dashboard, month: w.month || '', filters: w.filters, title: w.title || '' };
    window.TQChat.workbookContext().then(function (context) {
      return fetch('/api/design', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ op: w.kind.dashboard ? 'dashboard' : 'wizard', args: args, context: context }) });
    }).then(function (r) { return r.json(); }).then(function (data) {
      if (!data.ok) throw new Error(data.error || 'Request failed');
      if (data.spec) window.TQChat.setSpec(data.spec);
      return window.TQChat.applyActions(data.actions).then(function () {
        closeWizard();
        notice = data.reply.split('\n')[0];
        if (w.kind.dashboard && data.actions && data.actions[0] && window.TQPack) TQPack.placeLogo(data.actions[0].sheetName);
        if (data.design && window.TQDesigner) TQDesigner.select(data.design.name);
        showCard();
      });
    }).catch(function (e) { $('wMessage').textContent = e.message; $('wMessage').className = 'status-line bad'; });
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

  function start() {
    if (!$('homeCard')) return;
    startCharts();
    showCard();
    if (inExcel()) {
      try {
        Excel.run(function (context) {
          context.workbook.worksheets.onActivated.add(function () { notice = ''; if ($('wizard').hidden) showCard(); });
          return context.sync();
        }).catch(function () {});
      } catch (e) { /* an Excel without sheet events: the card updates when the panel opens */ }
    }
    if (window.Office && Office.actions && Office.actions.associate) {
      Office.actions.associate('newReport', viaPanel(openWizard));
      Office.actions.associate('newDashboard', viaPanel(function () {
        openWizard();
        w.kind = KINDS.filter(function (x) { return x.dashboard === 'executive'; })[0];
        w.step = 0;
        draw();
      }));
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
      if (b) b.addEventListener('click', function () {
        openWizard();
        w.kind = KINDS.filter(function (x) { return x.dashboard === k; })[0];
        w.step = 1;
        draw();
      });
    });
  }

  window.TQHome = { openWizard: openWizard, showCard: showCard };

  if (window.Office && Office.onReady) Office.onReady(function () { start(); });
  else if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
