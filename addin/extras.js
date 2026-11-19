/* TQ Excel Connector - helpers that make the panel easier to use:
 *   - saved filter sets: name the filters of a report ("Riyadh hotels") and put them on any report later;
 *   - collect comments: every comment typed on the reports, and the big changes still without one, on one sheet;
 *   - help: a short tip for the screen you are on (the ? button), and a tour of a report sheet;
 *   - quick search on Home: type a few letters to find a report to make, an action or a sheet;
 *   - move to another PC: export the setup to one file, import it on the new computer. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  function el(tag, text, cls) {
    var n = document.createElement(tag);
    if (text) n.textContent = text;
    if (cls) n.className = cls;
    return n;
  }
  function inExcel() { return typeof Excel !== 'undefined' && window.Office && Office.context && Office.context.document; }
  function say(text, bad, action) { if (window.TQPanel) TQPanel.say(text, bad, action); }
  function post(path, body) {
    return fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) })
      .then(function (r) { return r.json(); })
      .then(function (d) { if (!d || !d.ok) throw new Error((d && d.error) || 'Request failed'); return d; });
  }

  // ------------------------------------------------------------ saved filter sets --

  var sets = [];
  function drawSets() {
    var pick = $('fsPick');
    if (!pick) return;
    pick.textContent = '';
    var first = el('option', sets.length ? 'Saved filter sets (' + sets.length + ')…' : 'No saved filter sets yet');
    first.value = '';
    pick.appendChild(first);
    sets.forEach(function (s) {
      var o = el('option', s.name + ': ' + Object.keys(s.cells).map(function (k) { return s.cells[k]; }).join(' · '));
      o.value = s.name;
      pick.appendChild(o);
    });
  }
  function loadSets() {
    return fetch('/api/prefs').then(function (r) { return r.json(); }).then(function (d) { sets = d.filterSets || []; drawSets(); }).catch(function () {});
  }
  function applySet() {
    var name = $('fsPick').value;
    var set = sets.filter(function (s) { return s.name === name; })[0];
    if (!set) { say('Pick a saved filter set first.'); return; }
    var cells = { subsidiary: '', department: '', class: '', location: '', filters: '' };
    Object.keys(set.cells).forEach(function (k) { cells[k] = set.cells[k]; });
    TQFilters.write(cells).then(function () { say('Filters "' + set.name + '" put on the report. The numbers are being refreshed.'); },
      function (e) { say(e.message, true); });
  }
  function saveSet() {
    var name = $('fsName').value.trim();
    if (!name) { say('Give the filter set a name.', true); return; }
    TQFilters.read().then(function (cells) {
      return post('/api/prefs', { op: 'saveSet', name: name, cells: cells });
    }).then(function (d) {
      sets = d.filterSets || [];
      drawSets();
      $('fsName').value = '';
      $('fsSaveRow').hidden = true;
      say('Saved "' + name + '". Pick it on any report to put the same filters there.');
    }).catch(function (e) { say(e.message, true); });
  }

  // ------------------------------------------------------------- collect comments --

  /* Reads every visible sheet with a Comments column and writes "TQ Commentary": the line, its change %, the
   * comment, and whether a change of at least the chosen size still needs one. */
  function collectComments() {
    if (!inExcel()) { say('Open the panel in Excel to collect comments.'); return; }
    var limit = Number($('rpCollectAt').value || 10) / 100;
    say('Collecting the comments…');
    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name,items/visibility');
      return context.sync().then(function () {
        var shown = sheets.items.filter(function (s) { return s.visibility === 'Visible' && s.name !== 'TQ Commentary'; });
        var used = shown.map(function (s) { var u = s.getUsedRangeOrNullObject(true); u.load('values'); return u; });
        return context.sync().then(function () {
          var out = [];
          shown.forEach(function (s, n) {
            if (used[n].isNullObject) return;
            var v = used[n].values;
            var head = -1, col = -1;
            for (var r = 0; r < Math.min(v.length, 40) && head < 0; r++) {
              var at = v[r].indexOf('Comments');
              if (at !== -1) { head = r; col = at; }
            }
            if (head < 0) return;
            var pct = v[head].findIndex(function (h) { return /%/.test(String(h)); });
            for (var i = head + 1; i < v.length; i++) {
              var line = String(v[i][1] || '').trim();
              if (!line) continue;
              var note = String(v[i][col] || '').trim();
              var change = pct >= 0 && typeof v[i][pct] === 'number' ? v[i][pct] : '';
              var big = change !== '' && Math.abs(change) >= limit;
              if (note || big) out.push([s.name, line, change, note, note ? 'Explained' : 'Comment needed']);
            }
          });
          var ws = context.workbook.worksheets.getItemOrNullObject('TQ Commentary');
          return context.sync().then(function () {
            if (ws.isNullObject) ws = context.workbook.worksheets.add('TQ Commentary');
            else ws.getRange().clear();
            var rows = [['Commentary', '', '', '', ''], ['Comments typed on the reports, and changes of ' + Math.round(limit * 100) + '% or more still without one. Made ' + new Date().toLocaleString() + '.', '', '', '', ''], ['', '', '', '', ''],
              ['Sheet', 'Line', 'Change %', 'Comment', 'Status']].concat(out.length ? out : [['', 'Nothing to show: no comments and no big changes.', '', '', '']]);
            var range = ws.getRange('A1').getResizedRange(rows.length - 1, 4);
            range.values = rows;
            ws.getRange('A1').format.font.set({ bold: true, size: 14 });
            ws.getRange('A2').format.font.color = '#605e5c';
            var head = ws.getRange('A4:E4');
            head.format.font.bold = true;
            head.format.fill.color = '#1F3A5F';
            head.format.font.color = '#FFFFFF';
            if (out.length) {
              ws.getRange('C5').getResizedRange(out.length - 1, 0).numberFormat = [['0.0%']];
              out.forEach(function (row, i) {
                if (row[4] === 'Comment needed') ws.getRange('E' + (5 + i)).format.fill.color = '#FFF4CE';
              });
            }
            ws.getRange('A:A').format.columnWidth = 110;
            ws.getRange('B:B').format.columnWidth = 220;
            ws.getRange('C:C').format.columnWidth = 70;
            ws.getRange('D:D').format.columnWidth = 320;
            ws.getRange('D:D').format.wrapText = true;
            ws.getRange('E:E').format.columnWidth = 110;
            ws.activate();
            return context.sync().then(function () { return out; });
          });
        });
      });
    }).then(function (out) {
      var needed = out.filter(function (r) { return r[4] === 'Comment needed'; }).length;
      say('Commentary sheet ready: ' + (out.length - needed) + ' comment(s)' + (needed ? ', ' + needed + ' big change(s) still need one.' : '.'));
    }, function (e) { say(e.message, true); });
  }

  // ------------------------------------------------------------------------ help --

  var TIPS = {
    home: 'Start with New report, or type in the search box. The tiles open the other tools. The reports already in this workbook are listed below: click one to open it.',
    report: 'This is the report on the open sheet. Change the month with ‹ ›, add filters with + Add filter, and save filters you use often. Refresh brings the latest numbers. Restore puts back the sheet as it was before the last rebuild.',
    filter: 'Pick a field at the top, tick values (or type them), then Apply. "With everything under it" includes the sub-departments or sub-locations. Keep in the layout makes the filter part of the report for good.',
    layout: 'A layout is the recipe of a report: its lines (rows) and columns. Change it here or on its TQ Layout sheet, then Update report. Save to the library shares it with your whole company.',
    wizard: 'Pick what you need, then answer a few questions. Every report it makes is a layout, so you can change it later. Hospitality and Planning reports roll forward with the month cell.',
    dash: 'A dashboard is a sheet of tiles and charts over live NetSuite numbers. It has the same month cell (B8) and filters as a report.',
    send: 'Set up for PDF puts your logo, the company name and page numbers on the pages; then save as PDF in Excel. Copy as values makes a version for people without the add-in.',
    upload: 'Make the sheet, fill it in Excel, then Check: it writes nothing and lists anything wrong. Upload only works after a clean check.',
    formulas: 'Pick an account or total, a period and filters; the formula goes into the selected cell. =TQ.BALANCE works anywhere in your own sheets too.',
    data: 'Data tables bring any NetSuite records (customers, invoices, saved searches) onto their own sheets. Refresh data runs them all again.',
    drill: 'The transactions behind a number. Open large shows them in a big window with sorting and search; Send to a new sheet copies them.',
    chat: 'Ask in plain words: "P&L for this year by quarter", "balance sheet last month vs last year", then refine: "by department", "add column Q1 2025".',
    transfer: 'Export on the old PC, copy the file (USB, email, OneDrive), Import on the new PC. Passwords and keys are never in the file: sign in with NetSuite again on the new PC.',
    keys: 'These shortcuts work while you are in the Excel grid. The same commands are on the right-click menu of a cell.',
    welcome: 'These checks run every time the panel opens; this screen only shows the first time.'
  };
  function toggleHelp(force) {
    var box = $('tipBox');
    var open = force === undefined ? box.hidden : force;
    box.hidden = !open;
    $('helpButton').setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) showTip();
  }
  function showTip() {
    var box = $('tipBox');
    if (box.hidden) return;
    var name = window.TQNav ? TQNav.current() : 'home';
    box.textContent = '';
    box.appendChild(el('span', TIPS[name] || TIPS.home, 'grow'));
    var x = el('button', '✕', 'x');
    x.setAttribute('aria-label', 'Close help');
    x.addEventListener('click', function () { toggleHelp(false); });
    box.appendChild(x);
  }

  /* A tour of the report on the open sheet: selects each part in Excel and says what it does. */
  function tour() {
    if (!inExcel()) { say('Open the panel in Excel for the tour.'); return; }
    Excel.run(function (context) {
      var used = context.workbook.worksheets.getActiveWorksheet().getUsedRange(true);
      used.load('formulas,rowIndex,columnIndex');
      return context.sync().then(function () {
        var f = used.formulas, cell = '';
        for (var r = 0; r < f.length && !cell; r++) {
          for (var c = 0; c < f[r].length && !cell; c++) {
            if (/TQ\.BALANCE/i.test(String(f[r][c]))) cell = String.fromCharCode(65 + used.columnIndex + c) + (used.rowIndex + r + 1);
          }
        }
        return cell;
      });
    }).then(function (numberCell) {
      var steps = [
        ['B8', 'Step 1 of 4: B8 is the report month. Change it (or use ‹ › in the panel) and every column rolls with it.'],
        ['B3:B7', 'Step 2 of 4: B3 to B7 are the filters: other filters, subsidiary, department, class and location. Type or pick values; the numbers follow.'],
        [numberCell || 'C12', 'Step 3 of 4: each number is a live NetSuite formula. Right-click it > TQ Connector > Drill down to see the transactions behind it.'],
        ['', 'Step 4 of 4: Refresh (in the panel, the ribbon or Ctrl+Alt+Shift+R) brings the latest numbers. To change lines or columns, open Lines, columns and options.']
      ];
      var go = function (n) {
        var s = steps[n];
        if (s[0]) {
          Excel.run(function (context) { context.workbook.worksheets.getActiveWorksheet().getRange(s[0]).select(); return context.sync(); }).catch(function () {});
        }
        say(s[1], false, n + 1 < steps.length ? { label: 'Next', run: function () { go(n + 1); } } : { label: 'Done', run: function () { say(''); } });
      };
      go(0);
    }).catch(function (e) { say(e.message, true); });
  }

  // --------------------------------------------------------------- quick search --

  function commands() {
    var list = [
      ['New report', 'Start the report wizard', function () { TQHome.openWizard(); }],
      ['Refresh all', 'Every TQ number in the workbook', function () { TQPanel.refreshAll(); }],
      ['Refresh this sheet', 'Only the open sheet', function () { TQPanel.refreshSheet(); }],
      ['Dashboard', 'Tiles and charts', function () { TQNav.go('dash'); }],
      ['Send / PDF', 'Logo, values copy, PDF pages', function () { TQNav.go('send'); }],
      ['Upload budget or journal', 'From a sheet to NetSuite', function () { TQNav.go('upload'); }],
      ['Insert a formula', 'One NetSuite number in a cell', function () { TQNav.go('formulas'); }],
      ['Data tables', 'Customers, invoices, saved searches', function () { TQNav.go('data'); }],
      ['Report layout', 'Lines, columns, options, library', function () { TQNav.go('layout'); }],
      ['Company library', 'Layouts shared in your company', function () { TQNav.go('layout'); var f = $('libFold'); if (f) f.open = true; }],
      ['Drill down', 'The transactions behind the selected cell', function () { TQPanel.drillSelected(); }],
      ['Demo pack', 'Three example sheets', function () { TQLibrary.demoPack(); }],
      ['Collect comments', 'All comments onto one sheet', collectComments],
      ['Move to another PC', 'Export or import your setup', function () { TQNav.go('transfer'); }],
      ['Keyboard shortcuts', 'Ctrl+Alt+Shift+…', function () { TQNav.go('keys'); }],
      ['Chat', 'Ask in your own words', function () { TQChat.showTab('Chat'); }],
      ['Setup', 'Connection, licence, logo, seats, diagnostics', function () { location.href = 'setup.html'; }]
    ];
    (window.TQHome && TQHome.kinds ? TQHome.kinds() : []).forEach(function (k) {
      list.push([k.title, 'New report: ' + k.group, function () { TQHome.openKind(k.id); }]);
    });
    (window.TQHome && TQHome.reports ? TQHome.reports() : []).forEach(function (r) {
      list.push([r.title, 'Open sheet ' + r.name, function () { TQHome.openSheet(r.name); }]);
    });
    return list;
  }
  var hits = [];
  function search() {
    var q = $('qsInput').value.trim().toLowerCase();
    var box = $('qsList');
    box.textContent = '';
    if (!q) { box.hidden = true; return; }
    var words = q.split(/\s+/);
    hits = commands().filter(function (c) {
      var hay = (c[0] + ' ' + c[1]).toLowerCase();
      return words.every(function (w) { return hay.indexOf(w) !== -1; });
    }).slice(0, 8);
    if (!hits.length) box.appendChild(el('div', 'Nothing found. Try "aging", "hotel", "refresh"…', 'empty'));
    hits.forEach(function (c, n) {
      var b = el('button', '', 'it');
      b.type = 'button';
      b.setAttribute('role', 'option');
      var t = el('div', '', 'grow');
      t.appendChild(el('b', c[0]));
      t.appendChild(el('small', c[1]));
      b.appendChild(t);
      b.addEventListener('click', function () { run(n); });
      box.appendChild(b);
    });
    box.hidden = false;
  }
  function run(n) {
    var c = hits[n];
    $('qsInput').value = '';
    $('qsList').hidden = true;
    if (c) c[2]();
  }

  // ---------------------------------------------------------- move to another PC --

  function exportSetup() {
    var msg = $('xMessage');
    msg.className = 'status-line';
    msg.textContent = 'Gathering…';
    var layouts = [], queries = [];
    var chain = Promise.resolve();
    if ($('xLayouts').checked) {
      chain = chain.then(function () { return TQChat.workbookContext(); }).then(function (context) {
        var names = (context.layouts || []).map(function (l) { return l.name; });
        return names.reduce(function (p, name) {
          return p.then(function () {
            return post('/api/design', { op: 'libraryExport', args: { layout: name }, context: context })
              .then(function (d) { layouts.push(d.design); }, function () { /* a layout with a mistake is left out */ });
          });
        }, Promise.resolve());
      });
    }
    if ($('xQueries').checked && inExcel()) {
      chain = chain.then(function () {
        return Excel.run(function (context) {
          var ws = context.workbook.worksheets.getItemOrNullObject('TQ Queries');
          return context.sync().then(function () {
            if (ws.isNullObject) return null;
            var u = ws.getUsedRange(true);
            u.load('values');
            return context.sync().then(function () { queries.push({ sheet: 'TQ Queries', rows: u.values }); });
          });
        });
      });
    }
    chain.then(function () {
      return post('/api/export', { name: $('xName').value.trim(), layouts: layouts, queries: queries });
    }).then(function (d) {
      var s = d.summary;
      msg.textContent = 'Saved: ' + d.path + '\nIt holds ' + [s.settings ? 'the NetSuite connection' : '', s.licence ? 'the licence' : '',
        s.logo ? 'the logo' : '', s.filterSets ? s.filterSets + ' filter set(s)' : '', s.layouts ? s.layouts + ' layout(s)' : '',
        s.queries ? 'the data tables' : ''].filter(Boolean).join(', ') + '. Copy it to the other PC (USB, email, OneDrive) and Import it there.';
    }).catch(function (e) { msg.textContent = e.message; msg.className = 'status-line bad'; });
  }

  var bundle = null;
  function readFile() {
    var file = $('iFile').files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () {
      var msg = $('iMessage');
      msg.className = 'status-line';
      try { bundle = JSON.parse(String(reader.result)); } catch (e) { bundle = null; }
      if (!bundle || bundle.kind !== 'tq-connector-setup') {
        msg.textContent = 'That is not a TQ Connector setup file.';
        msg.className = 'status-line bad';
        $('iChoices').hidden = true;
        return;
      }
      var list = $('iList');
      list.textContent = '';
      list.appendChild(el('p', 'From "' + (bundle.name || 'unnamed') + '", made ' + bundle.made + ' with version ' + bundle.connector + '. Bring in:', 'hint'));
      [['settings', 'The NetSuite connection (account ' + ((bundle.settings || {}).accountId || '?') + ')', !!(bundle.settings && bundle.settings.accountId)],
        ['licence', 'The licence', !!bundle.licence],
        ['logo', 'The company logo', !!bundle.logo],
        ['filterSets', (bundle.filterSets || []).length + ' saved filter set(s)', (bundle.filterSets || []).length > 0],
        ['layouts', (bundle.layouts || []).length + ' report layout(s), built in this workbook', (bundle.layouts || []).length > 0],
        ['queries', 'The data tables (TQ Queries sheet)', (bundle.queries || []).length > 0]].forEach(function (c) {
        if (!c[2]) return;
        var l = el('label', '', 'check');
        var box = el('input');
        box.type = 'checkbox';
        box.checked = true;
        box.setAttribute('data-use', c[0]);
        l.appendChild(box);
        l.appendChild(document.createTextNode(' ' + c[1]));
        list.appendChild(l);
      });
      $('iChoices').hidden = false;
      msg.textContent = '';
    };
    reader.readAsText(file);
  }

  function importSetup() {
    if (!bundle) return;
    var use = {};
    Array.prototype.forEach.call(document.querySelectorAll('#iList [data-use]'), function (b) { use[b.getAttribute('data-use')] = b.checked; });
    var msg = $('iMessage');
    msg.className = 'status-line';
    msg.textContent = 'Importing…';
    var done = [], notes = [];
    post('/api/import', { bundle: bundle, use: use }).then(function (d) {
      done = done.concat(d.done);
      notes = notes.concat(d.notes);
      if (!use.layouts || !(bundle.layouts || []).length) return null;
      return bundle.layouts.reduce(function (p, design) {
        return p.then(function () {
          return TQChat.workbookContext().then(function (context) {
            return post('/api/design', { op: 'libraryImport', args: { design: design }, context: context });
          }).then(function (data) { return TQChat.applyActions(data.actions); })
            .then(function () { done.push('layout "' + (design.title || 'report') + '"'); },
              function (e) { notes.push('Layout "' + (design.title || '?') + '": ' + e.message); });
        });
      }, Promise.resolve());
    }).then(function () {
      if (!use.queries || !(bundle.queries || []).length || !inExcel()) return null;
      var q = bundle.queries[0];
      return Excel.run(function (context) {
        var ws = context.workbook.worksheets.getItemOrNullObject('TQ Queries');
        return context.sync().then(function () {
          if (!ws.isNullObject) { notes.push('This workbook already has a TQ Queries sheet: the data tables were left as they are.'); return null; }
          ws = context.workbook.worksheets.add('TQ Queries');
          ws.getRange('A1').getResizedRange(q.rows.length - 1, q.rows[0].length - 1).values = q.rows;
          return context.sync().then(function () { done.push('the data tables (press Refresh data on the Data table screen)'); });
        });
      });
    }).then(function () {
      msg.textContent = (done.length ? 'Imported ' + done.join(', ') + '.' : 'Nothing was imported.') + (notes.length ? '\n' + notes.join('\n') : '');
      if (window.TQPanel) TQPanel.showStatus();
      loadSets();
      if (window.TQHome) TQHome.showCard();
    }).catch(function (e) { msg.textContent = e.message; msg.className = 'status-line bad'; });
  }

  // ----------------------------------------------------------------------- start --

  function start() {
    if ($('fsApply')) {
      $('fsApply').addEventListener('click', applySet);
      $('fsSaveOpen').addEventListener('click', function () { $('fsSaveRow').hidden = !$('fsSaveRow').hidden; if (!$('fsSaveRow').hidden) $('fsName').focus(); });
      $('fsSave').addEventListener('click', saveSet);
      $('fsName').addEventListener('keydown', function (e) { if (e.key === 'Enter') saveSet(); });
      loadSets();
    }
    if ($('rpCollect')) $('rpCollect').addEventListener('click', collectComments);
    if ($('rpTour')) $('rpTour').addEventListener('click', tour);
    if ($('helpButton')) $('helpButton').addEventListener('click', function () { toggleHelp(); });
    document.addEventListener('tq:screen', showTip);
    document.addEventListener('tq:screen', function (e) { if (e.detail === 'report') loadSets(); });
    if ($('qsInput')) {
      $('qsInput').addEventListener('input', search);
      $('qsInput').addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); run(0); }
        if (e.key === 'Escape') { $('qsInput').value = ''; search(); }
      });
      document.addEventListener('click', function (e) { if (!e.target.closest('.qs')) $('qsList').hidden = true; });
    }
    if ($('xExport')) $('xExport').addEventListener('click', exportSetup);
    if ($('iFile')) $('iFile').addEventListener('change', readFile);
    if ($('iImport')) $('iImport').addEventListener('click', importSetup);
  }

  window.TQExtras = { collectComments: collectComments, tour: tour, loadSets: loadSets };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
