/* TQ Excel Connector - Designer tab. Builds a report layout by clicking: its lines (rows) and its
 * columns. The layout lives on a "TQ Layout" sheet in the workbook; every change here is sent to
 * the local server, which checks it and writes the sheet back. "Update report" rebuilds the report
 * on its own sheet, in place. */
(function () {
  'use strict';

  var KINDS = ['subsidiary', 'department', 'class', 'location'];
  var design = null;      // the layout being edited, as the server last returned it
  var accounts = null;    // [{number, name, type}] for the account search
  var listsLoaded = false;
  var busy = false;
  var BUILDS = ['build', 'setPages', 'setDetail', 'setComments', 'setCommentAt', 'setHighlight', 'addRule', 'removeRule'];  // requests that rebuild the report as well
  var $ =function (id) { return document.getElementById(id); };

  function say(text, bad) {
    var el = $('dMessage');
    el.textContent = text || '';
    el.className = bad ? 'bad' : '';
  }

  function el(tag, text, className) {
    var node = document.createElement(tag);
    if (text) node.textContent = text;
    if (className) node.className = className;
    return node;
  }

  function option(select, value, text) {
    var o = el('option', text === undefined ? value : text);
    o.value = value;
    select.appendChild(o);
  }

  /* Sends one request about the current layout; applies what comes back to the workbook and this tab. */
  function request(op, args, working) {
    if (busy) return Promise.resolve(null);
    busy = true;
    say(working || 'Working…');
    var wanted = Object.assign({ layout: $('dLayout').value || (design && design.name) || '' }, args || {});
    return window.TQChat.workbookContext()
      .then(function (context) {
        if (op !== 'new' && !(context.layouts || []).length) {
          throw new Error(context.layouts ? 'This workbook has no layout yet. Press New layout to start one.'
            : 'Open this panel inside Excel to design a report.');
        }
        return fetch('/api/design', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ op: op, args: wanted, context: context })
        }).then(function (response) { return response.json(); }).then(function (data) {
          if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
          return window.TQChat.applyActions(data.actions).then(function (notes) {
            if (data.design) design = data.design;
            if (data.spec) window.TQChat.setSpec(data.spec);
            var names = (context.layouts || []).map(function (l) { return l.name; });
            if (design && names.indexOf(design.name) === -1) names.push(design.name);
            render(names);
            say([data.reply].concat(BUILDS.indexOf(op) !== -1 ? notes : []).filter(Boolean).join('\n'));
            return data;
          });
        });
      })
      .catch(function (error) { say(error.message, true); return null; })
      .then(function (result) { busy = false; return result; });
  }

  function itemRow(text, detail, what, label) {
    var row = el('div', '', 'item');
    var span = el('span', text);
    span.title = text + (detail ? '  ' + detail : '');
    if (detail) { span.appendChild(document.createTextNode(' ')); span.appendChild(el('small', detail)); }
    row.appendChild(span);
    [['↑', 'Move up', -1], ['↓', 'Move down', 1], ['✕', 'Remove', 0]].forEach(function (b) {
      var button = el('button', b[0]);
      button.title = b[1] + ' ' + label;
      button.setAttribute('aria-label', b[1] + ' ' + label);
      button.addEventListener('click', function () {
        if (b[2]) request('move' + what, { label: label, by: b[2] });
        else request('remove' + what, { label: label });
      });
      row.appendChild(button);
    });
    return row;
  }

  function chips(box, input, labels) {
    box.textContent = '';
    labels.forEach(function (label) {
      var chip = el('button', label);
      chip.title = 'Put [' + label + '] in the formula';
      chip.addEventListener('click', function () {
        input.value = (input.value || '=') + '[' + label + ']';
        input.focus();
      });
      box.appendChild(chip);
    });
  }

  function render(names) {
    var select = $('dLayout');
    if (names) {
      select.textContent = '';
      names.forEach(function (name) { option(select, name); });
    }
    $('dBody').hidden = !design;
    if (!design) return;
    select.value = design.name;
    $('dTitle').value = design.title;
    $('dPages').value = design.pages || '';
    $('dDetail').checked = !!design.detail;
    $('dComments').checked = !!design.comments;
    var h = design.highlight ? String(Number(design.highlight)) : '';
    if (h && !Array.prototype.some.call($('dHighlight').options, function (o) { return o.value === h; })) option($('dHighlight'), h, h + '%');
    $('dHighlight').value = h;
    if ($('dCommentAt')) $('dCommentAt').value = design.commentAt ? String(Number(design.commentAt)) : '';
    var rules = $('rList');
    rules.textContent = '';
    (design.rules || []).forEach(function (text) {
      var row = el('div', '', 'item');
      row.appendChild(el('span', text));
      var x = el('button', '✕');
      x.title = 'Remove this rule';
      x.addEventListener('click', function () { request('removeRule', { rule: text }, 'Building the report…'); });
      row.appendChild(x);
      rules.appendChild(row);
    });
    if (!(design.rules || []).length) rules.appendChild(el('div', 'No rules yet.', 'empty'));
    var cols = $('rColumn');
    var keep = cols.value;
    cols.textContent = '';
    option(cols, '', 'Every number column');
    design.columns.forEach(function (c) { option(cols, c.label); });
    cols.value = keep;

    var lines = $('dLines');
    lines.textContent = '';
    var section = null;
    var sections = [];
    design.lines.forEach(function (line) {
      if (line.section !== section) {
        section = line.section;
        if (section) { lines.appendChild(el('div', section, 'group')); sections.push(section); }
      }
      var detail = line.kind === 'accounts' ? line.source + (line.sign < 0 ? ' (negative)' : '')
        : line.kind === 'formula' ? line.source : 'subtotal';
      lines.appendChild(itemRow(line.label, detail, 'Line', line.label));
    });
    if (!design.lines.length) lines.appendChild(el('div', 'No lines yet.', 'empty'));
    var datalist = $('lSections');
    datalist.textContent = '';
    sections.forEach(function (name) { option(datalist, name); });

    var columns = $('dColumns');
    columns.textContent = '';
    design.columns.forEach(function (column) {
      var filters = KINDS.concat(['filters']).map(function (k) { return column.filters[k]; }).filter(Boolean).join(', ');
      var detail = column.type === 'period'
        ? (column.data ? 'budget, ' : '') + (column.from === column.to ? column.from : column.from + ' to ' + column.to) + (filters ? ', ' + filters : '')
        : column.formula;
      columns.appendChild(itemRow(column.label, detail, 'Column', column.label));
    });
    if (!design.columns.length) columns.appendChild(el('div', 'No columns yet.', 'empty'));

    chips($('lRefs'), $('lFormula'), design.lines.map(function (l) { return l.label; }).concat(['days']).slice(0, 60));
    chips($('cRefs'), $('cFormula'), design.columns.map(function (c) { return c.label; }).slice(0, 60));
  }

  /* Periods, segment lists and accounts for the pickers. Loaded once, when the tab first opens. */
  function loadLists() {
    if (listsLoaded) return;
    listsLoaded = true;
    TQ.callApi({ action: 'periods' }).then(function (data) {
      [$('cFrom'), $('cTo')].forEach(function (select, n) {
        select.textContent = '';
        if (n) option(select, '', 'Same as From');
        data.periods.forEach(function (p) { option(select, p.name); });
      });
    }).catch(function () { listsLoaded = false; });
    var box = $('cFilters');
    box.textContent = '';
    KINDS.forEach(function (kind) {
      var label = el('label', 'Only ' + kind + ' (optional)');
      label.setAttribute('for', 'cFilter_' + kind);
      var select = el('select', '', 'field');
      select.id = 'cFilter_' + kind;
      option(select, '', 'All');
      box.appendChild(label);
      box.appendChild(select);
      TQ.callApi({ action: 'segments', kind: kind }).then(function (data) {
        data.segments.forEach(function (s) { option(select, s.name); });
        if (!data.segments.length) { label.hidden = true; select.hidden = true; }
      }).catch(function () { label.hidden = true; select.hidden = true; });
    });
    var otherLabel = el('label', 'Other filters (optional)');
    otherLabel.setAttribute('for', 'cFilter_filters');
    var other = el('input', '', 'field');
    other.type = 'text';
    other.id = 'cFilter_filters';
    other.placeholder = 'e.g. Brand: Premium; Name: Al Noor';
    box.appendChild(otherLabel);
    box.appendChild(other);
    TQ.callApi({ action: 'accounts' }).then(function (data) { accounts = data.accounts; }).catch(function () {});
  }

  function searchAccounts() {
    var box = $('lResults');
    box.textContent = '';
    var words = $('lSearch').value.toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length || !accounts) return;
    accounts.filter(function (a) {
      var text = (a.number + ' ' + a.name + ' ' + a.type).toLowerCase();
      return words.every(function (w) { return text.indexOf(w) !== -1; });
    }).slice(0, 40).forEach(function (a) {
      var button = el('button', a.number + '  ' + a.name);
      button.title = 'Add account ' + a.number + ' to this line';
      button.addEventListener('click', function () {
        var source = $('lSource');
        var have = source.value.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
        if (have.indexOf(a.number) === -1) have.push(a.number);
        source.value = have.join(', ');
        if (!$('lLabel').value) $('lLabel').value = a.name.split(' : ').pop();
      });
      box.appendChild(button);
    });
  }

  function showLineBoxes() {
    var kind = $('lKind').value;
    $('lAccountsBox').hidden = kind !== 'accounts';
    $('lFormulaBox').hidden = kind !== 'formula';
  }

  function showColumnBoxes() {
    var type = $('cType').value;
    $('cFormulaBox').hidden = type === 'period';
    $('cPeriodBox').hidden = type === 'formula';
    $('cFilters').hidden = type !== 'period';
    $('cBudget').parentNode.hidden = type !== 'period';
    $('cFormula').placeholder = type === 'total' ? '=SUM([Jan 2026]:[Dec 2026])' : '=[YTD 2026]-[YTD 2025]';
  }

  function addLine() {
    var kind = $('lKind').value;
    var format = $('lFormat').value;
    request('addLine', {
      section: $('lSection').value, label: $('lLabel').value, kind: kind,
      source: kind === 'accounts' ? $('lSource').value : kind === 'formula' ? $('lFormula').value : '',
      negative: kind === 'accounts' && $('lNegative').checked, format: format
    }).then(function (data) {
      if (!data) return;
      ['lLabel', 'lSource', 'lFormula', 'lSearch'].forEach(function (id) { $(id).value = ''; });
      $('lResults').textContent = '';
    });
  }

  function addColumn() {
    var type = $('cType').value;
    var filters = {};
    KINDS.concat(['filters']).forEach(function (kind) {
      var select = $('cFilter_' + kind);
      if (type === 'period' && select && select.value) filters[kind] = select.value.trim();
    });
    request('addColumn', {
      label: $('cLabel').value, type: type, from: $('cFrom').value, to: $('cTo').value,
      filters: filters, formula: type === 'period' ? '' : $('cFormula').value, format: $('cFormat').value,
      budget: type === 'period' && $('cBudget').checked
    }).then(function (data) {
      if (!data) return;
      ['cLabel', 'cFormula', 'cFilter_filters'].forEach(function (id) { if ($(id)) $(id).value = ''; });
    });
  }

  /* New layout: from a ready-made template, or from a format the user typed on a sheet. */
  var OWN = '(sheet)';
  var templates = [];

  function loadTemplates() {
    fetch('/api/design', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ op: 'templates', args: {}, context: {} })
    }).then(function (response) { return response.json(); }).then(function (data) {
      if (!data || !data.templates) return;
      templates = data.templates;
      var select = $('dTemplate');
      select.textContent = '';
      templates.forEach(function (t) { option(select, t.id, t.name); });
      option(select, OWN, 'My own format, typed on a sheet');
      showTemplate();
    }).catch(function () {});
  }

  function showTemplate() {
    var id = $('dTemplate').value;
    var found = templates.filter(function (t) { return t.id === id; })[0];
    $('dTemplateAbout').textContent = found ? found.about : '';
    $('dSourceBox').hidden = id !== OWN;
    if (id !== OWN) return;
    window.TQChat.workbookContext().then(function (context) {
      var select = $('dSource');
      var keep = select.value;
      select.textContent = '';
      (context.sheets || []).filter(function (name) {
        return name.toLowerCase().indexOf('tq layout') !== 0 && name.toLowerCase() !== 'tq queries';
      }).forEach(function (name) { option(select, name); });
      if (keep) select.value = keep;
      if (context.activeSheet && !keep) select.value = context.activeSheet;
    });
  }

  /* The cells of one sheet, for the import: up to 400 rows by 30 columns. */
  function readSheet(name) {
    if (window.TQChat.readSheet) return window.TQChat.readSheet(name);  // a stand-in, outside Excel
    return Excel.run(function (context) {
      var used = context.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
      used.load('values');
      return context.sync().then(function () {
        return used.isNullObject ? [] : used.values.slice(0, 400).map(function (row) { return row.slice(0, 30); });
      });
    });
  }

  function createLayout() {
    var id = $('dTemplate').value;
    var done = function (data) { if (data) $('dNewName').value = ''; };
    if (id !== OWN) {
      request('new', { name: $('dNewName').value, template: id }, 'Building the layout…').then(done);
      return;
    }
    var sheet = $('dSource').value;
    if (!sheet) { say('Choose the sheet that holds your format.', true); return; }
    say('Reading sheet ' + sheet + '…');
    readSheet(sheet).then(function (rows) {
      if (!rows.length) throw new Error('Sheet "' + sheet + '" is empty.');
      return request('new', { name: $('dNewName').value, sheet: sheet, rows: rows }, 'Matching the lines to accounts…').then(done);
    }).catch(function (error) { say(error.message, true); });
  }

  /* Called when the tab is shown, and by Reload: reads the layout sheet again. */
  function load() {
    loadLists();
    return request('get', {}, 'Reading the layout…').then(function (data) {
      if (data) say('');
      else { design = null; render(); }
    });
  }

  function start() {
    $('dReload').addEventListener('click', load);
    $('dLayout').addEventListener('change', load);
    $('dNew').addEventListener('click', createLayout);
    $('dTemplate').addEventListener('change', showTemplate);
    loadTemplates();
    $('dTitleSave').addEventListener('click', function () { request('setTitle', { title: $('dTitle').value }); });
    $('dPages').addEventListener('change', function () {
      request('setPages', { pages: $('dPages').value }, $('dPages').value ? 'Building a sheet for each ' + $('dPages').value + '…' : 'Building the report…');
    });
    $('dDetail').addEventListener('change', function () { request('setDetail', { detail: $('dDetail').checked }, 'Building the report…'); });
    $('dComments').addEventListener('change', function () { request('setComments', { comments: $('dComments').checked }, 'Building the report…'); });
    $('dHighlight').addEventListener('change', function () { request('setHighlight', { highlight: $('dHighlight').value }, 'Building the report…'); });
    if ($('dCommentAt')) $('dCommentAt').addEventListener('change', function () { request('setCommentAt', { commentAt: $('dCommentAt').value }, 'Building the report…'); });
    $('rOp').addEventListener('change', function () { $('rV2').hidden = $('rOp').value !== 'between'; });
    $('rAdd').addEventListener('click', function () {
      var target = $('rTarget').value.trim() || 'all';
      var text = target + ($('rColumn').value ? ' in ' + $('rColumn').value : '') + ' ' + $('rOp').value + ' ' +
        $('rV1').value.trim() + ($('rOp').value === 'between' ? ' and ' + $('rV2').value.trim() : '') + ' ' + $('rColour').value;
      request('addRule', { rule: text }, 'Building the report…').then(function (data) {
        if (data) { $('rTarget').value = ''; $('rV1').value = ''; $('rV2').value = ''; }
      });
    });
    $('lKind').addEventListener('change', showLineBoxes);
    $('cType').addEventListener('change', showColumnBoxes);
    $('lSearch').addEventListener('input', searchAccounts);
    $('lAdd').addEventListener('click', addLine);
    $('cAdd').addEventListener('click', addColumn);
    $('dBuild').addEventListener('click', function () { request('build', {}, 'Building the report…'); });
    showLineBoxes();
    showColumnBoxes();
  }

  /* A lasting filter from the Filters section: saved in the layout, then the report is rebuilt. */
  function setFilters(filters) {
    return request('setFilters', { filters: filters }, 'Saving the filter in the layout…').then(function (data) {
      return data && request('build', {}, 'Building the report…');
    });
  }

  window.TQDesigner = { load: load, current: function () { return design; }, setFilters: setFilters,
    build: function () { return request('build', {}, 'Building the report…'); },
    select: function (name) { option($('dLayout'), name); $('dLayout').value = name; return load(); } };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
