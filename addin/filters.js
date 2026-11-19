/* TQ Excel Connector - Filters section of the Reports tab.
 *
 * Every report has filter cells at the top: B3 "Other filters" (any field: custom segments such as
 * Brand, Name, Item) and B4:B7 subsidiary, department, class, location. This section shows the
 * filters of the active report sheet and adds or removes them by clicking, so nobody has to know
 * the syntax:  "Sales, Bar"  "not Sales"  "Food*" (with everything under it)  "Brand: Premium". */
(function () {
  'use strict';

  var ROWS = { filters: 3, subsidiary: 4, department: 5, 'class': 6, location: 7 };
  var LABELS = { filters: 'Other filters', subsidiary: 'Subsidiary', department: 'Department', 'class': 'Class', location: 'Location' };
  var fields = [];        // [{id, label, kind}] from NetSuite
  var values = [];        // names of the chosen field's values
  var loaded = false;
  var $ = function (id) { return document.getElementById(id); };

  function say(text, bad) {
    $('fMessage').textContent = text || '';
    $('fMessage').className = 'status-line' + (bad ? ' bad' : '');
  }

  function el(tag, text, className) {
    var node = document.createElement(tag);
    if (text) node.textContent = text;
    if (className) node.className = className;
    return node;
  }

  function inExcel() { return typeof Excel !== 'undefined' && window.TQFilters && window.TQFilters.excel; }

  /* "Brand: Premium; Name: Al Noor" -> [{field: 'Brand', text: 'Premium'}, ...] */
  function splitOther(text) {
    return String(text || '').split(';').map(function (part) {
      var at = part.indexOf(':');
      return at > 0 ? { field: part.slice(0, at).trim(), text: part.slice(at + 1).trim() } : null;
    }).filter(Boolean);
  }

  function joinOther(parts) {
    return parts.map(function (p) { return p.field + ': ' + p.text; }).join('; ');
  }

  /* The filter cells of the active sheet, or an error when it is not a TQ report. */
  function readCells() {
    if (window.TQFilters.stand) return Promise.resolve(window.TQFilters.stand.read());
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange('A3:F7');
      range.load('values,formulas');
      sheet.load('name');
      return context.sync().then(function () {
        var v = range.values;
        if (String(v[0][0]) !== 'Other filters' || String(v[1][0]) !== 'Subsidiary') {
          throw new Error('Select a report sheet made by TQ Connector (one with "Other filters" in A3). ' +
            'A report built before this version: build it again with Update report.');
        }
        var out = { sheet: sheet.name, block: [] };
        Object.keys(ROWS).forEach(function (key) { out[key] = String(v[ROWS[key] - 3][1] || '').trim(); });
        // Dropdown filter rows (E4:F7): B3 is then a formula that joins them, and is left alone.
        if (String(v[0][4]) === 'More filters') {
          for (var r = 1; r <= 4; r++) {
            if (String(v[r][4]).trim()) out.block.push({ field: String(v[r][4]).trim(), value: String(v[r][5] || '').trim(), cell: 'F' + (r + 3) });
          }
          out.fixed = String(range.formulas[0][1]).indexOf('=') === 0;
        }
        return out;
      });
    });
  }

  function writeCells(cells) {
    if (window.TQFilters.stand) return Promise.resolve(window.TQFilters.stand.write(cells));
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      Object.keys(cells).forEach(function (key) {
        var cell = sheet.getRange(/^[A-Z]+\d+$/.test(key) ? key : 'B' + ROWS[key]);
        cell.numberFormat = [['@']];
        cell.values = [[cells[key]]];
      });
      return context.sync();
    }).then(function () {
      if (window.TQ && TQ.clearCache) TQ.clearCache();
      return Excel.run(function (context) {
        context.workbook.application.calculate(Excel.CalculationType.full);
        return context.sync();
      });
    });
  }

  function showCurrent() {
    var box = $('fCurrent');
    box.textContent = '';
    if (!inExcel() && !window.TQFilters.stand) {
      box.appendChild(el('div', 'Open this panel inside Excel to see the filters of a report.', 'empty'));
      return Promise.resolve(null);
    }
    return readCells().then(function (cells) {
      var count = 0;
      ['subsidiary', 'department', 'class', 'location'].forEach(function (key) {
        if (!cells[key]) return;
        count += 1;
        box.appendChild(row(LABELS[key], cells[key], function () {
          var change = {};
          change[key] = '';
          return change;
        }));
      });
      (cells.block || []).forEach(function (b) {
        if (!b.value) return;
        count += 1;
        box.appendChild(row(b.field + ' (dropdown)', b.value, function () {
          var change = {};
          change[b.cell] = '';
          return change;
        }));
      });
      var inBlock = (cells.block || []).map(function (b) { return b.field.toLowerCase(); });
      splitOther(cells.filters).forEach(function (part, n, all) {
        if (cells.fixed && inBlock.indexOf(part.field.toLowerCase()) !== -1) return;  // shown as its dropdown row
        count += 1;
        box.appendChild(row(part.field, part.text, cells.fixed ? null : function () {
          return { filters: joinOther(all.filter(function (p, i) { return i !== n; })) };
        }));
      });
      if (!count) box.appendChild(el('div', 'No filters on sheet "' + cells.sheet + '": every amount is shown.', 'empty'));
      return cells;
    }).catch(function (error) {
      box.appendChild(el('div', error.message, 'empty'));
      return null;
    });
  }

  function row(field, text, change) {
    var item = el('div', '', 'item');
    var span = el('span', field + ': ');
    span.appendChild(el('small', text));
    span.title = field + ': ' + text;
    item.appendChild(span);
    if (!change) {  // part of the joined Other filters of a report with dropdown rows: kept in the layout
      span.title += ' (change it in the layout, or with Keep in the layout)';
      return item;
    }
    var remove = el('button', '✕');
    remove.title = 'Remove this filter';
    remove.setAttribute('aria-label', 'Remove filter ' + field);
    remove.addEventListener('click', function () {
      writeCells(change()).then(showCurrent).then(function () { say('Filter removed. The numbers are being refreshed.'); },
        function (error) { say(error.message, true); });
    });
    item.appendChild(remove);
    return item;
  }

  function field() {
    var id = $('fField').value;
    return fields.filter(function (f) { return f.id === id; })[0] || null;
  }

  function showValues() {
    var box = $('fValues');
    box.textContent = '';
    var words = $('fSearch').value.toLowerCase().split(/\s+/).filter(Boolean);
    var shown = values.filter(function (name) {
      var t = name.toLowerCase();
      return words.every(function (w) { return t.indexOf(w) !== -1; });
    }).slice(0, 300);
    shown.forEach(function (name) {
      var line = el('div', '', 'fieldrow');
      var label = el('label', '', 'check');
      var box2 = document.createElement('input');
      box2.type = 'checkbox';
      box2.value = name;
      label.appendChild(box2);
      label.appendChild(document.createTextNode(' ' + name));
      label.title = name;
      line.appendChild(label);
      box.appendChild(line);
    });
    if (!values.length) box.appendChild(el('div', 'Type the values below (names or internal IDs, separated by commas).', 'hint'));
  }

  function loadValues() {
    var f = field();
    values = [];
    $('fUnderBox').hidden = !f || f.kind !== 'segment';
    $('fUnder').nextSibling.textContent = f && f.id === 'subsidiary'
      ? ' With every subsidiary under it, consolidated in its currency'
      : ' With everything under it';
    showValues();
    if (!f || f.kind === 'name' || f.kind === 'item') return;
    say('Loading the ' + f.label.toLowerCase() + ' list…');
    TQ.callApi({ action: 'segments', kind: f.id }).then(function (data) {
      values = data.segments.map(function (s) { return s.name; });
      say(values.length ? '' : 'No values found: type them below.');
      showValues();
    }).catch(function (error) { say(error.message, true); });
  }

  function load() {
    showCurrent();
    if (loaded) return;
    loaded = true;
    TQ.callApi({ action: 'filterFields' }).then(function (data) {
      fields = data.fields || [];
      var select = $('fField');
      select.textContent = '';
      fields.forEach(function (f) {
        var o = el('option', f.label + (f.kind === 'custom' ? ' (custom segment)' : f.kind === 'name' ? ' (customer, vendor, employee)' : ''));
        o.value = f.id;
        select.appendChild(o);
      });
      loadValues();
    }).catch(function (error) { loaded = false; say(error.message, true); });
  }

  /* The chosen values as filter text: "A, B", "not A", "Food*". */
  function chosenText() {
    var picked = Array.prototype.map.call($('fValues').querySelectorAll('input:checked'), function (b) { return b.value; });
    var typed = $('fTyped').value.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
    var all = picked.concat(typed.filter(function (t) { return picked.indexOf(t) === -1; }));
    if (!all.length) return '';
    var under = !$('fUnderBox').hidden && $('fUnder').checked;
    return ($('fExclude').checked ? 'not ' : '') + all.map(function (t) { return under ? t + '*' : t; }).join(', ');
  }

  /* What the cells become with the chosen filter added: a segment replaces its own cell, any other
   * field replaces its own part of "Other filters". */
  function merged(cells, f, text) {
    var change = {};
    var row = (cells.block || []).filter(function (b) {
      return b.field.toLowerCase() === f.label.toLowerCase() || b.field.toLowerCase() === f.id;
    })[0];
    if (f.kind === 'segment') {
      change[f.id] = text;
    } else if (row) {
      change[row.cell] = text;           // the field's dropdown row
    } else if (cells.fixed) {
      throw new Error('On this report the other filters come from the dropdown rows (E4:F7). Use "Add as a dropdown ' +
        'on the sheet" for ' + f.label + ', or "Keep in the layout".');
    } else {
      var parts = splitOther(cells.filters).filter(function (p) {
        return p.field.toLowerCase() !== f.label.toLowerCase() && p.field.toLowerCase() !== f.id;
      });
      parts.push({ field: f.label, text: text });
      change.filters = joinOther(parts);
    }
    return change;
  }

  function apply() {
    var f = field();
    var text = chosenText();
    if (!f) { say('Choose a field first.', true); return; }
    if (!text) { say('Tick at least one value or type one.', true); return; }
    say('Applying…');
    readCells().then(function (cells) {
      return writeCells(merged(cells, f, text));
    }).then(function () {
      reset();
      return showCurrent();
    }).then(function () {
      say('Filter applied: ' + f.label + ' ' + text + '. The numbers are being refreshed. ' +
        'Rebuilding a layout report puts its own filters back; use "Keep in the layout" for a lasting filter.');
    }).catch(function (error) { say(error.message, true); });
  }

  function keep() {
    var f = field();
    var text = chosenText();
    if (!f || !text) { say('Choose a field and at least one value first.', true); return; }
    if (!window.TQDesigner || !window.TQDesigner.current()) {
      say('There is no layout open in the Layout section below.', true);
      return;
    }
    var design = window.TQDesigner.current();
    var cells = { filters: (design.filters || {}).filters || '' };
    var change = merged(cells, f, text);
    var filters = Object.assign({}, design.filters || {}, change);
    window.TQDesigner.setFilters(filters).then(function (data) {
      if (!data) return;
      reset();
      say('Kept in layout "' + design.name + '" and the report was rebuilt.');
      showCurrent();
    });
  }

  /* A dropdown row for the chosen field on the report itself: the report is rebuilt with it, and the
   * row comes back on every rebuild (it is saved with the layout, or with the report's settings). */
  function addDropdown() {
    var f = field();
    if (!f) { say('Choose a field first.', true); return; }
    if (f.kind === 'segment') {
      say(f.label + ' already has its own cell with a dropdown (B4:B7) on every report.');
      return;
    }
    say('Adding a dropdown for ' + f.label + '…');
    window.TQChat.workbookContext().then(function (context) {
      context.last = window.TQChat.getSpec ? window.TQChat.getSpec() : null;
      return fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: 'add dropdown filter ' + f.label, context: context })
      }).then(function (r) { return r.json(); }).then(function (data) {
        if (!data || data.ok === false || !(data.actions || []).length) {
          throw new Error((data && (data.error || data.reply)) || 'Build a report first, then add a dropdown to it.');
        }
        return window.TQChat.applyActions(data.actions).then(function () {
          if (data.spec) window.TQChat.setSpec(data.spec);
          say('Dropdown for ' + f.label + ' added in E4:F7 of the report: pick a value there, then Refresh all. ' +
            data.reply.split('\n')[0]);
          return showCurrent();
        });
      });
    }).catch(function (error) { say(error.message, true); });
  }

  function clearAll() {
    readCells().then(function (cells) {
      var change = { subsidiary: '', department: '', 'class': '', location: '' };
      if (cells.fixed) (cells.block || []).forEach(function (b) { change[b.cell] = ''; });
      else change.filters = '';
      return writeCells(change);
    })
      .then(showCurrent)
      .then(function () { say('All filters cleared on this sheet.'); }, function (error) { say(error.message, true); });
  }

  function reset() {
    $('fTyped').value = '';
    $('fSearch').value = '';
    $('fExclude').checked = false;
    $('fUnder').checked = false;
    showValues();
  }

  function start() {
    $('fField').addEventListener('change', loadValues);
    $('fSearch').addEventListener('input', showValues);
    $('fApply').addEventListener('click', apply);
    $('fKeep').addEventListener('click', keep);
    $('fClear').addEventListener('click', clearAll);
    $('fDropdown').addEventListener('click', addDropdown);
    $('fRead').addEventListener('click', showCurrent);
  }

  window.TQFilters = { load: load, excel: false, stand: null, splitOther: splitOther, merged: merged,
    read: readCells, write: function (cells) { return writeCells(cells).then(showCurrent); } };

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(function (info) { window.TQFilters.excel = !!(info && info.host); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
