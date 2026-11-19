/* TQ Excel Connector - Insert a formula. Builds =TQ.BALANCE(...) from pickers (account, period, filters) and
 * writes it into the selected cell, or reads the selected cell's formula back into the pickers. */
(function () {
  'use strict';

  var KINDS = ['subsidiary', 'department', 'class', 'location'];
  var RELATIVE = [
    ['CM', 'This month (CM)'], ['CM-1', 'Last month (CM-1)'], ['LY', 'Same month last year (LY)'],
    ['FYS', 'First month of this year (FYS)'], ['FYE', 'Last month of this year (FYE)'],
    ['QS', 'First month of this quarter (QS)'], ['QE', 'Last month of this quarter (QE)'],
    ['FYS-12', 'First month of last year (FYS-12)'], ['FYS-1', 'Last year end (FYS-1)']
  ];
  var accounts = null;
  var loaded = false;
  var $ = function (id) { return document.getElementById(id); };

  function say(text, bad) {
    $('bMessage').textContent = text || '';
    $('bMessage').className = 'status-line' + (bad ? ' bad' : '');
  }

  function option(select, value, text) {
    var o = document.createElement('option');
    o.value = value;
    o.textContent = text === undefined ? value : text;
    select.appendChild(o);
  }

  function quote(text) { return '"' + String(text).replace(/"/g, '""') + '"'; }
  function isRelative(value) { return /^(CM|LY|FYS|FYE|QS|QE)([+-]\d+)?$/i.test(value); }

  function periodArg(value) {
    if (!value) return '""';
    if ($('bLink').checked && isRelative(value)) {
      var cell = ($('bCell').value || 'B8').replace(/[^$A-Za-z0-9]/g, '').toUpperCase();
      return 'TQ.PERIOD(' + cell.replace(/^\$?([A-Z]+)\$?(\d+)$/, '$$$1$$$2') + ',' + quote(value.toUpperCase()) + ')';
    }
    return quote(value);
  }

  function formula() {
    var account = $('bAccount').value.trim();
    if (!account) return '';
    if ($('bBudget').checked && !/^budget/i.test(account)) account = 'budget:' + account;
    var to = $('bTo').value;
    var args = [quote(account), periodArg($('bFrom').value), to ? periodArg(to) : '""'];
    KINDS.forEach(function (kind) {
      var select = $('bSeg_' + kind);
      args.push(select && select.value ? quote(select.value) : '""');
    });
    args.push($('bOther').value.trim() ? quote($('bOther').value.trim()) : '""');
    while (args.length > 2 && args[args.length - 1] === '""') args.pop();   // no empty arguments at the end
    return '=TQ.BALANCE(' + args.join(',') + ')';
  }

  function preview() { $('bPreview').textContent = formula() || 'Pick an account first.'; }

  function load() {
    if (loaded) return;
    loaded = true;
    [$('bFrom'), $('bTo')].forEach(function (select, n) {
      select.textContent = '';
      if (n) option(select, '', 'Same as Period (one period)');
      RELATIVE.forEach(function (r) { option(select, r[0], r[1]); });
    });
    TQ.callApi({ action: 'periods' }).then(function (data) {
      [$('bFrom'), $('bTo')].forEach(function (select) {
        var group = document.createElement('optgroup');
        group.label = 'Periods';
        data.periods.forEach(function (p) { option(group, p.name); });
        select.appendChild(group);
      });
    }).catch(function () { loaded = false; });
    var box = $('bSegments');
    box.textContent = '';
    KINDS.forEach(function (kind) {
      var label = document.createElement('label');
      label.textContent = kind.charAt(0).toUpperCase() + kind.slice(1) + ' (optional)';
      label.setAttribute('for', 'bSeg_' + kind);
      var select = document.createElement('select');
      select.className = 'field';
      select.id = 'bSeg_' + kind;
      option(select, '', 'All');
      select.addEventListener('change', preview);
      box.appendChild(label);
      box.appendChild(select);
      TQ.callApi({ action: 'segments', kind: kind }).then(function (data) {
        data.segments.forEach(function (s) { option(select, s.name); });
        if (!data.segments.length) { label.hidden = true; select.hidden = true; }
      }).catch(function () { label.hidden = true; select.hidden = true; });
    });
    TQ.callApi({ action: 'accounts' }).then(function (data) { accounts = data.accounts; }).catch(function () {});
    preview();
  }

  function search() {
    var box = $('bResults');
    box.textContent = '';
    var words = $('bSearch').value.toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length || !accounts) return;
    accounts.filter(function (a) {
      var text = (a.number + ' ' + a.name + ' ' + a.type).toLowerCase();
      return words.every(function (w) { return text.indexOf(w) !== -1; });
    }).slice(0, 40).forEach(function (a) {
      var button = document.createElement('button');
      button.textContent = a.number + '  ' + a.name;
      button.addEventListener('click', function () { $('bAccount').value = a.number; preview(); });
      box.appendChild(button);
    });
  }

  function inExcel() { return typeof Excel !== 'undefined' && window.Office && Office.context && Office.context.document; }

  function insert() {
    var text = formula();
    if (!text) { say('Pick an account first.', true); return; }
    if (!inExcel()) { say('Open the panel in Excel to insert the formula.', true); return; }
    Excel.run(function (context) {
      var cell = context.workbook.getSelectedRange().getCell(0, 0);
      cell.formulas = [[text]];
      cell.numberFormat = [['#,##0.00;(#,##0.00)']];
      cell.load('address');
      return context.sync().then(function () { return cell.address; });
    }).then(function (address) { say('Written to ' + address.split('!').pop() + '.'); },
      function (e) { say(e.message, true); });
  }

  /* '"4000",TQ.PERIOD($B$8,"FYS"),"x, y"' -> ['"4000"', 'TQ.PERIOD($B$8,"FYS")', '"x, y"'] */
  function splitArgs(text) {
    var out = [], depth = 0, inText = false, start = 0;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (ch === '"') inText = !inText;
      else if (!inText && ch === '(') depth++;
      else if (!inText && ch === ')') depth--;
      else if (!inText && depth === 0 && ch === ',') { out.push(text.slice(start, i).trim()); start = i + 1; }
    }
    out.push(text.slice(start).trim());
    return out;
  }

  function unquote(arg) {
    var m = /^"((?:[^"]|"")*)"$/.exec(arg || '');
    return m ? m[1].replace(/""/g, '"') : null;
  }

  function setPeriod(select, arg) {
    var linked = /^TQ\.PERIOD\(([^,]+),"([^"]*)"\)$/i.exec(arg || '');
    var value = linked ? linked[2] : unquote(arg);
    if (linked) { $('bLink').checked = true; $('bCell').value = linked[1].replace(/\$/g, ''); }
    if (value === null) return false;
    if (value && !Array.prototype.some.call(select.options, function (o) { return o.value === value; })) option(select, value);
    select.value = value;
    return true;
  }

  function read() {
    if (!inExcel()) { say('Open the panel in Excel to read a cell.', true); return; }
    Excel.run(function (context) {
      var cell = context.workbook.getSelectedRange().getCell(0, 0);
      cell.load('formulas');
      return context.sync().then(function () { return String(cell.formulas[0][0]); });
    }).then(function (text) {
      var m = /^=TQ\.BALANCE\((.*)\)$/i.exec(text.trim());
      if (!m) throw new Error('The selected cell has no TQ.BALANCE formula.');
      var args = splitArgs(m[1]);
      var account = unquote(args[0]);
      var unread = [];
      if (account === null) unread.push('account'); else {
        $('bBudget').checked = /^budget:/i.test(account);
        $('bAccount').value = account.replace(/^budget:/i, '');
      }
      $('bLink').checked = false;
      if (!setPeriod($('bFrom'), args[1])) unread.push('period');
      if (!setPeriod($('bTo'), args[2] || '""')) unread.push('to period');
      KINDS.forEach(function (kind, n) {
        var value = unquote(args[3 + n] || '""');
        var select = $('bSeg_' + kind);
        if (value === null) { unread.push(kind); return; }
        if (value && select && !Array.prototype.some.call(select.options, function (o) { return o.value === value; })) option(select, value);
        if (select) select.value = value;
      });
      var other = unquote(args[7] || '""');
      if (other === null) unread.push('other filters'); else $('bOther').value = other;
      preview();
      say(unread.length ? 'Read. These point at other cells and were left as they are in the pickers: ' + unread.join(', ') + '.'
        : 'Read. Change what you need, then Insert in selected cell.');
    }).catch(function (e) { say(e.message, true); });
  }

  function start() {
    if (!$('secFormula')) return;
    $('secFormula').addEventListener('toggle', function () { if ($('secFormula').open) load(); });
    $('bSearch').addEventListener('input', search);
    ['bAccount', 'bOther', 'bCell'].forEach(function (id) { $(id).addEventListener('input', preview); });
    ['bFrom', 'bTo', 'bBudget', 'bLink'].forEach(function (id) { $(id).addEventListener('change', preview); });
    Array.prototype.forEach.call(document.querySelectorAll('[data-total]'), function (button) {
      button.addEventListener('click', function () { $('bAccount').value = button.getAttribute('data-total'); preview(); });
    });
    $('bInsert').addEventListener('click', insert);
    $('bRead').addEventListener('click', read);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
