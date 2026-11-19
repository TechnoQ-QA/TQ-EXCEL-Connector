/* TQ Excel Connector - Budget upload. Makes a budget sheet, then checks it against NetSuite (writes nothing) and
 * uploads it. Upload stays disabled until a check of the same sheet found no problem. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var loaded = false;
  var checked = null;   // the cells that passed the check, as JSON: Upload sends only exactly these

  function say(text, bad) {
    $('gMessage').textContent = text || '';
    $('gMessage').className = 'status-line' + (bad ? ' bad' : '');
  }

  function post(payload) {
    return fetch('/api/budget', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
    }).then(function (r) { return r.json(); }).then(function (data) {
      if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
      return data;
    });
  }

  function option(select, value) {
    var o = document.createElement('option');
    o.value = value;
    o.textContent = value;
    select.appendChild(o);
  }

  function load() {
    if (loaded) return;
    loaded = true;
    say('Reading the fiscal years and budget categories…');
    post({ op: 'info' }).then(function (data) {
      $('gYear').textContent = '';
      data.years.forEach(function (y) { option($('gYear'), y.name); });
      var thisYear = String(new Date().getFullYear());
      var next = data.years.filter(function (y) { return y.name.indexOf(String(Number(thisYear) + 1)) !== -1; })[0] ||
        data.years.filter(function (y) { return y.name.indexOf(thisYear) !== -1; })[0];
      if (next) $('gYear').value = next.name;
      $('gCategory').textContent = '';
      if (!data.categories.length) option($('gCategory'), '');
      data.categories.forEach(function (c) { option($('gCategory'), c); });
      say(data.canWrite ? '' : 'You can make and check budget sheets; uploading needs a NetSuite role that may set up budgets.');
    }).catch(function (e) { loaded = false; say(e.message, true); });
  }

  function inExcel() { return typeof Excel !== 'undefined' && window.Office && Office.context && Office.context.document; }

  function make() {
    say('Building the budget sheet…');
    post({ op: 'sheet', year: $('gYear').value, category: $('gCategory').value, fill: $('gFill').checked })
      .then(function (data) {
        return window.TQChat.applyActions(data.actions).then(function () { say(data.reply); });
      }).catch(function (e) { say(e.message, true); });
  }

  function readSheet() {
    return Excel.run(function (context) {
      var used = context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true);
      used.load('values');
      return context.sync().then(function () { return used.isNullObject ? [] : used.values; });
    });
  }

  function send(op) {
    if (!inExcel()) { say('Open the panel in Excel and the budget sheet first.', true); return; }
    say(op === 'check' ? 'Checking the budget sheet with NetSuite…' : 'Uploading to NetSuite…');
    $('gUpload').disabled = true;
    readSheet().then(function (rows) {
      var cells = JSON.stringify(rows);
      if (op === 'upload' && cells !== checked) throw new Error('The sheet changed since the check. Press Check again.');
      return post({ op: op, rows: rows }).then(function (data) {
        checked = op === 'check' && !data.problems ? cells : null;
        $('gUpload').disabled = !checked;
        say(data.reply, data.problems > 0);
      });
    }).catch(function (e) { checked = null; say(e.message, true); });
  }

  /* Journal upload: the same Check-then-Upload steps, on /api/journal. */
  var journalChecked = null;
  function jsay(text, bad) {
    $('jMessage').textContent = text || '';
    $('jMessage').className = 'status-line' + (bad ? ' bad' : '');
  }
  function journalPost(payload) {
    return fetch('/api/journal', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
    }).then(function (r) { return r.json(); }).then(function (data) {
      if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
      return data;
    });
  }
  function journalMake() {
    jsay('Making the journal sheet…');
    journalPost({ op: 'sheet' }).then(function (data) {
      return window.TQChat.applyActions(data.actions).then(function () { jsay(data.reply); });
    }).catch(function (e) { jsay(e.message, true); });
  }
  function journalSend(op) {
    if (!inExcel()) { jsay('Open the panel in Excel and the journal sheet first.', true); return; }
    jsay(op === 'check' ? 'Checking the journal lines with NetSuite…' : 'Creating the journal entries…');
    $('jUpload').disabled = true;
    readSheet().then(function (rows) {
      var cells = JSON.stringify(rows);
      if (op === 'upload' && cells !== journalChecked) throw new Error('The sheet changed since the check. Press Check again.');
      return journalPost({ op: op, rows: rows }).then(function (data) {
        journalChecked = op === 'check' && !data.problems ? cells : null;
        $('jUpload').disabled = !journalChecked;
        jsay(data.reply, data.problems > 0);
      });
    }).catch(function (e) { journalChecked = null; jsay(e.message, true); });
  }

  function start() {
    if (!$('secBudget')) return;
    $('secBudget').addEventListener('toggle', function () { if ($('secBudget').open) load(); });
    $('gMake').addEventListener('click', make);
    $('gCheck').addEventListener('click', function () { send('check'); });
    $('gUpload').addEventListener('click', function () { send('upload'); });
    if ($('secJournal')) {
      $('jMake').addEventListener('click', journalMake);
      $('jCheck').addEventListener('click', function () { journalSend('check'); });
      $('jUpload').addEventListener('click', function () { journalSend('upload'); });
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
