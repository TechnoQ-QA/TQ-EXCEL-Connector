/* TQ Excel Connector - the company report library and the demo pack.
 *   - Library: layouts saved in NetSuite (tq_library.json next to the TQ script) so everyone in the company can
 *     start from "Group monthly pack". Save the open layout from the Layout screen; use one from there or from
 *     New report. Using one adds its layout sheet and builds its report in this workbook.
 *   - Demo pack: a monthly P&L, a balance sheet and an executive dashboard in one click, for demos and training. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var listed = null;   // the last library list: { items, canManage }

  function el(tag, text, cls) {
    var n = document.createElement(tag);
    if (text) n.textContent = text;
    if (cls) n.className = cls;
    return n;
  }
  function design(op, args) {
    return window.TQChat.workbookContext().then(function (context) {
      return fetch('/api/design', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ op: op, args: args || {}, context: context }) });
    }).then(function (r) { return r.json(); }).then(function (data) {
      if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
      return data;
    });
  }
  function apply(data) {
    if (data.spec) window.TQChat.setSpec(data.spec);
    return window.TQChat.applyActions(data.actions).then(function () {
      if (data.design && window.TQDesigner) TQDesigner.select(data.design.name);
      return data;
    });
  }
  function note(text, bad) {
    var m = $('libMessage');
    if (!m) return;
    m.textContent = text || '';
    m.className = 'status-line' + (bad ? ' bad' : '');
  }

  function load() {
    return TQ.callApi({ action: 'libraryList' }).then(function (d) { listed = d; return d; });
  }

  /* The Layout screen's list: Use, and Remove for your own (or any, for an administrator). */
  function render() {
    var box = $('libList');
    if (!box) return;
    box.textContent = '';
    box.appendChild(el('div', 'Reading the library…', 'empty'));
    load().then(function (d) {
      box.textContent = '';
      if (!d.items.length) { box.appendChild(el('div', 'Nothing shared yet. Save a layout below to start the library.', 'empty')); return; }
      d.items.forEach(function (it) {
        var row = el('div', '', 'item');
        var span = el('span', it.name + ' ');
        span.appendChild(el('small', (it.description ? it.description + ' · ' : '') + it.by + ', ' + it.at));
        span.title = it.name + (it.description ? ': ' + it.description : '');
        row.appendChild(span);
        var use = el('button', 'Use');
        use.title = 'Add this layout to the workbook and build its report';
        use.addEventListener('click', function () { useItem(it.id); });
        row.appendChild(use);
        if (it.mine || d.canManage) {
          var del = el('button', '✕');
          del.title = 'Remove "' + it.name + '" from the library';
          del.setAttribute('aria-label', del.title);
          del.addEventListener('click', function () {
            if (del.textContent !== 'Remove?') {   // Office panels have no confirm(): a second click confirms
              del.textContent = 'Remove?';
              note('Click Remove? again to take "' + it.name + '" out of the library. Workbooks that use it keep their copy.');
              setTimeout(function () { del.textContent = '✕'; }, 5000);
              return;
            }
            TQ.callApi({ action: 'libraryDelete', id: it.id }).then(function () { note('Removed "' + it.name + '".'); render(); },
              function (e) { note(e.message, true); });
          });
          row.appendChild(del);
        }
        box.appendChild(row);
      });
    }).catch(function (e) {
      box.textContent = '';
      box.appendChild(el('div', /Unknown action|libraryList/i.test(e.message)
        ? 'The library needs the TQ script version 7 in NetSuite. Ask your administrator to update it.' : e.message, 'empty'));
    });
  }

  /* Saves the layout open on the Layout screen under a name everyone sees. */
  function save() {
    var name = $('libName').value.trim();
    var layout = $('dLayout').value;
    if (!layout) { note('Open or create a layout first.', true); return; }
    note('Saving…');
    design('libraryExport', { layout: layout }).then(function (data) {
      return TQ.callApi({ action: 'librarySave', name: name || data.title, description: $('libDesc').value.trim(), design: data.design });
    }).then(function (d) {
      note((d.replaced ? 'Replaced' : 'Saved') + ' in the company library. Everyone with the add-in can use it from New report.');
      $('libName').value = '';
      $('libDesc').value = '';
      render();
    }).catch(function (e) { note(e.message, true); });
  }

  /* Adds a library layout to this workbook and builds its report. */
  function useItem(id) {
    var say = window.TQPanel ? TQPanel.say : function () {};
    say('Adding the layout from the library…');
    note('Adding…');
    return TQ.callApi({ action: 'libraryLoad', id: id }).then(function (d) {
      return design('libraryImport', { design: d.item.design, name: d.item.name });
    }).then(apply).then(function (data) {
      note('');
      say(data.reply.split('\n')[0]);
      window.TQNav.replace('report');
      if (window.TQHome) TQHome.showCard();
    }).catch(function (e) { note(e.message, true); say(e.message, true); });
  }

  /* New report: a "Company library" group under the built-in choices (nothing when the library is empty). */
  function fillWizard(box) {
    load().then(function (d) {
      if (!d.items.length) return;
      box.appendChild(el('div', 'Company library', 'group-title'));
      d.items.forEach(function (it) {
        var b = el('button');
        b.type = 'button';
        b.appendChild(el('strong', it.name));
        b.appendChild(el('small', (it.description ? it.description + ' · ' : '') + 'shared by ' + it.by));
        b.addEventListener('click', function () { useItem(it.id); });
        box.appendChild(b);
      });
    }).catch(function () { /* no library (older NetSuite script): the built-in choices are enough */ });
  }

  /* Demo pack: three sheets that show what the add-in does, for the last closed month. */
  function demoPack() {
    var say = window.TQPanel ? TQPanel.say : function () {};
    var month = '';
    var steps = [
      ['wizard', function () { return { template: 'monthytd', month: month, title: 'Demo - Monthly P&L', filters: {}, dropdowns: [] }; }],
      ['wizard', function () { return { template: 'balance', month: month, title: 'Demo - Balance sheet', filters: {}, dropdowns: [] }; }],
      ['dashboard', function () { return { kind: 'executive', month: month, filters: {}, title: 'Demo - Dashboard' }; }]
    ];
    say('Building the demo pack (3 sheets)…');
    var chain = TQ.period('CM', -1).then(function (m) { month = m; }, function () { month = ''; });
    steps.forEach(function (s, n) {
      chain = chain.then(function () {
        say('Building the demo pack: sheet ' + (n + 1) + ' of 3…');
        return design(s[0], s[1]()).then(apply).then(function (data) {
          if (s[0] === 'dashboard' && data.actions && data.actions[0] && window.TQPack) TQPack.placeLogo(data.actions[0].sheetName);
        });
      });
    });
    return chain.then(function () {
      say('Demo pack ready: a monthly P&L, a balance sheet and an executive dashboard. Change the month on any of them to roll it.');
      window.TQNav.replace('report');
      if (window.TQHome) TQHome.showCard();
    }).catch(function (e) { say(e.message, true); });
  }

  function start() {
    if ($('libSave')) $('libSave').addEventListener('click', save);
    if ($('libFold')) $('libFold').addEventListener('toggle', function () { if ($('libFold').open) render(); });
    if ($('menuDemo')) $('menuDemo').addEventListener('click', demoPack);
  }

  window.TQLibrary = { render: render, use: useItem, fillWizard: fillWizard, demoPack: demoPack };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
