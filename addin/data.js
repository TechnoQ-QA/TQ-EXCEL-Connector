/* TQ Excel Connector - Data tab. Brings any NetSuite data into the workbook: a saved search as it
 * is, or a query over record tables built by clicking (or typed as SuiteQL). Each query is listed
 * on the "TQ Queries" sheet and fills its own output sheet, rewritten in place on every run. */
(function () {
  'use strict';

  var state = null;       // { queries, params } as the server last returned it
  var catalog = null;     // common tables: [{ table, label, fields, links }]
  var templates = [];     // ready-made queries: [{ name, note, sql }]
  var fields = {};        // table -> field names shown in the builder
  var busy = false;
  var $ = function (id) { return document.getElementById(id); };

  function say(text, bad) {
    var el = $('qMessage');
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
    return o;
  }

  function post(op, args, context) {
    return fetch('/api/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ op: op, args: args || {}, context: context || {} })
    }).then(function (response) { return response.json(); }).then(function (data) {
      if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
      return data;
    });
  }

  /* A request that reads or changes the workbook's queries; what comes back is written to Excel. */
  function request(op, args, working) {
    if (busy) return Promise.resolve(null);
    busy = true;
    say(working || 'Working…');
    return window.TQChat.workbookContext()
      .then(function (context) { return post(op, args, context); })
      .then(function (data) {
        return window.TQChat.applyActions(data.actions).then(function () {
          if (data.state) state = data.state;
          render();
          say(data.reply);
          return data;
        });
      })
      .catch(function (error) { say(error.message, true); return null; })
      .then(function (result) { busy = false; return result; });
  }

  /* A lookup that does not touch the workbook (search list, table fields, query text). */
  function ask(op, args, working) {
    say(working || 'Working…');
    return post(op, args).then(function (data) { say(''); return data; })
      .catch(function (error) { say(error.message, true); return null; });
  }

  function render() {
    var list = $('qList');
    list.textContent = '';
    var queries = (state && state.queries) || [];
    queries.forEach(function (q) {
      var row = el('div', '', 'item');
      var span = el('span', q.name);
      var detail = (q.type === 'search' ? 'saved search' : 'query') + (q.last ? ', ' + (q.rows === '' ? q.last : q.rows + ' rows, ' + q.last) : '');
      span.appendChild(document.createTextNode(' '));
      span.appendChild(el('small', detail));
      span.title = q.source;
      row.appendChild(span);
      [['Run', 'run'], ['Edit', 'edit'], ['✕', 'remove']].forEach(function (b) {
        var button = el('button', b[0]);
        button.setAttribute('aria-label', b[0] === '✕' ? 'Remove ' + q.name : b[0] + ' ' + q.name);
        button.addEventListener('click', function () {
          if (b[1] === 'edit') {
            $('qName').value = q.name;
            $('qSql').value = q.source;
            $('qMax').value = q.max;
            $('qSqlForm').open = true;
            $('qSql').focus();
          } else {
            request(b[1], { name: q.name }, b[1] === 'run' ? 'Asking NetSuite…' : 'Working…');
          }
        });
        row.appendChild(button);
      });
      list.appendChild(row);
    });
    if (!queries.length) list.appendChild(el('div', 'No data queries in this workbook yet. Add one below.', 'empty'));
    $('qRefresh').disabled = !queries.length;

    var params = $('qParams');
    params.textContent = '';
    ((state && state.params) || []).forEach(function (p) {
      var row = el('div', '', 'item');
      var span = el('span', '{' + p.name + '}');
      span.appendChild(document.createTextNode(' '));
      span.appendChild(el('small', p.value));
      row.appendChild(span);
      var button = el('button', '✕');
      button.setAttribute('aria-label', 'Remove parameter ' + p.name);
      button.addEventListener('click', function () { request('setParam', { name: p.name, value: '' }); });
      row.appendChild(button);
      params.appendChild(row);
    });
    if (!params.children.length) params.appendChild(el('div', 'None. A query uses a parameter as {name}.', 'empty'));
  }

  // ------------------------------------------------------------ saved searches --

  function findSearches() {
    var box = $('qSearchResults');
    box.textContent = '';
    ask('searches', { text: $('qSearchText').value }, 'Reading the saved searches…').then(function (data) {
      if (!data) return;
      data.searches.slice(0, 80).forEach(function (s) {
        var button = el('button', s.title + '  (' + s.type + ')');
        button.title = 'Bring "' + s.title + '" (' + s.id + ') onto its own sheet';
        button.addEventListener('click', function () {
          request('save', { name: s.title, type: 'search', source: s.id }, 'Running the saved search…');
        });
        box.appendChild(button);
      });
      if (!data.searches.length) box.appendChild(el('div', 'No saved search matched.', 'empty'));
    });
  }

  // ------------------------------------------------------------- query builder --

  function currentTable() {
    return ($('qTableOther').value || $('qTable').value || '').trim().toLowerCase();
  }

  function entry(table) {
    return (catalog || []).filter(function (c) { return c.table === table; })[0] || { table: table, label: table, fields: [], links: [] };
  }

  /* One tick-box row per field, with what to show for it. */
  function fieldRows(box, table, names) {
    box.textContent = '';
    names.forEach(function (name) {
      var row = el('div', '', 'fieldrow');
      var id = 'f_' + table + '_' + name;
      var check = el('input');
      check.type = 'checkbox';
      check.id = id;
      check.value = table + '.' + name;
      var label = el('label', name, 'check');
      label.setAttribute('for', id);
      var how = el('select');
      how.setAttribute('aria-label', 'How to show ' + name);
      [['', 'as is'], ['name', 'show name'], ['sum', 'sum'], ['count', 'count'], ['min', 'smallest'], ['max', 'largest']]
        .forEach(function (o) { option(how, o[0], o[1]); });
      row.appendChild(check);
      row.appendChild(label);
      row.appendChild(how);
      box.appendChild(row);
    });
    if (!names.length) box.appendChild(el('div', 'No fields known yet. Press "Read all fields".', 'empty'));
  }

  function allFieldNames() {
    var table = currentTable();
    var joined = $('qJoin').value ? JSON.parse($('qJoin').value).table : '';
    var out = (fields[table] || []).map(function (f) { return table + '.' + f; });
    if (joined) out = out.concat((fields[joined] || []).map(function (f) { return joined + '.' + f; }));
    return out;
  }

  function fillFieldSelects() {
    ['qFilterField1', 'qFilterField2', 'qSort'].forEach(function (id) {
      var select = $(id);
      var keep = select.value;
      select.textContent = '';
      option(select, '', id === 'qSort' ? '(no sorting)' : '(no filter)');
      allFieldNames().forEach(function (name) { option(select, name); });
      select.value = keep;
    });
  }

  function showTable() {
    var table = currentTable();
    var info = entry(table);
    if (!fields[table]) fields[table] = info.fields.slice();
    fieldRows($('qFields'), table, fields[table]);
    var join = $('qJoin');
    join.textContent = '';
    option(join, '', '(none)');
    (info.links || []).forEach(function (link) {
      option(join, JSON.stringify({ table: link[1], left: link[0], right: link[2] }), entry(link[1]).label + '  (' + link[0] + ')');
    });
    showJoin();
  }

  function showJoin() {
    var join = $('qJoin').value ? JSON.parse($('qJoin').value) : null;
    $('qJoinFields').hidden = !join;
    if (join) {
      if (!fields[join.table]) fields[join.table] = entry(join.table).fields.slice();
      fieldRows($('qJoinFields'), join.table, fields[join.table]);
    }
    fillFieldSelects();
  }

  function readAllFields() {
    var table = currentTable();
    if (!table) return;
    ask('columns', { table: table }, 'Reading the fields of ' + table + '…').then(function (data) {
      if (!data) return;
      if (!data.columns.length) { say('That table has no rows yet, so its fields cannot be read.', true); return; }
      fields[table] = data.columns;
      showTable();
    });
  }

  function compose() {
    var picked = [];
    Array.prototype.forEach.call(document.querySelectorAll('#qFields .fieldrow, #qJoinFields .fieldrow'), function (row) {
      var check = row.querySelector('input');
      var how = row.querySelector('select').value;
      if (!check.checked || row.closest('[hidden]')) return;
      picked.push({ field: check.value, show: how === 'name' ? 'name' : '', total: how === 'name' ? '' : how });
    });
    var filters = [1, 2].map(function (n) {
      return { field: $('qFilterField' + n).value, op: $('qFilterOp' + n).value, value: $('qFilterValue' + n).value };
    }).filter(function (f) { return f.field; });
    var spec = {
      table: currentTable(), fields: picked, filters: filters,
      joins: $('qJoin').value ? [JSON.parse($('qJoin').value)] : [],
      sort: $('qSort').value ? [{ field: $('qSort').value, desc: $('qDesc').checked }] : []
    };
    ask('compose', { spec: spec }).then(function (data) {
      if (!data) return;
      $('qSql').value = data.sql;
      if (!$('qName').value) $('qName').value = entry(spec.table).label;
      $('qSqlForm').open = true;
      say('The query is written below. Check it, give it a name, then press Add and run.');
    });
  }

  function showTemplate() {
    var chosen = templates[Number($('qTemplate').value)];
    $('qTemplateNote').textContent = chosen ? chosen.note : '';
  }

  function useTemplate() {
    var chosen = templates[Number($('qTemplate').value)];
    if (!chosen) return;
    $('qName').value = chosen.name;
    $('qSql').value = chosen.sql;
    $('qSqlForm').open = true;
    var needs = (chosen.sql.match(/\{\w+\}/g) || []).filter(function (p, n, all) { return all.indexOf(p) === n; });
    say('The query is below. ' + (needs.length ? 'It uses ' + needs.join(' and ') + ': set them under Parameters (dates as 2026-01-31), then ' : 'Check it, then ') + 'press Add and run.');
  }

  function saveQuery() {
    request('save', { name: $('qName').value, type: 'sql', source: $('qSql').value, max: $('qMax').value }, 'Asking NetSuite…')
      .then(function (data) { if (data) $('qName').value = ''; });
  }

  function loadCatalog() {
    if (catalog) return;
    post('catalog').then(function (data) {
      catalog = data.catalog;
      templates = data.templates || [];
      var pick = $('qTemplate');
      pick.textContent = '';
      templates.forEach(function (t, n) { option(pick, String(n), t.name); });
      showTemplate();
      var select = $('qTable');
      select.textContent = '';
      catalog.forEach(function (c) { option(select, c.table, c.label); });
      showTable();
    }).catch(function (e) { TQPanel.say('The ready-made queries could not be loaded: ' + e.message, true); });
  }

  /* Called when the tab is shown: reads the workbook's queries. */
  function load() {
    loadCatalog();
    return request('get', {}, 'Reading the queries…');
  }

  function start() {
    $('qRefresh').addEventListener('click', function () { request('run', {}, 'Asking NetSuite…'); });
    $('qSearchGo').addEventListener('click', findSearches);
    $('qSearchText').addEventListener('keydown', function (event) { if (event.key === 'Enter') findSearches(); });
    $('qTable').addEventListener('change', function () { $('qTableOther').value = ''; showTable(); });
    $('qTableOther').addEventListener('change', showTable);
    $('qJoin').addEventListener('change', showJoin);
    $('qMoreFields').addEventListener('click', readAllFields);
    $('qCompose').addEventListener('click', compose);
    $('qSave').addEventListener('click', saveQuery);
    $('qTemplate').addEventListener('change', showTemplate);
    $('qTemplateUse').addEventListener('click', useTemplate);
    $('pSave').addEventListener('click', function () {
      request('setParam', { name: $('pName').value, value: $('pValue').value }).then(function (data) {
        if (data) { $('pName').value = ''; $('pValue').value = ''; }
      });
    });
  }

  /* New report > Operational reports: one of the ready-made queries on its own sheet, with its dates set as
   * parameters first ({from}, {to} as 2026-01-31). */
  function runTemplate(name, params) {
    var go = catalog ? Promise.resolve() : post('catalog').then(function (data) { catalog = data.catalog; templates = data.templates || []; });
    return go.then(function () {
      var chosen = templates.filter(function (t) { return t.name === name; })[0];
      if (!chosen) throw new Error('"' + name + '" is not one of the ready-made reports.');
      var chain = Promise.resolve();
      Object.keys(params || {}).forEach(function (k) {
        if (chosen.sql.indexOf('{' + k + '}') === -1) return;
        chain = chain.then(function () { return request('setParam', { name: k, value: params[k] }, 'Setting ' + k + '…'); });
      });
      return chain.then(function () {
        return request('save', { name: chosen.name, type: 'sql', source: chosen.sql, max: '5000' }, 'Asking NetSuite…');
      });
    });
  }

  window.TQData = { load: load, runTemplate: runTemplate };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
