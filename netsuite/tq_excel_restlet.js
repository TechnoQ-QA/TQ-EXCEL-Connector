/**
 * TQ Excel Connector - RESTlet entry point.
 *
 * Two kinds of request arrive here:
 *
 * - Balances and the lists behind them. A RESTlet always runs as the caller's
 *   own role, which hides transaction types that role cannot view and so
 *   understates balances. These are handed to the elevated Suitelet
 *   (tq_excel_elevated.js), where the work and the access check happen
 *   (tq_excel_core.js).
 *
 * - Data queries for the add-in's Data tab: saved searches and SuiteQL
 *   ('searchList' | 'searchRun' | 'sql' | 'tableColumns'). These run right
 *   here, as the caller's own role, so each user gets exactly the records and
 *   fields their NetSuite role lets them see, as in NetSuite itself.
 *
 * @NApiVersion 2.1
 * @NScriptType Restlet
 */
define(['N/https', 'N/query', 'N/search'], (https, query, search) => {
  const ACTIONS = ['searchList', 'searchRun', 'sql', 'tableColumns'];
  const MAX_ROWS = 20000;
  const DEFAULT_ROWS = 5000;
  const MAX_SEARCHES = 4000;

  const clean = (v) => (v === null || v === undefined ? '' : String(v).trim());
  const rowLimit = (n) => Math.max(1, Math.min(MAX_ROWS, Math.floor(Number(n)) || DEFAULT_ROWS));

  /** Saved searches the caller can see: [{ id: 'customsearch_x', title, type }]. */
  const searchList = (request) => {
    const words = clean(request.text).toLowerCase().split(/\s+/).filter(Boolean);
    const out = [];
    const paged = search
      .create({ type: search.Type.SAVED_SEARCH, columns: ['title', 'id', 'recordtype'] })
      .runPaged({ pageSize: 1000 });
    let seen = 0;
    paged.pageRanges.forEach((range) => {
      if (seen >= MAX_SEARCHES) return;
      paged.fetch({ index: range.index }).data.forEach((r) => {
        seen += 1;
        const item = {
          id: clean(r.getValue('id')) || clean(r.id),
          title: clean(r.getValue('title')),
          type: clean(r.getText('recordtype')) || clean(r.getValue('recordtype'))
        };
        const hay = (item.id + ' ' + item.title + ' ' + item.type).toLowerCase();
        if (words.every((w) => hay.indexOf(w) !== -1)) out.push(item);
      });
    });
    out.sort((a, b) => a.title.localeCompare(b.title));
    return { searches: out.slice(0, 500), truncated: out.length > 500 };
  };

  /** Runs a saved search as it is defined: its own columns, filters and sorting. */
  const searchRun = (request) => {
    const id = clean(request.id);
    if (!id) throw new Error('Saved search ID is required');
    let loaded;
    try {
      loaded = search.load({ id });
    } catch (e) {
      throw new Error('Saved search "' + id + '" could not be opened: ' + (e.message || String(e)));
    }
    const columns = loaded.columns;
    const max = rowLimit(request.maxRows);
    const out = [];
    let truncated = false;
    const paged = loaded.runPaged({ pageSize: 1000 });
    for (let p = 0; p < paged.pageRanges.length && !truncated; p += 1) {
      paged.fetch({ index: paged.pageRanges[p].index }).data.forEach((r) => {
        if (out.length >= max) {
          truncated = true;
          return;
        }
        out.push(
          columns.map((c) => {
            const shown = r.getText(c);
            return shown !== null && shown !== undefined && shown !== '' ? shown : r.getValue(c);
          })
        );
      });
    }
    return {
      columns: columns.map((c) => clean(c.label) || clean(c.name)),
      rows: out,
      truncated,
      total: paged.count
    };
  };

  const selectOnly = (text) => {
    const sql = clean(text).replace(/;\s*$/, '');
    if (!/^(select|with)\b/i.test(sql)) throw new Error('A query must start with SELECT');
    if (sql.indexOf(';') !== -1) throw new Error('Only one query at a time: remove the ";"');
    return sql;
  };

  /** Runs one SuiteQL SELECT. params fill the ? placeholders in order. */
  const sql = (request) => {
    const text = selectOnly(request.query);
    const max = rowLimit(request.maxRows);
    const params = Array.isArray(request.params) ? request.params : [];
    const paged = query.runSuiteQLPaged({ query: text, params, pageSize: 1000 });
    const out = [];
    let columns = null;
    let truncated = false;
    for (let p = 0; p < paged.pageRanges.length && !truncated; p += 1) {
      paged.fetch({ index: paged.pageRanges[p].index }).data.asMappedResults().forEach((row) => {
        if (!columns) columns = Object.keys(row);
        if (out.length >= max) {
          truncated = true;
          return;
        }
        out.push(columns.map((k) => row[k]));
      });
    }
    return { columns: columns || [], rows: out, truncated, total: paged.count };
  };

  /** The column names of a table, read from one of its rows. */
  const tableColumns = (request) => {
    const table = clean(request.table);
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(table)) throw new Error('Table name "' + table + '" is not valid');
    const found = query.runSuiteQL({ query: 'SELECT * FROM ' + table + ' WHERE ROWNUM <= 1' }).asMappedResults();
    return { table: table.toLowerCase(), columns: found.length ? Object.keys(found[0]) : [], empty: !found.length };
  };

  /** Saved searches and SuiteQL, run here as the caller's own role. */
  const runDataQuery = (body) => {
    try {
      const request = body || {};
      switch (request.action) {
        case 'searchList':
          return Object.assign({ ok: true }, searchList(request));
        case 'searchRun':
          return Object.assign({ ok: true }, searchRun(request));
        case 'sql':
          return Object.assign({ ok: true }, sql(request));
        case 'tableColumns':
          return Object.assign({ ok: true }, tableColumns(request));
        default:
          throw new Error('Unknown action "' + clean(request.action) + '"');
      }
    } catch (e) {
      return { ok: false, error: e.message || String(e) };
    }
  };

  const post = (body) => {
    const request = body || {};
    if (ACTIONS.indexOf(request.action) !== -1) {
      return runDataQuery(request);
    }
    try {
      const response = https.requestSuitelet({
        scriptId: 'customscript_tq_excel_elevated',
        deploymentId: 'customdeploy_tq_excel_elevated',
        method: https.Method.POST,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(request)
      });
      return JSON.parse(response.body);
    } catch (e) {
      return { ok: false, error: 'Elevated runner failed: ' + (e.message || String(e)) };
    }
  };

  return { post };
});
