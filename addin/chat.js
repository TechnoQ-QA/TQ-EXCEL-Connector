/* TQ Excel Connector - chat tab. Sends each message to the local server's built-in
 * command reader and writes the report or list it returns into the workbook. */
(function () {
  'use strict';

  var lastSpec = null; // the previous report, so "by quarter" or "now for 2025" can change it
  var busy = false;
  var $ = function (id) { return document.getElementById(id); };
  var inExcel = function () { return typeof Excel !== 'undefined' && typeof Office !== 'undefined' && !!Office.context && !!Office.context.document; };
  var supports = function (version) {
    try { return Office.context.requirements.isSetSupported('ExcelApi', version); } catch (e) { return false; }
  };

  function bubble(kind, text) {
    var hint = $('chatHint');
    if (hint) hint.remove();
    var el = document.createElement('div');
    el.className = 'bubble ' + kind;
    el.textContent = text;
    var log = $('chatLog');
    log.appendChild(el);
    log.scrollTop = log.scrollHeight;
    return el;
  }

  var QUERY_SHEET = 'tq queries'; // the sheet that lists the workbook's data queries
  var CHUNK_CELLS = 30000; // a larger table is written in pieces, or Excel refuses the request
  var LAYOUT_PREFIX = 'tq layout'; // sheets whose name starts like this describe a custom report format

  /* The selection, the open sheet and the cells of every layout sheet, for the server to read. */
  function workbookContext() {
    if (!inExcel()) return Promise.resolve({});
    return Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      range.load('address');
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      var active = sheets.getActiveWorksheet();
      active.load('name');
      return context.sync().then(function () {
        var read = function (s) {
          var used = s.getUsedRangeOrNullObject(true);
          used.load('values,rowIndex');
          return { name: s.name, used: used };
        };
        var found = sheets.items.filter(function (s) {
          return s.name.toLowerCase().indexOf(LAYOUT_PREFIX) === 0;
        }).slice(0, 10).map(read);
        var queries = sheets.items.filter(function (s) { return s.name.toLowerCase() === QUERY_SHEET; }).map(read)[0];
        return context.sync().then(function () {
          return {
            selection: range.address,
            activeSheet: active.name,
            sheets: sheets.items.map(function (s) { return s.name; }),
            queries: queries && !queries.used.isNullObject
              ? { row: queries.used.rowIndex, rows: queries.used.values.slice(0, 300).map(function (row) { return row.slice(0, 12); }) }
              : null,
            layouts: found.filter(function (f) { return !f.used.isNullObject; }).map(function (f) {
              return {
                name: f.name,
                row: f.used.rowIndex,
                rows: f.used.values.slice(0, 1000).map(function (row) { return row.slice(0, 24); })
              };
            })
          };
        });
      });
    }).catch(function () { return {}; });
  }

  function freeName(wanted, taken) {
    var base = (wanted || 'TQ Report').slice(0, 31);
    var name = base;
    for (var n = 2; taken.indexOf(name.toLowerCase()) !== -1; n++) {
      var suffix = ' ' + n;
      name = base.slice(0, 31 - suffix.length) + suffix;
    }
    taken.push(name.toLowerCase());
    return name;
  }

  function matrix(rows, cols, value) {
    var out = [];
    for (var r = 0; r < rows; r++) {
      var line = [];
      for (var c = 0; c < cols; c++) line.push(value);
      out.push(line);
    }
    return out;
  }

  /* A block of the written area: [first row, first column, last row, last column], counted from 0. */
  function block(start, b) {
    return start.getOffsetRange(b[0], b[1]).getResizedRange(b[2] - b[0], b[3] - b[1]);
  }

  function edge(range, side, kind) {
    var border = range.format.borders.getItem(side);
    border.style = kind === 'double' ? 'Double' : 'Continuous';
  }

  /* Fonts, borders, number formats, widths and frozen panes. Runs after the cells are in. */
  function decorate(start, sheet, action) {
    var rows = action.rows.length;
    var cols = action.rows[0].length;
    (action.formats || []).forEach(function (f) {
      var range = block(start, [f.r0, f.c0, f.r1, f.c1]);
      if (f.numberFormat) range.numberFormat = matrix(f.r1 - f.r0 + 1, f.c1 - f.c0 + 1, f.numberFormat);
      if (f.bold) range.format.font.bold = true;
      if (f.italic) range.format.font.italic = true;
      if (f.size) range.format.font.size = f.size;
      if (f.color) range.format.font.color = f.color;
      if (f.fill) range.format.fill.color = f.fill;
      if (f.align === 'left') range.format.horizontalAlignment = 'Left';
      if (f.align === 'right') range.format.horizontalAlignment = 'Right';
      if (f.indent && supports('1.9')) range.format.indentLevel = f.indent;
      if (f.wrap) range.format.wrapText = true;
      if (f.merge) range.merge(false);
      if (f.top) edge(range, 'EdgeTop', f.top);
      if (f.bottom) edge(range, 'EdgeBottom', f.bottom);
    });
    if (action.boldFirstRow) block(start, [0, 0, 0, cols - 1]).format.font.bold = true;
    if (action.widths) {
      action.widths.forEach(function (w) { block(start, [0, w.c0, 0, w.c1]).format.columnWidth = w.width; });
    } else {
      block(start, [0, 0, rows - 1, cols - 1]).format.autofitColumns();
    }
    if (action.gridlines === false && supports('1.8')) sheet.showGridlines = false;
    var freeze = action.freeze;
    if (freeze && action.target !== 'selection' && freeze.rows && supports('1.7')) {
      if (freeze.cols) sheet.freezePanes.freezeAt(sheet.getRange('A1').getResizedRange(freeze.rows - 1, freeze.cols - 1));
      else sheet.freezePanes.freezeRows(freeze.rows);
    }
  }

  /* Dropdown lists on the filter cells and collapsible row groups. Runs last, on its own, so a
   * failure here (an older Excel) leaves the report and its formatting in place. */
  function extras(start, action, sheet) {
    var added = false;
    // Charts of a dashboard: Excel's own, reading the sheet's cells, so they follow Refresh and the month in B8.
    if (sheet && supports('1.1')) {
      (action.charts || []).forEach(function (c) {
        var chart = sheet.charts.add(c.type, sheet.getRange(c.range), 'Columns');
        chart.title.text = c.title;
        chart.name = 'TQ Chart ' + (c.slot + 1);
        chart.setPosition(c.topLeft, c.bottomRight);
        chart.legend.position = 'Bottom';
        if (c.type === 'Pie') chart.legend.position = 'Right';
        added = true;
      });
    }
    if (supports('1.8')) {
      (action.dropdowns || []).forEach(function (d) {
        var cell = block(start, [d.r, d.c, d.r1 || d.r, d.c]);   // r1: the same list down a column
        cell.dataValidation.rule = { list: { inCellDropDown: true, source: d.formula } };
        // A filter cell also takes typed text such as "A, B" or "not A", so a value not in the list is allowed.
        if (d.free) cell.dataValidation.errorAlert = { showAlert: false, style: 'Information', title: '', message: '' };
        added = true;
      });
    }
    if (supports('1.6')) {
      // Variance highlights: green at or above the limit, red at or below minus it.
      (action.conditional || []).forEach(function (c) {
        var range = block(start, [c.r0, c.c0, c.r1, c.c1]);
        if (c.formula) {   // a highlight rule of the layout: its own test and colours
          var own = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          own.custom.rule.formula = c.formula;
          own.custom.format.font.color = c.font;
          own.custom.format.fill.color = c.fill;
          added = true;
          return;
        }
        // A formula rule, so blank ("") cells stay plain: Excel ranks text above every number.
        [['>=' + c.limit, '#1A7F4B', '#E6F4EA'], ['<=' + (-c.limit), '#B3261E', '#FCE8E6']].forEach(function (rule) {
          var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cf.custom.rule.formula = '=AND(ISNUMBER(' + c.cell + '),' + c.cell + rule[0] + ')';
          cf.custom.format.font.color = rule[1];
          cf.custom.format.fill.color = rule[2];
        });
        added = true;
      });
    }
    if (supports('1.10')) {
      (action.groups || []).forEach(function (g) {
        block(start, [g[0], 0, g[1], 0]).getEntireRow().group('ByRows');
        added = true;
      });
    }
    return added;
  }

  /* One small change in its own request, so that if this Excel cannot do it nothing else is lost. */
  function quiet(change) {
    return Excel.run(function (context) { change(context); return context.sync(); }).catch(function () {});
  }

  /* Before a sheet is rewritten in place: drop its frozen panes, dropdown lists and row groups.
   * Cell contents and formats are cleared with the write itself. */
  function cleanSheet(name) {
    var sheet = function (context) { return context.workbook.worksheets.getItem(name); };
    return quiet(function (c) { if (supports('1.7')) sheet(c).freezePanes.unfreeze(); })
      .then(function () { return quiet(function (c) { if (supports('1.6')) sheet(c).getRange().conditionalFormats.clearAll(); }); })
      .then(function () { return quiet(function (c) { sheet(c).getRange().unmerge(); }); })
      .then(function () {   // a dashboard rebuilt in place: its old charts go first
        return Excel.run(function (c) {
          var charts = sheet(c).charts;
          charts.load('items/name');
          return c.sync().then(function () {
            charts.items.forEach(function (ch) { if (/^TQ Chart /.test(ch.name)) ch.delete(); });
            return c.sync();
          });
        }).catch(function () {});
      })
      .then(function () { return quiet(function (c) { if (supports('1.8')) sheet(c).getRange().dataValidation.clear(); }); })
      .then(function () { return quiet(function (c) { if (supports('1.10')) sheet(c).getUsedRange().getEntireRow().ungroup('ByRows'); }); });
  }

  /* Sheets that a "sheet" write will replace, cleaned one after another. */
  function cleanTargets(writes) {
    var wanted = writes.filter(function (a) { return a.target === 'sheet'; }).map(function (a) { return a.sheetName.toLowerCase(); });
    if (!wanted.length) return Promise.resolve();
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        return sheets.items.map(function (s) { return s.name; }).filter(function (n) { return wanted.indexOf(n.toLowerCase()) !== -1; });
      });
    }).then(function (names) {
      return names.reduce(function (chain, name) { return chain.then(function () { return cleanSheet(name); }); }, Promise.resolve());
    }).catch(function () {});
  }

  /* Applies the writes in order; resolves with a line describing each one.
   * target: "new_sheet" (always a new sheet), "sheet" (the sheet of that name, rewritten in place,
   * or created) or "selection". */
  function applyActions(actions) {
    var writes = (actions || []).filter(function (a) { return a && a.type === 'write' && a.rows && a.rows.length; });
    if (!writes.length) return Promise.resolve([]);
    if (!inExcel()) {
      return Promise.resolve(writes.map(function (a) {
        var where = a.target === 'new_sheet' ? 'a new sheet' : a.target === 'sheet' ? 'sheet "' + a.sheetName + '"' : 'the selection';
        return 'Would write ' + a.rows.length + ' row(s) to ' + where + ' (open the panel in Excel).';
      }));
    }
    return backupSheets(writes).then(function () { return keepComments(writes); }).then(function () { return cleanTargets(writes); })
      .then(function () { return makeSheets(writes); }).then(function () { return writeAll(writes); });
  }

  /* Rebuilding a report rewrites its sheet, and Excel cannot undo that. So first each report sheet that is about
   * to be rewritten is copied to a hidden sheet ("TQ Backup ..."), one per report, replacing the older copy. The
   * This report screen offers Restore previous version, which puts the copy's cells back in place (the sheet
   * itself stays, so formulas on other sheets that point at it keep working). The list is in the workbook's
   * settings: tqBackups = { sheet name: { backup, at } }. */
  var BACKUP_MAX = 12;   // sheets copied per rebuild (a pack with many location sheets keeps the first ones)
  function backups() {
    try { return Office.context.document.settings.get('tqBackups') || {}; } catch (e) { return {}; }
  }
  function saveBackups(map) {
    try {
      Office.context.document.settings.set('tqBackups', map);
      Office.context.document.settings.saveAsync(function () {});
    } catch (e) { /* outside Excel */ }
  }
  function backupSheets(writes) {
    if (!supports('1.7')) return Promise.resolve();
    var targets = writes.filter(function (a) { return a.target === 'sheet' && !/^tq (layout|backup|queries)/i.test(a.sheetName); });
    if (!targets.length) return Promise.resolve();
    var map = backups();
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var names = sheets.items.map(function (s) { return s.name; });
        var lower = names.map(function (n) { return n.toLowerCase(); });
        var made = [];
        targets.slice(0, BACKUP_MAX).forEach(function (a, n) {
          var at = lower.indexOf(a.sheetName.toLowerCase());
          if (at === -1) return;   // a new sheet: nothing to keep
          var old = map[names[at]];
          if (old && lower.indexOf(String(old.backup).toLowerCase()) !== -1) sheets.getItem(old.backup).delete();
          var copy = sheets.getItem(names[at]).copy('End');
          var name = ('TQ Backup ' + Date.now().toString(36) + n).slice(0, 31);
          copy.name = name;
          copy.visibility = 'Hidden';
          made.push([names[at], name]);
        });
        return context.sync().then(function () {
          made.forEach(function (m) { map[m[0]] = { backup: m[1], at: Date.now() }; });
          if (made.length) saveBackups(map);
        });
      });
    }).catch(function () { /* a backup that cannot be made never stops the rebuild */ });
  }

  /* Puts the hidden copy's cells back on the sheet. Returns a promise of a sentence for the panel. */
  function restoreSheet(name) {
    var map = backups();
    var found = map[name] || map[Object.keys(map).filter(function (k) { return k.toLowerCase() === String(name).toLowerCase(); })[0]];
    if (!found) return Promise.reject(new Error('There is no earlier version of this sheet.'));
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getItem(name);
      var copy = context.workbook.worksheets.getItem(found.backup);
      var used = copy.getUsedRange(true);
      used.load('address');
      return context.sync().then(function () {
        sheet.getRange().clear();
        try { sheet.getRange().conditionalFormats.clearAll(); } catch (e) { /* older Excel */ }
        sheet.getRange(used.address.split('!').pop()).copyFrom(used, 'All');
        return context.sync();
      });
    }).then(function () {
      var when = new Date(found.at);
      return 'Restored the version from ' + when.toLocaleDateString([], { day: 'numeric', month: 'short' }) + ' ' +
        when.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) + '. Refresh to bring its numbers up to date.';
    });
  }

  /* A report with a Comments column: what was typed there is read before the sheet is rewritten and put back
   * on the row of the same line name. A line that was removed takes its comment with it. */
  function keepComments(writes) {
    var keeping = writes.filter(function (a) { return a.keep && a.target === 'sheet'; });
    if (!keeping.length) return Promise.resolve();
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var names = sheets.items.map(function (s) { return s.name.toLowerCase(); });
        var reads = keeping.filter(function (a) { return names.indexOf(a.sheetName.toLowerCase()) !== -1; }).map(function (a) {
          var used = sheets.getItem(a.sheetName).getUsedRange(true);
          used.load('values');
          return { action: a, used: used };
        });
        return context.sync().then(function () {
          reads.forEach(function (r) {
            var old = r.used.values, k = r.action.keep, rows = r.action.rows;
            var head = -1, col = -1;
            for (var i = 0; i < old.length && head < 0; i++) {
              var at = old[i].indexOf(k.title);
              if (at !== -1) { head = i; col = at; }
            }
            var target = rows[k.header].indexOf(k.title);
            if (head < 0 || target < 0) return;
            var notes = {};
            for (var n = head + 1; n < old.length; n++) {
              var label = String(old[n][k.key] || '').trim().toLowerCase();
              if (label && String(old[n][col] || '').trim() && !(label in notes)) notes[label] = old[n][col];
            }
            for (var m = k.header + 1; m < rows.length; m++) {
              var name = String(rows[m][k.key] || '').trim().toLowerCase();
              if (name in notes) rows[m][target] = notes[name];
            }
          });
        });
      });
    }).catch(function (e) {   // rebuilding without them would lose what people typed: stop instead
      throw new Error('The comments on the report could not be read, so it was not rebuilt (the comments are safe). ' +
        'Try again. ' + (e && e.message ? e.message : ''));
    });
  }

  /* Every named sheet is made, in order, before any formula is written: a summary sheet points at the
   * location sheets that come after it, and a formula written before its sheet exists stays #REF! for good. */
  function makeSheets(writes) {
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var taken = sheets.items.map(function (s) { return s.name.toLowerCase(); });
        writes.forEach(function (a) {
          if (a.target === 'sheet' && taken.indexOf(a.sheetName.toLowerCase()) === -1) {
            sheets.add(a.sheetName);
            taken.push(a.sheetName.toLowerCase());
            a.created = true;
          }
        });
        return context.sync();
      });
    });
  }

  function writeAll(writes) {
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var taken = sheets.items.map(function (s) { return s.name.toLowerCase(); });
        var done = [];
        var placed = [];
        var pieces = []; // parts of large tables, written one request each after the sheets exist
        writes.forEach(function (a) {
          var start;
          var sheet;
          var where;
          if (a.target === 'new_sheet') {
            var name = freeName(a.sheetName, taken);
            sheet = sheets.add(name);
            start = sheet.getRange(a.startCell || 'A1');
            sheet.activate();
            where = 'new sheet "' + name + '"';
          } else if (a.target === 'sheet') {
            var exists = !a.created;
            sheet = sheets.getItem(a.sheetName);
            if (exists) sheet.getRange().clear();
            start = sheet.getRange('A1');
            if (a.activate !== false) sheet.activate();
            where = (exists ? 'sheet "' : 'new sheet "') + a.sheetName + (exists ? '" (updated in place)' : '"');
          } else {
            sheet = context.workbook.worksheets.getActiveWorksheet();
            start = a.startCell ? sheet.getRange(a.startCell) : context.workbook.getSelectedRange().getCell(0, 0);
            where = a.startCell ? a.startCell + ' on the active sheet' : 'the selected cell';
          }
          // Period names and account numbers must stay text, or Excel turns "Jan 2026" into a date.
          (a.textRanges || []).forEach(function (b) {
            block(start, b).numberFormat = matrix(b[2] - b[0] + 1, b[3] - b[1] + 1, '@');
          });
          var width = a.rows[0].length;
          if (a.rows.length * width <= CHUNK_CELLS) {
            block(start, [0, 0, a.rows.length - 1, width - 1]).formulas = a.rows;
          } else {
            var step = Math.max(1, Math.floor(CHUNK_CELLS / width));
            for (var at = 0; at < a.rows.length; at += step) {
              pieces.push({ start: start, at: at, rows: a.rows.slice(at, at + step), width: width });
            }
          }
          placed.push({ start: start, sheet: sheet, action: a });
          if (a.activate !== false) done.push('Wrote ' + a.rows.length + ' row(s) × ' + a.rows[0].length + ' column(s) to ' + where + '.');
        });
        return context.sync().then(function () {
          return pieces.reduce(function (chain, piece) {
            return chain.then(function () {
              block(piece.start, [piece.at, 0, piece.at + piece.rows.length - 1, piece.width - 1]).formulas = piece.rows;
              return context.sync();
            });
          }, Promise.resolve());
        }).then(function () {
          placed.forEach(function (p) { decorate(p.start, p.sheet, p.action); });
          return context.sync().then(
            function () {
              var any = placed.map(function (p) { return extras(p.start, p.action, p.sheet); }).some(Boolean);
              if (!any) return done;
              return context.sync().then(
                function () { return done; },
                function () { return done.concat('The report is in, but the dropdown lists or row groups could not be added.'); }
              );
            },
            function () { return done.concat('The numbers are in, but the formatting could not be applied.'); }
          );
        });
      });
    });
  }

  function setBusy(value) {
    busy = value;
    $('chatSend').disabled = value;
    $('chatSend').textContent = value ? '…' : 'Send';
  }

  /* A picture for the AI add-on: attached with the Picture button or pasted into the message box. */
  var attachment = null; // { mediaType, data (base64), name }
  var MAX_SIDE = 1568;

  function showAttachment() {
    $('chatAttachment').hidden = !attachment;
    $('chatAttachmentName').textContent = attachment ? 'Picture attached: ' + attachment.name : '';
  }

  function attach(file) {
    if (!file || !/^image\//.test(file.type)) { bubble('error', 'That file is not a picture.'); return; }
    var reader = new FileReader();
    reader.onerror = function () { bubble('error', 'The picture could not be read.'); };
    reader.onload = function () {
      var image = new Image();
      image.onerror = function () { bubble('error', 'The picture could not be read.'); };
      image.onload = function () {
        // Large screenshots are scaled down: enough to read the text, and quicker to send.
        var scale = Math.min(1, MAX_SIDE / Math.max(image.width, image.height));
        var canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(image.width * scale));
        canvas.height = Math.max(1, Math.round(image.height * scale));
        var pen = canvas.getContext('2d');
        pen.fillStyle = '#ffffff';
        pen.fillRect(0, 0, canvas.width, canvas.height);
        pen.drawImage(image, 0, 0, canvas.width, canvas.height);
        var url = canvas.toDataURL('image/jpeg', 0.9);
        attachment = { mediaType: 'image/jpeg', data: url.slice(url.indexOf(',') + 1), name: file.name || 'pasted picture' };
        showAttachment();
      };
      image.src = reader.result;
    };
    reader.readAsDataURL(file);
  }

  function send(text) {
    text = (text || '').trim();
    if ((!text && !attachment) || busy) return;
    var picture = attachment;
    attachment = null;
    showAttachment();
    $('chatInput').value = '';
    bubble('user', (text || 'Build this report format') + (picture ? '\n[picture: ' + picture.name + ']' : ''));
    var waiting = bubble('note', 'Working on it…');
    setBusy(true);

    workbookContext()
      .then(function (context) {
        context.last = lastSpec;
        return fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(picture
            ? { message: text, context: context, image: { mediaType: picture.mediaType, data: picture.data } }
            : { message: text, context: context })
        });
      })
      .then(function (response) { return response.json(); })
      .then(function (data) {
        if (!data || !data.ok) throw new Error((data && data.error) || 'Request failed');
        if (data.spec) lastSpec = data.spec;
        return applyActions(data.actions).then(
          function (notes) { return { reply: data.reply, notes: notes, ai: data.ai }; },
          function (error) { return { reply: data.reply, notes: [], writeError: error.message, ai: data.ai }; }
        );
      })
      .then(function (result) {
        waiting.remove();
        bubble('assistant', result.reply);
        if (result.ai) bubble('note', 'Answered with the AI add-on. Check the accounts on the layout sheet before relying on the numbers.');
        result.notes.forEach(function (note) { bubble('note', note); });
        if (result.writeError) bubble('error', 'Could not write to the workbook: ' + result.writeError);
      })
      .catch(function (error) {
        waiting.remove();
        bubble('error', error.message);
      })
      .then(function () {
        setBusy(false);
        $('chatInput').focus();
      });
  }

  /* The chat is a screen of its own (the speech-bubble button in the header); 'Reports' goes back Home. */
  function showTab(name) {
    if (name === 'Chat') { window.TQNav.go('chat'); setTimeout(function () { $('chatInput').focus(); }, 0); }
    else { window.TQNav.go('home'); loadReports(); }
  }

  function loadReports() {
    if ($('secLayout').open && window.TQDesigner) window.TQDesigner.load();
    if ($('secFilters').open && window.TQFilters) window.TQFilters.load();
    if ($('secData').open && window.TQData) window.TQData.load();
  }

  function start() {
    if ($('menuHelp')) $('menuHelp').addEventListener('click', function () { showTab('Chat'); send('help'); });
    [['secLayout', 'TQDesigner'], ['secFilters', 'TQFilters'], ['secData', 'TQData']].forEach(function (pair) {
      $(pair[0]).addEventListener('toggle', function () {
        if ($(pair[0]).open && window[pair[1]]) window[pair[1]].load();
      });
    });
    var ready = function () { setTimeout(loadReports, 0); };
    if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(ready); else ready();
    $('chatSend').addEventListener('click', function () { send($('chatInput').value); });
    $('chatInput').addEventListener('keydown', function (event) {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        send($('chatInput').value);
      }
    });
    $('chatAttach').addEventListener('click', function () { $('chatFile').click(); });
    $('chatFile').addEventListener('change', function () {
      attach($('chatFile').files[0]);
      $('chatFile').value = '';
    });
    $('chatAttachmentRemove').addEventListener('click', function () { attachment = null; showAttachment(); });
    $('chatInput').addEventListener('paste', function (event) {
      var items = (event.clipboardData && event.clipboardData.items) || [];
      for (var n = 0; n < items.length; n++) {
        if (items[n].kind === 'file' && /^image\//.test(items[n].type)) {
          event.preventDefault();
          attach(items[n].getAsFile());
          return;
        }
      }
    });
    $('chatNew').addEventListener('click', function () {
      if (busy) return;
      lastSpec = null;
      $('chatLog').textContent = '';
      bubble('note', 'New conversation. Type help to see what I understand.');
    });
    Array.prototype.forEach.call(document.querySelectorAll('[data-example]'), function (button) {
      button.addEventListener('click', function () { send(button.getAttribute('data-example')); });
    });
  }

  window.TQChat = {
    hasBackup: function (name) {
      var map = backups();
      return Object.keys(map).some(function (k) { return k.toLowerCase() === String(name || '').toLowerCase(); });
    },
    restoreSheet: restoreSheet,
    showTab: showTab,
    send: send,
    applyActions: applyActions,
    workbookContext: workbookContext,
    setSpec: function (spec) { if (spec) lastSpec = spec; }, // so "by quarter" in the chat follows a Designer build
    getSpec: function () { return lastSpec; }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
