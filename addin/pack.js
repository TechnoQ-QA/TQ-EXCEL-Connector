/* TQ Excel Connector - Send the report. A copy of the report (or its whole pack) with every formula replaced by
 * its value, for people without the add-in, and page setup so the pack prints or saves as a PDF cleanly. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var LAYOUT_RE = /layout sheet "([^"]+)"/;
  var SUMMARY = ' - summary';
  var COPY_NOTE = 'Fixed values, copied ';

  function say(text, bad) {
    $('pMessage').textContent = text || '';
    $('pMessage').className = 'status-line' + (bad ? ' bad' : '');
  }

  function inExcel() { return typeof Excel !== 'undefined' && window.Office && Office.context && Office.context.document; }
  function supports(version) {
    try { return Office.context.requirements.isSetSupported('ExcelApi', version); } catch (e) { return false; }
  }

  /* The sheets of the active report: [{name, header}] in workbook order. A report says in A2 which layout sheet
   * it was built from; its summary's A1 is the report title followed by " - summary". */
  function packSheets(context, whole) {
    var sheets = context.workbook.worksheets;
    sheets.load('items/name');
    var active = context.workbook.worksheets.getActiveWorksheet();
    active.load('name');
    return context.sync().then(function () {
      var tops = sheets.items.map(function (s) {
        var range = s.getRange('A1:B40');
        range.load('values');
        return { name: s.name, range: range };
      });
      return context.sync().then(function () {
        var info = tops.map(function (t) {
          var v = t.range.values;
          var header = 0;
          for (var r = 0; r < v.length; r++) {
            if (String(v[r][0]) === 'Accounts' && String(v[r][1]) === 'Name') { header = r + 1; break; }
            if (String(v[r][0]) && r === 3 && String(v[0][0]).slice(-SUMMARY.length) === SUMMARY) { header = 4; break; }
          }
          var layout = LAYOUT_RE.exec(String(v[1][0]));
          return { name: t.name, title: String(v[0][0]), layout: layout ? layout[1] : '', header: header,
            subsidiary: String(v[3][0]) === 'Subsidiary' ? String(v[3][1] || '').trim() : '',
            copy: String(v[1][0]).indexOf(COPY_NOTE) === 0 };
        });
        var me = info.filter(function (i) { return i.name === active.name; })[0] || { name: active.name, header: 0, title: '' };
        var isSummary = me.title.slice(-SUMMARY.length) === SUMMARY;
        if (!whole || (!me.layout && !isSummary)) return [me];   // a single sheet, or a report that is not a pack
        var title = isSummary ? me.title.slice(0, -SUMMARY.length) : me.title;
        var layout = me.layout || (info.filter(function (i) { return i.title === title && i.layout; })[0] || {}).layout;
        return info.filter(function (i) {
          if (i.copy) return false;   // copies made earlier stay out
          return (layout && i.layout === layout) || i.title === title + SUMMARY;
        });
      });
    });
  }

  /* "Doha - P&L 18 Nov" : the copy says when it was taken (sheet names allow 31 characters, no slashes). */
  function stamp() {
    var d = new Date();
    return d.getDate() + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()];
  }

  function freeName(base, taken) {
    var tail = ' ' + stamp();
    var stem = base.slice(0, 31 - tail.length);
    var name = stem + tail;
    for (var n = 2; taken.indexOf(name.toLowerCase()) !== -1; n++) {
      var more = tail + ' (' + n + ')';
      name = base.slice(0, 31 - more.length) + more;
    }
    taken.push(name.toLowerCase());
    return name;
  }

  function copyAsValues() {
    if (!inExcel()) { say('Open the panel in Excel to copy the report.', true); return; }
    if (!supports('1.10')) { say('This Excel is too old to copy sheets (needs Excel 2019 or Microsoft 365).', true); return; }
    say('Copying…');
    Excel.run(function (context) {
      return packSheets(context, $('pWhole').checked).then(function (pack) {
        var sheets = context.workbook.worksheets;
        sheets.load('items/name');
        return context.sync().then(function () {
          var taken = sheets.items.map(function (s) { return s.name.toLowerCase(); });
          var copies = pack.map(function (p) {
            var copy = sheets.getItem(p.name).copy('End');
            copy.name = freeName(p.name, taken);
            return copy;
          });
          return context.sync().then(function () {
            var used = copies.map(function (c) { var u = c.getUsedRange(); u.load('values,address'); return u; });
            return context.sync().then(function () {
              var waiting = 0;
              used.forEach(function (u) {
                u.values.forEach(function (row) { row.forEach(function (v) { if (/^#(BUSY|GETTING_DATA)/.test(String(v))) waiting++; }); });
                var values = u.values;   // every formula becomes the value it shows
                if (values.length > 1) values[1][0] = COPY_NOTE + new Date().toLocaleString() + ' from TQ Connector. ' +
                  String(values[1][0]).replace(LAYOUT_RE, 'layout "$1"');   // A2 says when; it is no longer a live report
                u.values = values;
              });
              copies.forEach(function (c) {
                if (supports('1.8')) c.getRange().dataValidation.clear();
                c.tabColor = '#808080';
              });
              copies[0].activate();
              return context.sync().then(function () {
                return { count: copies.length, waiting: waiting, names: copies.map(function (c) { return c.name; }) };
              });
            });
          });
        });
      });
    }).then(function (r) {
      say(r.count + ' sheet(s) copied with fixed values (grey tabs): ' + r.names.join(', ') + '.' +
        (r.waiting ? ' ' + r.waiting + ' cell(s) were still loading when copied: wait for Refresh all to finish and copy again.' : '') +
        ' To send them on their own: right-click the tabs, Move or Copy, To book: (new book).');
    }).catch(function (e) { say(e.message, true); });
  }

  function setUpPrint() {
    if (!inExcel()) { say('Open the panel in Excel to set up the pages.', true); return; }
    if (!supports('1.9')) { say('This Excel is too old for page setup from the panel (needs Excel 2019 or Microsoft 365).', true); return; }
    say('Setting up the pages…');
    var withLogo = !$('pLogo') || $('pLogo').checked;
    var note = '';
    Excel.run(function (context) { return packSheets(context, $('pWhole').checked); }).then(function (pack) {
      // The company (or the report's one subsidiary) name and logo from NetSuite, asked once per subsidiary.
      var subs = pack.map(function (p) { return p.subsidiary; }).filter(function (s, i, all) { return all.indexOf(s) === i; });
      return Promise.all(subs.map(function (s) {
        return TQ.callApi({ action: 'branding', subsidiary: s }).catch(function (e) { note = e.message; return null; });
      })).then(function (found) {
        var brand = {};
        subs.forEach(function (s, i) { brand[s] = found[i]; });
        return Excel.run(function (context) {
          pack.forEach(function (p) {
            var sheet = context.workbook.worksheets.getItem(p.name);
            var b = brand[p.subsidiary] || {};
            var page = sheet.pageLayout;
            page.orientation = Excel.PageOrientation.landscape;
            page.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
            page.centerHorizontally = true;
            page.leftMargin = 28; page.rightMargin = 28; page.topMargin = 46; page.bottomMargin = 40;
            if (p.header) page.setPrintTitleRows('$1:$' + p.header);
            var foot = page.headersFooters.defaultForAllPages;
            var amp = function (t) { return String(t || '').replace(/&/g, '&&'); };   // "&" is Excel's code character
            foot.leftHeader = '&B' + amp(b.subsidiary || b.company || '');           // who the report is for
            foot.centerHeader = '&B' + amp(p.title || p.name);                        // the report title on every page
            foot.rightHeader = b.subsidiary && b.company ? amp(b.company) : '';
            foot.leftFooter = '&A';
            foot.centerFooter = 'Page &P of &N';
            foot.rightFooter = '&D';
          });
          return context.sync().then(function () {
            if (!withLogo) return 0;
            // The logo: Excel's page header takes no pictures from an add-in, so it goes on the sheet, top right,
            // above the report (it prints on the first page and stays on copies made with Copy as values).
            var placed = 0;
            var used = pack.map(function (p) {
              // the report's own width: its heading row (the lookup lists further right stay out)
              var row = p.header || 1;
              var whole = context.workbook.worksheets.getItem(p.name).getRange(row + ':' + row).getUsedRange(true);
              whole.load('width');
              return { p: p, whole: whole };
            });
            return context.sync().then(function () {
              used.forEach(function (u) {
                var b = brand[u.p.subsidiary] || {};
                if (!b.logo || !b.logo.data || !/PNG|JPG|JPEG|GIF/i.test(b.logo.type + ' ' + b.logo.name)) return;
                var sheet = context.workbook.worksheets.getItem(u.p.name);
                var old = sheet.shapes.getItemOrNullObject('TQ Logo');
                old.load('isNullObject');
                u.old = old;
              });
              return context.sync().then(function () {
                used.forEach(function (u) {
                  var b = brand[u.p.subsidiary] || {};
                  if (!u.old) return;
                  if (!u.old.isNullObject) u.old.delete();
                  var img = context.workbook.worksheets.getItem(u.p.name).shapes.addImage(b.logo.data);
                  img.name = 'TQ Logo';
                  img.lockAspectRatio = true;
                  img.height = 34;
                  img.top = 2;
                  img.left = Math.max(260, u.whole.width - 140);
                  placed += 1;
                });
                return context.sync().then(function () { return placed; });
              });
            });
          }).then(function (logos) { return { n: pack.length, logos: logos }; });
        });
      });
    }).then(function (r) {
      say(r.n + ' sheet(s) set to landscape, one page wide, with the company name, report title and page numbers' +
        (r.logos ? ', and the logo from NetSuite' : '') + '. To make the PDF: File > Export > Create PDF (Options: Entire ' +
        'workbook, or select the tabs first).' + (withLogo && !r.logos ? ' No logo was found in NetSuite (Setup > Company > ' +
        'Company Information, Company Logo (Forms))' + (note ? ': ' + note : '') + '.' : ''));
    }).catch(function (e) { say(e.message, true); });
  }

  /* The company logo (uploaded on the setup page, else NetSuite's) at the top right of a sheet, e.g. a dashboard. */
  function placeLogo(sheetName, left) {
    if (!inExcel() || !supports('1.9')) return Promise.resolve(false);
    return fetch('/logo').then(function (r) {
      if (!r.ok) return null;
      return r.blob().then(function (blob) {
        return new Promise(function (done) {
          var reader = new FileReader();
          reader.onload = function () { done(String(reader.result).split(',')[1]); };
          reader.readAsDataURL(blob);
        });
      });
    }).then(function (data) {
      if (!data) return false;
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(sheetName);
        var old = sheet.shapes.getItemOrNullObject('TQ Logo');
        old.load('isNullObject');
        return context.sync().then(function () {
          if (!old.isNullObject) old.delete();
          var img = sheet.shapes.addImage(data);
          img.name = 'TQ Logo';
          img.lockAspectRatio = true;
          img.height = 34;
          img.top = 2;
          img.left = left || 470;
          return context.sync().then(function () { return true; });
        });
      });
    }).catch(function () { return false; });
  }

  window.TQPack = { copyAsValues: copyAsValues, setUpPrint: setUpPrint, placeLogo: placeLogo };

  function start() {
    if (!$('secPack')) return;
    $('pValues').addEventListener('click', copyAsValues);
    $('pPrint').addEventListener('click', setUpPrint);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
