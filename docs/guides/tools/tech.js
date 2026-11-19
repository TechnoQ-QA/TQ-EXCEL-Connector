// docs/TECHNICAL.md -> Word, in the style of the guides. node tech.js <TECHNICAL.md> <images dir> <output.docx>
const fs = require('fs');
const path = require('path');
const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle } = require('docx');
const c = require('./common');
const [MD, IMG, OUT] = process.argv.slice(2);

const text = fs.readFileSync(MD, 'utf8').replace(/\r/g, '');
const lines = text.split('\n');
const clean = (s) => s.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1').replace(/\\\|/g, '|');

function codeBlock(codeLines) {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: c.TEXT_W, type: WidthType.DXA }, columnWidths: [c.TEXT_W],
    borders: { top: none, bottom: none, right: none, insideHorizontal: none, insideVertical: none,
      left: { style: BorderStyle.SINGLE, size: 18, color: c.NAVY } },
    rows: [new TableRow({ cantSplit: codeLines.length < 25, children: [new TableCell({
      width: { size: c.TEXT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: 'F3F2F1', color: 'auto' },
      margins: { top: 80, bottom: 80, left: 160, right: 120 },
      children: codeLines.map((l) => new Paragraph({ spacing: { after: 0 },
        children: [new TextRun({ text: l || ' ', font: 'Consolas', size: 17, color: '201F1E' })] })),
    })] })],
  });
}

function mdTable(rows) {
  const cells = rows.map((r) => r.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split(/(?<!\\)\|/).map((x) => clean(x.trim())));
  const body = cells.filter((r, i) => !(i === 1 && r.every((x) => /^:?-{2,}:?$/.test(x))));
  const n = body[0].length;
  // each column: at least its heading and its longest single word, at most 40 characters' share
  const longestWord = (j) => Math.max(...body.map((r) => Math.max(...String(r[j] || '').replace(/[*`]/g, '').split(/\s+/).map((w) => w.length))));
  const weight = Array.from({ length: n }, (_, j) => Math.min(40, Math.max(body[0][j].length + 2, Math.min(24, longestWord(j)),
    ...body.map((r) => Math.min(40, (r[j] || '').length * 0.6)))));
  const total = weight.reduce((a, b) => a + b, 0);
  const widths = weight.map((w) => Math.floor(c.TEXT_W * w / total));
  widths[n - 1] += c.TEXT_W - widths.reduce((a, b) => a + b, 0);
  return c.table(body.map((r) => Array.from({ length: n }, (_, j) => r[j] || '')), widths);
}

const out = [];
let para = [];
let diagram = 0;
const flush = () => { if (para.length) { out.push(c.P(clean(para.join(' ')))); para = []; } };

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  if (/^```/.test(line)) {
    flush();
    const kind = line.slice(3).trim();
    const block = [];
    for (i += 1; i < lines.length && !/^```/.test(lines[i]); i++) block.push(lines[i]);
    if (kind === 'mermaid') {
      diagram += 1;
      out.push(...c.picture(path.join(IMG, 'diagram' + diagram + '.png'), diagram === 1 ? 560 : 620, ''));
    } else out.push(codeBlock(block));
    continue;
  }
  if (/^\|/.test(line)) {
    flush();
    const rows = [];
    for (; i < lines.length && /^\|/.test(lines[i]); i++) rows.push(lines[i]);
    i -= 1;
    out.push(mdTable(rows));
    continue;
  }
  if (/^# /.test(line)) { flush(); continue; }                 // the title is on the cover
  if (/^---\s*$/.test(line)) { flush(); continue; }
  if (/^## /.test(line)) { flush(); out.push(c.H1(clean(line.slice(3)).replace(/`/g, ''))); continue; }
  if (/^### /.test(line)) { flush(); out.push(c.H2(clean(line.slice(4)).replace(/`/g, ''))); continue; }
  if (/^\s*- /.test(line)) {
    flush();
    const level = /^\s{2,}- /.test(line) ? 1 : 0;
    let item = line.replace(/^\s*- /, '');
    while (i + 1 < lines.length && /^\s{2,}\S/.test(lines[i + 1]) && !/^\s*- /.test(lines[i + 1])) { i += 1; item += ' ' + lines[i].trim(); }
    out.push(...c.bullets([clean(item)], level));
    continue;
  }
  if (/^\d+\. /.test(line)) {
    flush();
    const items = [];
    for (; i < lines.length && (/^\d+\. /.test(lines[i]) || /^\s{2,}\S/.test(lines[i])); i++) {
      if (/^\d+\. /.test(lines[i])) items.push(lines[i].replace(/^\d+\. /, ''));
      else if (/^\s*- /.test(lines[i])) items.push([clean(lines[i].replace(/^\s*- /, ''))]);
      else if (typeof items[items.length - 1] === 'string') items[items.length - 1] += ' ' + lines[i].trim();
    }
    i -= 1;
    out.push(...c.steps(items.map((x) => (typeof x === 'string' ? clean(x) : x))));
    continue;
  }
  if (!line.trim()) { flush(); continue; }
  para.push(line.trim());
}
flush();

const body = [
  ...c.cover('Technical Reference', 'How TQ Connector is built: NetSuite scripts, connector, add-in, installer and security',
    ['Version 1.4  ·  November 2026', 'Internal: Techno Q', 'Connector 1.4.0  ·  Add-in 1.4.0.0  ·  NetSuite script 7']),
  ...c.toc(),
  ...out,
];
c.save(c.doc('Technical Reference', body), OUT);
