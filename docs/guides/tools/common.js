// Shared building blocks for the TQ Connector Word guides (docx 9).
const fs = require('fs');
const path = require('path');
const d = require('docx');
const {
  Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, ImageRun, LevelFormat, TableLayoutType, Header, Footer, PageNumber, TableOfContents, PageBreak, TabStopType,
} = d;

const ORANGE = 'C2571A';
const NAVY = '1F3A5F';
const GREY = '605E5C';
const LINE = 'D9D9D9';
const PAGE_W = 11906;     // A4
const MARGIN = 1134;      // 2 cm
const TEXT_W = PAGE_W - 2 * MARGIN;
const FONT = 'Calibri';
const PAGE_H = 16838;     // A4, twips (1 twip = 635 EMU)
const BRAND = path.join(__dirname, 'brand');
const BANNER_W = 7562855; // EMU, the full page width
const BANNER_H = 789874;

/** A letterhead picture floating at a fixed place on the page (sizes and offsets in EMU, as in the scope document). */
function brandImage(file, type, cx, cy, horizontal, top, behind = false) {
  return new ImageRun({ type, data: fs.readFileSync(path.join(BRAND, file)),
    transformation: { width: Math.round(cx / 9525), height: Math.round(cy / 9525) },
    floating: { horizontalPosition: horizontal,
      verticalPosition: { relative: d.VerticalPositionRelativeFrom.PAGE, offset: Math.round(top) },
      behindDocument: behind, allowOverlap: true, wrap: { type: d.TextWrappingType.NONE } } });
}

let listInstance = 0;

function runs(text, base) {
  // **bold** and `code` inside a sentence
  const out = [];
  const parts = String(text).split(/(\*\*[^*]+\*\*|`[^`]+`)/);
  parts.forEach((p) => {
    if (!p) return;
    if (p.startsWith('**')) out.push(new TextRun(Object.assign({ text: p.slice(2, -2), bold: true }, base)));
    else if (p.startsWith('`')) out.push(new TextRun(Object.assign({ text: p.slice(1, -1), font: 'Consolas', size: 19, color: NAVY }, base)));
    else out.push(new TextRun(Object.assign({ text: p }, base)));
  });
  return out;
}

const P = (text, opts = {}) => new Paragraph(Object.assign({ children: runs(text, opts.run), spacing: { after: 120 } }, opts.para || {}));
const H1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
const BR = () => new Paragraph({ children: [new PageBreak()] });

function bullets(items, level = 0) {
  return items.map((t) => new Paragraph({ numbering: { reference: 'bullets', level }, children: runs(t), spacing: { after: 60 } }));
}
function steps(items) {
  listInstance += 1;
  const inst = listInstance;
  const out = [];
  items.forEach((t) => {
    if (Array.isArray(t)) out.push(...t.map((s) => new Paragraph({ numbering: { reference: 'bullets', level: 1 }, children: runs(s), spacing: { after: 40 } })));
    else out.push(new Paragraph({ numbering: { reference: 'steps', level: 0, instance: inst }, children: runs(t), spacing: { after: 80 } }));
  });
  return out;
}

function cell(content, width, opts = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    c instanceof Paragraph ? c : new Paragraph({ children: runs(c, opts.run), spacing: { after: 40 } }));
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: paras,
  });
}
function table(rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,   // long words (script IDs, paths) wrap instead of pushing the table off the page
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0 && opts.header !== false,
      cantSplit: true,
      children: r.map((c, j) => cell(c, widths[j], i === 0 && opts.header !== false
        ? { fill: NAVY, run: { bold: true, color: 'FFFFFF' } } : (opts.firstCol && j === 0 ? { run: { bold: true } } : {}))),
    })),
  });
}
const spacer = () => new Paragraph({ children: [], spacing: { after: 60 } });

/** A shaded box: kind note (blue), warn (amber), tip (green). */
function box(kind, title, lines) {
  const fill = { note: 'EFF6FC', warn: 'FFF4CE', tip: 'E6F4EA', key: 'FDF1E9' }[kind] || 'EFF6FC';
  const bar = { note: '0F6CBD', warn: 'C19C00', tip: '107C10', key: ORANGE }[kind] || '0F6CBD';
  const paras = [new Paragraph({ children: [new TextRun({ text: title, bold: true, color: bar })], spacing: { after: 60 } })]
    .concat((Array.isArray(lines) ? lines : [lines]).map((l) => new Paragraph({ children: runs(l), spacing: { after: 40 } })));
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA },
    columnWidths: [TEXT_W],
    borders: { top: none, bottom: none, right: none, insideHorizontal: none, insideVertical: none,
      left: { style: BorderStyle.SINGLE, size: 24, color: bar } },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: TEXT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
      margins: { top: 100, bottom: 100, left: 180, right: 180 },
      children: paras })] })],
  });
}

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
/** A screenshot with a thin frame and a caption; widthPx at 96 dpi. */
function picture(file, widthPx, caption) {
  const s = pngSize(file);
  const h = Math.round(widthPx * s.h / s.w);
  const out = [new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 40 }, keepNext: true,
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(file), transformation: { width: widthPx, height: h },
      altText: { title: caption || 'Screenshot', description: caption || 'Screenshot', name: path.basename(file) } })],
  })];
  if (caption) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 },
    children: [new TextRun({ text: caption, italics: true, size: 18, color: GREY })] }));
  return out;
}
/** Two screenshots side by side (no borders). */
function pictures2(a, b, widthPx) {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const half = Math.floor(TEXT_W / 2);
  const col = (p) => new TableCell({ width: { size: half, type: WidthType.DXA }, children: picture(p[0], widthPx, p[1]) });
  return new Table({ width: { size: half * 2, type: WidthType.DXA }, columnWidths: [half, half],
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    rows: [new TableRow({ children: [col(a), col(b)] })] });
}

function cover(title, subtitle, lines) {
  return [
    new Paragraph({ spacing: { before: 2200, after: 0 }, children: [new TextRun({ text: 'TECHNO Q', bold: true, color: ORANGE, size: 28, characterSpacing: 40 })] }),
    new Paragraph({ spacing: { before: 120, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 6 } },
      children: [new TextRun({ text: 'TQ Connector for NetSuite', size: 30, color: NAVY })] }),
    new Paragraph({ spacing: { before: 360, after: 120 }, children: [new TextRun({ text: title, bold: true, size: 56, color: '201F1E' })] }),
    new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: subtitle, size: 28, color: GREY })] }),
    ...lines.map((l) => new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: l, size: 22, color: GREY })] })),
    BR(),
  ];
}

/** Word merges two tables that touch; a small empty paragraph between them keeps them apart. */
function separate(list) {
  const out = [];
  list.forEach((x, i) => {
    if (i > 0 && x instanceof Table && list[i - 1] instanceof Table) out.push(new Paragraph({ children: [], spacing: { after: 0 } }));
    out.push(x);
  });
  return out;
}

function doc(title, children) {
  return new d.Document({
    creator: 'Techno Q', title, description: title,
    styles: {
      default: { document: { run: { font: FONT, size: 21, color: '201F1E' }, paragraph: { spacing: { line: 276 } } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 34, bold: true, color: NAVY },
          paragraph: { spacing: { before: 360, after: 140 }, keepNext: true, keepLines: true, outlineLevel: 0,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ORANGE, space: 4 } } } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 27, bold: true, color: ORANGE },
          paragraph: { spacing: { before: 260, after: 100 }, keepNext: true, keepLines: true, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 23, bold: true, color: NAVY },
          paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, keepLines: true, outlineLevel: 2 } },
      ],
    },
    numbering: { config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } }] },
      { reference: 'steps', levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } }, run: { bold: true, color: ORANGE } } }] },
    ] },
    sections: [{
      // Techno Q letterhead (as in the Techno Q scope documents): the "30 Dedicated Years" badge centred at the top
      // and the company banner across the bottom of every page, cover included.
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 2150, bottom: 1800, left: MARGIN, right: MARGIN, header: 560, footer: 1300 } } },
      headers: {
        default: new Header({ children: [
          new Paragraph({ children: [brandImage('badge.jpg', 'jpg', 1085850, 1085850,
            { relative: d.HorizontalPositionRelativeFrom.PAGE, align: d.HorizontalPositionAlign.CENTER }, 172085)] }),
          new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }], spacing: { before: 120 },
            children: [new TextRun({ text: 'TQ Connector for NetSuite', color: GREY, size: 17 }), new TextRun({ text: '\t' + title, color: GREY, size: 17 })] })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
          children: [
            brandImage('banner.png', 'png', BANNER_W, BANNER_H, { relative: d.HorizontalPositionRelativeFrom.PAGE, offset: 0 },
              PAGE_H * 635 - BANNER_H - 40000, true),
            new TextRun({ text: 'Confidential', color: GREY, size: 17 }),
            new TextRun({ children: ['\tPage ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], color: GREY, size: 17 })] })] }),
      },
      children: separate(children.flat(Infinity)),   // lists come back as arrays
    }],
  });
}

function toc() {
  return [new Paragraph({ children: [new TextRun({ text: 'Contents', bold: true, size: 34, color: NAVY })], spacing: { after: 120 } }),
    new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }), BR()];
}

async function save(document, file) {
  const buf = await d.Packer.toBuffer(document);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, buf);
  console.log('wrote', file, buf.length, 'bytes');
}

module.exports = { P, H1, H2, H3, BR, bullets, steps, table, box, picture, pictures2, cover, doc, toc, save, spacer, TEXT_W, ORANGE, NAVY };
