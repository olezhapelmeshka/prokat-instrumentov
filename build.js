// Сборка ТЗ из tz.md: docx по ГОСТ 7.32-2017 / 34.602-2020 и сайт index.html.
// Запуск: npm install && node build.js
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, LevelFormat, LevelSuffix, PageNumber,
  Footer, TableOfContents, HeadingLevel, ExternalHyperlink, VerticalAlign,
} = require('docx');

const ROOT = __dirname;
const DOCX_NAME = 'TZ_Prokat_instrumentov.docx';
const md = fs.readFileSync(path.join(ROOT, 'tz.md'), 'utf8');

// ---------- разбор markdown ----------
function parse(src) {
  const lines = src.split('\n');
  const blocks = [];
  let i = 0;
  let pendingCaption = null;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }
    let m;
    if ((m = line.match(/^(#{1,2}) (.+)$/))) {
      blocks.push({ type: 'h', level: m[1].length, text: m[2].trim() });
      i++;
    } else if ((m = line.match(/^Table: (.+)$/))) {
      pendingCaption = m[1].trim();
      i++;
    } else if (line.startsWith('|')) {
      const rows = [];
      while (i < lines.length && lines[i].startsWith('|')) {
        const cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map(s => s.trim());
        if (!cells.every(c => /^-+$/.test(c))) rows.push(cells);
        i++;
      }
      blocks.push({ type: 'table', caption: pendingCaption, rows });
      pendingCaption = null;
    } else if ((m = line.match(/^!\[(.+)\]\((.+)\)$/))) {
      blocks.push({ type: 'fig', caption: m[1], src: m[2] });
      i++;
    } else if (/^- /.test(line)) {
      const items = [];
      while (i < lines.length && /^- /.test(lines[i])) items.push(lines[i++].slice(2).trim());
      blocks.push({ type: 'ul', items });
    } else if (/^\d+\. /.test(line)) {
      const items = [];
      while (i < lines.length && /^\d+\. /.test(lines[i])) items.push(lines[i++].replace(/^\d+\. /, '').trim());
      blocks.push({ type: 'ol', items });
    } else {
      blocks.push({ type: 'p', text: line.trim() });
      i++;
    }
  }
  return blocks;
}

const LINK = /\[([^\]]+)\]\(([^)]+)\)/g;
const blocks = parse(md);

// ---------- DOCX ----------
const FONT = 'Times New Roman';
const TEXT_W = 11906 - 1701 - 850; // A4 минус поля 30 и 15 мм, в DXA
const INDENT = 709; // 1,25 см

function runs(text, opts = {}, { sources = false } = {}) {
  const out = [];
  let last = 0;
  for (const m of text.matchAll(LINK)) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...opts }));
    if (sources) {
      out.push(new TextRun({ text: `${m[1]} [Электронный ресурс]. — URL: `, ...opts }));
      out.push(new ExternalHyperlink({ link: m[2], children: [new TextRun({ text: m[2], ...opts })] }));
    } else {
      out.push(new TextRun({ text: m[1], ...opts }));
    }
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...opts }));
  return out;
}

function columnWidths(rows) {
  const n = rows[0].length;
  const plain = s => s.replace(LINK, '$1');
  const w = Array.from({ length: n }, (_, c) => {
    const longest = Math.max(...rows.map(r => plain(r[c] || '').length));
    return Math.min(Math.max(longest, 5), 70);
  });
  const sum = w.reduce((a, b) => a + b, 0);
  const res = w.map(x => Math.round((x / sum) * TEXT_W));
  res[res.length - 1] += TEXT_W - res.reduce((a, b) => a + b, 0);
  return res;
}

const thin = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const cellBorders = { top: thin, bottom: thin, left: thin, right: thin };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

function docxTable(rows) {
  const widths = columnWidths(rows);
  const size = rows[0].length > 5 ? 20 : 24; // 10 или 12 пт
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      cantSplit: true,
      children: r.map((cell, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA },
        borders: cellBorders,
        verticalAlign: ri === 0 ? VerticalAlign.CENTER : VerticalAlign.TOP,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [new Paragraph({
          alignment: ri === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          spacing: { line: 240, before: 0, after: 0 },
          children: runs(cell, { font: FONT, size, bold: ri === 0 }),
        })],
      })),
    })),
  });
}

function p(text, extra = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: INDENT },
    children: runs(text),
    ...extra,
  });
}

function center(text, opts = {}, extra = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [new TextRun({ text, ...opts })],
    ...extra,
  });
}

// титульный лист по ГОСТ 34.602-2020 (приложение В, упрощённо)
function signBlock(title) {
  return [
    new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun({ text: title, bold: true })] }),
    new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('Преподаватель')] }),
    new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('_________ [ФИО]')] }),
    new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('«___» __________ 2026 г.')] }),
  ];
}

const half = Math.round(TEXT_W / 2);
const titlePage = [
  center('Министерство науки и высшего образования Российской Федерации', {}, { spacing: { line: 276 } }),
  center('[Полное наименование университета]', {}, { spacing: { line: 276 } }),
  center('Кафедра [название кафедры]', {}, { spacing: { line: 276, after: 600 } }),
  new Table({
    width: { size: TEXT_W, type: WidthType.DXA },
    columnWidths: [half, TEXT_W - half],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: half, type: WidthType.DXA }, borders: noBorders, children: signBlock('СОГЛАСОВАНО') }),
      new TableCell({ width: { size: TEXT_W - half, type: WidthType.DXA }, borders: noBorders, children: signBlock('УТВЕРЖДАЮ') }),
    ] })],
  }),
  center('', {}, { spacing: { before: 1200 } }),
  center('АВТОМАТИЗИРОВАННАЯ ИНФОРМАЦИОННАЯ СИСТЕМА', { bold: true }),
  center('УЧЁТА ПРОКАТА ИНСТРУМЕНТОВ', { bold: true }, { spacing: { after: 360 } }),
  center('ТЕХНИЧЕСКОЕ ЗАДАНИЕ', { bold: true, size: 32 }),
  center('на разработку макета системы', {}, { spacing: { after: 360 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER, indent: { firstLine: 0 },
    children: [new TextRun('На '), new TextRun({ children: [PageNumber.TOTAL_PAGES] }), new TextRun(' листах')],
  }),
  center('Действует с «___» __________ 2026 г.', {}, { spacing: { after: 1200 } }),
  new Table({
    width: { size: TEXT_W, type: WidthType.DXA },
    columnWidths: [half, TEXT_W - half],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: half, type: WidthType.DXA }, borders: noBorders, children: [new Paragraph('')] }),
      new TableCell({ width: { size: TEXT_W - half, type: WidthType.DXA }, borders: noBorders, children: [
        new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun({ text: 'РАЗРАБОТАЛИ', bold: true })] }),
        new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('Студенты группы [номер]')] }),
        new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('_________ [ФИО 1]')] }),
        new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('_________ [ФИО 2]')] }),
        new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('«___» __________ 2026 г.')] }),
      ] }),
    ] })],
  }),
  center('[Город] 2026', {}, { spacing: { before: 1400 } }),
];

const tocPage = [
  center('СОДЕРЖАНИЕ', { bold: true }, { pageBreakBefore: true, spacing: { after: 240 } }),
  new TableOfContents('Содержание', { hyperlink: true, headingStyleRange: '1-2' }),
];

const body = [];
let tableNo = 0, figNo = 0, olNo = 0, section = '';
for (const b of blocks) {
  if (b.type === 'h') {
    if (b.level === 1) section = b.text;
    const unnumbered = b.level === 1 && b.text === b.text.toUpperCase();
    body.push(new Paragraph({
      heading: b.level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
      pageBreakBefore: b.level === 1,
      alignment: unnumbered ? AlignmentType.CENTER : AlignmentType.LEFT,
      indent: { firstLine: unnumbered ? 0 : INDENT },
      keepNext: true,
      children: [new TextRun(b.text)],
    }));
  } else if (b.type === 'p') {
    body.push(p(b.text));
  } else if (b.type === 'ul') {
    b.items.forEach((it, k) => body.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      numbering: { reference: 'dash', level: 0 },
      children: runs(it),
    })));
  } else if (b.type === 'ol') {
    olNo++;
    const isSources = section.startsWith('8 ');
    b.items.forEach(it => body.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      numbering: { reference: isSources ? 'sources' : 'num', level: 0, instance: olNo },
      children: runs(it, {}, { sources: isSources }),
    })));
  } else if (b.type === 'table') {
    tableNo++;
    body.push(new Paragraph({
      indent: { firstLine: 0 },
      keepNext: true,
      spacing: { before: 120 },
      children: [new TextRun(`Таблица ${tableNo} – ${b.caption}`)],
    }));
    body.push(docxTable(b.rows));
    body.push(new Paragraph({ spacing: { line: 240 }, children: [] }));
  } else if (b.type === 'fig') {
    figNo++;
    const heightMm = 205;
    const widthMm = heightMm * (1344 / 2058);
    const px = mm => Math.round((mm / 25.4) * 96);
    body.push(new Paragraph({
      alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, keepNext: true,
      spacing: { before: 120, line: 240 },
      children: [new ImageRun({
        type: 'png', data: fs.readFileSync(path.join(ROOT, b.src)),
        transformation: { width: px(widthMm), height: px(heightMm) },
        altText: { title: b.caption, description: b.caption, name: 'process' },
      })],
    }));
    body.push(center(`Рисунок ${figNo} – ${b.caption}`, {}, { spacing: { after: 240 } }));
  }
}

const doc = new Document({
  creator: 'Разработчики АИС «Прокат инструментов»',
  title: 'Техническое задание: АИС «Прокат инструментов»',
  features: { updateFields: !process.env.NO_PROMPT },
  styles: {
    default: { document: { run: { font: FONT, size: 28 }, paragraph: { spacing: { line: 360, before: 0, after: 0 } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: '000000' },
        paragraph: { spacing: { before: 0, after: 240, line: 360 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: '000000' },
        paragraph: { spacing: { before: 240, after: 120, line: 360 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'dash', levels: [{ level: 0, format: LevelFormat.BULLET, text: '–', suffix: LevelSuffix.SPACE,
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 0, firstLine: INDENT } } } }] },
      { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1)', suffix: LevelSuffix.SPACE,
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 0, firstLine: INDENT } } } }] },
      { reference: 'sources', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', suffix: LevelSuffix.SPACE,
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 0, firstLine: INDENT } } } }] },
    ],
  },
  sections: [{
    properties: {
      titlePage: true,
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, bottom: 1134, left: 1701, right: 850, footer: 567 },
      },
    },
    footers: {
      first: new Footer({ children: [new Paragraph('')] }),
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER, indent: { firstLine: 0 },
        children: [new TextRun({ children: [PageNumber.CURRENT] })],
      })] }),
    },
    children: [...titlePage, ...tocPage, ...body],
  }],
});

// ---------- HTML ----------
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const inline = s => esc(s).replace(LINK, (_, t, u) => `<a href="${u}" target="_blank" rel="noopener">${t}</a>`);
const slug = s => 's-' + s.toLowerCase().replace(/[^a-zа-яё0-9]+/gi, '-').replace(/^-|-$/g, '');

let html = '', toc = '', hTable = 0, hFig = 0;
for (const b of blocks) {
  if (b.type === 'h') {
    const id = slug(b.text);
    const label = b.level === 1 && b.text === b.text.toUpperCase()
      ? b.text.charAt(0) + b.text.slice(1).toLowerCase() : b.text;
    html += `<h${b.level + 1} id="${id}">${esc(label)}</h${b.level + 1}>\n`;
    toc += `<li class="l${b.level}"><a href="#${id}">${esc(label)}</a></li>\n`;
  } else if (b.type === 'p') html += `<p>${inline(b.text)}</p>\n`;
  else if (b.type === 'ul') html += `<ul>${b.items.map(i => `<li>${inline(i)}</li>`).join('')}</ul>\n`;
  else if (b.type === 'ol') html += `<ol>${b.items.map(i => `<li>${inline(i)}</li>`).join('')}</ol>\n`;
  else if (b.type === 'table') {
    hTable++;
    const [head, ...rest] = b.rows;
    html += `<figure class="tbl"><figcaption>Таблица ${hTable} — ${esc(b.caption)}</figcaption><div class="scroll"><table>`
      + `<thead><tr>${head.map(c => `<th>${inline(c)}</th>`).join('')}</tr></thead>`
      + `<tbody>${rest.map(r => `<tr>${r.map(c => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')}</tbody>`
      + `</table></div></figure>\n`;
  } else if (b.type === 'fig') {
    hFig++;
    html += `<figure class="img"><img src="${b.src}" alt="${esc(b.caption)}" loading="lazy"><figcaption>Рисунок ${hFig} — ${esc(b.caption)}</figcaption></figure>\n`;
  }
}

const page = `<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ТЗ: АИС «Прокат инструментов»</title>
<meta name="description" content="Техническое задание на макет автоматизированной информационной системы учёта проката инструментов.">
<style>
:root {
  --bg: #fbfaf7; --surface: #ffffff; --text: #1d1d1b; --muted: #5f5e58; --line: #e3e1da;
  --accent: #2563c9; --accent-soft: #e8effb; --head: #f3f1ec;
}
@media (prefers-color-scheme: dark) {
  :root { --bg: #161615; --surface: #1e1e1c; --text: #ecebe6; --muted: #a4a29a; --line: #34332f;
    --accent: #7aa7f0; --accent-soft: #1d2a40; --head: #262523; }
  .img img { background: #fff; }
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { margin: 0; background: var(--bg); color: var(--text);
  font: 16px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }
a { color: var(--accent); }
header.hero { border-bottom: 1px solid var(--line); background: var(--surface); }
.hero-in { max-width: 1180px; margin: 0 auto; padding: 40px 24px 32px; }
.eyebrow { color: var(--muted); font-size: 14px; letter-spacing: .04em; text-transform: uppercase; margin: 0 0 8px; }
h1 { font-size: clamp(26px, 4vw, 38px); line-height: 1.2; margin: 0 0 12px; }
.lead { color: var(--muted); max-width: 720px; margin: 0 0 20px; }
.actions { display: flex; gap: 12px; flex-wrap: wrap; }
.btn { display: inline-block; padding: 10px 16px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 15px;
  border: 1px solid var(--line); color: var(--text); background: var(--surface); }
.btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
@media (prefers-color-scheme: dark) { .btn.primary { color: #10131a; } }
.layout { max-width: 1180px; margin: 0 auto; padding: 0 24px; display: grid; grid-template-columns: 260px minmax(0, 1fr); gap: 48px; }
nav.toc { position: sticky; top: 0; align-self: start; max-height: 100vh; overflow: auto; padding: 32px 0; font-size: 14px; }
nav.toc p { font-weight: 600; margin: 0 0 8px; }
nav.toc ul { list-style: none; margin: 0; padding: 0; }
nav.toc li { margin: 2px 0; }
nav.toc li.l2 { padding-left: 14px; }
nav.toc a { color: var(--muted); text-decoration: none; display: block; padding: 3px 8px; border-radius: 6px; }
nav.toc a:hover { background: var(--accent-soft); color: var(--text); }
main { padding: 16px 0 80px; min-width: 0; }
main h2 { font-size: 26px; margin: 56px 0 16px; padding-top: 8px; border-top: 1px solid var(--line); scroll-margin-top: 16px; }
main h3 { font-size: 19px; margin: 32px 0 10px; scroll-margin-top: 16px; }
main p, main li { max-width: 760px; }
figure { margin: 24px 0; }
figcaption { color: var(--muted); font-size: 14px; margin-bottom: 8px; }
.scroll { overflow-x: auto; border: 1px solid var(--line); border-radius: 10px; background: var(--surface); }
table { border-collapse: collapse; width: 100%; font-size: 14.5px; }
th, td { text-align: left; vertical-align: top; padding: 9px 12px; border-bottom: 1px solid var(--line); }
th { background: var(--head); font-weight: 600; white-space: nowrap; }
tr:last-child td { border-bottom: 0; }
.img img { width: 100%; max-width: 680px; display: block; border: 1px solid var(--line); border-radius: 10px; }
.img figcaption { margin: 8px 0 0; }
footer { border-top: 1px solid var(--line); color: var(--muted); font-size: 14px; }
footer div { max-width: 1180px; margin: 0 auto; padding: 24px; }
@media (max-width: 900px) {
  .layout { grid-template-columns: 1fr; gap: 0; padding: 0 16px; }
  nav.toc { position: static; max-height: none; padding: 16px 0 0; border-bottom: 1px solid var(--line); }
  .hero-in { padding: 28px 16px 24px; }
  th { white-space: normal; }
}
</style>
</head>
<body>
<header class="hero"><div class="hero-in">
  <p class="eyebrow">Введение в программную инженерию · Лабораторная работа 1</p>
  <h1>Техническое задание: АИС «Прокат инструментов»</h1>
  <p class="lead">Макет автоматизированной информационной системы учёта проката инструментов. Реализуемый бизнес-процесс — выдача и возврат инструмента.</p>
  <div class="actions">
    <a class="btn primary" href="${DOCX_NAME}" download>Скачать .docx по ГОСТ</a>
    <a class="btn" href="#${slug('3.3 Автоматизируемый бизнес-процесс')}">Схема процесса</a>
  </div>
</div></header>
<div class="layout">
<nav class="toc" aria-label="Содержание"><p>Содержание</p><ul>
${toc}</ul></nav>
<main>
${html}</main>
</div>
<footer><div>Оформлено по ГОСТ 34.602-2020. Обновлено 2 октября 2026 г.</div></footer>
</body>
</html>
`;

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(path.join(ROOT, DOCX_NAME), buf);
  fs.writeFileSync(path.join(ROOT, 'index.html'), page);
  console.log(`ok: ${DOCX_NAME}, index.html — таблиц ${tableNo}, рисунков ${figNo}`);
});
