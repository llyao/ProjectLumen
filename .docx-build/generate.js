const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  PageNumber, Footer, PageBreak
} = require('docx');

const sourcePath = path.resolve(__dirname, '..', '朵朵', '6年级', '记录', '英语', '2026-09-15-第1单元7天学练打印版.md');
const outDir = path.resolve(__dirname, '..', '朵朵', '6年级', '记录', '英语', '第1单元7天Word版');
fs.mkdirSync(outDir, { recursive: true });
const source = fs.readFileSync(sourcePath, 'utf8').replace(/\r/g, '');

const C = { black: '000000', grey: '666666', light: 'F0F0F0', line: 'B7B7B7' };
const fontCN = '宋体';
const fontEN = 'Times New Roman';
const run = (text, opts = {}) => new TextRun({ text: String(text), font: { eastAsia: fontCN, ascii: fontEN }, size: opts.size || 21, color: opts.color || C.black, bold: opts.bold || false, italics: opts.italics || false });
const para = (text = '', opts = {}) => new Paragraph({ alignment: opts.alignment || AlignmentType.LEFT, spacing: { before: opts.before ?? 80, after: opts.after ?? 80, line: opts.line || 360 }, indent: opts.indent ? { firstLine: 420 } : undefined, keepNext: opts.keepNext || false, children: [run(text, opts)] });
const heading = (text, level = 1) => new Paragraph({ heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3, spacing: { before: level === 1 ? 260 : 180, after: 100, line: 360 }, keepNext: true, children: [run(text, { size: level === 1 ? 32 : level === 2 ? 26 : 23, bold: true })] });
const borders = { top: { style: BorderStyle.SINGLE, size: 1, color: C.line }, bottom: { style: BorderStyle.SINGLE, size: 1, color: C.line }, left: { style: BorderStyle.SINGLE, size: 1, color: C.line }, right: { style: BorderStyle.SINGLE, size: 1, color: C.line }, insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: C.line }, insideVertical: { style: BorderStyle.SINGLE, size: 1, color: C.line } };
const tableFrom = (rows) => new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders, rows: rows.map((row, ri) => new TableRow({ tableHeader: ri === 0, cantSplit: true, children: row.map(cell => new TableCell({ width: { size: Math.floor(100 / row.length), type: WidthType.PERCENTAGE }, margins: { top: 70, bottom: 70, left: 90, right: 90 }, shading: ri === 0 ? { type: ShadingType.CLEAR, fill: C.light } : undefined, children: [new Paragraph({ spacing: { before: 0, after: 0, line: 300 }, children: [run(cell, { size: 19, bold: ri === 0 })] })] })) })) });
const clean = (s) => s.replace(/^\s+|\s+$/g, '').replace(/`/g, '').replace(/\*\*/g, '');

function renderBlock(lines) {
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }
    if (line.startsWith('|')) {
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        const cells = lineCells(lines[i]);
        if (!cells.every(c => /^[-: ]+$/.test(c))) rows.push(cells);
        i++;
      }
      if (rows.length) out.push(tableFrom(rows));
      continue;
    }
    if (/^### /.test(line)) { out.push(heading(clean(line.slice(4)), 3)); i++; continue; }
    if (/^## /.test(line)) { out.push(heading(clean(line.slice(3)), 2)); i++; continue; }
    if (/^# /.test(line)) { out.push(heading(clean(line.slice(2)), 1)); i++; continue; }
    if (/^---+$/.test(line.trim())) { i++; continue; }
    if (/^\s*[-*] /.test(line)) { out.push(para('□ ' + clean(line.replace(/^\s*[-*] /, '')), { indent: false })); i++; continue; }
    if (/^\s*\d+\. /.test(line)) { out.push(para(clean(line), { indent: false })); i++; continue; }
    if (/^\s*\d+\) /.test(line)) { out.push(para(clean(line), { indent: false })); i++; continue; }
    out.push(para(clean(line), { indent: false }));
    i++;
  }
  return out;
}
function lineCells(line) { return line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(clean); }
function sliceDay(n) {
  const start = source.indexOf(`# 第${n}天`);
  const end = source.indexOf(`# 第${n + 1}天`, start + 1);
  return source.slice(start, end < 0 ? source.indexOf('# 家长用答案') : end).trim().split('\n');
}
function answerSlice(n) {
  const start = source.indexOf(`## 第${n}天`, source.indexOf('# 家长用答案'));
  const next = source.indexOf(`## 第${n + 1}天`, start + 1);
  const end = next < 0 ? source.length : next;
  return source.slice(start, end).trim().split('\n');
}
function footer() { return { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('第1单元《School life》学习练习', { size: 16, color: C.grey }), run('　', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: { eastAsia: fontCN, ascii: fontEN }, size: 16, color: C.grey })] })] }) }; }

(async () => {
  for (let n = 1; n <= 7; n++) {
    const dayLines = sliceDay(n);
    const answerLines = answerSlice(n);
    const title = `朵朵英语第${n}天学练：${clean(dayLines[0].replace(/^# /, ''))}`;
    const children = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 180, line: 420 }, children: [run(title, { size: 30, bold: true })] }),
      para('教材：上海教育出版社五四学制六年级上册 Unit 1《School life》　记录对象：朵朵', { alignment: AlignmentType.CENTER, size: 18, color: C.grey }),
      para('日期：__________　完成用时：__________', { alignment: AlignmentType.RIGHT, size: 18, color: C.grey }),
      ...renderBlock(dayLines.slice(1)),
      new Paragraph({ children: [run('本页完成后再查看答案', { size: 18, color: C.grey, italics: true }), new PageBreak()] }),
      heading('家长用答案与订正提示', 1),
      ...renderBlock(answerLines.slice(1)),
    ];
    const doc = new Document({
      styles: {
        default: { document: { run: { font: { eastAsia: fontCN, ascii: fontEN }, size: 21, color: C.black }, paragraph: { spacing: { line: 360 } } } },
        heading1: { run: { font: { eastAsia: fontCN, ascii: fontEN }, size: 32, bold: true, color: C.black }, paragraph: { spacing: { before: 260, after: 100, line: 360 } } },
        heading2: { run: { font: { eastAsia: fontCN, ascii: fontEN }, size: 26, bold: true, color: C.black }, paragraph: { spacing: { before: 180, after: 100, line: 360 } } },
        heading3: { run: { font: { eastAsia: fontCN, ascii: fontEN }, size: 23, bold: true, color: C.black }, paragraph: { spacing: { before: 140, after: 80, line: 360 } } }
      },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 850, left: 1134, right: 1134 } } }, footers: footer(), children }]
    });
    const out = path.join(outDir, `第${n}天-朵朵英语第1单元学练.docx`);
    fs.writeFileSync(out, await Packer.toBuffer(doc));
    console.log(out);
  }
})();
