const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, ImageRun, PageBreak, AlignmentType } = require('docx');
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, '朵朵', '6年级', '记录', '英语', '原卷Word版');
fs.mkdirSync(outDir, { recursive: true });
const sets = [
  { name: '民乐学校6年级第一学期英语练习L009-原卷', dir: path.join(root, '朵朵', '6年级', '记录', '英语', '错题', '2026-09-13'), files: ['01-英语练习L009-第1页.jpg','02-英语练习L009-第2页.jpg','03-英语练习L009-第3页.jpg','04-英语练习L009-第4页.jpg'] },
  { name: '民乐学校6年级第一学期英语练习L008-原卷', dir: path.join(root, '朵朵', '6年级', '记录', '英语', '错题', '2026-09-14'), files: ['01-英语练习L008-第1页.jpg','02-英语练习L008-第2页.jpg','03-英语练习L008-第3页.jpg','04-英语练习L008-第4页.jpg'] },
  { name: '民乐学校6年级第一学期英语练习-2026-09-16原卷', dir: path.join(root, '朵朵', '6年级', '记录', '英语', '错题', '2026-09-16'), files: ['01-英语练习.jpg','02-英语练习.jpg'] },
];
function jpgSize(file) {
  const b = fs.readFileSync(file);
  let i = 2;
  while (i < b.length) {
    if (b[i] !== 0xff) { i++; continue; }
    const marker = b[i+1];
    const len = b.readUInt16BE(i+2);
    if ([0xc0,0xc1,0xc2,0xc3,0xc5,0xc6,0xc7,0xc9,0xca,0xcb,0xcd,0xce,0xcf].includes(marker)) return { width: b.readUInt16BE(i+7), height: b.readUInt16BE(i+5) };
    i += 2 + len;
  }
  throw new Error('无法读取图片尺寸: '+file);
}
(async () => {
  for (const set of sets) {
    const children = [];
    for (let i = 0; i < set.files.length; i++) {
      const file = path.join(set.dir, set.files[i]);
      const { width, height } = jpgSize(file);
      const maxW = 7600; // twips, within A4 width with narrow margins
      const maxH = 16100; // twips, within A4 height
      const scale = Math.min(maxW / width, maxH / height);
      const w = Math.round(width * scale);
      const h = Math.round(height * scale);
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new ImageRun({ data: fs.readFileSync(file), transformation: { width: Math.round(w / 15), height: Math.round(h / 15) }, type: 'jpg' })] }));
      if (i < set.files.length - 1) children.push(new Paragraph({ children: [new PageBreak()] }));
    }
    const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 300, bottom: 300, left: 300, right: 300 } } }, children }] });
    const out = path.join(outDir, `${set.name}.docx`);
    fs.writeFileSync(out, await Packer.toBuffer(doc));
    console.log(out);
  }
})();
