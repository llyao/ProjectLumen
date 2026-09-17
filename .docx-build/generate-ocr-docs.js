const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, PageBreak, AlignmentType, HeadingLevel } = require('docx');
const root = path.resolve(__dirname, '..');
const ocrDir = path.join(root, '.docx-build', 'ocr');
const outDir = path.join(root, '朵朵', '6年级', '记录', '英语', '原卷Word文字稿');
fs.mkdirSync(outDir, { recursive: true });
const files = [
  ['L009', '民乐学校6年级第一学期英语练习L009-原卷文字识别稿.docx'],
  ['L008', '民乐学校6年级第一学期英语练习L008-原卷文字识别稿.docx'],
  ['2026-09-16', '民乐学校6年级第一学期英语练习-2026-09-16原卷文字识别稿.docx'],
];
function run(text, opts={}) { return new TextRun({ text, font: { eastAsia: '宋体', ascii: 'Times New Roman' }, size: opts.size || 21, bold: !!opts.bold, color: opts.color || '000000' }); }
function p(text, opts={}) { return new Paragraph({ alignment: opts.align || AlignmentType.LEFT, spacing: { before: opts.before ?? 60, after: opts.after ?? 60, line: 330 }, children: [run(text, opts)] }); }
(async()=>{
 for (const [key, name] of files) {
  const txt=fs.readFileSync(path.join(ocrDir, `${key}.txt`),'utf8');
  const lines=txt.split(/\r?\n/);
  const children=[p(`民乐学校六年级第一学期英语练习${key==='2026-09-16'?'':key}`,{align:AlignmentType.CENTER,size:28,bold:true}), p('以下内容为原卷图片文字识别稿，按原卷页码分隔；个别字母、标点或手写内容可能需要对照原卷核验。',{align:AlignmentType.CENTER,size:17,color:'666666',after:180})];
  for (const line of lines) {
   if (!line.trim()) { children.push(p('')); continue; }
   if (line.startsWith('===== 原卷第')) { children.push(new Paragraph({ children:[new PageBreak()] })); children.push(p(line,{size:20,bold:true,color:'555555'})); }
   else children.push(p(line));
  }
  const doc=new Document({
   styles:{
    default:{
     document:{run:{font:{eastAsia:'宋体',ascii:'Times New Roman'},size:21,color:'000000'},paragraph:{spacing:{line:330}}}
    }
   },
   sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:900,bottom:900,left:1100,right:1100}}},children}]
  });
  fs.writeFileSync(path.join(outDir,name),await Packer.toBuffer(doc));
  console.log(name);
 }
})();
