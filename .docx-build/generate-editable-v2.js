const fs=require('fs'), path=require('path');
const {Document,Packer,Paragraph,TextRun,PageBreak,AlignmentType}=require('docx');
const root=path.resolve(__dirname,'..');
const outDir=path.join(root,'朵朵','6年级','记录','英语','原卷Word文字稿'); fs.mkdirSync(outDir,{recursive:true});
const sets=[['L009','民乐学校6年级第一学期英语练习L009-可编辑整理稿.docx'],['L008','民乐学校6年级第一学期英语练习L008-可编辑整理稿.docx'],['2026-09-16','民乐学校6年级第一学期英语练习L010-可编辑整理稿.docx']];
const hand=/^(朵朵|我的答案|订正|错题|[A-Z]?\s*[（(]?[A-D][）)]?$)/i;
function r(t,o={}){return new TextRun({text:t,font:{eastAsia:'宋体',ascii:'Times New Roman'},size:o.size||21,bold:!!o.bold,color:o.color||'000000'})}
function p(t,o={}){return new Paragraph({alignment:o.align||AlignmentType.LEFT,spacing:{before:o.before??40,after:o.after??40,line:330},children:[r(t,o)]})}
function clean(line){let s=line.trim(); if(!s)return ''; if(/^=====/.test(s))return s; s=s.replace(/\s{2,}/g,' '); return s;}
(async()=>{for(const [key,outName] of sets){const txt=fs.readFileSync(path.join(root,'.docx-build','ocr',key+'.txt'),'utf8');const lines=txt.split(/\r?\n/);const children=[p(key==='2026-09-16'?'民乐学校六年级第一学期英语练习L010':`民乐学校六年级第一学期英语练习${key}`,{align:AlignmentType.CENTER,size:28,bold:true}),p('可编辑文字整理稿：按原卷图片识别并整理，空白线、勾选符号和手写痕迹不作为题目内容。',{align:AlignmentType.CENTER,size:17,color:'666666',after:160})]; for(const raw of lines){const s=clean(raw); if(!s)continue; if(/^=====/.test(s)){children.push(new Paragraph({children:[new PageBreak()]}));children.push(p(s,{size:19,bold:true,color:'555555'}));continue;} if(hand.test(s))continue; if(/^[A-Z]?\s*\(?[A-D]\)?$/.test(s))continue; children.push(p(s));} const doc=new Document({styles:{default:{document:{run:{font:{eastAsia:'宋体',ascii:'Times New Roman'},size:21,color:'000000'},paragraph:{spacing:{line:330}}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:900,bottom:900,left:1100,right:1100}}},children}]});fs.writeFileSync(path.join(outDir,outName),await Packer.toBuffer(doc));console.log(outName)}})();
