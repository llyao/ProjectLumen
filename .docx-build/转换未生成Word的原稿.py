from pathlib import Path
import hashlib
import importlib.util
import json
import re
import sys

from docx.shared import Pt

sys.stdout.reconfigure(encoding='utf-8')
工作目录 = Path(__file__).resolve().parent
原稿目录 = 工作目录.parent / '朵朵/6年级/记录/英语/原稿'
规格 = importlib.util.spec_from_file_location('排版工具', 工作目录 / '生成英语空白试卷Word.py')
排版 = importlib.util.module_from_spec(规格)
规格.loader.exec_module(排版)
清单路径 = 工作目录 / '本次原稿转Word清单.json'


def 文件摘要(路径):
    return hashlib.sha256(路径.read_bytes()).hexdigest()


def 是题号(块):
    return bool(re.match(r'^\*\*\d+\.\*\*', 块))


def 转换(输入):
    卷号 = re.search(r'L\d{3}', 输入.name).group()
    文档 = 排版.创建文档(卷号)
    各块 = 排版.提取块(输入.read_text(encoding='utf-8-sig'))
    正在题目中 = False
    是阅读部分 = False
    正在阅读前文 = False
    索引 = 0
    while 索引 < len(各块):
        块 = 各块[索引]
        后块 = 各块[索引 + 1] if 索引 + 1 < len(各块) else ''
        各行 = [行.rstrip() for 行 in 块.splitlines()]
        新节点起点 = len(文档._element.body)
        if 块.startswith('## '):
            是阅读部分 = 块[3:] != '基础练习'
        if 块.startswith('### '):
            正在阅读前文 = 是阅读部分
        if 块.startswith('#') or 块 == '---':
            正在题目中 = False
        if 是题号(块):
            正在阅读前文 = False
            正在题目中 = len(re.findall(r'\*\*\d+\.\*\*', 块)) == 1
        后续同题 = bool(正在题目中 and 后块 and not 是题号(后块)
                        and not 后块.startswith('#') and 后块 != '---')
        if 块 == '---':
            pass
        elif 块.startswith('# '):
            排版.段落(文档, 块[2:], 字号=17, 加粗=True, 居中=True,
                    行高=24, 段后=7, 跟随=True)
        elif 块.startswith('## '):
            行 = 排版.段落(文档, 块[3:], 字号=13.5, 加粗=True,
                         行高=20, 段前=7, 段后=4, 跟随=True)
            if 块[3:] == '阅读训练' and 卷号 != 'L001':
                行.paragraph_format.page_break_before = True
        elif 块.startswith('### '):
            行 = 排版.段落(文档, 块[4:], 字号=12, 加粗=True,
                         行高=18, 段前=7, 段后=4, 跟随=True)
            if (卷号 in ('L002', 'L003') and 块.startswith('### 四、')) or (
                    卷号 == 'L007' and 块.startswith('### 三、')):
                行.paragraph_format.page_break_before = True
        elif 块.startswith('**Exercise for Grade'):
            排版.段落(文档, 块, 字号=11.5, 居中=True, 行高=17, 段后=6, 跟随=True)
        elif 块.startswith('班级：'):
            for 行 in 各行:
                排版.段落(文档, 行, 字号=10.5, 行高=17, 段后=1, 跟随=True)
        elif 块.startswith('|'):
            数据 = [[re.sub(r'<br\s*/?>', '\n', 格.strip(), flags=re.I)
                    for 格 in 行.strip('|').split('|')]
                   for 行 in 各行 if not re.fullmatch(r'[| :\-]+', 行)]
            宽度 = None
            if 数据[0] == ['题号', '音标', '单词', '中文释义', '词性']:
                宽度 = [12, 46, 38, 42, 36]
            elif 数据[0] == ['题号', 'Classroom instruction', 'Subject']:
                宽度 = [12, 126, 36]
            elif 数据[0][0] == '类型':
                宽度 = [24, 30, 55, 65]
            表 = 排版.添加表格(文档, 数据, 宽度=宽度)
            if 数据[0] == ['题目', '题目']:
                for 行号, 行 in enumerate(表.rows[1:], 1):
                    for 格 in 行.cells:
                        排版.设置段落(格.paragraphs[0], 行高=20, 段前=2, 段后=3,
                                    跟随=行号 < len(表.rows)-1)
            排版.段落(文档, '', 行高=3, 段后=1)
        elif len(各行) > 1 and all(re.match(r'^[A-D]\. ', 行) for 行 in 各行):
            排版.添加选项(文档, 各行)
        elif 是题号(块):
            if re.fullmatch(r'\*\*\d+\.\*\*\s*（\s*）', 块) and re.match(r'^[A-D]\. ', 后块):
                排版.添加选项(文档, [行.rstrip() for 行 in 后块.splitlines()], 题号=块)
                索引 += 1
            else:
                排版.段落(文档, 块.replace('  \n', '\n'), 行高=16,
                        段后=2 if 后续同题 else 4, 跟随=后续同题)
        elif re.fullmatch(r'(?:（\d+）)?＿+[!！?？]?', 块):
            答题文字 = re.sub(r'＿+', '_' * 74, 块)
            排版.段落(文档, 答题文字, 行高=22, 段后=4, 跟随=后续同题)
        elif 正在题目中:
            排版.段落(文档, 块, 行高=22, 段后=4, 跟随=后续同题)
        elif len(各行) > 1:
            for 行号, 行 in enumerate(各行):
                排版.段落(文档, 行, 行高=16, 段后=3, 不拆=True,
                        跟随=块.startswith('- ') and 行号 < len(各行)-1)
        else:
            是说明 = 块.startswith('**') and 块.endswith('**')
            排版.段落(文档, 块, 行高=16, 段后=4, 跟随=是说明, 不拆=True)
        if 正在阅读前文 and 后块 and not 后块.startswith('#') and 后块 != '---':
            # 阅读文章、表格和首题一起排，避免信件结尾或首段落单独跨页。
            for 节点 in list(文档._element.body)[新节点起点 - 1:-1]:
                各段 = [节点] if 节点.tag == 排版.qn('w:p') else 节点.findall('.//' + 排版.qn('w:p'))
                for 段节点 in 各段:
                    段节点.get_or_add_pPr().get_or_add_keepNext().val = True
        索引 += 1
    已到阅读 = False
    for 行 in 文档.paragraphs:
        格式 = 行.paragraph_format
        if 行.text == '阅读训练':
            已到阅读 = True
        if 卷号 == 'L007':
            if 格式.space_after is not None:
                格式.space_after = Pt(max(0, 格式.space_after.pt - 2))
            if 格式.space_before is not None:
                格式.space_before = Pt(max(0, 格式.space_before.pt - 2))
        elif 卷号 == 'L003' and 已到阅读:
            if 格式.line_spacing is not None and isinstance(格式.line_spacing, int):
                高度 = 格式.line_spacing.pt
                if 高度 in (16, 22):
                    格式.line_spacing = Pt(高度 - 1)
            if 格式.space_after is not None:
                格式.space_after = Pt(max(0, 格式.space_after.pt - 2))
    输出 = 输入.with_suffix('.docx')
    文档.save(输出)
    print('已生成：', 输出.name)


if __name__ == '__main__':
    if 清单路径.exists():
        清单 = json.loads(清单路径.read_text(encoding='utf-8'))
    else:
        待转换 = [路径 for 路径 in sorted(原稿目录.glob('*.md'))
               if not 路径.with_suffix('.docx').exists()]
        清单 = {
            '待转换': [路径.name for 路径 in 待转换],
            '原文件摘要': {路径.name: 文件摘要(路径) for 路径 in 原稿目录.iterdir()
                        if 路径.is_file() and 路径.suffix.lower() in ('.md', '.docx')},
        }
        清单路径.write_text(json.dumps(清单, ensure_ascii=False, indent=2), encoding='utf-8')
    for 文件名, 原摘要 in 清单['原文件摘要'].items():
        assert 文件摘要(原稿目录 / 文件名) == 原摘要, f'原文件发生变化，请核对：{文件名}'
    for 文件名 in 清单['待转换']:
        转换(原稿目录 / 文件名)
