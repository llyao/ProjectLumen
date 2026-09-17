from pathlib import Path
import json
import re
import sys

from PIL import Image, ImageDraw
import pypdfium2 as pdfium
from pypdf import PdfReader

sys.stdout.reconfigure(encoding='utf-8')
工作目录 = Path(__file__).resolve().parent
清单 = json.loads((工作目录 / '本次原稿转Word清单.json').read_text(encoding='utf-8'))
for 文件名 in 清单['待转换']:
    卷号 = re.search(r'L\d{3}', 文件名).group()
    路径 = 工作目录 / f'{卷号}-新增原稿排版核对.pdf'
    文档 = pdfium.PdfDocument(路径)
    越界 = []
    for 起页 in range(0, len(文档), 6):
        页数 = min(6, len(文档) - 起页)
        总览 = Image.new('RGB', (1800, 880 * ((页数 + 2) // 3)), '#e5e5e5')
        绘图 = ImageDraw.Draw(总览)
        for 序号 in range(页数):
            页号 = 起页 + 序号
            页面 = 文档[页号]
            图像 = 页面.render(scale=1.2).to_pil().convert('RGB')
            图像.thumbnail((584, 834))
            横, 纵 = 序号 % 3 * 600 + 8, 序号 // 3 * 880 + 26
            总览.paste(图像, (横, 纵))
            绘图.text((横, 纵 - 20), f'{卷号} {页号 + 1}', fill='black')
            文字页 = 页面.get_textpage()
            for 字号 in range(文字页.count_chars()):
                字符 = 文字页.get_text_range(字号, 1)
                if not 字符.strip():
                    continue
                左, 下, 右, 上 = 文字页.get_charbox(字号)
                if not (30 <= 左 <= 右 <= 570 and 15 <= 下 <= 上 <= 825):
                    越界.append((页号 + 1, 字符, 左, 下, 右, 上))
        总览.save(工作目录 / f'{卷号}-新增原稿版面-{起页 + 1:02d}.png')
    for 页号, 页面 in enumerate(PdfReader(路径).pages, 1):
        各行 = [行.strip() for 行 in 页面.extract_text().splitlines() if 行.strip()]
        print(卷号, '页', 页号, '行数', len(各行),
              '首', repr(' / '.join(各行[:3])), '尾', repr(' / '.join(各行[-4:])))
    print(卷号, '越界字符数', len(越界), 越界[:5])
