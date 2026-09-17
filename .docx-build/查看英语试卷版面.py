from pathlib import Path
from PIL import Image, ImageDraw
import pypdfium2 as pdfium
from pypdf import PdfReader
import sys

sys.stdout.reconfigure(encoding='utf-8')
目录 = Path(__file__).resolve().parent
for 卷号 in ['L008', 'L009', 'L010']:
    文件 = 目录 / f'{卷号}-排版核对.pdf'
    文档 = pdfium.PdfDocument(文件)
    总览 = Image.new('RGB', (1800, 1760), '#e5e5e5')
    绘图 = ImageDraw.Draw(总览)
    for 页号, 页面 in enumerate(文档):
        图像 = 页面.render(scale=1).to_pil().convert('RGB')
        图像.thumbnail((584, 834))
        横, 纵 = 页号 % 3 * 600 + 8, 页号 // 3 * 880 + 26
        总览.paste(图像, (横, 纵))
        绘图.text((横, 纵-20), f'{卷号} {页号+1}', fill='black')
        文字页 = 页面.get_textpage()
        for 字号 in range(文字页.count_chars()):
            字符 = 文字页.get_text_range(字号, 1)
            if not 字符.strip():
                continue
            左, 下, 右, 上 = 文字页.get_charbox(字号)
            assert 30 <= 左 <= 右 <= 570, (卷号, 页号+1, 字符, 左, 右)
            assert 15 <= 下 <= 上 <= 825, (卷号, 页号+1, 字符, 下, 上)
    总览.save(目录 / f'{卷号}-版面总览.png')
    for 页号, 页面 in enumerate(PdfReader(文件).pages):
        各行 = [行 for 行 in 页面.extract_text().splitlines() if 行.strip()]
        print(卷号, 页号+1, '行数', len(各行), '开头', repr(' / '.join(各行[2:4])), '结尾', repr(' / '.join(各行[-3:])))
    print(卷号, '全部文字边界检查通过')
