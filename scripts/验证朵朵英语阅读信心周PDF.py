from __future__ import annotations

import re
import tempfile
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageDraw, ImageFont


项目目录 = Path(__file__).resolve().parents[1]
PDF目录 = 项目目录 / "朵朵" / "6年级" / "记录" / "英语" / "英语阅读信心周PDF"
预期文件 = [
    "00-阅读解题方法.pdf",
    "01-第1天-我的学校生活.pdf",
    "02-第2天-星期一课程表.pdf",
    "03-第3天-学校阅读社团.pdf",
    "04-第4天-上海博物馆之行.pdf",
    "05-第5天-新同学的第一天.pdf",
    "06-第6天-校园开放日.pdf",
    "07-第7天-复习与进步记录.pdf",
]


def 提取文字(PDF路径: Path) -> tuple[int, str]:
    with fitz.open(PDF路径) as 文档:
        return len(文档), "\n".join(页面.get_text() for 页面 in 文档)


def 生成预览(PDF路径列表: list[Path]) -> Path:
    缩略图: list[tuple[str, Image.Image]] = []
    for PDF路径 in PDF路径列表:
        with fitz.open(PDF路径) as 文档:
            for 页码, 页面 in enumerate(文档):
                像素 = 页面.get_pixmap(matrix=fitz.Matrix(0.55, 0.55), alpha=False)
                图片 = Image.frombytes("RGB", [像素.width, 像素.height], 像素.samples)
                缩略图.append((f"{PDF路径.stem}　第{页码 + 1}页", 图片))
    列数 = 3
    标签高度 = 34
    单元宽度 = max(图片.width for _, 图片 in 缩略图) + 24
    单元高度 = max(图片.height for _, 图片 in 缩略图) + 标签高度 + 24
    行数 = (len(缩略图) + 列数 - 1) // 列数
    画布 = Image.new("RGB", (单元宽度 * 列数, 单元高度 * 行数), "#d9e2e8")
    绘图 = ImageDraw.Draw(画布)
    字体路径 = Path(r"C:\Windows\Fonts\msyh.ttc")
    字体 = ImageFont.truetype(str(字体路径), 16)
    for 序号, (标签, 图片) in enumerate(缩略图):
        横坐标 = (序号 % 列数) * 单元宽度 + 12
        纵坐标 = (序号 // 列数) * 单元高度 + 12
        绘图.text((横坐标, 纵坐标), 标签, fill="#243447", font=字体)
        画布.paste(图片, (横坐标, 纵坐标 + 标签高度))
    预览路径 = Path(tempfile.gettempdir()) / "朵朵英语阅读信心周PDF-版面预览.png"
    画布.save(预览路径)
    return 预览路径


def 主程序() -> None:
    结果: list[tuple[str, int, int]] = []
    路径列表 = [PDF目录 / 文件名 for 文件名 in 预期文件]
    缺失文件 = [路径.name for 路径 in 路径列表 if not 路径.exists()]
    if 缺失文件:
        raise RuntimeError(f"缺少PDF：{'、'.join(缺失文件)}")
    for 序号, PDF路径 in enumerate(路径列表):
        页数, 文字 = 提取文字(PDF路径)
        if "家长答案区" in 文字 or re.search(r"1\.B\s+2\.A", 文字):
            raise RuntimeError(f"PDF中发现答案内容：{PDF路径.name}")
        if "file:///" in 文字:
            raise RuntimeError(f"PDF中仍包含本地文件路径页脚：{PDF路径.name}")
        题数 = len(re.findall(r"(?m)^\s*\d+\.\s+(?:When|What|Which|Why|Where|Who|How)", 文字))
        if 1 <= 序号 <= 6 and 题数 != 5:
            raise RuntimeError(f"{PDF路径.name}题目数为{题数}，不是5题")
        if 页数 != 1:
            raise RuntimeError(f"{PDF路径.name}共有{页数}页，未达到单页目标")
        结果.append((PDF路径.name, 页数, 题数))
    预览路径 = 生成预览(路径列表)
    print("验证通过：PDF数量、题目数量和答案隔离均符合要求。")
    for 文件名, 页数, 题数 in 结果:
        print(f"{文件名}：{页数}页，{题数}道阅读题")
    print(f"版面预览：{预览路径}")


if __name__ == "__main__":
    主程序()
