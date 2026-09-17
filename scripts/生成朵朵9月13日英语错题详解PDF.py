from __future__ import annotations

import importlib.util
import tempfile
from pathlib import Path


项目目录 = Path(__file__).resolve().parents[1]
原始文档 = 项目目录 / "朵朵" / "6年级" / "记录" / "英语" / "错题" / "2026-09-13-英语错题逐题详解.md"
输出文件 = 项目目录 / "朵朵" / "6年级" / "记录" / "英语" / "错题" / "2026-09-13-英语错题逐题详解.pdf"
共用生成器路径 = 项目目录 / "scripts" / "生成朵朵英语阅读信心周PDF.py"


def 加载共用生成器():
    规格 = importlib.util.spec_from_file_location("朵朵英语PDF共用生成器", 共用生成器路径)
    if 规格 is None or 规格.loader is None:
        raise RuntimeError(f"无法加载共用生成器：{共用生成器路径}")
    模块 = importlib.util.module_from_spec(规格)
    规格.loader.exec_module(模块)
    return 模块


def 完整网页(正文: str, 基础样式: str) -> str:
    详解样式 = """
@page { size: A4 landscape; margin: 9mm 10mm 10mm 10mm; }
body { font-size: 10pt; line-height: 1.34; color: #243447; }
h1 { margin-bottom: 10px; font-size: 22pt; }
h2 { margin-top: 11px; padding: 5px 8px; border: 0; border-left: 5px solid #0e7490; background: #e6f6fa; font-size: 15pt; break-after: avoid; }
h3 { break-after: avoid; }
p { margin: 5px 0; }
table { table-layout: fixed; margin: 7px 0 12px; font-size: 9pt; }
thead { display: table-header-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
th, td { padding: 5px 6px; overflow-wrap: anywhere; }
th:nth-child(1), td:nth-child(1) { width: 15%; }
th:nth-child(2), td:nth-child(2) { width: 17%; }
th:nth-child(3), td:nth-child(3) { width: 22%; }
th:nth-child(4), td:nth-child(4) { width: 18%; }
th:nth-child(5), td:nth-child(5) { width: 28%; }
code { font-size: 8.8pt; }
"""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>朵朵9月13日英语错题逐题详解</title>
<style>{基础样式}{详解样式}</style>
</head>
<body>{正文}</body>
</html>"""


def 主程序() -> None:
    if not 原始文档.exists():
        raise FileNotFoundError(f"找不到原始文档：{原始文档}")
    if not 共用生成器路径.exists():
        raise FileNotFoundError(f"找不到共用生成器：{共用生成器路径}")
    共用生成器 = 加载共用生成器()
    文本 = 原始文档.read_text(encoding="utf-8")
    网页 = 完整网页(共用生成器.文本转网页(文本), 共用生成器.基础样式)
    输出文件.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="朵朵9月13日英语错题详解-") as 临时目录文本:
        临时目录 = Path(临时目录文本)
        网页路径 = 临时目录 / "朵朵9月13日英语错题逐题详解.html"
        网页路径.write_text(网页, encoding="utf-8")
        共用生成器.打印网页到PDF(网页路径, 输出文件, 临时目录 / "浏览器用户目录")
    print(f"已生成：{输出文件}")


if __name__ == "__main__":
    主程序()
