from __future__ import annotations

import html
import re
import subprocess
import tempfile
from pathlib import Path


项目目录 = Path(__file__).resolve().parents[1]
原始文档 = 项目目录 / "朵朵" / "6年级" / "记录" / "英语" / "2026-09-14-英语阅读信心周计划.md"
输出目录 = 项目目录 / "朵朵" / "6年级" / "记录" / "英语" / "英语阅读信心周PDF"
浏览器路径 = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")


def 行内格式(内容: str) -> str:
    结果 = html.escape(内容.strip())
    结果 = re.sub(r"`([^`]+)`", r"<code>\1</code>", 结果)
    结果 = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", 结果)
    return 结果


def 表格转网页(行列表: list[str]) -> str:
    数据行: list[list[str]] = []
    for 行 in 行列表:
        单元格 = [项目.strip() for 项目 in 行.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-+:?", 项目) for 项目 in 单元格):
            continue
        数据行.append(单元格)
    if not 数据行:
        return ""
    表头 = "".join(f"<th>{行内格式(项目)}</th>" for 项目 in 数据行[0])
    表体 = "".join(
        "<tr>" + "".join(f"<td>{行内格式(项目)}</td>" for 项目 in 行) + "</tr>"
        for 行 in 数据行[1:]
    )
    return f"<table><thead><tr>{表头}</tr></thead><tbody>{表体}</tbody></table>"


def 文本转网页(内容: str) -> str:
    行 = 内容.splitlines()
    结果: list[str] = []
    序号 = 0
    while 序号 < len(行):
        当前 = 行[序号].rstrip()
        if not 当前.strip():
            序号 += 1
            continue
        if 当前.startswith("|"):
            表格行: list[str] = []
            while 序号 < len(行) and 行[序号].lstrip().startswith("|"):
                表格行.append(行[序号])
                序号 += 1
            结果.append(表格转网页(表格行))
            continue
        标题 = re.match(r"^(#{1,3})\s+(.+)$", 当前)
        if 标题:
            级别 = len(标题.group(1))
            结果.append(f"<h{级别}>{行内格式(标题.group(2))}</h{级别}>")
            序号 += 1
            continue
        if 当前.strip() == "---":
            结果.append("<hr>")
            序号 += 1
            continue
        if 当前.startswith("- "):
            项目: list[str] = []
            while 序号 < len(行) and 行[序号].startswith("- "):
                项目.append(f"<li>{行内格式(行[序号][2:])}</li>")
                序号 += 1
            结果.append("<ul>" + "".join(项目) + "</ul>")
            continue
        if re.match(r"^\d+\.\s+", 当前):
            结果.append(f"<div class='numbered'>{行内格式(当前)}</div>")
            序号 += 1
            continue
        if re.match(r"^\s+[ABC]\.\s+", 当前):
            结果.append(f"<div class='option'>{行内格式(当前)}</div>")
            序号 += 1
            continue
        类名 = ""
        if 当前.startswith("完成后检查："):
            类名 = " class='check-note'"
        elif 当前.startswith("词语提示："):
            类名 = " class='word-note'"
        结果.append(f"<p{类名}>{行内格式(当前)}</p>")
        序号 += 1
    return "\n".join(结果)


基础样式 = """
@page { size: A4; margin: 9mm 12mm 9mm 12mm; }
* { box-sizing: border-box; }
body { margin: 0; color: #253247; font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif; font-size: 11.7pt; line-height: 1.32; }
h1 { margin: 0 0 8px; color: #164e63; font-size: 21pt; text-align: center; }
h2 { margin: 7px 0 4px; color: #155e75; font-size: 15pt; border-bottom: 1.5px solid #a5d8e6; padding-bottom: 2px; }
h3 { margin: 5px 0 3px; color: #334155; font-size: 12.8pt; }
p { margin: 4px 0; }
ul { margin: 4px 0 7px 22px; padding: 0; }
li { margin: 2px 0; }
code { padding: 0 3px; border-radius: 3px; background: #e8f4f7; color: #0f4c5c; font-family: Consolas, "Microsoft YaHei", monospace; }
table { width: 100%; margin: 6px 0 8px; border-collapse: collapse; font-size: 10.7pt; }
th, td { padding: 4px 6px; border: 1px solid #94a3b8; text-align: left; vertical-align: top; }
th { background: #dff3f7; color: #164e63; }
.paper-top { display: flex; justify-content: space-between; align-items: end; margin-bottom: 5px; color: #475569; font-size: 10.5pt; }
.method-strip { margin: 6px 0 8px; padding: 6px 9px; border: 1px solid #67b7c7; border-radius: 6px; background: #edf9fb; color: #164e63; font-weight: 700; text-align: center; }
.numbered { margin: 5px 0 1px; font-weight: 700; }
.option { margin: 1px 0 1px 20px; }
.check-note { margin-top: 8px; padding: 5px 8px; border-left: 4px solid #f4b942; background: #fff8e6; font-size: 10.8pt; }
.word-note { padding: 4px 7px; border-radius: 4px; background: #f1f5f9; font-size: 10.5pt; }
.answer-line { margin-top: 7px; padding-top: 5px; border-top: 1px dashed #94a3b8; color: #475569; font-size: 10.5pt; }
.review-box { margin: 8px 0; min-height: 36px; border-bottom: 1px solid #94a3b8; }
hr { margin: 8px 0; border: 0; border-top: 1px solid #cbd5e1; }
"""


def 完整网页(正文: str, 类型: str, 天数: int | None = None) -> str:
    顶部 = ""
    方法条 = ""
    底部 = ""
    if 类型 == "每日":
        顶部 = (
            "<div class='paper-top'><span>朵朵英语阅读信心周</span>"
            f"<span>第{天数}天　日期：________　用时：________</span></div>"
        )
        方法条 = "<div class='method-strip'>先看题 → 圈关键词 → 回原文 → 画依据 → 比意思 → 再选择</div>"
        if 天数 and 天数 <= 6:
            底部 = "<div class='answer-line'>第一次正确：____／5　□没有空题　□每题画出依据　今天的收获：________________</div>"
    额外样式 = ""
    if 类型 == "方法":
        额外样式 = "@page{margin:8mm 10mm} body{font-size:10.2pt;line-height:1.25} h1{font-size:18pt;margin-bottom:5px} h2{font-size:13.2pt;margin-top:5px} table{font-size:9.2pt;margin:4px 0 5px} th,td{padding:3px 5px}"
    return f"""<!doctype html>
<html lang='zh-CN'>
<head><meta charset='utf-8'><title>朵朵英语阅读信心周</title><style>{基础样式}{额外样式}</style></head>
<body>{顶部}{方法条}{正文}{底部}</body>
</html>"""


def 截取(全文: str, 开始: str, 结束: str | None) -> str:
    起点 = 全文.index(开始)
    终点 = 全文.index(结束, 起点) if 结束 else len(全文)
    return 全文[起点:终点].strip()


def 生成文件定义(全文: str) -> list[tuple[str, str, str, int | None]]:
    文件: list[tuple[str, str, str, int | None]] = []
    方法内容 = 截取(全文, "# 朵朵英语阅读信心周计划", "## 第1天：")
    文件.append(("00-阅读解题方法.pdf", 方法内容, "方法", None))
    标题和文件名 = [
        ("## 第1天：", "## 第2天：", "01-第1天-我的学校生活.pdf"),
        ("## 第2天：", "## 第3天：", "02-第2天-星期一课程表.pdf"),
        ("## 第3天：", "## 第4天：", "03-第3天-学校阅读社团.pdf"),
        ("## 第4天：", "## 第5天：", "04-第4天-上海博物馆之行.pdf"),
        ("## 第5天：", "## 第6天：", "05-第5天-新同学的第一天.pdf"),
        ("## 第6天：", "## 第7天：", "06-第6天-校园开放日.pdf"),
    ]
    for 天数, (开始, 结束, 文件名) in enumerate(标题和文件名, start=1):
        文件.append((文件名, 截取(全文, 开始, 结束), "每日", 天数))
    第七天 = 截取(全文, "## 第7天：", "## 家长陪伴方法")
    第七天 += """

## 本周进步记录

- 本周第一次答题共做对：____／30题。
- 第七天重做错题后做对：____题。
- 我已经会使用的方法：________________________________。
- 家长看到的具体进步：________________________________。
"""
    文件.append(("07-第7天-复习与进步记录.pdf", 第七天, "每日", 7))
    return 文件


def 打印网页到PDF(网页路径: Path, PDF路径: Path, 用户目录: Path) -> None:
    if PDF路径.exists():
        PDF路径.unlink()
    命令 = [
        str(浏览器路径),
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-pdf-header-footer",
        f"--user-data-dir={用户目录}",
        f"--print-to-pdf={PDF路径}",
        网页路径.as_uri(),
    ]
    结果 = subprocess.run(命令, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if 结果.returncode != 0 or not PDF路径.exists():
        raise RuntimeError(f"PDF生成失败：{PDF路径.name}\n{结果.stdout}\n{结果.stderr}")


def 主程序() -> None:
    if not 原始文档.exists():
        raise FileNotFoundError(f"找不到原始文档：{原始文档}")
    if not 浏览器路径.exists():
        raise FileNotFoundError(f"找不到浏览器：{浏览器路径}")
    输出目录.mkdir(parents=True, exist_ok=True)
    全文 = 原始文档.read_text(encoding="utf-8")
    文件定义 = 生成文件定义(全文)
    with tempfile.TemporaryDirectory(prefix="朵朵英语PDF-") as 临时目录文本:
        临时目录 = Path(临时目录文本)
        用户目录 = 临时目录 / "浏览器用户目录"
        for 序号, (文件名, 内容, 类型, 天数) in enumerate(文件定义):
            网页路径 = 临时目录 / f"页面-{序号}.html"
            网页路径.write_text(完整网页(文本转网页(内容), 类型, 天数), encoding="utf-8")
            打印网页到PDF(网页路径, 输出目录 / 文件名, 用户目录)
            print(f"已生成：{输出目录 / 文件名}")


if __name__ == "__main__":
    主程序()
