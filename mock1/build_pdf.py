"""Render the exam markup (exam.txt / answers.txt) as a two-column, exam-style A4 PDF.

Uses the same markup parser as build.py and prints the HTML with headless Chromium.
Fonts: expects Nanum Myeongjo/Gothic web fonts from Google Fonts in ./fonts (fonts.css plus the
woff2 files it references). The PDF embeds them, so the output does not depend on the viewer's fonts.
"""
import html
import re
import sys
from pathlib import Path

from build import BOX_TITLE, OPTION, parse, spans

from playwright.sync_api import sync_playwright

HERE = Path(__file__).parent
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
HANJA_ROW = re.compile(r"^[\u4e00-\u9fff]{2,}.*\s{2,}\S")
OLD_HANGUL = re.compile(r"[\u1100-\u11ff\ua960-\ua97f\ud7b0-\ud7ff]")
MULTI_CHOICE = re.compile(r"(?=[①②③④⑤])")

CSS = """
@page { size: A4; margin: 13mm 12mm 15mm 12mm; }
* { box-sizing: border-box; }
html { -webkit-print-color-adjust: exact; }
body { margin: 0; font-family: 'Nanum Myeongjo', serif; font-size: 9.4pt; line-height: 1.58;
       color: #000; background: #fff; word-break: keep-all; overflow-wrap: break-word; }
.head { border: 2px solid #000; margin-bottom: 3mm; }
.head .row1 { display: flex; align-items: center; justify-content: space-between;
              padding: 2mm 4mm 1.5mm; border-bottom: 1px solid #000; }
.head .code { font-family: 'Nanum Gothic', sans-serif; font-weight: 800; font-size: 9pt; line-height: 1.15; }
.head .code b { font-size: 20pt; margin-left: 1.5mm; vertical-align: -2pt; }
.head .subject { font-family: 'Nanum Gothic', sans-serif; font-weight: 800; font-size: 21pt; letter-spacing: 1pt; }
.head .date { font-family: 'Nanum Gothic', sans-serif; font-weight: 700; font-size: 9pt; text-align: right; }
.head .title { text-align: center; font-weight: 800; font-size: 12pt; padding: 1.5mm 0 0.5mm; }
.head .meta { text-align: center; font-size: 9.5pt; padding: 0 0 1.8mm; white-space: pre-wrap; }
.old { font-family: 'Noto Serif CJK KR', serif; }
.choices { display: flex; flex-wrap: wrap; column-gap: 1mm; break-inside: avoid; }
.cols { column-count: 2; column-gap: 9mm; column-rule: 0.8px solid #000; column-fill: auto; }
p { margin: 0; text-align: justify; }
.set { font-weight: 800; margin: 3.2mm 0 1.6mm; padding-bottom: 1mm; border-bottom: 1.2px solid #000;
       break-after: avoid; }
.set.first { margin-top: 0; }
.q { font-weight: 700; margin: 3.4mm 0 1.4mm; padding-left: 4.2mm; text-indent: -4.2mm; break-after: avoid; }
.opt { padding-left: 4.2mm; text-indent: -4.2mm; break-inside: avoid; }
.center { text-align: center; }
.right { text-align: right; font-size: 8.6pt; }
.box { border: 1px solid #000; padding: 1.6mm 2.6mm 1.8mm; margin: 1.2mm 0 1.8mm; }
.box.passage { border-left: none; border-right: none; border-top: 1.2px solid #000; border-bottom: 1.2px solid #000;
               padding: 1.8mm 0.5mm; }
.box.passage p.body { text-indent: 2.5mm; }
.box .btitle { text-align: center; font-weight: 700; margin-bottom: 0.6mm; letter-spacing: 1pt; }
.box p.body { margin: 0.2mm 0; }
.box .label { font-weight: 700; margin-top: 1.2mm; }
.row { display: flex; gap: 2mm; }
.row span:first-child { width: 47%; flex: none; letter-spacing: 0.2pt; }
hr { border: none; border-top: 1.2px solid #000; margin: 3mm 0 2mm; }
table.grid { border-collapse: collapse; width: 100%; margin: 1.5mm 0 2mm; break-inside: avoid; }
table.grid td { border: 1px solid #000; text-align: center; padding: 0.8mm 0; font-size: 9pt; }
table.grid tr:nth-child(odd) td { background: #eee; font-weight: 700; }
u { text-underline-offset: 2px; }
.end { text-align: center; font-weight: 700; margin-top: 4mm; }
"""


def inline(text):
    out = []
    for chunk, under in spans(text):
        chunk = html.escape(chunk)
        out.append(f"<u>{chunk}</u>" if under else chunk)
    return "".join(out)


def box_html(body):
    first = next((l for l in body if l.strip()), "")
    kind = "box" if BOX_TITLE.match(first.strip()) else "box passage"
    parts = [f'<div class="{kind}">']
    for raw in body:
        s = raw.strip()
        if not s:
            parts.append('<p class="body">&nbsp;</p>' if kind == "box" else "")
        elif raw.startswith("=R "):
            parts.append(f'<p class="right">{inline(raw[3:])}</p>')
        elif raw.startswith("=C "):
            parts.append(f'<p class="center">{inline(raw[3:])}</p>')
        elif BOX_TITLE.match(s):
            parts.append(f'<p class="btitle">{inline(s)}</p>')
        elif re.fullmatch(r"\([가-힣]\)", s):
            parts.append(f'<p class="label">{inline(s)}</p>')
        elif HANJA_ROW.search(raw):
            left, right = re.split(r"\s{2,}", raw.strip(), maxsplit=1)
            parts.append(f'<p class="row"><span>{inline(left)}</span><span>{inline(right)}</span></p>')
        elif OPTION.match(s):
            parts.append(f'<p class="opt">{inline(s)}</p>')
        else:
            old = " old" if OLD_HANGUL.search(raw) else ""
            parts.append(f'<p class="body{old}">{inline(raw)}</p>')
    parts.append("</div>")
    return "".join(parts)


def to_html(blocks, date_label):
    head, body, in_head, first_set = [], [], True, True
    for kind, payload in blocks:
        if in_head:
            if kind == "LINE":
                in_head = False
            elif kind == "H":
                head.append(f'<div class="title">{inline(payload)}</div>')
            elif kind == "C":
                head.append(f'<div class="meta">{inline(payload)}</div>')
            continue
        if kind == "S":
            body.append(f'<p class="set{" first" if first_set else ""}">{inline(payload)}</p>')
            first_set = False
        elif kind == "Q":
            body.append(f'<p class="q">{inline(payload)}</p>')
        elif kind == "C":
            cls = "end" if payload.startswith("※") else "center"
            body.append(f'<p class="{cls}">{inline(payload)}</p>')
        elif kind == "R":
            body.append(f'<p class="right">{inline(payload)}</p>')
        elif kind == "LINE":
            body.append("<hr>")
        elif kind == "box":
            body.append(box_html(payload))
        elif kind == "table":
            rows = "".join("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r.split("|")) + "</tr>"
                           for r in payload)
            body.append(f'<table class="grid">{rows}</table>')
        elif payload.strip():
            choices = [c.strip() for c in MULTI_CHOICE.split(payload) if c.strip()]
            if OPTION.match(payload) and len(choices) > 1:
                width = "19%" if max(len(c) for c in choices) <= 6 else "32%"
                body.append('<p class="choices">' + "".join(f'<span style="min-width:{width}">{inline(c)}</span>'
                                                          for c in choices) + "</p>")
            else:
                body.append(f'<p class="{"opt" if OPTION.match(payload) else ""}">{inline(payload)}</p>')
    header = f"""<div class="head">
  <div class="row1"><div class="code">과목<br>코드<b>03</b></div><div class="subject">공통국어2</div>
  <div class="date">{date_label}</div></div>
  {''.join(head)}</div>"""
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<link rel="stylesheet" href="fonts/fonts.css"><style>{CSS}</style></head>
<body>{header}<div class="cols">{''.join(body)}</div></body></html>"""


FOOTER = """<div style="width:100%;font-size:8px;padding:0 12mm;display:flex;justify-content:space-between;
font-family:'WenQuanYi Zen Hei',sans-serif;color:#000;">
<span>실전 모의고사 1회</span><span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>"""


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    jobs = (("exam.txt", "실전모의고사1_공통국어2", "대비: 2학기 중간고사"),
            ("answers.txt", "실전모의고사1_정답및해설", "정답 및 해설"))
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROME)
        page = browser.new_page()
        for src, name, date_label in jobs:
            html_path = HERE / f".{name}.html"
            html_path.write_text(to_html(parse(HERE / src), date_label), encoding="utf-8")
            page.goto(html_path.resolve().as_uri())
            page.evaluate("document.fonts.ready")
            page.wait_for_timeout(500)
            page.pdf(path=str(out_dir / f"{name}.pdf"), format="A4", print_background=True,
                     display_header_footer=True, header_template="<span></span>", footer_template=FOOTER,
                     margin={"top": "13mm", "bottom": "15mm", "left": "12mm", "right": "12mm"})
            html_path.unlink()
            print("wrote", out_dir / f"{name}.pdf")
        browser.close()


if __name__ == "__main__":
    main()
