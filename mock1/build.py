"""Build the mock exam and answer key as HWPX (한글) and DOCX from the simple markup files.

Markup (one directive per line):
  =H title        =C centered      =R right-aligned (source line)
  =S set header   =Q question stem =LINE horizontal rule
  =BOX ... =END   boxed text (<보 기> etc.)
  =TABLE ... =END pipe-separated grid
  __text__        underlined span
Other lines are plain paragraphs; lines starting with ①~⑤ get a hanging indent.
"""
import re
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=DeprecationWarning)

from hwpx.document import HwpxDocument
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

HERE = Path(__file__).parent
OPTION = re.compile(r"^[①②③④⑤]")
BOX_TITLE = re.compile(r"^<[^>]{1,8}>$|^\[.*풀이\]$")
COL_WIDTH = 23800  # HWPUNIT, a little under one of two columns on A4 with 15mm margins


def parse(path):
    blocks, lines = [], Path(path).read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if line in ("=BOX", "=TABLE"):
            kind, body = line[1:].lower(), []
            i += 1
            while lines[i].rstrip() != "=END":
                body.append(lines[i].rstrip())
                i += 1
            blocks.append((kind, body))
        elif line.startswith("=") and line[1:2].isupper() and line[1:].split(" ")[0] in ("H", "C", "R", "S", "Q", "LINE"):
            tag, _, text = line[1:].partition(" ")
            blocks.append((tag, text))
        else:
            blocks.append(("P", line))
        i += 1
    return blocks


def spans(text):
    """Split into (text, underline) pieces."""
    parts = re.split(r"(__.+?__)", text)
    return [(p[2:-2], True) if p.startswith("__") and p.endswith("__") else (p, False) for p in parts if p]


def line_kind(text):
    """Classify a line inside a box into an alignment/format key."""
    if text.startswith("=R "):
        return "right", text[3:]
    if text.startswith("=C "):
        return "center", text[3:]
    if BOX_TITLE.match(text.strip()):
        return "center_bold", text.strip()
    if OPTION.match(text):
        return "option", text
    return "body", text


# ---------------------------------------------------------------- HWPX
class HwpxBuilder:
    FORMATS = {
        "body": dict(alignment="JUSTIFY", line_spacing_percent=150),
        "center": dict(alignment="CENTER", line_spacing_percent=150),
        "center_bold": dict(alignment="CENTER", line_spacing_percent=150),
        "right": dict(alignment="RIGHT", line_spacing_percent=150),
        "option": dict(alignment="JUSTIFY", line_spacing_percent=150, indent_left_mm=4, first_line_indent_mm=-4),
        "question": dict(alignment="JUSTIFY", line_spacing_percent=150, spacing_before_pt=8, spacing_after_pt=3,
                         indent_left_mm=4, first_line_indent_mm=-4, keep_with_next=True),
        "set": dict(alignment="LEFT", line_spacing_percent=150, spacing_before_pt=10, spacing_after_pt=4,
                    keep_with_next=True),
        "title": dict(alignment="CENTER", line_spacing_percent=130, spacing_after_pt=4),
    }

    def __init__(self):
        self.doc = HwpxDocument.new()
        self.doc.set_page_margins(left=4252, right=4252, top=4252, bottom=3700, header=1800, footer=1800)
        self.ids = {}
        for key, fmt in self.FORMATS.items():
            self.doc.add_paragraph("")
            self.doc.set_paragraph_format(paragraph_index=len(self.doc.paragraphs) - 1, **fmt)
            self.ids[key] = self.doc.paragraphs[-1].para_pr_id_ref
        self.columns_set = False

    def para(self, key, text, bold=False, size=None, container=None):
        target = container if container is not None else self.doc
        p = target.add_paragraph("", para_pr_id_ref=self.ids[key])
        for chunk, under in spans(text):
            p.add_run(chunk, bold=bold or key in ("center_bold", "question", "set"), underline=under, size=size)
        return p

    def start_columns(self, paragraph):
        if not self.columns_set:
            self.doc.set_columns(2, same_gap=2268, separator_type="SOLID", separator_width="0.12 mm",
                                 separator_color="#000000", paragraph=paragraph)
            self.columns_set = True

    def box(self, body):
        table = self.doc.add_table(1, 1, width=COL_WIDTH)
        cell = table.cell(0, 0)
        cell.set_margins(left=560, right=560, top=420, bottom=420)
        for idx, raw in enumerate(body):
            key, text = line_kind(raw)
            p = self.para(key, text, container=cell)
        # drop the empty default paragraph the cell was created with
        first = cell.paragraphs[0]
        if not first.text:
            first.remove()
        self.doc.add_paragraph("", para_pr_id_ref=self.ids["body"])

    def grid(self, body):
        rows = [r.split("|") for r in body]
        cols = max(len(r) for r in rows)
        table = self.doc.add_table(len(rows), cols, width=COL_WIDTH)
        for r, row in enumerate(rows):
            for c in range(cols):
                table.set_cell_text(r, c, row[c] if c < len(row) else "")
        self.doc.add_paragraph("", para_pr_id_ref=self.ids["body"])

    def build(self, blocks, out):
        for kind, payload in blocks:
            if kind == "H":
                self.para("title", payload, bold=True, size=16)
            elif kind == "C":
                self.para("center", payload)
            elif kind == "R":
                self.para("right", payload)
            elif kind == "LINE":
                p = self.para("body", "")
                self.doc.set_paragraph_format(paragraph_index=len(self.doc.paragraphs) - 1, bottom_border=True,
                                              border_color="#000000", border_width="0.4 mm")
            elif kind == "S":
                p = self.para("set", payload)
                self.start_columns(p)
            elif kind == "Q":
                p = self.para("question", payload)
                self.start_columns(p)
            elif kind == "box":
                self.box(payload)
            elif kind == "table":
                self.grid(payload)
            else:
                key, text = line_kind(payload)
                self.para(key, text)
        # remove the helper paragraphs used to register paragraph formats; paragraph 0 carries the
        # section properties (page size, margins) and must stay
        for _ in self.FORMATS:
            self.doc.paragraphs[1].remove()
        report = self.doc.validate()
        self.doc.save_to_path(str(out))
        return report


# ---------------------------------------------------------------- DOCX
class DocxBuilder:
    FONT = "함초롬바탕"

    def __init__(self):
        self.doc = Document()
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Mm(210), Mm(297)
        sec.left_margin = sec.right_margin = sec.top_margin = Mm(15)
        sec.bottom_margin = Mm(13)
        style = self.doc.styles["Normal"]
        style.font.name = self.FONT
        style.font.size = Pt(10)
        style.element.rPr.rFonts.set(qn("w:eastAsia"), self.FONT)
        style.paragraph_format.space_after = Pt(0)
        style.paragraph_format.line_spacing = 1.3
        self.columns_set = False

    def para(self, key, text, bold=False, size=None, container=None):
        target = container if container is not None else self.doc
        p = target.add_paragraph()
        fmt = p.paragraph_format
        fmt.alignment = {
            "center": WD_ALIGN_PARAGRAPH.CENTER, "center_bold": WD_ALIGN_PARAGRAPH.CENTER,
            "title": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }.get(key, WD_ALIGN_PARAGRAPH.JUSTIFY)
        if key in ("option", "question"):
            fmt.left_indent, fmt.first_line_indent = Mm(4), Mm(-4)
        if key == "question":
            fmt.space_before, fmt.space_after, fmt.keep_with_next = Pt(8), Pt(3), True
        if key == "set":
            fmt.space_before, fmt.space_after, fmt.keep_with_next = Pt(10), Pt(4), True
        for chunk, under in spans(text):
            run = p.add_run(chunk)
            run.bold = bold or key in ("center_bold", "question", "set", "title")
            run.underline = under
            if size:
                run.font.size = Pt(size)
        return p

    def start_columns(self):
        if self.columns_set:
            return
        # everything so far stays single-column; a continuous section break starts the 2-column part
        new = self.doc.add_section(0)  # WD_SECTION.CONTINUOUS
        cols = new._sectPr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            new._sectPr.append(cols)
        cols.set(qn("w:num"), "2")
        cols.set(qn("w:space"), str(Mm(8).twips))
        cols.set(qn("w:sep"), "1")
        self.columns_set = True

    @staticmethod
    def borders(table):
        tbl_pr = table._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "6")
            el.set(qn("w:color"), "000000")
            b.append(el)
        tbl_pr.append(b)

    def box(self, body):
        table = self.doc.add_table(rows=1, cols=1)
        self.borders(table)
        cell = table.cell(0, 0)
        cell._element.remove(cell.paragraphs[0]._element)
        for raw in body:
            key, text = line_kind(raw)
            self.para(key, text, container=cell)
        self.doc.add_paragraph()

    def grid(self, body):
        rows = [r.split("|") for r in body]
        cols = max(len(r) for r in rows)
        table = self.doc.add_table(rows=len(rows), cols=cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        self.borders(table)
        for r, row in enumerate(rows):
            for c in range(cols):
                cell = table.cell(r, c)
                cell.text = row[c] if c < len(row) else ""
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph()

    def build(self, blocks, out):
        for kind, payload in blocks:
            if kind == "H":
                self.para("title", payload, size=16)
            elif kind == "C":
                self.para("center", payload)
            elif kind == "R":
                self.para("right", payload)
            elif kind == "LINE":
                p = self.para("body", "")
                pbdr = OxmlElement("w:pBdr")
                bottom = OxmlElement("w:bottom")
                for k, v in (("val", "single"), ("sz", "12"), ("space", "1"), ("color", "000000")):
                    bottom.set(qn(f"w:{k}"), v)
                pbdr.append(bottom)
                p._p.get_or_add_pPr().append(pbdr)
            elif kind in ("S", "Q"):
                self.start_columns()
                self.para("set" if kind == "S" else "question", payload)
            elif kind == "box":
                self.box(payload)
            elif kind == "table":
                self.grid(payload)
            else:
                key, text = line_kind(payload)
                self.para(key, text)
        self.doc.save(str(out))


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE
    out_dir.mkdir(parents=True, exist_ok=True)
    for src, name in (("exam.txt", "실전모의고사1_공통국어2"), ("answers.txt", "실전모의고사1_정답및해설")):
        blocks = parse(HERE / src)
        report = HwpxBuilder().build(blocks, out_dir / f"{name}.hwpx")
        DocxBuilder().build(blocks, out_dir / f"{name}.docx")
        print(name, "hwpx validation issues:", len(report.issues))


if __name__ == "__main__":
    main()
