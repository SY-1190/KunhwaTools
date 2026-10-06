from __future__ import annotations

import base64
import io
import math
import re
from pathlib import Path

import pymupdf as fitz
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

MAX_PAGES = 150
MAX_OPERATIONS = 500


def open_pdf(data: bytes):
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise ValueError("PDF 파일을 읽을 수 없습니다. 정상 PDF인지 확인해 주세요.") from exc
    if doc.needs_pass or not 0 < doc.page_count <= MAX_PAGES:
        doc.close()
        raise ValueError("암호가 없는 150페이지 이하 PDF를 선택해 주세요.")
    if not doc.permissions & fitz.PDF_PERM_MODIFY:
        doc.close()
        raise ValueError("편집이 허용되지 않은 PDF입니다.")
    if doc.get_sigflags() > 0:
        doc.close()
        raise ValueError("전자서명이 있는 PDF는 서명 무효화를 방지하기 위해 편집할 수 없습니다.")
    for page in doc:
        if page.rotation:
            page.remove_rotation()
    return doc


def color_hex(value: int) -> str:
    return f"#{value & 0xffffff:06x}"


def objects(page):
    result = []
    for bi, block in enumerate(page.get_text("dict")["blocks"]):
        if block["type"] != 0:
            continue
        for li, line in enumerate(block["lines"]):
            spans = line["spans"]
            text = "".join(span["text"] for span in spans)
            if not text.strip():
                continue
            first = spans[0]
            result.append({"id": f"t{bi}-{li}", "type": "text", "rect": list(line["bbox"]),
                           "text": text, "size": round(first["size"], 2), "color": color_hex(first["color"]),
                           "font": first["font"]})
    for index, item in enumerate(page.get_image_info(xrefs=True)):
        if item.get("xref"):
            extracted = page.parent.extract_image(item["xref"])
            if extracted.get("smask"):
                pixels = fitz.Pixmap(page.parent, item["xref"])
                mask = fitz.Pixmap(page.parent, extracted["smask"])
                extracted = {"ext": "png", "image": fitz.Pixmap(pixels, mask).tobytes("png")}
            image = "data:image/" + extracted["ext"] + ";base64," + base64.b64encode(extracted["image"]).decode("ascii")
        else:
            pix = page.get_pixmap(clip=fitz.Rect(item["bbox"]), alpha=False)
            image = "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode("ascii")
        result.append({"id": f"i{index}", "type": "image", "rect": list(item["bbox"]), "image": image})
    return result


def analyze_pdf(data: bytes) -> dict:
    with open_pdf(data) as doc:
        pages = []
        for page in doc:
            scale = min(1.5, 1200 / max(page.rect.width, page.rect.height))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            pages.append({"width": page.rect.width, "height": page.rect.height,
                          "image": "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode("ascii"),
                          "objects": objects(page)})
        return {"pages": pages, "pageCount": len(pages)}


def rgb(value):
    if not isinstance(value, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", value):
        raise ValueError("색상 값이 올바르지 않습니다.")
    return tuple(int(value[i:i + 2], 16) / 255 for i in (1, 3, 5))


def target_rect(value, page):
    if not isinstance(value, list) or len(value) != 4:
        raise ValueError("편집 영역이 올바르지 않습니다.")
    if any(not isinstance(v, (float, int)) or not math.isfinite(v) for v in value):
        raise ValueError("편집 좌표가 올바르지 않습니다.")
    rect = fitz.Rect(value)
    if rect.width < 1 or rect.height < 1 or not page.rect.contains(rect):
        raise ValueError("항목이 페이지 밖에 있습니다. 위치와 크기를 확인해 주세요.")
    return rect


def edit_pdf(data: bytes, operations: list) -> bytes:
    if not isinstance(operations, list) or len(operations) > MAX_OPERATIONS:
        raise ValueError("편집 항목은 500개 이하로 저장할 수 있습니다.")
    with open_pdf(data) as doc:
        grouped = {}
        for op in operations:
            if not isinstance(op, dict) or type(op.get("page")) is not int or not 0 <= op["page"] < doc.page_count:
                raise ValueError("편집할 페이지가 올바르지 않습니다.")
            grouped.setdefault(op["page"], []).append(op)
        fontpath = next((p for p in [Path("C:/Windows/Fonts/malgun.ttf"), Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf")] if p.is_file()), None)
        for number, edits in grouped.items():
            page = doc[number]
            sources = {item["id"]: item for item in objects(page)}
            seen = set()
            redactions = {"text": [], "image": []}
            for op in edits:
                if op.get("type") not in ("text", "image"):
                    raise ValueError("지원되지 않는 편집 항목입니다.")
                source_id = op.get("sourceId")
                if source_id:
                    source = sources.get(source_id)
                    if not source or source["type"] != op["type"] or source_id in seen:
                        raise ValueError("원본 항목을 찾을 수 없습니다. 파일을 다시 열어 주세요.")
                    seen.add(source_id)
                    rect = fitz.Rect(source["rect"]) & page.rect
                    # Narrow text redactions avoid erasing neighbouring lines; image redactions use their full area.
                    if source["type"] == "text":
                        rect = fitz.Rect(rect.x0, rect.y0 + rect.height * .2, rect.x1, rect.y1 - rect.height * .2)
                    redactions[source["type"]].append((rect, rgb(op.get("fill", "#ffffff"))))
            for kind, regions in redactions.items():
                if not regions:
                    continue
                for rect, fill in regions:
                    page.add_redact_annot(rect, fill=fill)
                page.apply_redactions(images=2 if kind == "image" else 0, graphics=0, text=1 if kind == "image" else 0)
            for op in edits:
                if op.get("deleted"):
                    continue
                rect = target_rect(op.get("rect"), page)
                if op["type"] == "image":
                    image = op.get("image", "")
                    if not isinstance(image, str) or not image.startswith("data:image/") or ";base64," not in image:
                        raise ValueError("교체할 이미지를 선택해 주세요.")
                    try:
                        content = base64.b64decode(image.split(",", 1)[1], validate=True)
                        if len(content) > 12 * 1024 * 1024:
                            raise ValueError()
                        page.insert_image(rect, stream=content, keep_proportion=False)
                    except Exception as exc:
                        raise ValueError("12MB 이하의 PNG/JPG/WebP 이미지를 선택해 주세요.") from exc
                else:
                    text = op.get("text", "")
                    size = op.get("size", 12)
                    if not isinstance(text, str) or len(text) > 10000 or not isinstance(size, (int, float)) or not 6 <= size <= 120:
                        raise ValueError("글자 내용 또는 크기가 올바르지 않습니다.")
                    fontname = "korean" if fontpath else "korea"
                    if fontpath:
                        page.insert_font(fontname=fontname, fontfile=str(fontpath))
                    spare = page.insert_textbox(rect, text, fontname=fontname, fontsize=size, color=rgb(op.get("color", "#17202a")), lineheight=1.15)
                    if spare < 0:
                        raise ValueError("글자가 영역보다 큽니다. 너비·높이를 늘리거나 글자 크기를 줄여 주세요.")
        return doc.tobytes(garbage=4, deflate=True)


def set_run_font(run, size=11, color=0, bold=False):
    run.font.name = "맑은 고딕"
    run.font.size = Pt(min(72, max(6, size)))
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(f"{color & 0xffffff:06X}")
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "맑은 고딕")


def convert_word(data: bytes, mode="editable") -> tuple[bytes, dict]:
    if mode not in ("editable", "appearance"):
        raise ValueError("Word 변환 방식이 올바르지 않습니다.")
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "맑은 고딕"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(3)
    report = {"pages": 0, "paragraphs": 0, "tables": 0, "images": 0, "scannedPages": [], "mode": mode}
    with open_pdf(data) as doc:
        for number, page in enumerate(doc):
            section = document.sections[0] if number == 0 else document.add_section(2)
            section.page_width, section.page_height = Pt(page.rect.width), Pt(page.rect.height)
            section.left_margin = section.right_margin = Pt(24)
            section.top_margin = section.bottom_margin = Pt(24)
            width = max(24, page.rect.width - 48)
            if mode == "appearance" or not page.get_text().strip():
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                # Both width and height are constrained so Word does not spill the picture onto a blank extra page.
                scale = min(width / page.rect.width, (page.rect.height - 66) / page.rect.height)
                document.add_picture(io.BytesIO(pix.tobytes("png")), width=Pt(page.rect.width * scale), height=Pt(page.rect.height * scale))
                report["images"] += 1
                if mode == "editable":
                    report["scannedPages"].append(number + 1)
                continue
            tables = page.find_tables().tables
            events = [(table.bbox[1], table.bbox[0], "table", table) for table in tables]
            for block in page.get_text("dict", sort=True)["blocks"]:
                box = fitz.Rect(block["bbox"])
                if block["type"] == 0:
                    for line in block["lines"]:
                        linebox = fitz.Rect(line["bbox"])
                        if any(fitz.Rect(t.bbox).contains(linebox.tl + (linebox.br - linebox.tl) * .5) for t in tables):
                            continue
                        events.append((linebox.y0, linebox.x0, "text", line))
                elif block["type"] == 1:
                    events.append((box.y0, box.x0, "image", block))
            for _, _, kind, content in sorted(events, key=lambda e: (round(e[0] / 4), e[1])):
                if kind == "text":
                    paragraph = document.add_paragraph()
                    for span in content["spans"]:
                        run = paragraph.add_run(span["text"])
                        set_run_font(run, span["size"], span["color"], bool(span["flags"] & 16))
                    report["paragraphs"] += 1
                elif kind == "image":
                    document.add_picture(io.BytesIO(content["image"]), width=Pt(min(width, content["bbox"][2] - content["bbox"][0])))
                    report["images"] += 1
                else:
                    values = content.extract()
                    table = document.add_table(rows=content.row_count, cols=content.col_count)
                    table.style = "Table Grid"
                    x_edges = sorted({round(r[i], 2) for r in content.cells if r for i in (0, 2)})
                    y_edges = sorted({round(r[i], 2) for r in content.cells if r for i in (1, 3)})
                    merged = set()
                    for row in range(content.row_count):
                        for col in range(content.col_count):
                            cellrect = content.rows[row].cells[col]
                            if cellrect is None or (row, col) in merged:
                                continue
                            rect = fitz.Rect(cellrect)
                            endrow, endcol = row, col
                            if len(x_edges) == content.col_count + 1 and len(y_edges) == content.row_count + 1:
                                endcol = min(content.col_count - 1, x_edges.index(round(rect.x1, 2)) - 1)
                                endrow = min(content.row_count - 1, y_edges.index(round(rect.y1, 2)) - 1)
                            cell = table.cell(row, col)
                            if endcol > col or endrow > row:
                                cell = cell.merge(table.cell(endrow, endcol))
                                merged.update((rr, cc) for rr in range(row, endrow + 1) for cc in range(col, endcol + 1))
                            cell.text = values[row][col] or ""
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    set_run_font(run)
                    report["tables"] += 1
        report["pages"] = doc.page_count
    output = io.BytesIO()
    document.save(output)
    return output.getvalue(), report
