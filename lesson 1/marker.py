import re
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


def has_soft_breaks(paragraph) -> bool:
    """Return True if the paragraph contains soft line breaks (<w:br/>)."""
    return bool(paragraph._p.findall(".//w:br", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}))


def has_bullet_numbering(paragraph) -> bool:
    """Return True if the paragraph has numbering/bullet XML formatting."""
    num_pr = paragraph._p.find(
        ".//w:numPr",
        namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )
    style_name = (paragraph.style.name or "").lower()
    return num_pr is not None or "list bullet" in style_name


def has_shape_inserted(doc: Document) -> bool:
    """Check body, headers, and footers for shapes/drawings/picts."""
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "v": "urn:schemas-microsoft-com:vml",
    }
    containers = [doc.element.body]
    for section in doc.sections:
        containers.append(section.header._element)
        containers.append(section.footer._element)

    for container in containers:
        if (
            container.findall(".//w:drawing", namespaces=ns)
            or container.findall(".//w:pict", namespaces=ns)
            or container.findall(".//v:shape", namespaces=ns)
            or container.findall(".//v:rect", namespaces=ns)
            or container.findall(".//v:oval", namespaces=ns)
            or container.findall(".//v:roundrect", namespaces=ns)
        ):
            return True
    return False


def has_page_number_or_footer(doc: Document) -> bool:
    """Check footers/headers for page number fields or non-empty footer text."""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for section in doc.sections:
        for part in (section.footer, section.header):
            text = " ".join(p.text.strip() for p in part.paragraphs if p.text.strip())
            xml_str = part._element.xml
            if text or "PAGE" in xml_str or "pgNum" in xml_str:
                return True
    return False


def grade_document(file_path: Path) -> dict:
    scores = {
        "file_setup": 0,      # / 2
        "text_entry": 0,      # / 4
        "home_formatting": 0, # / 10
        "insert_elements": 0, # / 5
        "server_bonus": 0,    # / 4 (manual/server verification placeholder)
    }
    feedback = []

    # 1. File Setup & Naming (2 marks)
    if not file_path.exists() or file_path.suffix.lower() != ".docx":
        return {
            "total": 0,
            "max_auto": 21,
            "max_total": 25,
            "scores": scores,
            "feedback": ["File missing or not a valid .docx document (0/2)."],
        }

    pattern = re.compile(r"^WordTask1_[A-Za-z]+_[A-Za-z]+$", re.IGNORECASE)
    if pattern.match(file_path.stem):
        scores["file_setup"] += 2
        feedback.append(f"Filename '{file_path.name}' matches required pattern (+2/2).")
    else:
        scores["file_setup"] += 1
        feedback.append(f"Valid .docx file, but filename '{file_path.name}' does not match WordTask1_Surname_Name (+1/2).")

    doc = Document(file_path)
    non_empty_paras = [p for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(p.text.strip() for p in non_empty_paras)

    # 2. Text Entry & Breaks (4 marks)
    title_ok = "smart habits for teens online in 2027" in full_text.lower()
    body_ok = "digital choices" in full_text.lower() and "lasting footprint" in full_text.lower()
    sub_ok = "three quick rules" in full_text.lower()
    rules_ok = all(
        phrase in full_text.lower()
        for phrase in ("protect your passwords", "pause before posting", "speak to a trusted adult")
    )

    content_score = sum([title_ok, body_ok, sub_ok, rules_ok])
    # Award up to 2 marks for accurate text presence
    text_pts = round((content_score / 4) * 2)
    scores["text_entry"] += text_pts
    feedback.append(f"Text content check: {content_score}/4 key sections found (+{text_pts}/2).")

    # Award 2 marks for hard/soft return structure (at least 3 paragraphs + soft breaks)
    soft_break_count = sum(1 for p in non_empty_paras if has_soft_breaks(p))
    if len(non_empty_paras) >= 3 and soft_break_count >= 1:
        scores["text_entry"] += 2
        feedback.append(f"Paragraph structure and soft line breaks detected (+2/2).")
    elif len(non_empty_paras) >= 2:
        scores["text_entry"] += 1
        feedback.append("Some paragraph breaks found, but soft line breaks (<Shift+Enter>) were not clearly detected (+1/2).")
    else:
        feedback.append("Missing clear paragraph/line break structure (+0/2).")

    # 3. Home Tab Formatting (10 marks)
    # 3a. Title style, 18 pt, centre aligned (3 marks)
    title_para = non_empty_paras[0] if non_empty_paras else None
    t_pts = 0
    if title_para:
        style_lower = (title_para.style.name or "").lower()
        if "title" in style_lower or "heading 1" in style_lower:
            t_pts += 1
        if title_para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            t_pts += 1
        run_sizes = [r.font.size.pt for r in title_para.runs if r.font.size is not None]
        if any(abs(sz - 18.0) < 0.5 for sz in run_sizes):
            t_pts += 1
    scores["home_formatting"] += t_pts
    feedback.append(f"Title formatting (style/centre/18pt): +{t_pts}/3.")

    # 3b. Body font/size + bold 'digital choices' + italic 'lasting footprint' (3 marks)
    b_pts = 0
    bold_digital = False
    italic_footprint = False
    font_ok = False

    for p in non_empty_paras[1:]:
        for r in p.runs:
            txt_low = r.text.lower()
            if r.font.name and r.font.name.lower() in {"calibri", "arial"}:
                font_ok = True
            if "digital choices" in txt_low and r.bold:
                bold_digital = True
            if "lasting footprint" in txt_low and r.italic:
                italic_footprint = True

    b_pts = sum([font_ok, bold_digital, italic_footprint])
    scores["home_formatting"] += b_pts
    feedback.append(
        f"Body formatting (font={font_ok}, bold 'digital choices'={bold_digital}, italic 'lasting footprint'={italic_footprint}): +{b_pts}/3."
    )

    # 3c. Subheading emphasis: 'Three Quick Rules:' bold + colour or highlight (2 marks)
    s_pts = 0
    for p in non_empty_paras:
        if "three quick rules" in p.text.lower():
            is_bold = any(r.bold for r in p.runs if r.text.strip()) or bool(p.style.font.bold)
            has_color_or_hl = any(
                (r.font.color and r.font.color.rgb is not None) or (r.font.highlight_color is not None)
                for r in p.runs
            )
            if is_bold:
                s_pts += 1
            if has_color_or_hl:
                s_pts += 1
            break
    scores["home_formatting"] += s_pts
    feedback.append(f"Subheading 'Three Quick Rules:' emphasis: +{s_pts}/2.")

    # 3d. Bulleted list on rules (2 marks)
    bullet_paras = [
        p for p in non_empty_paras
        if has_bullet_numbering(p)
        and any(k in p.text.lower() for k in ("password", "posting", "trusted adult"))
    ]
    if len(bullet_paras) >= 2:
        scores["home_formatting"] += 2
        feedback.append(f"Bulleted list applied to rules ({len(bullet_paras)} bulleted items): +2/2.")
    elif len(bullet_paras) == 1:
        scores["home_formatting"] += 1
        feedback.append("Partial bullet formatting detected on rules: +1/2.")
    else:
        feedback.append("Bulleted list not detected on rules: +0/2.")

    # 4. Insert Tab Elements (5 marks)
    shape_ok = has_shape_inserted(doc)
    footer_pg_ok = has_page_number_or_footer(doc)

    if shape_ok:
        scores["insert_elements"] += 3
        feedback.append("Inserted shape/drawing element detected: +3/3.")
    else:
        feedback.append("No inserted shape/drawing detected: +0/3.")

    if footer_pg_ok:
        scores["insert_elements"] += 2
        feedback.append("Page number or footer content detected: +2/2.")
    else:
        feedback.append("No page number or footer detected: +0/2.")

    auto_total = scores["file_setup"] + scores["text_entry"] + scores["home_formatting"] + scores["insert_elements"]
    return {
        "total": auto_total,
        "max_auto": 21,
        "max_total": 25,
        "scores": scores,
        "feedback": feedback,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python grade_word_task1.py WordTask1_Surname_Name.docx")
        sys.exit(1)

    target = Path(sys.argv[1])
    result = grade_document(target)

    print(f"\n=== Marking Report: {target.name} ===")
    for item in result["feedback"]:
        print(f"- {item}")
    print("\nSection Totals (Automated):")
    for k, v in result["scores"].items():
        if k != "server_bonus":
            print(f"  {k}: {v}")
    print(f"\nAutomated Score: {result['total']} / {result['max_auto']} (4 marks reserved for server submission check)")