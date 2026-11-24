#!/usr/bin/env python3
"""
Self-Executing GUI Marker for:
Word Task 1 – Home & Insert Tabs on Ribbon (Grade 8 – 2027)
"""

import csv
import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Dependency",
        "The 'python-docx' package is required.\n\nInstall it by running:\npip install python-docx",
    )
    raise SystemExit(1)


# =====================================================================
# MARKING ENGINE
# =====================================================================

def has_soft_breaks(paragraph) -> bool:
    return bool(
        paragraph._p.findall(
            ".//w:br",
            namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        )
    )


def has_bullet_numbering(paragraph) -> bool:
    num_pr = paragraph._p.find(
        ".//w:numPr",
        namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )
    style_name = (paragraph.style.name or "").lower()
    return num_pr is not None or "list bullet" in style_name


def has_shape_inserted(doc: Document) -> bool:
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
    for section in doc.sections:
        for part in (section.footer, section.header):
            text = " ".join(p.text.strip() for p in part.paragraphs if p.text.strip())
            xml_str = part._element.xml
            if text or "PAGE" in xml_str or "pgNum" in xml_str:
                return True
    return False


def grade_document(file_path: Path, award_server_bonus: bool = True) -> dict:
    scores = {
        "file_setup": 0,      # / 2
        "text_entry": 0,      # / 4
        "home_formatting": 0, # / 10
        "insert_elements": 0, # / 5
        "server_bonus": 4 if award_server_bonus else 0,  # / 4
    }
    feedback = []

    if not file_path.exists() or file_path.suffix.lower() != ".docx":
        return {
            "filename": file_path.name,
            "learner": file_path.stem,
            "auto_total": 0,
            "final_total": 0,
            "scores": scores,
            "feedback": ["File missing or not a valid .docx document (0/2)."],
        }

    pattern = re.compile(r"^WordTask1_([A-Za-z\-]+)_([A-Za-z\-]+)$", re.IGNORECASE)
    match = pattern.match(file_path.stem)
    if match:
        scores["file_setup"] += 2
        learner_name = f"{match.group(1)}, {match.group(2)}"
        feedback.append(f"Filename '{file_path.name}' matches required pattern (+2/2).")
    else:
        scores["file_setup"] += 1
        learner_name = file_path.stem
        feedback.append(
            f"Valid .docx file, but filename '{file_path.name}' does not match WordTask1_Surname_Name (+1/2)."
        )

    try:
        doc = Document(file_path)
    except Exception as exc:
        return {
            "filename": file_path.name,
            "learner": learner_name,
            "auto_total": scores["file_setup"],
            "final_total": scores["file_setup"] + scores["server_bonus"],
            "scores": scores,
            "feedback": feedback + [f"Could not read Word XML structure: {exc}"],
        }

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
    text_pts = round((content_score / 4) * 2)
    scores["text_entry"] += text_pts
    feedback.append(f"Text content check: {content_score}/4 key sections found (+{text_pts}/2).")

    soft_break_count = sum(1 for p in non_empty_paras if has_soft_breaks(p))
    if len(non_empty_paras) >= 3 and soft_break_count >= 1:
        scores["text_entry"] += 2
        feedback.append("Paragraph structure and soft line breaks detected (+2/2).")
    elif len(non_empty_paras) >= 2:
        scores["text_entry"] += 1
        feedback.append("Paragraph breaks found, but soft line breaks (<Shift+Enter>) not clearly detected (+1/2).")
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
    feedback.append(f"Title formatting (Style / Centre / 18pt): +{t_pts}/3.")

    # 3b. Body font/size + bold 'digital choices' + italic 'lasting footprint' (3 marks)
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
        f"Body formatting (Font={font_ok}, Bold 'digital choices'={bold_digital}, Italic 'lasting footprint'={italic_footprint}): +{b_pts}/3."
    )

    # 3c. Subheading emphasis (2 marks)
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
        p
        for p in non_empty_paras
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
    if has_shape_inserted(doc):
        scores["insert_elements"] += 3
        feedback.append("Inserted shape/drawing element detected: +3/3.")
    else:
        feedback.append("No inserted shape/drawing detected: +0/3.")

    if has_page_number_or_footer(doc):
        scores["insert_elements"] += 2
        feedback.append("Page number or footer content detected: +2/2.")
    else:
        feedback.append("No page number or footer detected: +0/2.")

    if award_server_bonus:
        feedback.append("Assessment server upload verified: +4/4.")
    else:
        feedback.append("Assessment server upload bonus not included: +0/4.")

    auto_total = (
        scores["file_setup"]
        + scores["text_entry"]
        + scores["home_formatting"]
        + scores["insert_elements"]
    )
    final_total = auto_total + scores["server_bonus"]

    return {
        "filename": file_path.name,
        "learner": learner_name,
        "auto_total": auto_total,
        "final_total": final_total,
        "scores": scores,
        "feedback": feedback,
    }


# =====================================================================
# DESKTOP GUI APPLICATION
# =====================================================================

class WordTaskMarkerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word Task 1 – Automated Class Marker (Grade 8 – 2027)")
        self.geometry("1040x660")
        self.minsize(900, 560)

        self.selected_folder: Path | None = None
        self.results: list[dict] = []
        self.include_server_bonus = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        # Header banner
        header = tk.Frame(self, bg="#1F4E79", padx=16, pady=12)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Word Task 1 – Home & Insert Tabs on Ribbon (Automated Marker)",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#1F4E79",
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Select a folder containing learner .docx submissions to grade all files and export marks.",
            font=("Segoe UI", 9),
            fg="#D9E2EC",
            bg="#1F4E79",
        ).pack(anchor="w", pady=(2, 0))

        # Controls bar
        controls = tk.Frame(self, padx=14, pady=10)
        controls.pack(fill="x")

        tk.Button(
            controls,
            text="1. Select Submissions Folder...",
            font=("Segoe UI", 10, "bold"),
            bg="#2B6CB0",
            fg="white",
            padx=12,
            pady=5,
            relief="flat",
            command=self.choose_folder,
        ).pack(side="left")

        tk.Button(
            controls,
            text="Re-Mark Folder",
            font=("Segoe UI", 10),
            padx=10,
            pady=5,
            command=self.run_marking,
        ).pack(side="left", padx=(8, 0))

        tk.Checkbutton(
            controls,
            text="Award +4 Server Submission Marks (Total /25)",
            variable=self.include_server_bonus,
            font=("Segoe UI", 10),
            command=self.run_marking,
        ).pack(side="left", padx=(16, 0))

        self.export_btn = tk.Button(
            controls,
            text="Export Results to CSV",
            font=("Segoe UI", 10, "bold"),
            bg="#2F855A",
            fg="white",
            padx=12,
            pady=5,
            relief="flat",
            state="disabled",
            command=self.export_csv,
        )
        self.export_btn.pack(side="right")

        # Folder path label
        self.folder_label = tk.Label(
            self,
            text="Selected Folder: None",
            font=("Segoe UI", 9, "italic"),
            fg="#4A5568",
            padx=14,
            anchor="w",
        )
        self.folder_label.pack(fill="x")

        # Main split pane (Table on top, Detailed Feedback on bottom)
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=14, pady=(6, 10))

        # Treeview frame
        tree_frame = ttk.Frame(paned)
        columns = ("learner", "filename", "setup", "text", "home", "insert", "server", "total", "pct")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=13)

        headings = {
            "learner": ("Learner", 160),
            "filename": ("File Name", 220),
            "setup": ("Setup (/2)", 75),
            "text": ("Text (/4)", 75),
            "home": ("Home (/10)", 85),
            "insert": ("Insert (/5)", 80),
            "server": ("Server (/4)", 80),
            "total": ("Total", 80),
            "pct": ("%", 65),
        }
        for col, (label, width) in headings.items():
            self.tree.heading(col, text=label)
            anchor = "w" if col in ("learner", "filename") else "center"
            self.tree.column(col, width=width, anchor=anchor)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        paned.add(tree_frame, weight=3)

        # Feedback frame
        fb_frame = ttk.LabelFrame(paned, text="Detailed Marking Breakdown (Click a learner row above)")
        self.fb_text = tk.Text(
            fb_frame,
            height=9,
            wrap="word",
            font=("Consolas", 10),
            bg="#F7FAFC",
            fg="#1A202C",
            padx=10,
            pady=8,
            state="disabled",
        )
        self.fb_text.pack(fill="both", expand=True)
        paned.add(fb_frame, weight=2)

        # Status bar
        self.status_bar = tk.Label(
            self,
            text="Ready. Click 'Select Submissions Folder...' to begin.",
            bd=1,
            relief="sunken",
            anchor="w",
            padx=10,
            pady=4,
            font=("Segoe UI", 9),
        )
        self.status_bar.pack(fill="x", side="bottom")

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing Learner .docx Files")
        if not folder:
            return
        self.selected_folder = Path(folder)
        self.folder_label.config(text=f"Selected Folder: {self.selected_folder}")
        self.run_marking()

    def run_marking(self):
        if not self.selected_folder or not self.selected_folder.exists():
            return

        # Find all .docx files (ignoring temporary ~$ Word lock files)
        docx_files = sorted(
            [
                f
                for f in self.selected_folder.glob("*.docx")
                if not f.name.startswith("~$")
            ],
            key=lambda p: p.name.lower(),
        )

        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self._set_feedback_text("")

        if not docx_files:
            self.export_btn.config(state="disabled")
            self.status_bar.config(text="No .docx files found in the selected folder.")
            messagebox.showinfo("No Files Found", "No .docx files were found in that folder.")
            return

        award_bonus = self.include_server_bonus.get()
        max_possible = 25 if award_bonus else 21

        for idx, file_path in enumerate(docx_files):
            res = grade_document(file_path, award_server_bonus=award_bonus)
            self.results.append(res)
            sc = res["scores"]
            pct = round((res["final_total"] / max_possible) * 100)
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    res["learner"],
                    res["filename"],
                    sc["file_setup"],
                    sc["text_entry"],
                    sc["home_formatting"],
                    sc["insert_elements"],
                    sc["server_bonus"],
                    f"{res['final_total']} / {max_possible}",
                    f"{pct}%",
                ),
            )

        avg_score = sum(r["final_total"] for r in self.results) / len(self.results)
        self.export_btn.config(state="normal")
        self.status_bar.config(
            text=f"Marked {len(self.results)} file(s) | Class Average: {avg_score:.1f} / {max_possible}"
        )

        # Select first item automatically
        first_id = self.tree.get_children()[0]
        self.tree.selection_set(first_id)
        self.on_row_select(None)

    def on_row_select(self, _event):
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        res = self.results[idx]
        max_possible = 25 if self.include_server_bonus.get() else 21

        lines = [
            f"Learner : {res['learner']}  ({res['filename']})",
            f"Score   : {res['final_total']} / {max_possible}",
            "-" * 60,
        ] + [f"• {item}" for item in res["feedback"]]
        self._set_feedback_text("\n".join(lines))

    def _set_feedback_text(self, text: str):
        self.fb_text.config(state="normal")
        self.fb_text.delete("1.0", "end")
        self.fb_text.insert("1.0", text)
        self.fb_text.config(state="disabled")

    def export_csv(self):
        if not self.results:
            return

        default_name = f"WordTask1_Marks_{self.selected_folder.name}.csv"
        save_path = filedialog.asksaveasfilename(
            title="Save Class Mark Sheet",
            initialdir=str(self.selected_folder),
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not save_path:
            return

        max_possible = 25 if self.include_server_bonus.get() else 21
        with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "Learner",
                    "File Name",
                    "Setup (/2)",
                    "Text & Breaks (/4)",
                    "Home Tab (/10)",
                    "Insert Tab (/5)",
                    "Server Bonus (/4)",
                    f"Total (/{max_possible})",
                    "Percentage",
                    "Feedback Summary",
                ]
            )
            for r in self.results:
                sc = r["scores"]
                pct = f"{round((r['final_total'] / max_possible) * 100)}%"
                writer.writerow(
                    [
                        r["learner"],
                        r["filename"],
                        sc["file_setup"],
                        sc["text_entry"],
                        sc["home_formatting"],
                        sc["insert_elements"],
                        sc["server_bonus"],
                        r["final_total"],
                        pct,
                        " | ".join(r["feedback"]),
                    ]
                )

        messagebox.showinfo("Export Complete", f"Class marks saved to:\n{save_path}")


if __name__ == "__main__":
    app = WordTaskMarkerApp()
    app.mainloop()