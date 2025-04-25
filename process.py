import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def classify_line(line):
    line = line.strip()

    if re.match(r'^Nepal.*Act.*\d{4}', line, re.I):
        return "title"
    elif re.match(r'^Date of Authentication.*', line, re.I):
        return "title"
    elif re.match(r'^Preamble\s*:?', line, re.I):
        return "header1"
    elif re.match(r'^Chapter\s*[-–]?\s*\d+', line, re.I):
        return "header2"
    elif re.match(r'^\d+\.\s+', line):
        return "header3"
    elif re.match(r'^\(\d+\)', line):
        return "header4"
    elif re.match(r'^\d+\.\s+.*Amendment.*', line, re.I):
        return "subtitle"
    else:
        return "text"


def add_styled_paragraph(doc, text, style_tag):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Times New Roman"

    if style_tag == "title":
        run.bold = True
        run.font.size = Pt(16)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    elif style_tag == "header1":
        run.bold = True
        run.font.size = Pt(13)
        p.paragraph_format.space_before = Pt(10)
    elif style_tag == "header2":
        run.bold = True
        run.font.size = Pt(12)
        p.paragraph_format.space_before = Pt(10)
    elif style_tag == "header3":
        run.bold = True
        run.font.size = Pt(11)
        p.paragraph_format.space_before = Pt(6)
    elif style_tag == "header4":
        run.font.size = Pt(11)
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "subtitle":
        run.italic = True
        run.font.size = Pt(11)
    else:
        run.font.size = Pt(11)
        p.paragraph_format.left_indent = Inches(0.3)


def convert_pdf_to_docx(pdf_path):
    doc = Document()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                lines = page.extract_text().split("\n")
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue

                    tag = classify_line(line)

                    if tag == "header3":
                        section_match = re.match(r'^(\d+)\.\s+(.*)', line)
                        if section_match:
                            sec_num, sec_rest = section_match.groups()
                            add_styled_paragraph(doc, f"section {sec_num}:", "header3")
                            subsection_split = re.split(r'(?=\(\d+\))', sec_rest, maxsplit=1)
                            section_title = subsection_split[0].strip()
                            if section_title:
                                add_styled_paragraph(doc, section_title, "text")
                            if len(subsection_split) > 1:
                                rest_subs = subsection_split[1]
                                subs = re.findall(r'\((\d+)\)\s*([^()]*(?=(\(\d+\)|$)))', rest_subs)
                                for sub_num, sub_text, _ in subs:
                                    add_styled_paragraph(doc, f"subsection ({sub_num}): {sub_text.strip()}", "header4")
                        continue

                    elif tag == "header4":
                        sub_match = re.match(r'^\((\d+)\)\s*(.*)', line)
                        if sub_match:
                            sub_num, sub_text = sub_match.groups()
                            add_styled_paragraph(doc, f"subsection ({sub_num}): {sub_text.strip()}", "header4")
                        else:
                            add_styled_paragraph(doc, line, "header4")
                        continue

                    elif tag == "header2":
                        chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line, re.I)
                        if chap_match:
                            chap_num, chap_title = chap_match.groups()
                            add_styled_paragraph(doc, f"chapter {chap_num.strip()}: {chap_title.strip()}", "header2")
                        else:
                            add_styled_paragraph(doc, line, "header2")
                        continue

                    add_styled_paragraph(doc, line, tag)

        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path

    except Exception as e:
        print("❌ Error:", e)
        return None


def open_file_dialog():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        return
    result = convert_pdf_to_docx(pdf_path)
    if result:
        messagebox.showinfo("Success", f"Saved to:\n{result}")
    else:
        messagebox.showerror("Error", "Conversion failed.")


def launch_gui():
    root = tk.Tk()
    root.title("PDF → Structured DOCX Converter")
    root.geometry("420x200")
    root.resizable(False, False)

    label = tk.Label(root, text="Select a PDF to convert it into a structured Word doc", font=("Arial", 12))
    label.pack(pady=30)

    btn = tk.Button(root, text="Select PDF File", font=("Arial", 12), command=open_file_dialog)
    btn.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
