import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

def classify_line(line):
    line = line.strip()
    if re.match(r'^NEPAL.*ACT.*\d{4}', line, re.I):
        return "Title"
    elif re.match(r'^AN ACT MADE TO.*', line, re.I):
        return "Subtitle"
    elif re.match(r'^Date of Authentication.*', line, re.I):
        return "Title"
    elif re.match(r'^Preamble\s*:?', line, re.I):
        return "Heading 1"
    elif re.match(r'^Chapter\s*[-–]?\s*\d+', line, re.I):
        return "Heading 2"
    elif re.match(r'^\d+\.\s+', line):
        return "Heading 3"
    elif re.match(r'^\(\d+\)', line):  # Only numeric parentheses
        return "Heading 4"
    else:
        return "Normal"

def add_styled_paragraph(doc, text, style_tag):
    p = doc.add_paragraph(text)
    p.style = style_tag
    if style_tag == "Heading 4":
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal":
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag in ["Title", "Subtitle"]:
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def convert_pdf_to_docx(pdf_path):
    doc = Document()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                lines = text.split("\n")
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue

                    tag = classify_line(line)

                    if tag == "Heading 3":
                        sec_match = re.match(r'^(\d+)\.\s+(.*)', line)
                        if sec_match:
                            sec_num, sec_body = sec_match.groups()
                            add_styled_paragraph(doc, f"section {sec_num}:", "Heading 3")

                            subsection_parts = re.split(r'(?=\(\d+\))', sec_body)
                            if subsection_parts:
                                add_styled_paragraph(doc, subsection_parts[0].strip(), "Normal")
                            for part in subsection_parts[1:]:
                                sub_match = re.match(r'^\((\d+)\)\s*(.*)', part)
                                if sub_match:
                                    sub_num, sub_text = sub_match.groups()
                                    add_styled_paragraph(doc, f"subsection ({sub_num}): {sub_text.strip()}", "Heading 4")
                        continue

                    elif tag == "Heading 4":
                        sub_match = re.match(r'^\((\d+)\)\s*(.*)', line)
                        if sub_match:
                            sub_num, sub_text = sub_match.groups()
                            add_styled_paragraph(doc, f"subsection ({sub_num}): {sub_text.strip()}", "Heading 4")
                        else:
                            add_styled_paragraph(doc, line, "Heading 4")
                        continue

                    elif tag == "Heading 2":
                        chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line, re.I)
                        if chap_match:
                            chap_num, chap_title = chap_match.groups()
                            full_title = f"chapter {chap_num.strip()}: {chap_title.strip()}"
                            add_styled_paragraph(doc, full_title, "Heading 2")
                        else:
                            add_styled_paragraph(doc, line, "Heading 2")
                        continue

                    add_styled_paragraph(doc, line, tag)

        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path

    except Exception as e:
        print("❌ Error processing:", pdf_path, "\n", e)
        return None

def select_files_and_convert(status_label):
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if not file_paths:
        return

    if len(file_paths) > 10:
        messagebox.showwarning("Limit Exceeded", "You can only upload up to 10 PDFs at a time.")
        return

    status_label.config(text="Converting...", fg="blue")
    successes, failures = [], []

    for path in file_paths:
        result = convert_pdf_to_docx(path)
        if result:
            successes.append(os.path.basename(result))
        else:
            failures.append(os.path.basename(path))

    status_label.config(text="Done ✔", fg="green")

    summary = f"{len(successes)} file(s) converted successfully:\n" + "\n".join(successes)
    if failures:
        summary += f"\n\n{len(failures)} failed:\n" + "\n".join(failures)

    messagebox.showinfo("Batch Conversion Complete", summary)

def launch_gui():
    root = tk.Tk()
    root.title("PDF → Structured DOCX Converter")
    root.geometry("540x260")
    root.resizable(False, False)
    root.configure(bg="#f5f5f5")

    title = tk.Label(root, text="Batch PDF to Structured Word Converter", font=("Helvetica", 14, "bold"), bg="#f5f5f5")
    title.pack(pady=(20, 10))

    instruction = tk.Label(root, text="Select up to 10 PDF files to convert to structured .docx format", font=("Helvetica", 11), bg="#f5f5f5")
    instruction.pack(pady=(0, 20))

    status_label = tk.Label(root, text="", font=("Helvetica", 10), bg="#f5f5f5")
    status_label.pack()

    btn = ttk.Button(root, text="Select PDF Files", command=lambda: select_files_and_convert(status_label))
    btn.pack(pady=10)

    note = tk.Label(root, text="Output files will be saved next to the original PDFs", font=("Helvetica", 9), fg="gray", bg="#f5f5f5")
    note.pack(side="bottom", pady=10)

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
