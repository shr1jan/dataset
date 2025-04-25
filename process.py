import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Classify lines based on structure
def classify_line(line):
    line = line.strip()

    if re.match(r'^Chapter\s*[-–]?\s*\d+', line, re.I):
        return "chapter"
    elif re.match(r'^\d+\.\s+', line):
        return "section"
    elif re.match(r'^\(\d+\)', line):
        return "subsection"
    else:
        return "text"

# Apply formatting to DOCX
def format_docx(doc, text, tag):
    p = doc.add_paragraph()
    run = p.add_run()
    font = run.font
    font.name = "Times New Roman"

    if tag == "chapter":
        run.text = text.strip()
        font.bold = True
        font.size = Pt(14)
        p.paragraph_format.space_before = Pt(12)

    elif tag == "section":
        match = re.match(r'^(\d+)\.\s+(.*)', text)
        if match:
            sec_num, sec_text = match.groups()
            # First line: section label
            p1 = doc.add_paragraph()
            run1 = p1.add_run(f"section {sec_num}:")
            run1.font.name = "Times New Roman"
            run1.font.bold = True
            run1.font.size = Pt(12)
            p1.paragraph_format.space_before = Pt(8)

            # Second line: section title (as normal text)
            p2 = doc.add_paragraph()
            run2 = p2.add_run(sec_text.strip())
            run2.font.name = "Times New Roman"
            run2.font.size = Pt(11)
            p2.paragraph_format.left_indent = Inches(0.3)
            
            # Return early since we've already added the paragraphs
            return
        else:
            run.text = text
            font.bold = True
            font.size = Pt(12)
            p.paragraph_format.space_before = Pt(8)

    elif tag == "subsection":
        match = re.match(r'^\((\d+)\)\s*(.*)', text)
        if match:
            sub_num, sub_text = match.groups()
            run.text = f"subsection ({sub_num}): {sub_text.strip()}"
        else:
            run.text = text
        font.size = Pt(11)
        p.paragraph_format.left_indent = Inches(0.3)

    else:  # plain text
        run.text = text
        font.size = Pt(11)
        p.paragraph_format.left_indent = Inches(0.3)

# Convert PDF to DOCX with formatting
def convert_pdf_to_docx(pdf_path):
    doc = Document()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                lines = page.extract_text().split("\n")
                for line in lines:
                    if not line.strip():
                        continue
                    tag = classify_line(line)
                    format_docx(doc, line.strip(), tag)

        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path
    except Exception as e:
        print("❌ Error:", e)
        return None

# GUI handlers
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

# Launch the app
if __name__ == "__main__":
    launch_gui()
