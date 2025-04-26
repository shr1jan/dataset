import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Frame, END, DISABLED, NORMAL
from tkinter import ttk
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
import logging
from datetime import datetime
import json
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message="CropBox missing from /Page, defaulting to MediaBox")

selected_pdf_paths = []
PRESERVE_AMENDMENTS = True
FORMAT_DATES = True

log_file = f"pdf_conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def classify_line(line):
    line = line.strip()
    # Check for Notes: pattern first
    if re.match(r'^Notes\s*:', line, re.I):
        return "Normal"  # Treat Notes: as normal text
    # Check for special section formats with symbols - EXPANDED PATTERN
    elif re.match(r'^[♦◉]\s*\d+[A-Za-z]?\.', line):
        return "Heading 3"
    elif re.match(r'^Schedule\b.*', line, re.I):
        return "Heading 5"
    elif re.match(r'^NEPAL.*ACT.*\d{4}', line, re.I):
        return "Title"
    elif re.match(r'^Date of (Authentication|Publication|Authentication and Publication|Royal Seal and Publication)\b.*', line, re.I):
        return "Subtitle"
    elif re.match(r'^AN ACT MADE TO.*', line, re.I):
        return "Subtitle"
    elif re.match(r'^Amendments\s*:?', line, re.I):
        return "Subtitle"
    elif re.match(r'^Preamble\s*:?', line, re.I):
        return "Heading 1"
    elif re.match(r'^Chapter\s*[-–]?\s*\d+', line, re.I):
        return "Heading 2"
    # New section formats
    elif re.match(r'^[♦◉]\s*\(\d+\)D?', line):
        return "Heading 3"
    elif re.match(r'^\d+[A-Za-z]?\.\s+', line):
        return "Heading 3"
    elif re.match(r'^\(\d+\)', line):
        return "Heading 4"
    elif re.match(r'^\d{4}\.\d{1,2}\.\d{1,2}.*', line):
         return "Normal"
    else:
        return "Normal"

def add_styled_paragraph(doc, text, style_tag, is_under_h5=False):
    p = doc.add_paragraph()
    try:
        p.style = doc.styles[style_tag]
    except KeyError:
        p.style = doc.styles['Normal']

    run = p.add_run(text)

    if style_tag == "Heading 5":
        pass
    elif style_tag == "Heading 4":
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal" and is_under_h5:
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal":
        p.paragraph_format.left_indent = Inches(0)
    elif style_tag == "Title":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run.font.size = Pt(16)
        run.font.bold = True
    elif style_tag == "Subtitle":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run.font.size = Pt(14)
        run.font.italic = True

    return p

def add_table_to_doc(doc, table_data):
    """Add a table to the document with the given data."""
    if not table_data or not table_data[0]:
        return None
    
    num_rows = len(table_data)
    num_cols = max(len(row) for row in table_data)
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    
    # Fill the table with data
    for i, row in enumerate(table_data):
        for j, cell_text in enumerate(row):
            if j < num_cols:  # Ensure we don't exceed the number of columns
                table.cell(i, j).text = cell_text
    
    return table

def is_likely_table_row(line):
    """Check if a line is likely part of a table based on patterns."""
    # Check for multiple cell separators like | or tab characters
    if '|' in line and line.count('|') >= 2:
        return True
    if '\t' in line and line.count('\t') >= 1:
        return True
    # Check for patterns like "Column1   Column2   Column3"
    if re.search(r'\S+\s{2,}\S+\s{2,}\S+', line):
        return True
    return False

def extract_table_from_lines(lines, start_idx):
    """Extract table data from consecutive lines that appear to be a table."""
    table_data = []
    i = start_idx
    
    while i < len(lines) and is_likely_table_row(lines[i]):
        row = lines[i]
        # Split by pipe if present
        if '|' in row:
            cells = [cell.strip() for cell in row.split('|')]
            # Remove empty cells at the beginning and end if they're just artifacts of the pipe splitting
            if cells and not cells[0]:
                cells = cells[1:]
            if cells and not cells[-1]:
                cells = cells[:-1]
        # Split by tabs if present
        elif '\t' in row:
            cells = [cell.strip() for cell in row.split('\t')]
        # Split by multiple spaces
        else:
            cells = re.split(r'\s{2,}', row.strip())
        
        table_data.append(cells)
        i += 1
    
    return table_data, i - start_idx

def convert_pdf_to_docx(pdf_path, output_dir=None):
    doc = Document()
    url_pattern = re.compile(r'(?:https?://|www\.)\S+')
    translation_pattern = re.compile(r'\s*\((Official|Unofficial)\s+Translation\)\s*', re.I)
    is_within_heading_5 = False
    last_p = None
    last_tag = None
    is_first_line_of_document = True
    is_within_amendments = False
    is_within_subsection = False  # Track if we're inside a subsection like (a), (b), etc.
    is_processing_table = False
    table_buffer = []
    all_tables = []  # Initialize all_tables
    
    try:
        print(f"Processing: {os.path.basename(pdf_path)}")
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages: {total_pages}")
            
            # Process each page for text
            for page_num, page in enumerate(pdf.pages, 1):
                if page_num % 5 == 0 or page_num == 1 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")
                text = page.extract_text()
                if not text:
                    continue
                lines = text.split("\n")
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    original_line = line
                    
                    # Remove translation markers
                    line = translation_pattern.sub('', line)
                    
                    if not line:
                        if is_within_heading_5 and not is_within_amendments:
                            current_p = add_styled_paragraph(doc, "", "Normal", is_under_h5=True)
                            last_p = current_p
                            last_tag = "Normal"
                        i += 1
                        continue
                    
                    if re.fullmatch(r'\d+', line):
                        i += 1
                        continue
                    
                    # Check if this line is likely the start of a table
                    if is_likely_table_row(line) and not is_processing_table:
                        is_processing_table = True
                        table_data, rows_consumed = extract_table_from_lines(lines, i)
                        if table_data and len(table_data) > 1:  # Ensure it's actually a table with multiple rows
                            add_table_to_doc(doc, table_data)
                            all_tables.append(table_data)  # Store table data for later reference
                            i += rows_consumed
                            is_processing_table = False
                            continue
                        else:
                            # Not a real table, process as normal text
                            is_processing_table = False
                    
                    # Process as normal text if not a table
                    line_after_url_removal = url_pattern.sub('', line).strip()
                    line_after_url_removal = translation_pattern.sub('', line_after_url_removal)
                    
                    if not line_after_url_removal:
                        i += 1
                        continue
                    
                    if is_first_line_of_document:
                        tag = "Title"
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                        last_p = current_p
                        last_tag = tag
                        is_first_line_of_document = False
                        is_within_amendments = False
                        is_within_heading_5 = False
                        i += 1
                        continue

                    tag = classify_line(line_after_url_removal)

                    # Check if this is a subsection marker like (a), (b), etc.
                    is_letter_subsection = re.match(r'^\([a-z]\)', line_after_url_removal)
                    if is_letter_subsection:
                        is_within_subsection = True
                    
                    # Check if this is a numbered item inside a subsection
                    is_numbered_item_in_subsection = is_within_subsection and re.match(r'^\(\d+\)', line_after_url_removal)
                    
                    # If it's a numbered item inside a subsection, treat it as normal text
                    if is_numbered_item_in_subsection:
                        tag = "Normal"

                    if is_within_amendments:
                        if tag in ["Title", "Heading 1", "Heading 2"] or \
                           (tag == "Subtitle" and not re.match(r'^Amendments\s*:?', line_after_url_removal, re.I)):
                            is_within_amendments = False
                        else:
                            current_p = add_styled_paragraph(doc, line_after_url_removal, "Subtitle")
                            last_p = current_p
                            last_tag = "Subtitle"
                            i += 1
                            continue

                    is_date_line = re.match(r'^\d{4}\.\d{1,2}\.\d{1,2}', original_line)
                    is_prev_date_subtitle = last_p is not None and last_tag == "Subtitle" and \
                                            re.match(r'^Date of (Authentication|Publication|Authentication and Publication|Royal Seal and Publication)', last_p.text.split('\n')[0], re.I)
                    if is_prev_date_subtitle and is_date_line:
                        last_p.add_run(f"\n{original_line}")
                        i += 1
                        continue

                    current_p = None

                    if tag == "Subtitle" and re.match(r'^Amendments\s*:?', line_after_url_removal, re.I):
                        is_within_amendments = True
                        is_within_heading_5 = False
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                    elif tag == "Heading 5":
                        is_within_heading_5 = True
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                    elif tag in ["Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3", "Heading 4"]:
                        is_within_heading_5 = False
                        
                        # Reset subsection tracking when we hit a new section
                        if tag in ["Heading 1", "Heading 2", "Heading 3"]:
                            is_within_subsection = False
                            
                        if tag == "Heading 3":
                            # Check for standard section format (number followed by dot)
                            sec_match = re.match(r'^(\d+[A-Za-z]?)\.\s*(.*)', line_after_url_removal)
                            # Check for special section formats with symbols - UPDATED PATTERN
                            symbol_sec_match = re.match(r'^([♦◉])\s*(\d+[A-Za-z]?)\.?\s*(.*)', line_after_url_removal)
                            
                            if sec_match:
                                sec_num, sec_body = sec_match.groups()
                                sec_body = sec_body.strip()
                                parts = re.split(r'\s*(?=\(\d+\))', sec_body, maxsplit=1)
                                section_title = parts[0].strip()
                                current_p = add_styled_paragraph(doc, f"Section {sec_num}: {section_title}", "Heading 3")
                                if len(parts) > 1:
                                    first_subsection_text = parts[1].strip()
                                    sub_match = re.match(r'^\((\d+)\)\s*(.*)', first_subsection_text)
                                    if sub_match:
                                        sub_num, sub_title = sub_match.groups()
                                        # Check if this is a lettered subsection like (a), (b)
                                        if re.match(r'^[a-z]$', sub_num):
                                            is_within_subsection = True
                                            current_p = add_styled_paragraph(doc, f"({sub_num}) {sub_title.strip()}", "Heading 4")
                                        else:
                                            # If we're inside a lettered subsection, treat numbered items as normal text
                                            if is_within_subsection:
                                                current_p = add_styled_paragraph(doc, f"({sub_num}) {sub_title.strip()}", "Normal")
                                            else:
                                                # Format long subsections with the number separated from content
                                                current_p = add_styled_paragraph(doc, f"Subsection ({sub_num}):", "Heading 4")
                                                current_p = add_styled_paragraph(doc, f"{sub_title.strip()}", "Normal")
                                                tag = "Normal"
                                    else:
                                        current_p = add_styled_paragraph(doc, first_subsection_text, "Normal")
                                        tag = "Normal"
                            elif symbol_sec_match:
                                symbol, sec_num, sec_body = symbol_sec_match.groups()
                                sec_body = sec_body.strip()
                                section_format = f"{symbol}{sec_num}"
                                
                                parts = re.split(r'\s*(?=\(\d+\))', sec_body, maxsplit=1)
                                section_title = parts[0].strip()
                                
                                current_p = add_styled_paragraph(doc, f"Section {section_format}: {section_title}", "Heading 3")
                                
                                if len(parts) > 1:
                                    first_subsection_text = parts[1].strip()
                                    sub_match = re.match(r'^\((\d+)\)\s*(.*)', first_subsection_text)
                                    if sub_match:
                                        sub_num, sub_text = sub_match.groups()
                                        # Check if we're inside a lettered subsection
                                        if is_within_subsection:
                                            current_p = add_styled_paragraph(doc, f"({sub_num}) {sub_text.strip()}", "Normal")
                                            tag = "Normal"
                                        else:
                                            # Format long subsections with the number separated from content
                                            current_p = add_styled_paragraph(doc, f"Subsection ({sub_num}):", "Heading 4")
                                            current_p = add_styled_paragraph(doc, f"{sub_text.strip()}", "Normal")
                                            tag = "Normal"
                                    else:
                                        current_p = add_styled_paragraph(doc, first_subsection_text, "Normal")
                                        tag = "Normal"
                            else:
                                current_p = add_styled_paragraph(doc, line_after_url_removal, "Heading 3")
                        elif tag == "Heading 2":
                            chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line_after_url_removal, re.I)
                            if chap_match:
                                chap_num, chap_title = chap_match.groups()
                                full_title = f"Chapter {chap_num.strip()}: {chap_title.strip()}"
                                current_p = add_styled_paragraph(doc, full_title, "Heading 2")
                            else:
                                current_p = add_styled_paragraph(doc, line_after_url_removal, "Heading 2")
                        else:
                            current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                    elif tag == "Normal":
                        current_p = add_styled_paragraph(doc, line_after_url_removal, "Normal", is_under_h5=is_within_heading_5)

                    if current_p:
                        last_p = current_p
                        last_tag = tag
                    
                    i += 1
            
            # Add any tables that were detected but not already processed
            if all_tables:
                doc.add_paragraph("").add_run("Tables from Document:").bold = True
                for table_data in all_tables:
                    # Filter out empty rows and cells
                    filtered_table = [[cell if cell else "" for cell in row] for row in table_data if any(cell for cell in row)]
                    if filtered_table:
                        add_table_to_doc(doc, filtered_table)
                        doc.add_paragraph()  # Add space after table

        if output_dir:
            output_filename = os.path.basename(os.path.splitext(pdf_path)[0]) + "_structured.docx"
            output_path = os.path.join(output_dir, output_filename)
        else:
            output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
            
        doc.save(output_path)
        logging.info(f"Successfully converted: {pdf_path}")
        return output_path
        
    except Exception as e:
        error_msg = f"Error processing: {pdf_path} - {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return None

def select_files():
    global selected_pdf_paths
    selected_pdf_paths.clear()
    status_label.config(text="")

    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files", 
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_paths:
        convert_button.config(state=DISABLED)
        return

    selected_pdf_paths.extend(file_paths)
    
    status_label.config(text=f"{len(selected_pdf_paths)} PDF(s) selected", fg="black")
    
    # Enable convert button when files are selected
    convert_button.config(state=NORMAL)

def convert_files():
    if not selected_pdf_paths:
        messagebox.showwarning("No Files", "Please select PDF files first.")
        return
        
    select_button.config(state=DISABLED)
    convert_button.config(state=DISABLED)
    
    status_label.config(text="Converting...", fg="blue")
    root.update_idletasks()
    
    successes, failures = [], []
    for path in selected_pdf_paths:
        result = convert_pdf_to_docx(path)
        if result:
            successes.append(os.path.basename(result))
        else:
            failures.append(os.path.basename(path))
    
    status_label.config(text="Done ✔", fg="green")
    
    if failures:
        messagebox.showinfo("Conversion Complete", 
                          f"{len(successes)} file(s) converted successfully.\n{len(failures)} file(s) failed.")
    else:
        messagebox.showinfo("Conversion Complete", 
                          f"All {len(successes)} file(s) converted successfully.")
    
    select_button.config(state=NORMAL)
    convert_button.config(state=NORMAL)
    selected_pdf_paths.clear()

# Create the main window
root = tk.Tk()
root.title("PDF to Structured DOCX Converter")
root.geometry("500x300")
root.resizable(True, True)

# Create a frame for the buttons
button_frame = Frame(root)
button_frame.pack(pady=20)

# Create buttons
select_button = tk.Button(button_frame, text="Select PDF Files", command=select_files)
select_button.pack(side=tk.LEFT, padx=5)

convert_button = tk.Button(button_frame, text="Convert to DOCX", command=convert_files, state=DISABLED)
convert_button.pack(side=tk.LEFT, padx=5)

# Create a status label
status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack(pady=10)

# Start the main loop
root.mainloop()