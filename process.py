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

# Global variables
selected_pdf_paths = []
PRESERVE_AMENDMENTS = True
FORMAT_DATES = True

# Set up logging
log_file = f"pdf_conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def classify_line(line):
    line = line.strip()
    if re.match(r'^Schedule\b.*', line, re.I): # Added check for Schedule
        return "Heading 5"
    elif re.match(r'^NEPAL.*ACT.*\d{4}', line, re.I):
        return "Title"
    # Change this block from Title to Subtitle
    elif re.match(r'^Date of (Authentication|Publication|Authentication and Publication)\b.*', line, re.I):
        return "Subtitle"
    elif re.match(r'^AN ACT MADE TO.*', line, re.I):
        return "Subtitle"
    elif re.match(r'^Amendments\s*:?', line, re.I): # Added this line to classify Amendments as Subtitle
        return "Subtitle"
    elif re.match(r'^Preamble\s*:?', line, re.I):
        return "Heading 1"
    elif re.match(r'^Chapter\s*[-–]?\s*\d+', line, re.I):
        return "Heading 2"
    elif re.match(r'^\d+\.\s+', line):
        return "Heading 3"
    elif re.match(r'^\(\d+\)', line):  # Only numeric parentheses
        return "Heading 4"
    # Add a check for the date line format itself, though we primarily use it contextually
    elif re.match(r'^\d{4}\.\d{1,2}\.\d{1,2}.*', line):
         # Tentatively classify as Normal, context will decide if it's appended
         return "Normal"
    else:
        return "Normal"

def add_styled_paragraph(doc, text, style_tag, is_under_h5=False): # Added is_under_h5 flag
    p = doc.add_paragraph() # Create paragraph first
    # Apply style first if it exists
    try:
        p.style = doc.styles[style_tag]
    except KeyError:
        # Apply basic formatting if style doesn't exist
        p.style = doc.styles['Normal'] # Default to Normal if style tag is unknown

    # Add text *after* potentially setting the style
    run = p.add_run(text)

    # Apply specific formatting based on tag or context
    if style_tag == "Heading 5":
        # Add any specific formatting for Heading 5 itself if needed
        pass
    elif style_tag == "Heading 4":
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal" and is_under_h5: # Indent Normal text under Heading 5
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal":
        # Reset indent for regular Normal text if needed
        p.paragraph_format.left_indent = Inches(0) # Explicitly set to 0
    elif style_tag == "Title":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run.font.size = Pt(16) # Example size for Title
        run.font.bold = True
    elif style_tag == "Subtitle":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run.font.size = Pt(14) # Example size for Subtitle
        run.font.italic = True # Example style for Subtitle

    return p # Return the created paragraph object

def convert_pdf_to_docx(pdf_path, output_dir=None):
    doc = Document()
    url_pattern = re.compile(r'(?:https?://|www\.)\S+')
    is_within_heading_5 = False # State variable to track if we are inside a Schedule block
    last_p = None # Track the last paragraph object added
    last_tag = None # Track the tag of the last paragraph added
    is_first_line_of_document = True # Flag to identify the very first valid line
    is_within_amendments = False # Flag to track if we are inside the Amendments block

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                lines = text.split("\n")
                for line in lines:
                    original_line = line.strip() # Keep original for date check
                    line = original_line # Use stripped line for processing

                    # --- Skip empty lines (unless under H5) ---
                    if not line:
                        # Don't add empty lines if inside amendments block either
                        if is_within_heading_5 and not is_within_amendments:
                             current_p = add_styled_paragraph(doc, "", "Normal", is_under_h5=True)
                             # Update last_p/tag only if a paragraph was actually added
                             last_p = current_p
                             last_tag = "Normal"
                        continue

                    # --- Skip page numbers ---
                    if re.fullmatch(r'\d+', line):
                        continue

                    # --- Remove URLs (but check original line for date pattern) ---
                    line_after_url_removal = url_pattern.sub('', line).strip()
                    if not line_after_url_removal:
                        continue

                    # --- Force First Line as Title ---
                    if is_first_line_of_document:
                        tag = "Title" # Force the tag
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                        last_p = current_p
                        last_tag = tag
                        is_first_line_of_document = False # We've processed the first line
                        is_within_amendments = False # Cannot be in amendments on first line
                        is_within_heading_5 = False
                        continue # Move to the next line

                    # --- Classify Line (only if not the first line) ---
                    tag = classify_line(line_after_url_removal)

                    # --- Handle Continuation of Amendments Block ---
                    if is_within_amendments:
                        # Check if the current line signals the end of the amendments
                        # End if it's any Heading (except maybe Normal/H4/H5 if allowed within)
                        # Or a Subtitle that ISN'T "Amendments:" itself (unlikely case)
                        if tag in ["Title", "Heading 1", "Heading 2"] or \
                           (tag == "Subtitle" and not re.match(r'^Amendments\s*:?', line_after_url_removal, re.I)):
                            is_within_amendments = False # End of amendments block
                            # Fall through to process this line normally below
                        else:
                            # Format amendments content as Subtitle instead of Normal
                            current_p = add_styled_paragraph(doc, line_after_url_removal, "Subtitle")
                            # Update last paragraph and tag tracking
                            last_p = current_p
                            last_tag = "Subtitle"
                            continue # Skip normal processing for this line

                    # --- Check for Date Appending (Only if NOT inside amendments) ---
                    is_date_line = re.match(r'^\d{4}\.\d{1,2}\.\d{1,2}', original_line)
                    is_prev_date_subtitle = last_p is not None and last_tag == "Subtitle" and \
                                            re.match(r'^Date of (Authentication|Publication|Authentication and Publication)', last_p.text.split('\n')[0], re.I)
                    if is_prev_date_subtitle and is_date_line:
                        last_p.add_run(f"\n{original_line}")
                        continue

                    # --- Process the line normally (if not first line, not appended date, not appended amendment) ---
                    current_p = None # Reset current paragraph tracker

                    # --- Check if this line *starts* the Amendments block ---
                    if tag == "Subtitle" and re.match(r'^Amendments\s*:?', line_after_url_removal, re.I):
                        is_within_amendments = True
                        is_within_heading_5 = False # Subtitles reset H5 state
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                    # --- Handle other tags ---
                    elif tag == "Heading 5":
                        is_within_heading_5 = True
                        # is_within_amendments = False # Already handled by the check above
                        current_p = add_styled_paragraph(doc, line_after_url_removal, tag)
                    elif tag in ["Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3", "Heading 4"]:
                        is_within_heading_5 = False # Reset H5 state
                        # is_within_amendments = False # Already handled by the check above
                        # Process these headings as before
                        if tag == "Heading 3":
                            # Restore H3 processing logic
                            sec_match = re.match(r'^(\d+)\.\s*(.*)', line_after_url_removal)
                            if sec_match:
                                sec_num, sec_body = sec_match.groups()
                                sec_body = sec_body.strip()
                                parts = re.split(r'\s*(?=\(\d+\))', sec_body, maxsplit=1)
                                section_title = parts[0].strip()
                                current_p = add_styled_paragraph(doc, f"Section {sec_num}: {section_title}", "Heading 3")
                                # If subsections are added, update current_p to the *last* one added in this block
                                if len(parts) > 1:
                                    first_subsection_text = parts[1].strip()
                                    sub_match = re.match(r'^\((\d+)\)\s*(.*)', first_subsection_text)
                                    if sub_match:
                                        sub_num, sub_text = sub_match.groups()
                                        current_p = add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_text.strip()}", "Heading 4") # Update current_p
                                        tag = "Heading 4" # Update tag if subsection added
                                    else:
                                        current_p = add_styled_paragraph(doc, first_subsection_text, "Normal") # Update current_p
                                        tag = "Normal" # Update tag
                            else:
                                current_p = add_styled_paragraph(doc, line_after_url_removal, "Heading 3")
                        elif tag == "Heading 4":
                            # Restore H4 processing logic
                            sub_match = re.match(r'^\((\d+)\)\s*(.*)', line_after_url_removal)
                            if sub_match:
                                sub_num, sub_title = sub_match.groups()
                                current_p = add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_title.strip()}", "Heading 4")
                            else:
                                current_p = add_styled_paragraph(doc, line_after_url_removal, "Heading 4")
                        elif tag == "Heading 2":
                            # Restore H2 processing logic
                            chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line_after_url_removal, re.I)
                            if chap_match:
                                chap_num, chap_title = chap_match.groups()
                                full_title = f"Chapter {chap_num.strip()}: {chap_title.strip()}"
                                current_p = add_styled_paragraph(doc, full_title, "Heading 2")
                            else:
                                current_p = add_styled_paragraph(doc, line_after_url_removal, "Heading 2")
                        else: # Title (non-first), Subtitle (non-date, non-amendments), Heading 1
                             current_p = add_styled_paragraph(doc, line_after_url_removal, tag)

                    elif tag == "Normal":
                        # is_within_amendments = False # Already handled by the check above
                        current_p = add_styled_paragraph(doc, line_after_url_removal, "Normal", is_under_h5=is_within_heading_5)

                    # --- Update last paragraph and tag tracking ---
                    if current_p: # Only update if a new paragraph was actually created in this iteration
                        last_p = current_p
                        last_tag = tag # Use the potentially updated tag

        # Modify output path based on output_dir
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

def select_and_convert():
    # Split this function into two separate functions
    pass  # This function can be removed

def select_files():
    # Clear previous selection
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
    
    # Update status label with file count
    status_label.config(text=f"{len(selected_pdf_paths)} PDF(s) selected", fg="black")
    
    # Enable convert button now that files are selected
    convert_button.config(state=NORMAL)

def convert_files():
    if not selected_pdf_paths:
        messagebox.showwarning("No Files", "Please select PDF files first.")
        return
        
    # Disable buttons during conversion
    select_button.config(state=DISABLED)
    convert_button.config(state=DISABLED)
    
    # Update status
    status_label.config(text="Converting...", fg="blue")
    root.update_idletasks()
    
    # Process files
    successes, failures = [], []
    for path in selected_pdf_paths:
        result = convert_pdf_to_docx(path)
        if result:
            successes.append(os.path.basename(result))
        else:
            failures.append(os.path.basename(path))
    
    # Show completion message
    status_label.config(text="Done ✔", fg="green")
    
    # Show simple summary
    if failures:
        messagebox.showinfo("Conversion Complete", 
                           f"{len(successes)} file(s) converted successfully.\n{len(failures)} file(s) failed.")
    else:
        messagebox.showinfo("Conversion Complete", 
                           f"All {len(successes)} file(s) converted successfully.")
    
    # Re-enable buttons
    select_button.config(state=NORMAL)
    convert_button.config(state=DISABLED)  # Disable convert until new files are selected
    
    # Clear the selection after conversion
    selected_pdf_paths.clear()

# Create simple UI
root = tk.Tk()
root.title("PDF → DOCX Converter")
root.geometry("400x200")  # Slightly taller to accommodate two buttons
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Main frame
main_frame = Frame(root, bg="#f0f0f0", padx=20, pady=20)
main_frame.pack(expand=True, fill="both")

# Title
title = tk.Label(main_frame, text="PDF to Structured DOCX Converter", 
                font=("Helvetica", 14, "bold"), bg="#f0f0f0")
title.pack(pady=(0, 20))

# Button frame to hold both buttons
button_frame = Frame(main_frame, bg="#f0f0f0")
button_frame.pack(pady=10)

# Select button
select_button = tk.Button(button_frame, text="Select PDFs", 
                         font=("Helvetica", 12),
                         command=select_files,
                         width=15, height=1)
select_button.pack(side="left", padx=10)

# Convert button (initially disabled)
convert_button = tk.Button(button_frame, text="Convert", 
                          font=("Helvetica", 12),
                          command=convert_files,
                          width=15, height=1,
                          state=DISABLED)
convert_button.pack(side="left", padx=10)

# Status label
status_label = tk.Label(main_frame, text="", font=("Helvetica", 10), bg="#f0f0f0")
status_label.pack(pady=10)

# Start the app
root.mainloop()