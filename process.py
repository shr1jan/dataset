import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Frame, END, DISABLED, NORMAL
from tkinter import ttk
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt # Add Pt import here

# Global variable to store selected file paths
selected_pdf_paths = []

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

def convert_pdf_to_docx(pdf_path):
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

        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path

    except Exception as e:
        print("❌ Error processing:", pdf_path, "\n", e)
        # Reset state variables in case of error within a file? Maybe not necessary.
        # is_within_amendments = False
        return None

# --- New function to handle file selection ---
def select_files(listbox, convert_button, status_label):
    global selected_pdf_paths
    # Clear previous selection
    listbox.delete(0, END)
    selected_pdf_paths.clear()
    status_label.config(text="", fg="black") # Reset status

    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files", # Removed limit from title
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_paths:
        convert_button.config(state=DISABLED) # Disable convert if no files selected
        return

    # Removed the check and truncation for the file limit
    # if len(file_paths) > 20: ... (This block is deleted)

    selected_pdf_paths.extend(file_paths)

    # Populate the listbox
    listbox.delete(0, END) # Clear listbox before repopulating
    for path in selected_pdf_paths:
        listbox.insert(END, os.path.basename(path))

    # Enable the convert button if files are selected
    if selected_pdf_paths:
        convert_button.config(state=NORMAL)
        status_label.config(text=f"{len(selected_pdf_paths)} file(s) selected.", fg="black") # Show count
    else:
        convert_button.config(state=DISABLED)

# --- New function to start the conversion process ---
def start_conversion(listbox, select_button, convert_button, status_label):
    global selected_pdf_paths
    if not selected_pdf_paths:
        messagebox.showwarning("No Files", "Please select PDF files first.")
        return

    # Disable buttons during conversion
    select_button.config(state=DISABLED)
    convert_button.config(state=DISABLED)

    all_successes, all_failures = [], []
    batch_size = 10
    total_files = len(selected_pdf_paths)
    num_batches = (total_files + batch_size - 1) // batch_size # Calculate total batches

    original_paths_to_process = list(selected_pdf_paths) # Create a copy to iterate over

    for i in range(num_batches):
        start_index = i * batch_size
        end_index = start_index + batch_size
        current_batch_paths = original_paths_to_process[start_index:end_index]
        current_batch_num = i + 1

        status_label.config(text=f"Converting batch {current_batch_num}/{num_batches} ({len(current_batch_paths)} files)...", fg="blue")
        # Force UI update
        listbox.master.update_idletasks()

        batch_successes, batch_failures = [], []
        for path in current_batch_paths:
            result = convert_pdf_to_docx(path)
            if result:
                batch_successes.append(os.path.basename(result))
            else:
                batch_failures.append(os.path.basename(path))

        all_successes.extend(batch_successes)
        all_failures.extend(batch_failures)

        # Optional: Short pause or update after each batch if needed
        # time.sleep(0.1) # Requires 'import time'

    status_label.config(text="Done ✔", fg="green")

    summary = f"{len(all_successes)} file(s) converted successfully across {num_batches} batches.\n"
    # Optionally list successes if not too many:
    # summary += "\n".join(all_successes)

    if all_failures:
        summary += f"\n\n{len(all_failures)} conversion(s) failed:\n" + "\n".join(all_failures)

    messagebox.showinfo("Batch Conversion Complete", summary)

    # Clear list and selection after conversion, re-enable select button
    listbox.delete(0, END)
    selected_pdf_paths.clear()
    select_button.config(state=NORMAL)
    # Keep convert button disabled until new files are selected


# --- Remove the old select_files_and_convert function ---
# def select_files_and_convert(status_label):
#    ... (delete this function) ...


def launch_gui():
    root = tk.Tk()
    root.title("PDF → Structured DOCX Converter")
    root.geometry("600x400") # Increased size for listbox
    root.resizable(True, True) # Allow resizing
    root.configure(bg="#f0f0f0") # Slightly different background

    # --- Main Frame ---
    main_frame = Frame(root, bg="#f0f0f0", padx=15, pady=15)
    main_frame.pack(expand=True, fill="both")

    # --- Title ---
    title = tk.Label(main_frame, text="Batch PDF to Structured Word Converter", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
    title.pack(pady=(0, 15))

    # --- File List Frame ---
    list_frame = Frame(main_frame, bd=1, relief="sunken")
    list_frame.pack(pady=10, fill="both", expand=True)

    scrollbar = Scrollbar(list_frame, orient="vertical")
    listbox = Listbox(list_frame, yscrollcommand=scrollbar.set, height=10, font=("Helvetica", 10))
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side="right", fill="y")
    listbox.pack(side="left", fill="both", expand=True)

    # --- Button Frame ---
    button_frame = Frame(main_frame, bg="#f0f0f0")
    button_frame.pack(pady=(15, 5), fill="x")

    # --- Status Label ---
    status_label = tk.Label(main_frame, text="", font=("Helvetica", 10, "italic"), bg="#f0f0f0", fg="gray")
    status_label.pack(pady=(0, 10))


    # --- Buttons ---
    # Use ttk for better styling if available
    style = ttk.Style()
    style.configure("TButton", font=("Helvetica", 10), padding=6)

    select_button = ttk.Button(button_frame, text="Select PDF Files")
    select_button.pack(side="left", padx=(0, 10), expand=True, fill="x")

    convert_button = ttk.Button(button_frame, text="Convert Selected Files", state=DISABLED)
    convert_button.pack(side="right", padx=(10, 0), expand=True, fill="x")

    # --- Configure Button Commands (pass necessary widgets) ---
    select_button.config(command=lambda: select_files(listbox, convert_button, status_label))
    convert_button.config(command=lambda: start_conversion(listbox, select_button, convert_button, status_label))


    # --- Note ---
    note = tk.Label(main_frame, text="Output files will be saved next to the original PDFs", font=("Helvetica", 9), fg="gray", bg="#f0f0f0")
    note.pack(side="bottom", pady=(5, 0))

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
