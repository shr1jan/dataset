import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Frame, END, DISABLED, NORMAL
from tkinter import ttk
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

# Global variable to store selected file paths
selected_pdf_paths = []

def classify_line(line):
    line = line.strip()
    if re.match(r'^Schedule\b.*', line, re.I): # Added check for Schedule
        return "Heading 5"
    elif re.match(r'^NEPAL.*ACT.*\d{4}', line, re.I):
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

def add_styled_paragraph(doc, text, style_tag, is_under_h5=False): # Added is_under_h5 flag
    p = doc.add_paragraph(text)
    # Use built-in styles if they match, otherwise apply formatting manually
    try:
        p.style = doc.styles[style_tag]
    except KeyError:
        # Apply basic formatting if style doesn't exist (though Heading 1-5 should)
        # For custom tags like "Normal Under H5", we handle formatting below
        p.style = doc.styles['Normal'] # Default to Normal if style tag is unknown

    # Apply specific formatting based on tag or context
    if style_tag == "Heading 5":
        # Add any specific formatting for Heading 5 itself if needed
        # For now, it will just use the built-in 'Heading 5' style
        pass
    elif style_tag == "Heading 4":
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal" and is_under_h5: # Indent Normal text under Heading 5
        p.paragraph_format.left_indent = Inches(0.3)
    elif style_tag == "Normal":
        # Reset indent for regular Normal text if needed, though default is usually 0
        p.paragraph_format.left_indent = Inches(0) # Explicitly set to 0
    elif style_tag in ["Title", "Subtitle"]:
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def convert_pdf_to_docx(pdf_path):
    doc = Document()
    url_pattern = re.compile(r'(?:https?://|www\.)\S+')
    is_within_heading_5 = False # State variable to track if we are inside a Schedule block

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
                        # Add empty lines if they are within H5 content, otherwise skip
                        if is_within_heading_5:
                             add_styled_paragraph(doc, "", "Normal", is_under_h5=True)
                        continue

                    # Check for page number first
                    if re.fullmatch(r'\d+', line):
                        continue

                    # Remove URLs
                    line = url_pattern.sub('', line).strip()
                    if not line:
                        continue

                    tag = classify_line(line)

                    # --- State Management for Heading 5 ---
                    if tag == "Heading 5":
                        is_within_heading_5 = True
                        add_styled_paragraph(doc, line, tag) # Add the Schedule line itself
                    elif tag in ["Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3", "Heading 4"]:
                        is_within_heading_5 = False # Reset state when any other heading is found
                        # Process these headings as before
                        if tag == "Heading 3":
                            sec_match = re.match(r'^(\d+)\.\s*(.*)', line)
                            if sec_match:
                                sec_num, sec_body = sec_match.groups()
                                sec_body = sec_body.strip()
                                parts = re.split(r'\s*(?=\(\d+\))', sec_body, maxsplit=1)
                                section_title = parts[0].strip()
                                add_styled_paragraph(doc, f"Section {sec_num}: {section_title}", "Heading 3")
                                if len(parts) > 1:
                                    first_subsection_text = parts[1].strip()
                                    sub_match = re.match(r'^\((\d+)\)\s*(.*)', first_subsection_text)
                                    if sub_match:
                                        sub_num, sub_text = sub_match.groups()
                                        add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_text.strip()}", "Heading 4")
                                    else:
                                        add_styled_paragraph(doc, first_subsection_text, "Normal")
                            else:
                                add_styled_paragraph(doc, line, "Heading 3")
                        elif tag == "Heading 4":
                            sub_match = re.match(r'^\((\d+)\)\s*(.*)', line)
                            if sub_match:
                                sub_num, sub_title = sub_match.groups()
                                add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_title.strip()}", "Heading 4")
                            else:
                                add_styled_paragraph(doc, line, "Heading 4")
                        elif tag == "Heading 2":
                            chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line, re.I)
                            if chap_match:
                                chap_num, chap_title = chap_match.groups()
                                full_title = f"Chapter {chap_num.strip()}: {chap_title.strip()}"
                                add_styled_paragraph(doc, full_title, "Heading 2")
                            else:
                                add_styled_paragraph(doc, line, "Heading 2")
                        else: # Title, Subtitle, Heading 1
                             add_styled_paragraph(doc, line, tag)

                    elif tag == "Normal":
                        # Add as normal text, applying indentation if inside Heading 5
                        add_styled_paragraph(doc, line, "Normal", is_under_h5=is_within_heading_5)

        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path

    except Exception as e:
        # Ensure state is reset even on error? Maybe not necessary per-file.
        print("❌ Error processing:", pdf_path, "\n", e)
        return None

# --- New function to handle file selection ---
def select_files(listbox, convert_button, status_label):
    global selected_pdf_paths
    # Clear previous selection
    listbox.delete(0, END)
    selected_pdf_paths.clear()
    status_label.config(text="", fg="black") # Reset status

    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files (up to 10)",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_paths:
        convert_button.config(state=DISABLED) # Disable convert if no files selected
        return

    if len(file_paths) > 10:
        messagebox.showwarning("Limit Exceeded", "You can only select up to 10 PDFs at a time.")
        # Keep only the first 10 files
        file_paths = file_paths[:10]
        # Optionally inform the user which files were kept, or just proceed

    selected_pdf_paths.extend(file_paths)

    # Populate the listbox
    for path in selected_pdf_paths:
        listbox.insert(END, os.path.basename(path))

    # Enable the convert button if files are selected
    if selected_pdf_paths:
        convert_button.config(state=NORMAL)
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
    status_label.config(text="Converting...", fg="blue")
    # Force UI update
    listbox.master.update_idletasks()


    successes, failures = [], []
    for path in selected_pdf_paths:
        result = convert_pdf_to_docx(path)
        if result:
            successes.append(os.path.basename(result))
        else:
            failures.append(os.path.basename(path))

    status_label.config(text="Done ✔", fg="green")

    summary = f"{len(successes)} file(s) converted successfully:\n" + "\n".join(successes)
    if failures:
        summary += f"\n\n{len(failures)} conversion(s) failed:\n" + "\n".join(failures)

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
