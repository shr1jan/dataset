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
    # Regex to find URLs (http, https, or www.)
    url_pattern = re.compile(r'(?:https?://|www\.)\S+')
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

                    # Check if the line consists only of digits (likely a page number)
                    if re.fullmatch(r'\d+', line):
                        continue # Skip this line as it's probably a page number

                    # Remove URLs from the line
                    line = url_pattern.sub('', line).strip()
                    # If the line becomes empty after removing URL, skip it
                    if not line:
                        continue

                    tag = classify_line(line)

                    # --- Handle Section/Subsection logic ---
                    if tag == "Heading 3":
                        sec_match = re.match(r'^(\d+)\.\s*(.*)', line)
                        if sec_match:
                            sec_num, sec_body = sec_match.groups()
                            sec_body = sec_body.strip()

                            # Try to split the section body at the first subsection marker like " (1)"
                            # Use maxsplit=1 to only split the first occurrence
                            parts = re.split(r'\s*(?=\(\d+\))', sec_body, maxsplit=1)
                            section_title = parts[0].strip()

                            # Add the main section title
                            add_styled_paragraph(doc, f"Section {sec_num}: {section_title}", "Heading 3")

                            # If there was a split, process the second part as the first subsection
                            if len(parts) > 1:
                                first_subsection_text = parts[1].strip()
                                sub_match = re.match(r'^\((\d+)\)\s*(.*)', first_subsection_text)
                                if sub_match:
                                    sub_num, sub_text = sub_match.groups()
                                    add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_text.strip()}", "Heading 4")
                                else:
                                    # If regex fails (unlikely), add the text as normal under the section
                                    add_styled_paragraph(doc, first_subsection_text, "Normal")
                        else:
                             # Fallback
                            add_styled_paragraph(doc, line, "Heading 3")

                    elif tag == "Heading 4":
                        # Match subsection number and the rest of the line as title
                        sub_match = re.match(r'^\((\d+)\)\s*(.*)', line)
                        if sub_match:
                            sub_num, sub_title = sub_match.groups()
                            # Add the subsection line, potentially including its title
                            add_styled_paragraph(doc, f"Subsection ({sub_num}): {sub_title.strip()}", "Heading 4")
                        else:
                            # Fallback if regex fails
                            add_styled_paragraph(doc, line, "Heading 4")

                    elif tag == "Heading 2":
                        chap_match = re.match(r'^Chapter\s*[-–]?\s*(\d+)\s*(.*)', line, re.I)
                        if chap_match:
                            chap_num, chap_title = chap_match.groups()
                            full_title = f"Chapter {chap_num.strip()}: {chap_title.strip()}" # Changed 'chapter' to 'Chapter' for consistency
                            add_styled_paragraph(doc, full_title, "Heading 2")
                        else:
                            add_styled_paragraph(doc, line, "Heading 2")

                    # Handle other tags or add as normal text if not a specific heading type handled above
                    elif tag not in ["Heading 3", "Heading 4", "Heading 2"]: # Check if not already handled
                         add_styled_paragraph(doc, line, tag)


        output_path = os.path.splitext(pdf_path)[0] + "_structured.docx"
        doc.save(output_path)
        return output_path

    except Exception as e:
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
