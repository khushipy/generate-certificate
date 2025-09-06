import os
import shutil
import re
from datetime import datetime
from docx import Document
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------- CONFIG ----------------
TEMPLATE_FILENAME = "certificate.docx"
EXCEL_FILENAME = "internship_details.xlsx"
OUTPUT_DIRNAME = "certificates"

# Paths relative to current working directory
current_directory = os.getcwd()
template_path = os.path.join(current_directory, TEMPLATE_FILENAME)
excel_path = os.path.join(current_directory, EXCEL_FILENAME)
output_dir = os.path.join(current_directory, OUTPUT_DIRNAME)
os.makedirs(output_dir, exist_ok=True)

# ---------------- HELPERS ----------------
PLACEHOLDER_RE = re.compile(r'\{\s*([^}]+?)\s*\}')

def format_value(val):
    """Format Excel values; special handling for dates (and '19th July 2025' style)."""
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d-%b-%Y")  # 19-Jul-2025
    if isinstance(val, str):
        cleaned = val.strip()
        try:
            c = (cleaned.replace("st", "").replace("nd", "")
                        .replace("rd", "").replace("th", ""))
            parsed = datetime.strptime(c.strip(), "%d %B %Y")
            return parsed.strftime("%d-%b-%Y")
        except Exception:
            try:
                c = (cleaned.replace("st", "").replace("nd", "")
                            .replace("rd", "").replace("th", ""))
                parsed = datetime.strptime(c.strip(), "%d %b %Y")
                return parsed.strftime("%d-%b-%Y")
            except Exception:
                return cleaned
    return str(val)

def copy_formatting(target_run, source_run):
    """Copy a few basic formatting attributes if available."""
    try:
        target_run.bold = source_run.bold
        target_run.italic = source_run.italic
        target_run.underline = source_run.underline
    except Exception:
        pass
    try:
        if source_run.font.name:
            target_run.font.name = source_run.font.name
    except Exception:
        pass
    try:
        if source_run.font.size:
            target_run.font.size = source_run.font.size
    except Exception:
        pass
    try:
        if source_run.font.color and source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb
    except Exception:
        pass

def replace_placeholders_in_paragraph(paragraph, mapping):
    """Replace placeholders {Key} with mapping values (bold) in a paragraph."""
    if not paragraph.runs:
        return
    full_text = ''.join(run.text for run in paragraph.runs)
    if "{" not in full_text:
        return
    matches = list(PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return

    # run_map: for each char index which run index it came from
    run_map = []
    for idx, run in enumerate(paragraph.runs):
        text = run.text or ""
        run_map.extend([idx] * len(text))

    # build segments (text / replace)
    segments = []
    last = 0
    for m in matches:
        s, e = m.start(), m.end()
        key = m.group(1).strip()
        if s > last:
            segments.append(("text", full_text[last:s], last))
        if key in mapping:
            segments.append(("replace", mapping[key], s))
        else:
            segments.append(("text", full_text[s:e], s))
        last = e
    if last < len(full_text):
        segments.append(("text", full_text[last:], last))

    # clear old runs
    for run in paragraph.runs:
        run.text = ""

    # insert new runs
    for seg_type, seg_text, seg_pos in segments:
        if run_map:
            src_idx = run_map[min(seg_pos, len(run_map) - 1)]
            src_run = paragraph.runs[src_idx]
        else:
            src_run = None
        new_run = paragraph.add_run(seg_text)
        if src_run:
            copy_formatting(new_run, src_run)
        if seg_type == "replace":
            new_run.bold = True

def process_document(doc, mapping):
    """Apply placeholder replacement across document body, tables, headers, footers."""
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)
    for section in doc.sections:
        for container in (section.header, section.footer):
            if container:
                for p in container.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)
                for table in container.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                replace_placeholders_in_paragraph(p, mapping)

# ---------------- DATA LOADING ----------------
wb = load_workbook(excel_path, data_only=True)
sheet = wb.active
headers = [str(cell.value).strip() if cell.value else "" for cell in sheet[1]]

# Ensure Certificate No. column exists
if "Certificate No." not in headers:
    raise ValueError("Excel must contain a column named 'Certificate No.'")

certno_index = headers.index("Certificate No.")
name_index = headers.index("name") if "name" in headers else None

# Build certificate list for UI
cert_data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    cert_val = row[certno_index]
    if not cert_val:
        continue
    cert_no = str(cert_val).strip()
    name = ""
    if name_index is not None and row[name_index]:
        name = str(row[name_index]).strip()
    cert_data.append((cert_no, name, row))

# ---------------- UI ----------------
def generate_selected():
    selected_indices = listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("No selection", "Please select at least one certificate.")
        return

    count = 0
    for idx in selected_indices:
        cert_no, name, row = cert_data[int(idx)]
        output_file = os.path.join(output_dir, f"{cert_no}.docx")

        try:
            shutil.copy2(template_path, output_file)
            doc = Document(output_file)
        except Exception as e:
            messagebox.showerror("File error", f"Error with template for {cert_no}: {e}")
            continue

        # Build mapping excluding Certificate No.
        mapping = {}
        for i, h in enumerate(headers):
            key = str(h).strip() if h else ""
            val = row[i] if i < len(row) else None
            mapping[key] = format_value(val)

        process_document(doc, mapping)
        try:
            doc.save(output_file)
            count += 1
        except Exception as e:
            messagebox.showerror("Save error", f"Could not save {output_file}: {e}")

    messagebox.showinfo("Done", f"Generated {count} certificate(s).")

# Tkinter window
root = tk.Tk()
root.title("Certificate Generator")

tk.Label(root, text="Select certificates to generate:").pack(pady=6)

listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=18)
for cert_no, name, _ in cert_data:
    display = f"{cert_no} - {name}" if name else cert_no
    listbox.insert(tk.END, display)
listbox.pack(padx=10, pady=6)

generate_button = ttk.Button(root, text="Generate Selected Certificates", command=generate_selected)
generate_button.pack(pady=8)

root.mainloop()
