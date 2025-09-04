from openpyxl import load_workbook
from docx import Document
from datetime import datetime
import os
import re

TEMPLATE_FILE = "certificate.docx"
EXCEL_FILE = "internship_details.xlsx"
OUTPUT_DIR = "certificates"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- Helper functions ----------------

def format_date(date_value):
    if not date_value:
        return ""
    if isinstance(date_value, datetime):
        return date_value.strftime("%d-%b-%Y")
    date_str = str(date_value)
    date_str = re.sub(r"(st|nd|rd|th)", "", date_str)
    try:
        dt = datetime.strptime(date_str.strip(), "%d %B %Y")
        return dt.strftime("%d-%b-%Y")
    except:
        return date_value

def sanitize_filename(name):
    name_str = "NA" if name is None else str(name)
    return re.sub(r"[^A-Za-z0-9._-]", "_", name_str)

def build_mapping(row, headers):
    mapping = {}
    for hdr, val in zip(headers, row):
        if hdr:
            placeholder = "{" + hdr + "}"
            if hdr in ["Internship Start date", "Internship End date", "Date"]:
                val = format_date(val)
            mapping[placeholder] = "" if val is None else str(val)
    return mapping

# ---------------- Replacement function ----------------

def replace_placeholders_in_paragraph(paragraph, mapping):
    # Scan the full paragraph text
    text = "".join(run.text for run in paragraph.runs)
    new_text = ""
    i = 0
    while i < len(text):
        if text[i] == '{':
            j = i + 1
            while j < len(text) and text[j] != '}':
                j += 1
            if j < len(text):
                key = text[i:j+1]  # include braces
                replacement = mapping.get(key, key)  # leave as-is if not found
                new_text += replacement
                i = j + 1
            else:
                new_text += text[i]
                i += 1
        else:
            new_text += text[i]
            i += 1

    # Update first run with bold replaced values
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_text
        # Bold any replaced values
        for k, v in mapping.items():
            if v:
                first_run.text = first_run.text.replace(v, v)  # placeholder replaced with value
                first_run.bold = True  # only values replaced will appear bold
        for r in paragraph.runs[1:]:
            r.text = ""

def apply_mapping(doc, mapping):
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

# ---------------- Main Script ----------------

wb = load_workbook(EXCEL_FILE)
sheet = wb.active
headers = ["" if c.value is None else str(c.value).strip() for c in sheet[1]]

try:
    slno_index = headers.index("Sl.No")
except ValueError:
    slno_index = 0

for row in sheet.iter_rows(min_row=2, values_only=True):
    mapping = build_mapping(row, headers)
    slno_value = row[slno_index]
    filename = f"{sanitize_filename(slno_value)}.docx"
    out_path = os.path.join(OUTPUT_DIR, filename)

    # Duplicate template
    doc_copy = Document(TEMPLATE_FILE)
    doc_copy.save(out_path)

    # Open duplicate and apply mapping
    doc_to_edit = Document(out_path)
    apply_mapping(doc_to_edit, mapping)
    doc_to_edit.save(out_path)

    print(f"✅ Generated: {filename}")

print("🎉 All certificates generated successfully!")
