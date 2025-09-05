import os
import shutil
import re
from docx import Document
from openpyxl import load_workbook

# ---------------- CONFIG ----------------
TEMPLATE_FILENAME = "certificate.docx"
EXCEL_FILENAME = "internship_details.xlsx"
OUTPUT_DIRNAME = "certificates"

# Paths relative to current directory
current_directory = os.getcwd()
template_path = os.path.join(current_directory, TEMPLATE_FILENAME)
excel_path = os.path.join(current_directory, EXCEL_FILENAME)
output_dir = os.path.join(current_directory, OUTPUT_DIRNAME)

os.makedirs(output_dir, exist_ok=True)

# ---------------- HELPERS ----------------

PLACEHOLDER_RE = re.compile(r'\{\s*([^}]+?)\s*\}')

def copy_formatting(target_run, source_run):
    """Copy simple formatting from one run to another."""
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
    """Replace placeholders {Key} with values (bold) in a paragraph."""
    if not paragraph.runs:
        return
    full_text = ''.join(run.text for run in paragraph.runs)
    if "{" not in full_text:
        return
    matches = list(PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return

    # Build run_map: for each char position, which run index it came from
    run_map = []
    for idx, run in enumerate(paragraph.runs):
        text = run.text or ""
        run_map.extend([idx] * len(text))

    # Segments: (type, text, pos)
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

    # Clear old runs
    for run in paragraph.runs:
        run.text = ""

    # Insert new runs
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
    """Replace placeholders everywhere in the document."""
    # Paragraphs
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    # Headers & Footers
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

# ---------------- MAIN ----------------

# Load Excel
wb = load_workbook(excel_path, data_only=True)
sheet = wb.active
headers = [str(cell.value).strip() if cell.value else "" for cell in sheet[1]]

# Find Sl.No column (fallback to first column)
try:
    slno_index = headers.index("Sl.No")
except ValueError:
    try:
        slno_index = headers.index("Sl No")
    except ValueError:
        slno_index = 0

for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
    sl = row[slno_index]
    if not sl or str(sl).strip() == "":
        print(f"Skipping row {row_idx} (empty Sl.No)")
        continue

    slno = str(sl).strip()
    output_file = os.path.join(output_dir, f"{slno}.docx")

    # Copy template
    shutil.copy2(template_path, output_file)
    doc = Document(output_file)

    # Build mapping (exclude Sl.No)
    mapping = {}
    for i, h in enumerate(headers):
        key = str(h).strip() if h else ""
        if key == "Sl.No" or key == "Sl No":
            continue
        val = row[i] if i < len(row) else None
        mapping[key] = "" if val is None else str(val)

    # Debug print
    print(f"\nRow {row_idx} -> Sl.No {slno} mapping:")
    for k, v in mapping.items():
        print(f"  {k} -> {v}")

    # Replace placeholders
    process_document(doc, mapping)

    # Save
    doc.save(output_file)
    print(f"Generated: {output_file}")

print("\nAll certificates generated successfully!")
