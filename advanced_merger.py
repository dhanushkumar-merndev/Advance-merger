import os
import pandas as pd
import json
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# ------------------ PATHS ------------------
INPUT_DIR = "input"
OUTPUT_DIR = "output"
DUPLICATE_DIR = os.path.join(OUTPUT_DIR, "Duplicated")
TEMPLATE_DIR = "templates"

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DUPLICATE_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)

# ---------------- FORMAT CODES ----------------
FORMAT_CODES = {
    "a": "DATE_DD-MM-YYYY",
    "b": "TIME_HH:MM",
    "c": "TIME_HH:MM:SS",
    "d": "LAST_10_DIGITS",
    "e": "ADD_+91",
    "f": "UPPER",
    "g": "LOWER",
    "h": "TITLE",
    "i": "INTEGER",
    "j": "TRIM_DASH",
    "u": "TRIM_UNDERSCORE",
    "x": "TRIM_DOT",
    "k": "DICT_LOOKUP"
}

ALIGN_CODES = {"l": "left", "r": "right"}

# ---------------- FILE READ ----------------
def read_file(path):
    if path.endswith(".xlsx"):
        df = pd.read_excel(path)
    else:
        try:
            df = pd.read_csv(path, sep="\t", encoding="utf-16")
        except:
            df = pd.read_csv(path, encoding="latin1")

    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .str.lower()
    )
    return df

# ---------------- FORMAT APPLY ----------------
def apply_format(val, code):
    try:
        if code == "a":
            return datetime.now().strftime("%d-%m-%Y")
        if code == "b":
            return datetime.now().strftime("%H:%M")
        if code == "c":
            return datetime.now().strftime("%H:%M:%S")
        if code == "d":
            digits = re.sub(r"\D", "", str(val))
            return digits[-10:] if len(digits) >= 10 else ""
        if code == "e":
            return "+91" + str(val)
        if code == "f":
            return str(val).upper()
        if code == "g":
            return str(val).lower()
        if code == "h":
            return str(val).title()
        if code == "i":
            digits = re.sub(r"\D", "", str(val))
            return int(digits) if digits else ""
        if code == "j":
            return str(val).replace("-", "")
        if code == "u":
            return str(val).replace("_", "")
        if code == "x":
            return str(val).replace(".", "")
    except:
        return ""
    return val

# ---------------- DICTIONARY INPUT ----------------
def read_dictionary_inline():
    print("\nEnter dictionary mapping (ENTER key to stop)")
    dk = {}
    while True:
        key = input("Key: ").strip().lower()
        if not key:
            break
        value = input(f"Value for '{key}': ").strip()
        dk[key] = value
    return dk

# ---------------- MENU ----------------
print("\n1. Use existing template")
print("2. Create new template")
print("3. Exit")

choice = input("\nSelect option: ").strip()
if choice == "3":
    exit()

# ---------------- INPUT SOURCE ----------------
dfs = []
column_index = {}
global_idx = 1

if choice == "2":
    print("\nSelect input source:")
    print("1. Use input folder")
    print("2. Use external Excel/CSV file")

    src = input("Choose (1/2): ").strip()

    if src == "2":
        path = input("Enter full file path: ").strip()
        df = read_file(path)
        dfs.append(df)
    else:
        for file in os.listdir(INPUT_DIR):
            if file.endswith((".csv", ".xlsx")):
                dfs.append(read_file(os.path.join(INPUT_DIR, file)))
else:
    for file in os.listdir(INPUT_DIR):
        if file.endswith((".csv", ".xlsx")):
            dfs.append(read_file(os.path.join(INPUT_DIR, file)))

# Build column index
for df in dfs:
    for col in df.columns:
        column_index[global_idx] = col
        global_idx += 1

merged = pd.concat(dfs, ignore_index=True)

# ---------------- SHOW COLUMNS (CREATE TEMPLATE) ----------------
if choice == "2":
    print("\nINPUT COLUMN INDEX\n")
    for i, c in column_index.items():
        print(f"{i:>3} → {c}")

# ---------------- TEMPLATE ----------------
if choice == "1":
    templates = os.listdir(TEMPLATE_DIR)
    for i, t in enumerate(templates, 1):
        print(f"{i}. {t}")
    tsel = int(input("\nSelect template number: "))
    with open(os.path.join(TEMPLATE_DIR, templates[tsel - 1])) as f:
        template = json.load(f)
else:
    print("\nFORMAT CODES:")
    for k, v in FORMAT_CODES.items():
        print(f"{k} → {v}")

    template = []
    while True:
        name = input("\nOutput column name (ENTER to finish): ").strip()
        if not name:
            break

        mapping = input("Mapping: ").split()
        rule = [name] + mapping

        if "k" in mapping:
            rule.append(read_dictionary_inline())

        template.append(rule)

    tname = input("\nSave template as (name.json): ")
    with open(os.path.join(TEMPLATE_DIR, tname), "w") as f:
        json.dump(template, f, indent=2)

# ---------------- OUTPUT FILE ----------------
output_name = input("\nEnter output file name (without .xlsx): ").strip() or "ADVANCED_MERGED_OUTPUT"
out_path = os.path.join(OUTPUT_DIR, f"{output_name}.xlsx")

# ---------------- APPLY TEMPLATE ----------------
output = {}
column_alignments = {}

for rule in template:
    col_name = rule[0]
    tokens = rule[1:]
    col_dict = {}
    align = "center"

    if isinstance(tokens[-1], dict):
        col_dict = tokens[-1]
        tokens = tokens[:-1]

    for t in tokens:
        if t in ALIGN_CODES:
            align = ALIGN_CODES[t]

    tokens = [t for t in tokens if t not in ALIGN_CODES]

    if tokens and tokens[0] in FORMAT_CODES:
        tokens = ["0"] + tokens

    values = []

    for _, row in merged.iterrows():
        val = ""
        ptr = 0

        if tokens[0] == "0":
            val = ""
        elif tokens[0].startswith("["):
            cols = list(map(int, tokens[0].strip("[]").split(",")))
            raw = [str(row[column_index[c]]).strip() for c in cols if c in column_index and str(row[column_index[c]]).strip()]
            val = " ".join(dict.fromkeys(raw))
            ptr = 1
        else:
            idx = int(tokens[0])
            val = row[column_index[idx]]
            ptr = 1

        for t in tokens[ptr:]:
            if t != "k":
                val = apply_format(val, t)

        if "k" in tokens and col_dict:
            matches = [v for k, v in col_dict.items() if k in str(val).lower()]
            val = ", ".join(dict.fromkeys(matches))

        values.append(val)

    output[col_name] = values
    column_alignments[col_name] = align

final_df = pd.DataFrame(output)

# ---------------- DEDUPLICATION ----------------
uniq = input("\nDo you need unique records? (y/n): ").lower().strip()
total_before = len(final_df)
duplicate_count = 0

if uniq == "y":
    print("\nSelect columns for duplicate check:")
    for i, col in enumerate(final_df.columns, 1):
        print(f"{i}. {col}")

    idxs = input("Enter column numbers: ")
    cols = [final_df.columns[int(i)-1] for i in idxs.split(",")]

    dup_mask = final_df.duplicated(subset=cols, keep="first")
    dup_df = final_df[dup_mask]
    duplicate_count = len(dup_df)
    final_df = final_df[~dup_mask]

    if not dup_df.empty:
        dup_df.to_excel(os.path.join(DUPLICATE_DIR, f"{output_name}_DUPLICATES.xlsx"), index=False)

# ---------------- SAVE & FORMAT ----------------
final_df.to_excel(out_path, index=False)
wb = load_workbook(out_path)
ws = wb.active
header_font = Font(bold=True)

for col in ws.columns:
    name = col[0].value
    align = column_alignments.get(name, "center")
    for cell in col:
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if cell.row == 1:
            cell.font = header_font
    ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value)) if c.value else 0 for c in col) + 4

wb.save(out_path)

# ---------------- SUMMARY ----------------
print("\n📊 MERGE SUMMARY")
print(f"Total rows before dedupe : {total_before}")
print(f"Unique rows kept         : {len(final_df)}")
print(f"Duplicates removed       : {duplicate_count}")
print(f"\n✅ Final output created: {out_path}")
