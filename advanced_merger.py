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
    "0": "BLANK/EMPTY",
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

# ---------------- GOOGLE SHEETS URL CONVERTER ----------------
def convert_google_sheets_url(url):
    """Convert Google Sheets sharing URL to export URL"""
    if "docs.google.com/spreadsheets" in url:
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            sheet_id = match.group(1)
            gid_match = re.search(r'[#&]gid=([0-9]+)', url)
            if gid_match:
                gid = gid_match.group(1)
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
            else:
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return url

# ---------------- FILE READ ----------------
def read_file(path):
    if "docs.google.com/spreadsheets" in path:
        path = convert_google_sheets_url(path)
        print(f"  → Converted to export URL")
    
    if path.startswith("http://") or path.startswith("https://"):
        df = pd.read_csv(path)
    elif path.endswith(".xlsx"):
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

# ---------------- LOAD FILES FUNCTION ----------------
def load_files():
    """Load files based on user choice"""
    dfs = []
    file_names = []
    print("\nSelect input source:")
    print("1. Use input folder")
    print("2. Use external Excel/CSV/Google Sheets file(s)")

    src = input("Choose (1/2): ").strip()

    if src == "2":
        print("\nEnter file paths or Google Sheets URLs (press ENTER without typing to finish):")
        print("Tip: For Google Sheets, paste the sharing link directly")
        file_count = 1
        while True:
            path = input(f"File {file_count} path: ").strip()
            if not path:
                break
            try:
                df = read_file(path)
                dfs.append(df)
                # Generate file name
                if "docs.google.com" in path:
                    fname = f"GoogleSheet_{file_count}"
                else:
                    fname = os.path.basename(path)
                file_names.append(fname)
                print(f"✓ Loaded: {len(df)} rows, {len(df.columns)} columns")
                file_count += 1
            except Exception as e:
                print(f"✗ Error loading file: {e}")
                print("  Please check the URL/path and try again, or press ENTER to skip")
    else:
        files = [f for f in os.listdir(INPUT_DIR) if f.endswith((".csv", ".xlsx"))]
        if not files:
            print(f"\n✗ No CSV or Excel files found in '{INPUT_DIR}' folder!")
            return [], []
        for file in files:
            df = read_file(os.path.join(INPUT_DIR, file))
            dfs.append(df)
            file_names.append(file)
            print(f"✓ Loaded: {file} ({len(df)} rows, {len(df.columns)} columns)")

    return dfs, file_names

# ---------------- MENU ----------------
print("\n" + "="*50)
print("   ADVANCED DATA MERGER")
print("="*50)
print("\n1. Use existing template")
print("2. Create new template")
print("3. Exit")

choice = input("\nSelect option: ").strip()
if choice == "3":
    exit()

# ---------------- LOAD INPUT FILES ----------------
dfs, file_names = load_files()

if not dfs:
    print("\n✗ No files loaded. Exiting.")
    exit()

# Build column index with both number and name
column_index = {}  # number -> column_name
column_list = []   # list of all unique columns
file_columns = {}  # file_name -> list of (col_num, col_name)
global_idx = 1

for df, fname in zip(dfs, file_names):
    file_columns[fname] = []
    for col in df.columns:
        column_index[global_idx] = col
        # Only add to column_list if it's the first occurrence of this column name
        if col not in column_list:
            column_list.append(col)
        file_columns[fname].append((global_idx, col))
        global_idx += 1

merged = pd.concat(dfs, ignore_index=True)

print(f"\n✓ Total rows merged: {len(merged)}")
print(f"✓ Total unique columns: {len(column_list)}")

# ---------------- SHOW COLUMNS (UNIQUE LIST) ----------------
if choice == "2":
    # Get terminal width (default 120 if can't detect)
    try:
        import shutil
        terminal_width = shutil.get_terminal_size().columns
    except:
        terminal_width = 120
    
    print("\n" + "="*terminal_width)
    print("INPUT COLUMNS (UNIQUE LIST)")
    print("="*terminal_width)
    
    # Build a consolidated view: column_name -> list of (file_name, col_num)
    column_to_files = {}
    first_occurrence = {}  # Track first column number for each unique column
    
    for fname, cols in file_columns.items():
        for col_num, col_name in cols:
            if col_name not in column_to_files:
                column_to_files[col_name] = []
                first_occurrence[col_name] = col_num  # Store first occurrence
            column_to_files[col_name].append((fname, col_num))
    
    # Display columns with their FIRST column number only
    print(f"\n{'No.':<6} │ {'Column Name':<45} │ {'Found In'}")
    print("-" * terminal_width)
    
    for col_name, file_info in column_to_files.items():
        # Show only the FIRST column number
        first_num = first_occurrence[col_name]
        
        # Get unique file names
        file_names_list = [fname[:20] for fname, num in file_info]
        unique_files = list(dict.fromkeys(file_names_list))
        
        # Build "Found In" text
        if len(file_info) == 1:
            found_in = unique_files[0]
        elif len(unique_files) == 1:
            found_in = f"{unique_files[0]} ({len(file_info)} times)"
        else:
            found_in = ", ".join(unique_files)
        
        # Truncate if too long
        col_name_display = col_name[:44]
        found_in_display = found_in[:terminal_width - 55] if terminal_width > 55 else found_in[:30]
        
        print(f"{first_num:<6} │ {col_name_display:<45} │ {found_in_display}")
    
    print("=" * terminal_width)
    print(f"\n💡 Tip: Use the column number shown above - it will access data from all occurrences")
    print()

# ---------------- TEMPLATE ----------------
if choice == "1":
    templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.json')]
    if not templates:
        print("\n✗ No templates found in templates folder!")
        print("Please create a template first using option 2.")
        exit()
    
    print("\n" + "="*50)
    print("AVAILABLE TEMPLATES")
    print("="*50)
    for i, t in enumerate(templates, 1):
        print(f"{i}. {t}")
    print("="*50)
    
    tsel = int(input("\nSelect template number: "))
    template_path = os.path.join(TEMPLATE_DIR, templates[tsel - 1])
    
    with open(template_path) as f:
        template = json.load(f)
    
    print(f"\n✓ Template loaded: {templates[tsel - 1]}")
    
    # Validate template columns
    missing_cols = []
    for rule in template:
        tokens = rule[1:]
        # Remove dict if present
        if isinstance(tokens[-1], dict):
            tokens = tokens[:-1]
        # Remove alignment codes
        tokens = [t for t in tokens if t not in ALIGN_CODES]
        
        for token in tokens:
            # Skip format codes and "0"
            if token in FORMAT_CODES or token == "0":
                continue
            # Check if it's a column list
            if token.startswith("[") and token.endswith("]"):
                cols = token.strip("[]").split(",")
                for col_name in cols:
                    col_name = col_name.strip()
                    if col_name not in column_list:
                        missing_cols.append(col_name)
            # Single column
            elif token not in column_list:
                missing_cols.append(token)
    
    if missing_cols:
        print("\n" + "="*50)
        print("❌ FAILED TO APPLY TEMPLATE")
        print("="*50)
        print("Reason: Column mismatch detected")
        print("\nMissing columns in your data:")
        for col in set(missing_cols):
            print(f"  ✗ {col}")
        print("\nPossible causes:")
        print("  - Wrong Excel/CSV file selected")
        print("  - Columns have been renamed or deleted")
        print("  - Template was created for different data")
        print("="*50)
        exit()
    
    print("✓ All template columns found in data")

else:  # Create new template
    print("\n" + "="*50)
    print("FORMAT CODES")
    print("="*50)
    for k, v in FORMAT_CODES.items():
        print(f"  {k} → {v}")
    print("\nALIGNMENT CODES")
    print("  l → left,  r → right")
    print("="*50)

    template = []
    while True:
        print("\n" + "-"*50)
        name = input("Output column name (ENTER to finish): ").strip()
        if not name:
            break

        print("\nMapping format:")
        print("  - Use 0 for blank/empty column")
        print("  - Enter column number(s) from the list above")
        print("  - Or use [col1,col2,col3] for multiple columns")
        print("  - Add format codes (a,b,c,d,e,f,g,h,i,j,u,x,k)")
        print("  - Add alignment (l or r)")
        print("Example 1: 0  (blank column)")
        print("Example 2: 5 d e  (column 5, last 10 digits, add +91)")
        print("Example 3: 0 a  (blank column with today's date)")
        
        mapping_input = input("\nMapping: ").split()
        
        # Convert column numbers to column names
        converted_mapping = []
        for token in mapping_input:
            # Check if it's "0" (blank column indicator)
            if token == "0":
                converted_mapping.append("0")
            # Check if it's a number
            elif token.isdigit():
                col_num = int(token)
                if col_num in column_index:
                    col_name = column_index[col_num]
                    # Check if this column name is already in converted_mapping
                    if col_name not in converted_mapping:
                        converted_mapping.append(col_name)
                    else:
                        print(f"ℹ Note: Column '{col_name}' already added (skipping duplicate)")
                else:
                    print(f"⚠ Warning: Column number {col_num} not found, skipping")
            # Check if it's a list of numbers [1,2,3]
            elif token.startswith("[") and token.endswith("]"):
                nums = token.strip("[]").split(",")
                col_names = []
                for num in nums:
                    if num.strip().isdigit():
                        col_num = int(num.strip())
                        if col_num in column_index:
                            col_name = column_index[col_num]
                            # Only add if not already in the list
                            if col_name not in col_names:
                                col_names.append(col_name)
                if col_names:
                    converted_mapping.append("[" + ",".join(col_names) + "]")
            else:
                # Keep format codes and alignment as-is
                converted_mapping.append(token)

        rule = [name] + converted_mapping

        if "k" in converted_mapping:
            rule.append(read_dictionary_inline())

        template.append(rule)
        print(f"✓ Added column: {name}")

    if not template:
        print("\n✗ No columns added. Exiting.")
        exit()

    tname = input("\nSave template as (name.json): ").strip()
    if not tname.endswith('.json'):
        tname += '.json'
    
    template_path = os.path.join(TEMPLATE_DIR, tname)
    with open(template_path, "w") as f:
        json.dump(template, f, indent=2)
    print(f"\n✓ Template saved: {tname}")

# ---------------- OUTPUT FILE ----------------
output_name = input("\nEnter output file name (without .xlsx): ").strip() or "ADVANCED_MERGED_OUTPUT"
out_path = os.path.join(OUTPUT_DIR, f"{output_name}.xlsx")

# ---------------- APPLY TEMPLATE ----------------
print("\n⚙ Processing data...")

output = {}
column_alignments = {}

for rule in template:
    col_name = rule[0]
    tokens = rule[1:]
    col_dict = {}
    align = "center"

    # Extract dictionary if present
    if isinstance(tokens[-1], dict):
        col_dict = tokens[-1]
        tokens = tokens[:-1]

    # Extract alignment
    for t in tokens:
        if t in ALIGN_CODES:
            align = ALIGN_CODES[t]

    tokens = [t for t in tokens if t not in ALIGN_CODES]

    # Add "0" prefix for format-only columns
    if tokens and tokens[0] in FORMAT_CODES:
        tokens = ["0"] + tokens

    values = []

    for _, row in merged.iterrows():
        val = ""
        ptr = 0

        if not tokens or tokens[0] == "0":
            val = ""
        elif tokens[0].startswith("["):
            # Multiple columns
            col_names = tokens[0].strip("[]").split(",")
            raw = []
            for cn in col_names:
                cn = cn.strip()
                if cn in merged.columns and str(row[cn]).strip():
                    raw.append(str(row[cn]).strip())
            val = " ".join(dict.fromkeys(raw))
            ptr = 1
        else:
            # Single column
            col_name_token = tokens[0]
            if col_name_token in merged.columns:
                val = row[col_name_token]
            ptr = 1

        # Apply format codes
        for t in tokens[ptr:]:
            if t != "k":
                val = apply_format(val, t)

        # Apply dictionary lookup
        if "k" in tokens and col_dict:
            matches = [v for k, v in col_dict.items() if k in str(val).lower()]
            val = ", ".join(dict.fromkeys(matches))

        values.append(val)

    output[col_name] = values
    column_alignments[col_name] = align

final_df = pd.DataFrame(output)
print(f"✓ Processed {len(final_df)} rows")

# ---------------- DEDUPLICATION ----------------
uniq = input("\nDo you need unique records? (y/n): ").lower().strip()
total_before = len(final_df)
duplicate_count = 0

if uniq == "y":
    print("\nSelect columns for duplicate check:")
    for i, col in enumerate(final_df.columns, 1):
        print(f"{i}. {col}")

    idxs = input("\nEnter column numbers (comma-separated): ").strip()
    cols = [final_df.columns[int(i)-1] for i in idxs.split(",")]

    dup_mask = final_df.duplicated(subset=cols, keep="first")
    dup_df = final_df[dup_mask]
    duplicate_count = len(dup_df)
    final_df = final_df[~dup_mask]

    if not dup_df.empty:
        dup_path = os.path.join(DUPLICATE_DIR, f"{output_name}_DUPLICATES.xlsx")
        dup_df.to_excel(dup_path, index=False)
        print(f"✓ Duplicates saved: {dup_path}")

# ---------------- SAVE & FORMAT ----------------
print("\n⚙ Formatting output file...")
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
print("\n" + "="*50)
print("📊 MERGE SUMMARY")
print("="*50)
print(f"Total rows before dedupe : {total_before}")
print(f"Unique rows kept         : {len(final_df)}")
print(f"Duplicates removed       : {duplicate_count}")
print(f"\n✅ Final output created: {out_path}")
print("="*50)