import os
import pandas as pd
import json
import re
import sys
import shutil
import io
import time
import urllib.request
from datetime import datetime
from openpyxl.styles import Alignment, Font

INPUT_DIR = "input"
OUTPUT_DIR = "output"
TEMPLATE_DIR = "templates"
CAMP_DIR = "campaigns"
SOURCE_GRP_DIR = "source_groups"
BASELINE_DIR = os.path.join(TEMPLATE_DIR, "baselines")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(os.path.join(OUTPUT_DIR, "Duplicated"), exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)
os.makedirs(CAMP_DIR, exist_ok=True)
os.makedirs(SOURCE_GRP_DIR, exist_ok=True)
os.makedirs(BASELINE_DIR, exist_ok=True)

FORMAT_CODES = {"0","a","b","c","d","e","f","g","h","i","j","u","x","k","q"}
ALIGN_CODES = {"l": "left", "r": "right", "z": "center"}
ALL_CODES = FORMAT_CODES | set(ALIGN_CODES.keys())

def find_template(tname):
    candidates = [tname, tname.replace(" ", "_"), tname.strip(), tname.strip().replace(" ", "_")]
    for c in candidates:
        p = os.path.join(TEMPLATE_DIR, c + ".json")
        if os.path.exists(p):
            return p
    available = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.json')]
    tname_lower = tname.lower().replace(" ", "").replace("_", "")
    for f in available:
        f_lower = f.replace(".json","").lower().replace(" ", "").replace("_", "")
        if f_lower == tname_lower:
            return os.path.join(TEMPLATE_DIR, f)
    return None

def convert_google_sheets_url(url):
    if "docs.google.com/spreadsheets" in url:
        if "export?format=csv" in url:
            return url
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            sheet_id = match.group(1)
            gid_match = re.search(r'[#&?]gid=([0-9]+)', url)
            gid = gid_match.group(1) if gid_match else "0"
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
            print(f"  Sheet  →  {csv_url}")
            return csv_url
    return url

def read_file(path):
    if "docs.google.com/spreadsheets" in path:
        path = convert_google_sheets_url(path)

    if path.startswith("http://") or path.startswith("https://"):
        last_err = None
        for attempt in range(3):
            try:
                req = urllib.request.Request(path, headers={
                    'User-Agent': 'Mozilla/5.0',
                    'Accept': 'text/csv,*/*'
                })
                with urllib.request.urlopen(req, timeout=30) as resp:
                    content = resp.read().decode('utf-8', errors='replace')
                df = pd.read_csv(io.StringIO(content))
                # ← return is here, inside the try after successful read
                df.columns = (
                    df.columns.astype(str)
                    .str.replace("\ufeff", "", regex=False)
                    .str.strip()
                    .str.lower()
                )
                print(f"  Fetched  {len(df)} rows")
                return df
            except Exception as e:
                last_err = e
                print(f"  Fetch attempt {attempt+1} failed: {e}")
                time.sleep(2)
        raise Exception(f"Failed to fetch after 3 attempts: {last_err}")

    elif path.endswith(".xlsx"):
        df = pd.read_excel(path)
    else:
        try:
            df = pd.read_csv(path, sep="\t", encoding="utf-16")
        except:
            try:
                df = pd.read_csv(path, encoding="utf-8")
            except:
                df = pd.read_csv(path, encoding="latin1")

    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .str.lower()
    )
    return df

def apply_format_series(series, code):
    try:
        if code == "a":
            return pd.Series([datetime.now().strftime("%d-%m-%Y")] * len(series), index=series.index)
        if code == "b":
            return pd.Series([datetime.now().strftime("%H:%M")] * len(series), index=series.index)
        if code == "c":
            return pd.Series([datetime.now().strftime("%H:%M:%S")] * len(series), index=series.index)
        if code == "d":
            return series.astype(str).str.replace(r"\D", "", regex=True).str[-10:]
        if code == "e":
            return "+91" + series.astype(str)
        if code == "f":
            return series.astype(str).str.upper()
        if code == "g":
            return series.astype(str).str.lower()
        if code == "h":
            return series.astype(str).str.title()
        if code == "i":
            return pd.to_numeric(series.astype(str).str.replace(r"\D", "", regex=True), errors='coerce').fillna(0).astype(int)
        if code == "j":
            return series.astype(str).str.replace("-", "", regex=False)
        if code == "u":
            return series.astype(str).str.replace("_", "", regex=False)
        if code == "x":
            return series.astype(str).str.replace(".", "", regex=False)
    except:
        return series
    return series

def process_df(df_list, template, output_name, quick_mode, template_unique_cols, sort_col=None, sort_order="asc"):
    merged = pd.concat(df_list, ignore_index=True)
    output = {}
    column_alignments = {}

    for rule in template:
        col_name = rule[0]
        tokens = list(rule[1:])
        col_dict = {}
        align = "center" # Default

        if tokens and isinstance(tokens[-1], dict):
            col_dict = tokens[-1]
            tokens = tokens[:-1]

        for t in tokens:
            if t in ALIGN_CODES:
                align = ALIGN_CODES[t]

        if not tokens:
            s = pd.Series([""] * len(merged), index=merged.index)
            fmt_tokens = []
        elif tokens[0] == "0":
            s = pd.Series([""] * len(merged), index=merged.index)
            fmt_tokens = tokens[1:]
        elif tokens[0] in ALL_CODES:
            s = pd.Series([""] * len(merged), index=merged.index)
            fmt_tokens = tokens
        elif tokens[0].startswith("["):
            col_names = [cn.strip().lower() for cn in tokens[0].strip("[]").split(",")]
            valid_cols = [cn for cn in col_names if cn in merged.columns]
            if valid_cols:
                s = merged[valid_cols].astype(str).apply(
                    lambda x: " ".join(dict.fromkeys(v for v in x if v not in ("", "nan"))), axis=1
                )
            else:
                s = pd.Series([""] * len(merged), index=merged.index)
            fmt_tokens = tokens[1:]
        else:
            src_lower = tokens[0].lower()
            s = merged[src_lower].fillna("") if src_lower in merged.columns else pd.Series([""] * len(merged), index=merged.index)
            fmt_tokens = tokens[1:]

        for t in fmt_tokens:
            if t in ALIGN_CODES or t in ("k", "q"):
                continue
            s = apply_format_series(s, t)

        if ("k" in fmt_tokens or "q" in fmt_tokens) and col_dict:
            use_default = "q" in fmt_tokens
            def lookup(val):
                norm_val = re.sub(r"[\s_\-]", "", str(val).lower())
                matches = [v for k, v in col_dict.items() if k != "__default__" and re.sub(r"[\s_\-]", "", str(k).lower()) in norm_val]
                if matches:
                    return ", ".join(dict.fromkeys(matches))
                if use_default and "__default__" in col_dict:
                    return col_dict["__default__"]
                return "" if not use_default else val
            s = s.map(lookup)

        output[col_name] = s
        column_alignments[col_name] = align

    final_df = pd.DataFrame(output)
    selected_unique_cols = [c for c in template_unique_cols if c in final_df.columns] if quick_mode else []
    if selected_unique_cols:
        final_df = final_df.drop_duplicates(subset=selected_unique_cols, keep="first")
    if sort_col and sort_col in final_df.columns:
        is_asc = (str(sort_order).lower() != "desc")
        final_df = final_df.sort_values(by=sort_col, ascending=is_asc)
    return final_df, column_alignments

def style_sheet(ws, aligns):
    header_font = Font(bold=True)
    for col in ws.columns:
        name = col[0].value
        h_align = aligns.get(name, "center")
        for cell in col:
            cell.alignment = Alignment(horizontal=h_align, vertical="center")
            if cell.row == 1:
                cell.font = header_font
        ws.column_dimensions[col[0].column_letter].width = max(
            len(str(c.value)) if c.value else 0 for c in col
        ) + 4

def run_campaign(config, output_name, incremental=False):
    ts = datetime.now().strftime("%y%m%d_%H%M")
    out_path = os.path.join(OUTPUT_DIR, f"{output_name}_{ts}.xlsx")
    baseline_path = os.path.join(BASELINE_DIR, f"{output_name}.xlsx")

    baseline_sheets = {}
    if incremental and os.path.exists(baseline_path):
        try:
            baseline_sheets = pd.read_excel(baseline_path, sheet_name=None)
            total_bl = sum(len(v) for v in baseline_sheets.values())
            print(f"Baseline loaded  {total_bl} rows")
        except:
            print("Baseline unreadable, running full merge.")

    writer = pd.ExcelWriter(out_path, engine='openpyxl')
    sheets_added = 0
    total_new = 0

    for group in config.get("groups", []):
        gname = group.get("name", "Sheet")
        sources = group.get("sources", [])
        tname = group.get("template", "")

        tpl_path = find_template(tname)
        if not tpl_path:
            print(f"Template not found  '{tname}'")
            print(f"  Available: {[f.replace('.json','') for f in os.listdir(TEMPLATE_DIR) if f.endswith('.json')]}")
            continue

        with open(tpl_path, encoding='utf-8') as f:
            tdata = json.load(f)
            template = tdata["columns"] if isinstance(tdata, dict) else tdata
            u_cols = tdata.get("unique_columns", []) if isinstance(tdata, dict) else []
            s_col = tdata.get("sort_column") if isinstance(tdata, dict) else None
            s_order = tdata.get("sort_order", "asc") if isinstance(tdata, dict) else "asc"

        print(f"Processing  {gname}...")
        try:
            group_dfs = [read_file(s) for s in sources if s]
            if not group_dfs:
                print(f"  No sources loaded for {gname}")
                continue
        except Exception as e:
            print(f"  Source load failed: {e}")
            continue

        df, aligns = process_df(group_dfs, template, output_name, True, u_cols, sort_col=s_col, sort_order=s_order)

        if incremental and gname in baseline_sheets and u_cols:
            b_df = baseline_sheets[gname].copy()
            b_df.columns = b_df.columns.astype(str).str.strip()
            valid_u_cols = [c for c in u_cols if c in df.columns and c in b_df.columns]
            if valid_u_cols:
                df_key = df[valid_u_cols].astype(str).apply(lambda x: x.str.strip().str.lower())
                b_key  = b_df[valid_u_cols].astype(str).apply(lambda x: x.str.strip().str.lower())
                df_key_str = df_key.apply(lambda r: "|".join(r.values), axis=1)
                b_key_str  = b_key.apply(lambda r: "|".join(r.values), axis=1)
                df = df[~df_key_str.isin(set(b_key_str))].reset_index(drop=True)

        row_count = len(df)
        total_new += row_count
        print(f"  {row_count} new rows  ->  {gname}")

        sheet_name = gname[:30]
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        style_sheet(writer.sheets[sheet_name], aligns)
        sheets_added += 1

    if sheets_added == 0:
        writer.close()
        try:
            os.remove(out_path)
        except:
            pass
        print("No sheets written.")
        return None, baseline_path

    writer.close()
    print(f"Done  {out_path}  ({total_new} rows)")
    return out_path, baseline_path

# ── CLI MODE ──────────────────────────────────────────────────────────────────
if len(sys.argv) > 2:
    mode = sys.argv[1]
    config = json.loads(sys.argv[2])
    output_name = config.get("output", "MERGED_OUTPUT")
    incremental = config.get("incremental", False)
    if mode == "4":
        run_campaign({"groups": config.get("groups", [])}, output_name, incremental)
    exit()

# ── INTERACTIVE MODE ──────────────────────────────────────────────────────────
print("\n=== DATA MERGER ===")
print("1. Standard Merge")
print("2. Campaign Mode")
print("3. Create Template")
print("4. Exit")
choice = input("\nSelect option: ").strip()
if choice == "4" or not choice:
    exit()

if choice == "2":
    camps = sorted([f for f in os.listdir(CAMP_DIR) if f.endswith('.json')], key=str.lower)
    if not camps:
        print("No campaigns found.")
        exit()
    print("\nCAMPAIGNS:")
    [print(f"{i+1}. {c.replace('.json','')}") for i, c in enumerate(camps)]
    sel = int(input("Select: "))
    with open(os.path.join(CAMP_DIR, camps[sel-1]), encoding='utf-8') as f:
        config = json.load(f)
    output_name = config.get("name", "CAMPAIGN_OUTPUT")
    print("\nMode:")
    print("1. Full merge  (all rows)")
    print("2. Incremental  (only new rows vs baseline)")
    inc_choice = input("Choose (1/2): ").strip()
    incremental = (inc_choice == "2")
    groups = []
    for gname in config.get("groups", []):
        g_path = os.path.join(SOURCE_GRP_DIR, gname + ".json")
        if not os.path.exists(g_path):
            g_path = os.path.join(SOURCE_GRP_DIR, gname.replace(" ", "_") + ".json")
        if not os.path.exists(g_path):
            print(f"Source group not found  {gname}")
            continue
        with open(g_path, encoding='utf-8') as f:
            gdata = json.load(f)
        groups.append({
            "name": gname,
            "sources": [s.get("path") for s in gdata.get("sources", [])],
            "template": gdata.get("templateName", "")
        })
    out_path, baseline_path = run_campaign({"groups": groups}, output_name, incremental)
    if out_path:
        save_bl = input("\nSave baseline? (y/n): ").strip().lower()
        if save_bl == "y":
            shutil.copy2(out_path, baseline_path)
            print(f"Baseline saved  {baseline_path}")
    exit()

dfs = []
print("\nInput source:")
print("1. input/ folder")
print("2. External files/URLs")
src = input("Choose (1/2): ").strip()
if src == "2":
    while True:
        path = input(f"File {len(dfs)+1} (ENTER to finish): ").strip()
        if not path:
            break
        try:
            df = read_file(path)
            dfs.append(df)
            print(f"Loaded  {len(df)} rows")
        except Exception as e:
            print(f"Failed  {e}")
else:
    for f in sorted([f for f in os.listdir(INPUT_DIR) if f.endswith((".csv", ".xlsx"))]):
        df = read_file(os.path.join(INPUT_DIR, f))
        dfs.append(df)
        print(f"Loaded  {f}")

if not dfs:
    print("No files loaded.")
    exit()

merged = pd.concat(dfs, ignore_index=True)
column_index = {i+1: c for i, c in enumerate([c for df in dfs for c in df.columns])}

if choice == "3":
    print("\nCOLUMNS:")
    [print(f"{i:2}. {c}") for i, c in column_index.items()]
    template = []
    while True:
        name = input("\nOutput column name (ENTER to finish): ").strip()
        if not name:
            break
        mapping = input("Mapping (e.g. 5 d e): ").split()
        rule = [name]
        for t in mapping:
            if t.isdigit() and int(t) in column_index:
                rule.append(column_index[int(t)])
            elif t.startswith("["):
                rule.append("[" + ",".join([column_index[int(n)] for n in t.strip("[]").split(",") if n.strip().isdigit()]) + "]")
            else:
                rule.append(t)
        template.append(rule)
        print(f"Added  {name}")
    s_col = input("\nSort by column name (ENTER for none): ").strip()
    tname = input("Save template as: ").strip()
    if not tname.endswith('.json'):
        tname += '.json'
    with open(os.path.join(TEMPLATE_DIR, tname), "w", encoding='utf-8') as f:
        json.dump({"columns": template, "unique_columns": [], "sort_column": s_col}, f, indent=2)
    print(f"Saved  {tname}")
    exit()

if choice == "1":
    templates = sorted([f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.json')], key=str.lower)
    [print(f"{i+1}. {t}") for i, t in enumerate(templates)]
    tsel = int(input("Select template: "))
    with open(os.path.join(TEMPLATE_DIR, templates[tsel-1]), encoding='utf-8') as f:
        tdata = json.load(f)
        template = tdata["columns"] if isinstance(tdata, dict) else tdata
        u_cols = (tdata.get("unique_columns") or tdata.get("dedupCols") or []) if isinstance(tdata, dict) else []
        s_col = tdata.get("sort_column") if isinstance(tdata, dict) else None
        s_order = (tdata.get("sort_order") or "asc") if isinstance(tdata, dict) else "asc"
    print("\nMode:")
    print("1. Quick  (auto dedup + auto filename)")
    print("2. Advanced  (custom name & dedup)")
    p_mode = input("Choose (1/2): ").strip()
    out_name = os.path.splitext(templates[tsel-1])[0]
    if p_mode == "2":
        out_name = input(f"Output filename [{out_name}]: ").strip() or out_name
        print("\nDedup:")
        print("1. Saved columns")
        print("2. Custom columns")
        print("3. Skip")
        d_choice = input("Select (1/2/3): ").strip()
        if d_choice == "2":
            u_cols = [c.strip() for c in input("Columns (comma separated): ").split(",") if c.strip()]
        elif d_choice == "3":
            u_cols = []
    print("Processing...")
    df, aligns = process_df(dfs, template, out_name, True, u_cols, sort_col=s_col, sort_order=s_order)
    ts = datetime.now().strftime("%y%m%d_%H%M")
    out_path = os.path.join(OUTPUT_DIR, f"{out_name}_{ts}.xlsx")
    df.to_excel(out_path, index=False)
    print(f"Done  {out_path}")