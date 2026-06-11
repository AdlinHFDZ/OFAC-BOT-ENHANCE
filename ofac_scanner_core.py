# ofac_scanner_core.py
"""
OFAC Scanner Core – fully fixed with:
- no age‑based discarding (minors kept)
- missing‑name counting
- CSV splitting for very large outputs
- content‑based column detection
- detailed debug logging
- output folder and log file named with YYYYDDMM format
- BUGFIX: large text file slicing – no rows lost after first chunk
- NEW: multiple‑table detection in a single Excel sheet (via column clusters)
- NEW: numeric column filter – numeric IDs are no longer treated as names
- FIX: column renaming before numeric filter so name columns work correctly
  (all original multi‑name, shift‑adjustment, etc. preserved)
"""

import os, sys, csv, json, subprocess, time, shutil, re, io, difflib
from datetime import datetime
from ctypes import WinDLL

import polars as pl
import msoffcrypto
import chardet
import clevercsv
from dateutil import parser as dateutil_parser
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ==================== CONSTANTS ====================
ENV_VAR_FOLDER = "OFAC_INPUT_FOLDER"
ENV_VAR_CSV = "OFAC_COMPANY_PASSWORD_CSV"
APP_NAME = "OFAC_Scanner_GUI"
SCRIPT_PATH = os.path.abspath(__file__)


COMPANY_HEADERS_PATH = os.environ.get(
    "OFAC_HEADERS_PATH",
    r"C:\Users\s0055198\OneDrive - RGA Reinsurance Company\Documents\OFAC CODE\header")
BASE_OUTPUT_FOLDER = os.environ.get(
    "OFAC_OUTPUT_FOLDER",
    r"C:\Users\s0055198\OneDrive - RGA Reinsurance Company\Documents\OFAC CODE\output")
SEVEN_ZIP_PATH = os.environ.get(
    "SEVEN_ZIP_PATH", r"C:\Program Files\7-zip\7z.exe")

FILE_EXTENSIONS_EXCEL = ["xlsx", "xls", "xlsm", "xlsb"]
FILE_EXTENSIONS_TEXT = ["csv", "txt", "rpt"]
FILE_EXTENSIONS_ARCHIVE = ["zip", "zipx", "tar", "7z", "rar"]
EXCEL_MAX_ROWS = 1_000_000
SIZE_TO_CHUNK = 100 * 1024 * 1024
EXCEL_SIZE_TO_CHUNK = 200 * 1024 * 1024
CHUNK_SIZE = 10000
MAX_SEARCH_ROWS = 50
EXCEL_ROW_CHUNK = 50000

# Max rows per single output CSV before splitting (0 = no splitting)
MAX_ROWS_PER_OUTPUT_CSV = 500_000

OUTPUT_COLUMNS = ["SURNAME", "FIRST_NAME", "COMPLETE_NAME", "SEX", "DATE_OF_BIRTH",
                  "CMPY_NO", "POLICY_NUMBER", "FILE_PATH", "SHEET"]

LOG_HEADERS = ["File Path", "File Name", "Scan Date", "Extension", "Company Code",
               "Password", "Sheet Name", "Error Msg", "Identified Headers",
               "Multiple Name", "Row Count", "Output Row Count", "Output CSV",
               "First Last Name Header", "Full Name Header", "Policy Number Header",
               "DOB Header", "Sex Header", "Remarks"]

# ==================== PERSISTENT SETTINGS ====================
_settings_file_path = None

def set_settings_file(path):
    global _settings_file_path
    _settings_file_path = path

def load_settings():
    if _settings_file_path and os.path.exists(_settings_file_path):
        try:
            with open(_settings_file_path, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_settings(settings_dict):
    if _settings_file_path:
        with open(_settings_file_path, 'w') as f:
            json.dump(settings_dict, f, indent=4)

# ==================== HELPERS ====================
def instance_check_or_focus(window_title):
    user32 = WinDLL("user32", use_last_error=True)
    hwnd = user32.FindWindowW(None, window_title)
    if hwnd:
        user32.ShowWindow(hwnd, 5)
        user32.SetForegroundWindow(hwnd)
        sys.exit(0)
    return True

def set_user_env_var(name, value):
    try:
        import win32api, win32con
        key = __import__("winreg").OpenKey(__import__("winreg").HKEY_CURRENT_USER,
                                          r"Environment", 0, __import__("winreg").KEY_SET_VALUE)
        __import__("winreg").SetValueEx(key, name, 0, __import__("winreg").REG_SZ, value)
        __import__("winreg").CloseKey(key)
        win32api.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "Environment")
        return True
    except:
        return False

def get_env_var(name):
    val = os.environ.get(name)
    if val:
        return val
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ)
        val, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return val
    except:
        return None

def get_unique_save_path(file_path):
    folder = os.path.dirname(file_path)
    base, ext = os.path.splitext(os.path.basename(file_path))
    safe_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', base)[:200]
    new_path = os.path.join(folder, f"{safe_name}{ext}")
    count = 1
    while os.path.exists(new_path):
        new_path = os.path.join(folder, f"{safe_name}_{count}{ext}")
        count += 1
    return new_path

def get_all_files(directory):
    return {os.path.join(root, f) for root, _, files in os.walk(directory) for f in files}

def move_file_to_archived(src, archived_folder, retries=3):
    dest = get_unique_save_path(os.path.join(archived_folder, os.path.basename(src)))
    for attempt in range(retries):
        try:
            shutil.move(src, dest)
            return dest
        except (PermissionError, OSError):
            time.sleep(10)
    raise Exception(f"Failed to move {src} after {retries} attempts")

# ==================== SCANNER JOB ====================
class ScannerJob:
    def __init__(self, input_folder, company_code, passwords, email_received_date, file_names, email=None):
        self.input_folder = input_folder
        self.company_code = company_code
        self.passwords = passwords
        self.email_received_date = email_received_date   # original YYYY-MM-DD
        self.file_names = file_names
        self.email = email
        self.today = datetime.now().strftime("%Y%m%d")

        # Convert GUI date (YYYY-MM-DD) to YYYYDDMM for folder/file naming
        try:
            dt = datetime.strptime(email_received_date, "%Y-%m-%d")
            self.date_display = dt.strftime("%Y%m%d")   # YYYYDDMM
        except:
            # fallback: use as‑is with dashes removed
            self.date_display = email_received_date.replace("-", "")

        # Output paths now use YYYYDDMM format
        self.output_root = os.path.join(BASE_OUTPUT_FOLDER, self.date_display)
        self.csv_folder = os.path.join(self.output_root, "CSVs")
        self.archived_folder = os.path.join(self.output_root, "Archived")
        self.unzipped_folder = os.path.join(self.output_root, "Unzipped")
        self.compiled_folder = os.path.join(self.output_root, "Compiled")
        self.log_file = os.path.join(self.output_root, f"Log_{self.date_display}.csv")

        for d in [self.csv_folder, self.archived_folder, self.unzipped_folder, self.compiled_folder]:
            os.makedirs(d, exist_ok=True)

# ==================== HEADER HELPERS ====================
def get_company_header(company_code):
    sheets = ["name", "firstlastname", "sex", "dob", "policynum"]
    company_path = os.path.join(COMPANY_HEADERS_PATH, "by_company", f"header_{company_code.upper().strip()}.xlsx")
    default_path = os.path.join(COMPANY_HEADERS_PATH, "header.xlsx")
    file_path = company_path if os.path.exists(company_path) else default_path
    header_dict = {}
    for sheet in sheets:
        try:
            df = pl.read_excel(file_path, sheet_name=sheet, has_header=False, raise_if_empty=False)
            if df.is_empty() or len(df.columns) == 0:
                header_dict[sheet] = set()
                continue
            first_col = df.columns[0]
            values = df.select(pl.col(first_col).str.strip_chars().str.to_lowercase().str.replace_all(r'[^a-z]', '')).to_series().to_list()
            header_dict[sheet] = set(values)
        except:
            header_dict[sheet] = set()
    return header_dict, file_path

def clean_for_match(text):
    if text is None:
        return ''
    return re.sub(r'[^a-z]', '', str(text).lower())

def normalize_sex(value):
    if value is None:
        return ''
    v = str(value).strip().lower()
    if v in ('male', 'm'):
        return 'M'
    if v in ('female', 'f'):
        return 'F'
    return v.upper()

def parse_date_to_mmddyyyy(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%m/%d/%Y')
    s = str(value).strip()
    if not s:
        return ''
    try:
        dt = dateutil_parser.parse(s, fuzzy=False)
        return dt.strftime('%m/%d/%Y')
    except:
        pass
    formats = ['%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%b-%Y', '%d-%b-%y', '%b %d, %Y', '%Y%m%d']
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%m/%d/%Y')
        except:
            continue
    return ''

def clean_string_expr(col_name):
    return pl.col(col_name).cast(pl.Utf8).str.strip_chars().str.replace_all(r'[^a-zA-Z0-9\s]', '').fill_null('')

# --- Numeric column filter ---
def is_column_numeric(series: pl.Series, sample_size=100, threshold=0.9) -> bool:
    if len(series) == 0:
        return False
    sample = series.drop_nulls().head(sample_size)
    if len(sample) == 0:
        return False
    digit_count = sample.cast(pl.Utf8).str.contains(r'^\d+$').sum()
    return (digit_count / len(sample)) >= threshold

def filter_numeric_columns(df: pl.DataFrame, col_names, threshold=0.9):
    if not col_names:
        return col_names
    return [c for c in col_names if c in df.columns and not is_column_numeric(df[c], threshold=threshold)]

# --- Duplicate filter ---
def filter_duplicates_only(df):
    if df.is_empty():
        return df
    return df.unique()

# ==================== SMART DETECTION ====================
def fuzzy_match_word(word, candidates, threshold=0.7):
    for cand in candidates:
        if difflib.SequenceMatcher(None, word, cand).ratio() >= threshold:
            return True
    return False

def infer_column_type_by_content(df: pl.DataFrame, col_index: int) -> str:
    if df.is_empty() or col_index >= df.width:
        return ''
    col = df[:, col_index]
    non_null = col.drop_nulls()
    sample = non_null.head(200).to_list()
    if not sample:
        return ''
    str_vals = [str(v).strip() for v in sample if str(v).strip() != '']
    if not str_vals:
        return ''

    sex_indicators = {'m', 'f', 'male', 'female'}
    lower_vals = [v.lower() for v in str_vals]
    sex_count = sum(1 for v in lower_vals if v in sex_indicators)
    if sex_count / len(lower_vals) > 0.7:
        return 'sex'

    date_count = 0
    for v in str_vals:
        if parse_date_to_mmddyyyy(v) != '':
            date_count += 1
    if date_count / len(str_vals) > 0.7:
        return 'dob'

    policy_pattern = re.compile(r'^[A-Za-z0-9\-./]{6,30}$')
    policy_count = sum(1 for v in str_vals if policy_pattern.match(v))
    if policy_count / len(str_vals) > 0.7:
        return 'policynum'

    name_pattern = re.compile(r"^[A-Za-zÀ-ÿ'\-\. ]{2,50}$")
    name_count = sum(1 for v in str_vals if name_pattern.match(v))
    if name_count / len(str_vals) > 0.8:
        return 'name'
    return ''

def detect_columns_by_content(df: pl.DataFrame):
    if df.is_empty():
        return [], [], [], [], []
    sample = df.head(min(df.height, 100))
    col_types = {}
    for i in range(sample.width):
        col_types[i] = infer_column_type_by_content(sample, i)

    name_indices = [i for i, t in col_types.items() if t == 'name']
    firstlast_indices = []
    sex_indices = [i for i, t in col_types.items() if t == 'sex']
    dob_indices = [i for i, t in col_types.items() if t == 'dob']
    pol_indices = [i for i, t in col_types.items() if t == 'policynum']

    if len(name_indices) >= 2:
        firstlast_indices = name_indices[:2]
        name_indices = name_indices[2:]

    return name_indices, firstlast_indices, sex_indices, dob_indices, pol_indices

# ==================== HEADER ROW DETECTION ====================
def detect_header_row_and_columns(df_sample, company_headers):
    sanitized_company = {key: {clean_for_match(v) for v in values} for key, values in company_headers.items()}
    for row_idx in range(min(len(df_sample), MAX_SEARCH_ROWS)):
        row_vals = df_sample.row(row_idx)
        row_sanitized_map = {}
        for v in row_vals:
            clean = clean_for_match(v)
            if clean:
                row_sanitized_map.setdefault(clean, []).append(v)
        row_set = set(row_sanitized_map.keys())
        name_match = sanitized_company['name'].intersection(row_set)
        firstlast_match = sanitized_company['firstlastname'].intersection(row_set)
        if name_match or firstlast_match:
            def get_raw(match_set):
                result = []
                for m in match_set:
                    result.extend(row_sanitized_map[m])
                return result
            name_cols = get_raw(name_match)
            firstlast_cols = get_raw(firstlast_match)
            sex_cols = get_raw(sanitized_company['sex'].intersection(row_set))
            dob_cols = get_raw(sanitized_company['dob'].intersection(row_set))
            pol_cols = get_raw(sanitized_company['policynum'].intersection(row_set))
            return row_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols, list(row_vals)

    for row_idx in range(min(len(df_sample), MAX_SEARCH_ROWS)):
        row_vals = df_sample.row(row_idx)
        name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols = [], [], [], [], []
        for v in row_vals:
            clean_v = clean_for_match(v)
            if not clean_v:
                continue
            if fuzzy_match_word(clean_v, company_headers['name']):
                name_cols.append(v)
            if fuzzy_match_word(clean_v, company_headers['firstlastname']):
                firstlast_cols.append(v)
            if fuzzy_match_word(clean_v, company_headers['sex']):
                sex_cols.append(v)
            if fuzzy_match_word(clean_v, company_headers['dob']):
                dob_cols.append(v)
            if fuzzy_match_word(clean_v, company_headers['policynum']):
                pol_cols.append(v)
        if name_cols or firstlast_cols:
            for v in row_vals:
                clean_v = clean_for_match(v)
                if not clean_v:
                    continue
                if not sex_cols and fuzzy_match_word(clean_v, company_headers['sex']):
                    sex_cols.append(v)
                if not dob_cols and fuzzy_match_word(clean_v, company_headers['dob']):
                    dob_cols.append(v)
                if not pol_cols and fuzzy_match_word(clean_v, company_headers['policynum']):
                    pol_cols.append(v)
            return row_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols, list(row_vals)
    return None, [], [], [], [], [], []

# ==================== COLUMN RENAMING HELPER ====================
def rename_columns_with_header(df, header_row_idx):
    """Rename columns using the header row, then return the DataFrame with that row removed."""
    raw_headers = df.row(header_row_idx)
    raw_headers = [str(h) if h is not None else '' for h in raw_headers]
    unique_headers = []
    counts = {}
    for h in raw_headers:
        if h in counts:
            counts[h] += 1
            unique_headers.append(f"{h}.{counts[h]}")
        else:
            counts[h] = 0
            unique_headers.append(h)
    data_df = df.slice(header_row_idx + 1)
    data_df.columns = unique_headers
    return data_df

# ==================== OUTPUT BUILDERS (ORIGINAL, WITH BUGFIX) ====================
def build_output_df(df, header_row_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                    file_path, sheet_name, company_code):
    # Use rename_columns_with_header if header_row_idx >= 0, else assume already renamed
    if header_row_idx >= 0:
        df = rename_columns_with_header(df, header_row_idx)
        # now no header row
    data_df = df
    multiple_name = (len(name_cols) >= 2) or (len(firstlast_cols) > 2)

    def finalize(result):
        result = result.with_columns(
            pl.when(pl.col('COMPLETE_NAME_RAW').is_not_null() & (pl.col('COMPLETE_NAME_RAW') != ''))
            .then(pl.col('COMPLETE_NAME_RAW'))
            .otherwise(pl.concat_str([pl.col('SURNAME'), pl.col('FIRST_NAME')], separator=' ').str.strip_chars())
            .alias('COMPLETE_NAME')
        )
        missing_name_count = result.filter(pl.col('COMPLETE_NAME') == '').height
        result = result.filter(pl.col('COMPLETE_NAME') != '')
        return result.select(OUTPUT_COLUMNS), missing_name_count

    if multiple_name:
        valid_name_cols = [c for c in name_cols if c in data_df.columns]
        if valid_name_cols:
            clean_exprs = [clean_string_expr(c) for c in valid_name_cols]
            df_with_names = data_df.select(
                pl.concat_list(clean_exprs).list.eval(pl.element().filter(pl.element() != "")).alias('names_list'),
                pl.all().exclude(valid_name_cols)
            )
            df_multi = df_with_names.filter(pl.col('names_list').list.len() > 1).explode('names_list')
            if not df_multi.is_empty():
                result = df_multi.select([
                    pl.lit(None).alias('SURNAME'), pl.lit(None).alias('FIRST_NAME'),
                    pl.col('names_list').alias('COMPLETE_NAME_RAW'), pl.lit('U').alias('SEX'),
                    pl.lit('').alias('DATE_OF_BIRTH'), pl.lit(company_code).alias('CMPY_NO'),
                    (pl.col(pol_cols[0]) if pol_cols else pl.lit(None)).alias('POLICY_NUMBER'),
                    pl.lit(file_path).alias('FILE_PATH'), pl.lit(sheet_name).alias('SHEET')
                ])
                return finalize(result)
        if len(firstlast_cols) > 2:
            pair_dfs = []
            for i in range(0, len(firstlast_cols)-1, 2):
                sur = firstlast_cols[i]
                fn = firstlast_cols[i+1]
                if sur in data_df.columns and fn in data_df.columns:
                    pair = data_df.select([
                        clean_string_expr(sur).alias('SURNAME'),
                        clean_string_expr(fn).alias('FIRST_NAME'),
                        (pl.col(pol_cols[0]) if pol_cols else pl.lit(None)).alias('POLICY_NUMBER')
                    ]).filter((pl.col('SURNAME') != '') | (pl.col('FIRST_NAME') != ''))
                    pair_dfs.append(pair)
            if pair_dfs:
                combined = pl.concat(pair_dfs)
                result = combined.select([
                    pl.col('SURNAME'), pl.col('FIRST_NAME'),
                    pl.concat_str([pl.col('SURNAME'), pl.col('FIRST_NAME')], separator=' ').str.strip_chars().alias('COMPLETE_NAME_RAW'),
                    pl.lit('U').alias('SEX'), pl.lit('').alias('DATE_OF_BIRTH'),
                    pl.lit(company_code).alias('CMPY_NO'), pl.col('POLICY_NUMBER'),
                    pl.lit(file_path).alias('FILE_PATH'), pl.lit(sheet_name).alias('SHEET')
                ])
                return finalize(result)

    sur_expr = clean_string_expr(firstlast_cols[0]) if firstlast_cols else pl.lit(None)
    fn_expr = clean_string_expr(firstlast_cols[1]) if len(firstlast_cols) > 1 else pl.lit(None)
    name_expr = clean_string_expr(name_cols[0]) if name_cols else pl.lit(None)
    sex_expr = pl.col(sex_cols[0]).map_elements(normalize_sex, return_dtype=pl.Utf8) if sex_cols else pl.lit('')
    dob_expr = pl.col(dob_cols[0]).map_elements(parse_date_to_mmddyyyy, return_dtype=pl.Utf8) if dob_cols else pl.lit('')
    result = data_df.select([
        sur_expr.alias('SURNAME'), fn_expr.alias('FIRST_NAME'), name_expr.alias('COMPLETE_NAME_RAW'),
        sex_expr.alias('SEX'), dob_expr.alias('DATE_OF_BIRTH'), pl.lit(company_code).alias('CMPY_NO'),
        (pl.col(pol_cols[0]) if pol_cols else pl.lit(None)).alias('POLICY_NUMBER'),
        pl.lit(file_path).alias('FILE_PATH'), pl.lit(sheet_name).alias('SHEET')
    ])
    return finalize(result)

def build_output_df_safe(df, header_row_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                         file_path, sheet_name, company_code):
    if header_row_idx >= 0:
        df = rename_columns_with_header(df, header_row_idx)
    data_df = df

    sur_expr = clean_string_expr(firstlast_cols[0]) if firstlast_cols else pl.lit(None)
    fn_expr = clean_string_expr(firstlast_cols[1]) if len(firstlast_cols) > 1 else pl.lit(None)
    name_expr = clean_string_expr(name_cols[0]) if name_cols else pl.lit(None)
    sex_expr = pl.col(sex_cols[0]).map_elements(normalize_sex, return_dtype=pl.Utf8) if sex_cols else pl.lit('')
    dob_expr = pl.col(dob_cols[0]).map_elements(parse_date_to_mmddyyyy, return_dtype=pl.Utf8) if dob_cols else pl.lit('')

    result = data_df.select([
        sur_expr.alias('SURNAME'), fn_expr.alias('FIRST_NAME'), name_expr.alias('COMPLETE_NAME_RAW'),
        sex_expr.alias('SEX'), dob_expr.alias('DATE_OF_BIRTH'), pl.lit(company_code).alias('CMPY_NO'),
        (pl.col(pol_cols[0]) if pol_cols else pl.lit(None)).alias('POLICY_NUMBER'),
        pl.lit(file_path).alias('FILE_PATH'), pl.lit(sheet_name).alias('SHEET')
    ])
    result = result.with_columns(
        pl.when(pl.col('COMPLETE_NAME_RAW').is_not_null() & (pl.col('COMPLETE_NAME_RAW') != ''))
        .then(pl.col('COMPLETE_NAME_RAW'))
        .otherwise(pl.concat_str([pl.col('SURNAME'), pl.col('FIRST_NAME')], separator=' ').str.strip_chars())
        .alias('COMPLETE_NAME')
    )
    missing_name_count = result.filter(pl.col('COMPLETE_NAME') == '').height
    result = result.filter(pl.col('COMPLETE_NAME') != '')
    return result.select(OUTPUT_COLUMNS), missing_name_count

def build_output_df_safe_without_header(df, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                                        file_path, sheet_name, company_code):
    # This is kept for headerless fallback; it doesn't handle multi-name,
    # but it's only used when no header was found at all.
    data_df = df.clone()
    sur_expr = clean_string_expr(firstlast_cols[0]) if firstlast_cols and firstlast_cols[0] in data_df.columns else pl.lit(None)
    fn_expr = clean_string_expr(firstlast_cols[1]) if len(firstlast_cols) > 1 and firstlast_cols[1] in data_df.columns else pl.lit(None)
    name_expr = clean_string_expr(name_cols[0]) if name_cols and name_cols[0] in data_df.columns else pl.lit(None)
    sex_expr = pl.col(sex_cols[0]).map_elements(normalize_sex, return_dtype=pl.Utf8) if sex_cols and sex_cols[0] in data_df.columns else pl.lit('')
    dob_expr = pl.col(dob_cols[0]).map_elements(parse_date_to_mmddyyyy, return_dtype=pl.Utf8) if dob_cols and dob_cols[0] in data_df.columns else pl.lit('')
    pol_expr = pl.col(pol_cols[0]) if pol_cols and pol_cols[0] in data_df.columns else pl.lit(None)

    result = data_df.select([
        sur_expr.alias('SURNAME'), fn_expr.alias('FIRST_NAME'), name_expr.alias('COMPLETE_NAME_RAW'),
        sex_expr.alias('SEX'), dob_expr.alias('DATE_OF_BIRTH'), pl.lit(company_code).alias('CMPY_NO'),
        pol_expr.alias('POLICY_NUMBER'),
        pl.lit(file_path).alias('FILE_PATH'), pl.lit(sheet_name).alias('SHEET')
    ])
    result = result.with_columns(
        pl.when(pl.col('COMPLETE_NAME_RAW').is_not_null() & (pl.col('COMPLETE_NAME_RAW') != ''))
        .then(pl.col('COMPLETE_NAME_RAW'))
        .otherwise(pl.concat_str([pl.col('SURNAME'), pl.col('FIRST_NAME')], separator=' ').str.strip_chars())
        .alias('COMPLETE_NAME')
    )
    missing_name_count = result.filter(pl.col('COMPLETE_NAME') == '').height
    return result.filter(pl.col('COMPLETE_NAME') != '').select(OUTPUT_COLUMNS), missing_name_count

# ==================== EXCEL PROCESSING (ORIGINAL LOGIC, WITH FIXES) ====================
def process_excel_file_standard(job, file_path, file_name, company_headers, password):
    try:
        sheets = pl.read_excel(file_path if isinstance(file_path, str) else file_path,
                               has_header=False, sheet_id=0, raise_if_empty=False)
        for sheet_name, df in sheets.items():
            sample_df = df.head(min(df.height, MAX_SEARCH_ROWS))

            # Multi‑table detection
            tables = detect_tables_in_sheet(sample_df, company_headers, MAX_SEARCH_ROWS)
            if not tables:
                # Fallback to single‑table
                tables = [{
                    'col_start': 0,
                    'col_end': df.width - 1,
                    'header_row_idx': -1,
                    'header_cols': {
                        'name': [], 'firstlastname': [], 'sex': [], 'dob': [], 'policynum': []
                    },
                    'identified_headers': []
                }]
                hdr_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols, identified = \
                    detect_header_row_and_columns(sample_df, company_headers)
                if hdr_idx is not None:
                    tables[0]['header_row_idx'] = hdr_idx
                    tables[0]['header_cols']['name'] = name_cols
                    tables[0]['header_cols']['firstlastname'] = firstlast_cols
                    tables[0]['header_cols']['sex'] = sex_cols
                    tables[0]['header_cols']['dob'] = dob_cols
                    tables[0]['header_cols']['policynum'] = pol_cols
                    tables[0]['identified_headers'] = identified
                else:
                    name_idx, firstlast_idx, sex_idx, dob_idx, pol_idx = detect_columns_by_content(df)
                    if name_idx or firstlast_idx:
                        tables[0]['header_row_idx'] = -1
                        tables[0]['header_cols']['name'] = [df.columns[i] for i in name_idx] if name_idx else []
                        tables[0]['header_cols']['firstlastname'] = [df.columns[i] for i in firstlast_idx] if firstlast_idx else []
                        tables[0]['header_cols']['sex'] = [df.columns[i] for i in sex_idx] if sex_idx else []
                        tables[0]['header_cols']['dob'] = [df.columns[i] for i in dob_idx] if dob_idx else []
                        tables[0]['header_cols']['policynum'] = [df.columns[i] for i in pol_idx] if pol_idx else []
                    else:
                        write_log_row(job, {
                            'File Path': file_path, 'File Name': file_name,
                            'Sheet Name': sheet_name,
                            'Error Msg': 'No identifiable headers or content patterns',
                            'Remarks': ''
                        })
                        continue

            # Process each table
            for table_idx, table in enumerate(tables):
                col_start = table['col_start']
                col_end = table['col_end']
                hdr_idx_rel = table['header_row_idx']
                header_cols = table['header_cols']
                identified = table.get('identified_headers', [])

                # Slice the table
                if col_start == 0 and col_end == df.width - 1:
                    table_df = df
                else:
                    table_df = df[:, col_start:col_end+1]

                # Prepare column names from header values
                name_cols_raw = header_cols['name']
                firstlast_cols_raw = header_cols['firstlastname']
                sex_cols_raw = header_cols['sex']
                dob_cols_raw = header_cols['dob']
                pol_cols_raw = header_cols['policynum']

                # ---------- RENAME COLUMNS IF HEADER FOUND ----------
                if hdr_idx_rel >= 0:
                    # Save the original header values before renaming
                    original_header_vals = [str(v) if v is not None else '' for v in table_df.row(hdr_idx_rel)]
                    table_df = rename_columns_with_header(table_df, hdr_idx_rel)
                    # Map raw column names to new unique column names
                    def map_to_new_names(raw_names):
                        new = []
                        for raw in raw_names:
                            # The raw name came from the header row; find which new column corresponds to it
                            # Since renaming preserved order, we can match by index: find indices of raw in original_header_vals
                            indices = [i for i, oh in enumerate(original_header_vals) if oh == raw]
                            for idx in indices:
                                new.append(table_df.columns[idx])
                        return new
                    name_cols = map_to_new_names(name_cols_raw)
                    firstlast_cols = map_to_new_names(firstlast_cols_raw)
                    sex_cols = map_to_new_names(sex_cols_raw)
                    dob_cols = map_to_new_names(dob_cols_raw)
                    pol_cols = map_to_new_names(pol_cols_raw)
                    hdr_idx_rel = -1  # No header row remaining
                else:
                    name_cols = name_cols_raw
                    firstlast_cols = firstlast_cols_raw
                    sex_cols = sex_cols_raw
                    dob_cols = dob_cols_raw
                    pol_cols = pol_cols_raw

                # ---------- NUMERIC FILTER (NOW ON RENAMED COLUMNS) ----------
                name_cols = filter_numeric_columns(table_df, name_cols)
                firstlast_cols = filter_numeric_columns(table_df, firstlast_cols)

                # ---------- BUILD OUTPUT ----------
                table_sheet_name = f"{sheet_name}_table{table_idx+1}" if len(tables) > 1 else sheet_name
                remarks = []

                if hdr_idx_rel >= 0:
                    # This shouldn't happen after renaming, but just in case
                    output_df, missing_name = build_output_df_safe(
                        table_df, hdr_idx_rel, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                        file_path, table_sheet_name, job.company_code)
                else:
                    # Use the safe builder (which handles multi‑name for headerless? No,
                    # but multi‑name is in build_output_df, not in build_output_df_safe.
                    # So we call build_output_df_safe with hdr_idx=-1, which now avoids slicing.
                    output_df, missing_name = build_output_df_safe(
                        table_df, -1, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                        file_path, table_sheet_name, job.company_code)

                remarks.append(f"Missing name rows: {missing_name}")
                before_filter = output_df.height
                output_df = filter_duplicates_only(output_df)
                after_filter = output_df.height
                remarks.append(f"Rows before filter: {before_filter}, after: {after_filter}")

                if output_df.is_empty():
                    write_log_row(job, {
                        'File Path': file_path, 'File Name': file_name,
                        'Sheet Name': table_sheet_name,
                        'Error Msg': 'All rows filtered out (age/duplicates)',
                        'Remarks': '; '.join(remarks)
                    })
                    continue

                # CSV splitting
                max_rows = MAX_ROWS_PER_OUTPUT_CSV if MAX_ROWS_PER_OUTPUT_CSV > 0 else output_df.height
                part = 1
                chunk_start = 0
                multiple_name = (len(name_cols) >= 2 or len(firstlast_cols) > 2)

                while chunk_start < output_df.height:
                    chunk = output_df.slice(chunk_start, max_rows)
                    out_path = os.path.join(
                        job.csv_folder,
                        f"{file_name}[{table_sheet_name}]_part{part}_OFAC_OUTPUT.csv"
                    )
                    out_path = get_unique_save_path(out_path)
                    chunk.write_csv(out_path)

                    row_dict = {
                        'File Path': file_path, 'File Name': file_name, 'Scan Date': job.today,
                        'Extension': os.path.splitext(file_name)[1].replace('.', ''),
                        'Company Code': job.company_code,
                        'Password': password,
                        'Sheet Name': f"{table_sheet_name}_part{part}" if max_rows < output_df.height else table_sheet_name,
                        'Identified Headers': ', '.join(str(x) for x in identified) if identified else None,
                        'Multiple Name': multiple_name,
                        'Row Count': table_df.height,
                        'Output Row Count': chunk.height,
                        'Output CSV': out_path,
                        'First Last Name Header': ', '.join(firstlast_cols) if firstlast_cols else None,
                        'Full Name Header': ', '.join(name_cols) if name_cols else None,
                        'Policy Number Header': pol_cols[0] if pol_cols else None,
                        'DOB Header': dob_cols[0] if dob_cols else None,
                        'Sex Header': sex_cols[0] if sex_cols else None,
                        'Remarks': '; '.join(remarks) + (f' (part {part})' if max_rows < output_df.height else '')
                    }
                    write_log_row(job, row_dict)
                    chunk_start += max_rows
                    part += 1

    except Exception as e:
        write_log_row(job, {'File Path': file_path, 'File Name': file_name, 'Error Msg': str(e)})

# --- Chunked processing disabled ---
def process_excel_chunked(job, file_path, file_name, company_headers, password):
    raise NotImplementedError("Chunked Excel processing requires implementation of extract_table_data_from_chunk.")

# ==================== TEXT FILE PROCESSING (RESTORED WITH FIX) ====================
def process_text_file(job, file_path, file_name, company_headers):
    try:
        with open(file_path, 'rb') as f:
            raw = f.read(10000)
            enc = chardet.detect(raw)['encoding'] or 'utf-8'
        with open(file_path, 'r', encoding=enc) as f:
            sample = f.read(10000)
            dialect = clevercsv.Sniffer().sniff(sample)
        delim, quote = dialect.delimiter, dialect.quotechar
        size = os.path.getsize(file_path)
        remarks = []

        if size < SIZE_TO_CHUNK:
            df = pl.read_csv(file_path, has_header=False, encoding=enc, separator=delim, quote_char=quote)
            sample_df = df.head(min(df.height, MAX_SEARCH_ROWS))
            hdr_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols, identified = \
                detect_header_row_and_columns(sample_df, company_headers)
            if hdr_idx is not None:
                # Rename columns first, then numeric filter
                df = rename_columns_with_header(df, hdr_idx)
                name_cols = filter_numeric_columns(df, name_cols)
                firstlast_cols = filter_numeric_columns(df, firstlast_cols)

                output_df, missing_name = build_output_df(
                    df, -1, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                    file_path, '', job.company_code)
                remarks.append(f"Missing name rows: {missing_name}")
                output_df = filter_duplicates_only(output_df)
                # ... rest same (splitting, logging) ...
                if not output_df.is_empty():
                    max_rows = MAX_ROWS_PER_OUTPUT_CSV if MAX_ROWS_PER_OUTPUT_CSV > 0 else output_df.height
                    part = 1
                    chunk_start = 0
                    out_rows = output_df.height
                    first_out_path = None
                    while chunk_start < output_df.height:
                        chunk = output_df.slice(chunk_start, max_rows)
                        out_path = get_unique_save_path(
                            os.path.join(job.csv_folder, f"{file_name}_part{part}_OFAC_OUTPUT.csv")
                        )
                        chunk.write_csv(out_path)
                        if part == 1:
                            first_out_path = out_path
                        row_dict = {
                            'File Path': file_path, 'File Name': file_name, 'Scan Date': job.today,
                            'Extension': os.path.splitext(file_name)[1].replace('.',''),
                            'Company Code': job.company_code,
                            'Sheet Name': f"part{part}" if max_rows < output_df.height else '',
                            'Identified Headers': ', '.join(str(x) for x in identified) if identified else None,
                            'Multiple Name': (len(name_cols) >= 2 or len(firstlast_cols) > 2),
                            'Row Count': df.height,
                            'Output Row Count': chunk.height,
                            'Output CSV': out_path,
                            'First Last Name Header': ', '.join(firstlast_cols) if firstlast_cols else None,
                            'Full Name Header': ', '.join(name_cols) if name_cols else None,
                            'Policy Number Header': pol_cols[0] if pol_cols else None,
                            'DOB Header': dob_cols[0] if dob_cols else None,
                            'Sex Header': sex_cols[0] if sex_cols else None,
                            'Remarks': '; '.join(remarks) + (f' (part {part})' if max_rows < output_df.height else '')
                        }
                        write_log_row(job, row_dict)
                        chunk_start += max_rows
                        part += 1
                else:
                    out_path, out_rows = None, 0
            else:
                # Header not found – log error
                row_dict = {
                    'File Path': file_path, 'File Name': file_name, 'Scan Date': job.today,
                    'Extension': os.path.splitext(file_name)[1].replace('.',''),
                    'Company Code': job.company_code,
                    'Sheet Name': '',
                    'Error Msg': 'No identifiable headers found',
                    'Output Row Count': 0,
                    'Remarks': '; '.join(remarks)
                }
                write_log_row(job, row_dict)

        else:
            # Large file processing with batching – updated to rename first batch and use -1 for others
            reader = pl.read_csv_batched(file_path, has_header=False, encoding=enc,
                                         separator=delim, quote_char=quote, batch_size=CHUNK_SIZE)
            first_batch = reader.next_batches(1)[0]
            sample_df = first_batch.head(min(first_batch.height, MAX_SEARCH_ROWS))
            hdr_idx, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols, identified = \
                detect_header_row_and_columns(sample_df, company_headers)

            if hdr_idx is None:
                write_log_row(job, {
                    'File Path': file_path, 'File Name': file_name,
                    'Error Msg': 'No identifiable headers found in large file',
                    'Remarks': 'Header match failed'
                })
                return

            # Rename first batch and apply numeric filter
            first_batch = rename_columns_with_header(first_batch, hdr_idx)
            name_cols = filter_numeric_columns(first_batch, name_cols)
            firstlast_cols = filter_numeric_columns(first_batch, firstlast_cols)

            multiple_name = (len(name_cols) >= 2 or len(firstlast_cols) > 2)
            total_out_rows = 0
            total_missing_name = 0
            part = 1
            current_part_rows = 0
            current_chunk_df = pl.DataFrame()
            first_out_path = None

            def process_one_batch(batch_df):
                nonlocal total_out_rows, total_missing_name, part, current_part_rows, current_chunk_df, first_out_path
                output_df, missing_name = build_output_df(
                    batch_df, -1, name_cols, firstlast_cols, sex_cols, dob_cols, pol_cols,
                    file_path, '', job.company_code)
                total_missing_name += missing_name
                output_df = filter_duplicates_only(output_df)
                if output_df.is_empty():
                    return

                max_rows = MAX_ROWS_PER_OUTPUT_CSV if MAX_ROWS_PER_OUTPUT_CSV > 0 else output_df.height
                chunk_start = 0
                while chunk_start < output_df.height:
                    chunk = output_df.slice(chunk_start, max_rows)
                    if current_chunk_df.is_empty():
                        current_chunk_df = chunk
                    else:
                        if current_part_rows + chunk.height > max_rows:
                            out_path = get_unique_save_path(
                                os.path.join(job.csv_folder, f"{file_name}_part{part}_OFAC_OUTPUT.csv")
                            )
                            current_chunk_df.write_csv(out_path)
                            if part == 1:
                                first_out_path = out_path
                            remarks_part = [f"Missing name rows (total so far): {total_missing_name}",
                                            f"Rows in this part: {current_part_rows}"]
                            write_log_row(job, {
                                'File Path': file_path, 'File Name': file_name, 'Scan Date': job.today,
                                'Extension': os.path.splitext(file_name)[1].replace('.',''),
                                'Company Code': job.company_code,
                                'Sheet Name': f"part{part}",
                                'Identified Headers': ', '.join(str(x) for x in identified) if identified else None,
                                'Multiple Name': multiple_name,
                                'Row Count': None,
                                'Output Row Count': current_part_rows,
                                'Output CSV': out_path,
                                'First Last Name Header': ', '.join(firstlast_cols) if firstlast_cols else None,
                                'Full Name Header': ', '.join(name_cols) if name_cols else None,
                                'Policy Number Header': pol_cols[0] if pol_cols else None,
                                'DOB Header': dob_cols[0] if dob_cols else None,
                                'Sex Header': sex_cols[0] if sex_cols else None,
                                'Remarks': '; '.join(remarks_part) + ' (Chunked processing)'
                            })
                            part += 1
                            current_part_rows = 0
                            current_chunk_df = chunk
                        else:
                            current_chunk_df = pl.concat([current_chunk_df, chunk], how='diagonal_relaxed')
                    current_part_rows += chunk.height
                    total_out_rows += chunk.height
                    chunk_start += max_rows

            process_one_batch(first_batch)
            for batch in reader:
                process_one_batch(batch)

            if not current_chunk_df.is_empty():
                out_path = get_unique_save_path(
                    os.path.join(job.csv_folder, f"{file_name}_part{part}_OFAC_OUTPUT.csv")
                )
                current_chunk_df.write_csv(out_path)
                if part == 1:
                    first_out_path = out_path
                remarks_part = [f"Missing name rows (total): {total_missing_name}",
                                f"Rows in this part: {current_part_rows}"]
                write_log_row(job, {
                    'File Path': file_path, 'File Name': file_name, 'Scan Date': job.today,
                    'Extension': os.path.splitext(file_name)[1].replace('.',''),
                    'Company Code': job.company_code,
                    'Sheet Name': f"part{part}",
                    'Identified Headers': ', '.join(str(x) for x in identified) if identified else None,
                    'Multiple Name': multiple_name,
                    'Row Count': None,
                    'Output Row Count': current_part_rows,
                    'Output CSV': out_path,
                    'First Last Name Header': ', '.join(firstlast_cols) if firstlast_cols else None,
                    'Full Name Header': ', '.join(name_cols) if name_cols else None,
                    'Policy Number Header': pol_cols[0] if pol_cols else None,
                    'DOB Header': dob_cols[0] if dob_cols else None,
                    'Sex Header': sex_cols[0] if sex_cols else None,
                    'Remarks': '; '.join(remarks_part) + ' (Chunked processing)'
                })

    except Exception as e:
        write_log_row(job, {'File Path': file_path, 'File Name': file_name, 'Error Msg': str(e)})

# ==================== ARCHIVE, LOGGING, COMPILATION, WATCHER (UNCHANGED) ====================
def process_archive_file(job, archive_path, archive_name):
    extracted_files = []
    try:
        files_before = get_all_files(job.unzipped_folder)
        success = False
        used_pwd = None
        for pwd in job.passwords:
            result = subprocess.run(
                [SEVEN_ZIP_PATH, 'x', archive_path, '-aou', f'-o{job.unzipped_folder}', f'-p{pwd}'],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                success = True
                used_pwd = pwd
                break
        if success:
            files_after = get_all_files(job.unzipped_folder)
            extracted_files = list(files_after - files_before)
    except:
        pass
    write_log_row(job, {
        'File Path': archive_path, 'File Name': archive_name,
        'Extension': os.path.splitext(archive_name)[1].replace('.',''),
        'Password': used_pwd if success else None,
        'Error Msg': None if success else "Extraction failed",
        'Remarks': f"Extracted {len(extracted_files)} files" if success else None
    })
    return extracted_files

def write_log_row(job, row_dict):
    row = {h: None for h in LOG_HEADERS}
    row.update(row_dict)
    try:
        file_exists = os.path.exists(job.log_file) and os.path.getsize(job.log_file) > 0
        with open(job.log_file, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=LOG_HEADERS)
            if not file_exists:
                writer.writeheader()
            writer.writerow(row)
    except Exception as e:
        print(f"Failed to write log: {e}")

def compile_outputs(job):
    if not os.path.exists(job.log_file):
        return
    log_df = pl.read_csv(job.log_file)
    compiled_row_count = 0
    current_df = pl.DataFrame()
    output_files = []
    base_out = os.path.join(job.compiled_folder, f"OFAC_ABS_Log_{job.date_display}_{job.company_code}")
    file_idx = 1
    for row in log_df.iter_rows(named=True):
        csv_path = row.get('Output CSV')
        if not csv_path or not os.path.exists(csv_path):
            continue
        out_rows_str = row.get('Output Row Count')
        if out_rows_str is None:
            continue
        try:
            out_rows = int(out_rows_str)
        except (ValueError, TypeError):
            continue
        if compiled_row_count + out_rows > EXCEL_MAX_ROWS and not current_df.is_empty():
            out_path = get_unique_save_path(f"{base_out}_{file_idx}.xlsx")
            current_df.write_excel(out_path)
            output_files.append(out_path)
            current_df = pl.DataFrame()
            compiled_row_count = 0
            file_idx += 1
        try:
            df_chunk = pl.read_csv(csv_path, schema_overrides={'POLICY_NUMBER': pl.Utf8})
            if not df_chunk.is_empty():
                current_df = pl.concat([current_df, df_chunk], how='diagonal_relaxed')
                compiled_row_count += out_rows
        except Exception as e:
            print(f"Error reading {csv_path}: {e}")
    if not current_df.is_empty():
        out_path = get_unique_save_path(f"{base_out}_{file_idx}.xlsx")
        current_df.write_excel(out_path)
        output_files.append(out_path)

def process_files_direct(job, progress_callback=None, stop_flag=None, progress_update=None):
    company_headers, used_header_file = get_company_header(job.company_code)
    debug_header_info = f"Header file: {used_header_file}. Sets: name={len(company_headers['name'])} keywords, firstlastname={len(company_headers['firstlastname'])}, sex={len(company_headers['sex'])}, dob={len(company_headers['dob'])}, policynum={len(company_headers['policynum'])}"
    write_log_row(job, {'Remarks': debug_header_info})

    if not os.path.exists(job.log_file) or os.path.getsize(job.log_file) == 0:
        with open(job.log_file, 'w', newline='', encoding='utf-8') as f:
            csv.DictWriter(f, fieldnames=LOG_HEADERS).writeheader()
    all_file_paths = [os.path.join(job.input_folder, f) for f in job.file_names if os.path.exists(os.path.join(job.input_folder, f))]
    idx = 0
    while idx < len(all_file_paths):
        if stop_flag and stop_flag():
            write_log_row(job, {'Remarks': 'Scan stopped by user'})
            break
        if progress_update:
            progress_update(idx + 1, len(all_file_paths))
        src_path = all_file_paths[idx]
        file_name = os.path.basename(src_path)
        if progress_callback:
            progress_callback(f"Processing {file_name}...")
        try:
            archived_path = move_file_to_archived(src_path, job.archived_folder)
        except Exception as e:
            write_log_row(job, {'File Path': src_path, 'File Name': file_name, 'Error Msg': str(e)})
            idx += 1
            continue
        ext = os.path.splitext(file_name)[1].lower().replace('.', '')
        if ext in FILE_EXTENSIONS_EXCEL:
            process_excel_file(job, archived_path, file_name, company_headers)
        elif ext in FILE_EXTENSIONS_TEXT:
            process_text_file(job, archived_path, file_name, company_headers)
        elif ext in FILE_EXTENSIONS_ARCHIVE:
            new_files = process_archive_file(job, archived_path, file_name)
            all_file_paths.extend(new_files)
        idx += 1
    compile_outputs(job)
    return True

def process_excel_file(job, file_path, file_name, company_headers):
    use_chunking = False
    temp_path = None
    try:
        decrypted, password = unlock_excel(file_path, job.passwords)
        if isinstance(decrypted, io.BytesIO):
            temp_path = os.path.join(job.unzipped_folder, f"temp_{int(time.time())}_{file_name}")
            with open(temp_path, 'wb') as f:
                f.write(decrypted.getvalue())
            actual_path = temp_path
        else:
            actual_path = decrypted
        if use_chunking:
            process_excel_chunked(job, actual_path, file_name, company_headers, password)
        else:
            process_excel_file_standard(job, actual_path, file_name, company_headers, password)
    except Exception as e:
        write_log_row(job, {'File Path': file_path, 'File Name': file_name, 'Error Msg': str(e)})
    finally:
        if temp_path:
            try:
                os.remove(temp_path)
            except:
                pass

# ==================== TABLE DETECTION & SHIFT ADJUSTMENT (ORIGINAL) ====================
def is_empty_cell(val):
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

def find_non_empty_columns(df_sample, max_rows=10):
    if df_sample.is_empty():
        return []
    n_rows = min(df_sample.height, max_rows)
    non_empty = set()
    for col_idx in range(df_sample.width):
        col_data = df_sample[:, col_idx].to_list()[:n_rows]
        if any(not is_empty_cell(v) for v in col_data):
            non_empty.add(col_idx)
    return sorted(non_empty)

def detect_tables_in_sheet(df_sample, company_headers, max_search_rows):
    """Scan for multiple tables in one sheet by column clusters.
    Returns column names (header cell values) for each category."""
    tables = []
    non_empty_cols = find_non_empty_columns(df_sample, max_rows=10)
    if not non_empty_cols:
        return tables
    col_ranges = []
    start = non_empty_cols[0]
    prev = non_empty_cols[0]
    for c in non_empty_cols[1:]:
        if c > prev + 1:
            col_ranges.append((start, prev))
            start = c
        prev = c
    col_ranges.append((start, prev))

    sanitized_company = {key: {clean_for_match(v) for v in values} for key, values in company_headers.items()}
    for col_start, col_end in col_ranges:
        sub_sample = df_sample[:, col_start:col_end+1]
        for row_idx in range(min(sub_sample.height, max_search_rows)):
            row_vals = sub_sample.row(row_idx)
            row_sanitized_map = {}
            for v in row_vals:
                clean = clean_for_match(v)
                if clean:
                    row_sanitized_map.setdefault(clean, []).append(v)
            row_set = set(row_sanitized_map.keys())
            name_match = sanitized_company['name'].intersection(row_set)
            firstlast_match = sanitized_company['firstlastname'].intersection(row_set)
            if name_match or firstlast_match:
                def get_indices(match_set):
                    indices = []
                    for m in match_set:
                        for orig_val in row_sanitized_map[m]:
                            for i, val in enumerate(row_vals):
                                if val == orig_val:
                                    indices.append(i)
                    return sorted(set(indices))
                name_cols_idx = get_indices(name_match)
                firstlast_cols_idx = get_indices(firstlast_match)
                sex_cols_idx = get_indices(sanitized_company['sex'].intersection(row_set))
                dob_cols_idx = get_indices(sanitized_company['dob'].intersection(row_set))
                pol_cols_idx = get_indices(sanitized_company['policynum'].intersection(row_set))

                tables.append({
                    'header_row_idx': row_idx,
                    'col_start': col_start,
                    'col_end': col_end,
                    'header_cols': {
                        'name': [row_vals[i] for i in name_cols_idx],
                        'firstlastname': [row_vals[i] for i in firstlast_cols_idx],
                        'sex': [row_vals[i] for i in sex_cols_idx],
                        'dob': [row_vals[i] for i in dob_cols_idx],
                        'policynum': [row_vals[i] for i in pol_cols_idx]
                    },
                    'identified_headers': list(row_vals)
                })
                break   # only one header per cluster
    return tables

def adjust_for_shifts(df, table_info):
    # (original code kept for potential future use, but no longer called in the new flow)
    header_row = table_info['header_row_idx']
    col_start = table_info['col_start']
    col_end = table_info['col_end']
    data_start_row = header_row + 1
    for r in range(header_row+1, min(df.height, header_row+20)):
        row_slice = df[r, col_start:col_end+1]
        if any(not is_empty_cell(v) for v in row_slice):
            data_start_row = r
            break
    adjusted_cols = {}
    for category, indices in table_info['header_cols'].items():
        adjusted = []
        for col_in_table in indices:
            abs_col = col_start + col_in_table
            col_data = df[data_start_row:, abs_col].to_list()
            if any(not is_empty_cell(v) for v in col_data):
                adjusted.append(col_in_table)
            else:
                found = False
                for offset in (1, -1):
                    neighbor = col_in_table + offset
                    if 0 <= neighbor < (col_end - col_start + 1):
                        neigh_abs = col_start + neighbor
                        neigh_data = df[data_start_row:, neigh_abs].to_list()
                        if any(not is_empty_cell(v) for v in neigh_data):
                            adjusted.append(neighbor)
                            found = True
                            break
                if not found:
                    adjusted.append(col_in_table)
        adjusted_cols[category] = adjusted
    table_info['data_start_row'] = data_start_row
    table_info['adjusted_header_cols'] = adjusted_cols
    return table_info

def unlock_excel(file_path, passwords):
    with open(file_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return file_path, ''
        for pwd in passwords:
            try:
                office_file.load_key(password=pwd)
                decrypted = io.BytesIO()
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                return decrypted, pwd
            except Exception:
                continue
        raise ValueError(f"Could not decrypt {file_path} with any password")

# ==================== WATCHER ====================
class FolderHandler(FileSystemEventHandler):
    def __init__(self, watch_path):
        self.watch_path = watch_path
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith('.json'):
            print(f"Config detected: {event.src_path} – launching scanner")
            try:
                subprocess.Popen([sys.executable, sys.argv[0], "--process-config", event.src_path],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as e:
                print(f"Failed: {e}")

def start_watching():
    folder = get_env_var(ENV_VAR_FOLDER)
    if not folder or not os.path.isdir(folder):
        print(f"Invalid watch folder: {folder}")
        return
    event_handler = FolderHandler(folder)
    observer = Observer()
    observer.schedule(event_handler, folder, recursive=False)
    observer.start()
    print(f"Watching {folder}...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def process_config_and_scan(config_path):
    if not os.path.exists(config_path):
        return
    with open(config_path, 'r') as f:
        config = json.load(f)
    input_folder = get_env_var(ENV_VAR_FOLDER)
    if not input_folder:
        print("Input folder not configured.")
        return
    company_code = config.get('company_code')
    passwords = config.get('passwords', [])
    email_received_date = config.get('email_received_date')
    files = config.get('files', [])
    if not company_code or not passwords or not email_received_date or not files:
        print("Incomplete configuration.")
        return
    job = ScannerJob(input_folder, company_code, passwords, email_received_date, files)
    def log_callback(msg):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
    process_files_direct(job, progress_callback=log_callback, stop_flag=None)
    print("Scan completed.")
    try:
        os.remove(config_path)
    except:
        pass
