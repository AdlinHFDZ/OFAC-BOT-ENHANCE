"""
header_extractor_core.py
Enhanced Header Extractor backend – mirrors the scanner core in file handling,
archive extraction and decryption, but collects column headers with rich metadata:
- File Path, Sheet Name, Table, Column Index, Row Index, Category.
Categories are determined by analysing the actual data in each column.
"""

import os
import sys
import csv
import subprocess
import time
import shutil
import re
import io
import tempfile
from datetime import datetime

import polars as pl
import msoffcrypto

# Reuse constants and helpers from scanner core
from ofac_scanner_core import (
    BASE_OUTPUT_FOLDER,
    SEVEN_ZIP_PATH,
    FILE_EXTENSIONS_EXCEL,
    FILE_EXTENSIONS_TEXT,
    FILE_EXTENSIONS_ARCHIVE,
    ENV_VAR_FOLDER,
    ENV_VAR_CSV,
    get_env_var,
    get_unique_save_path,
    get_all_files,
    move_file_to_archived,
    unlock_excel,
    infer_column_type_by_content,
    MAX_SEARCH_ROWS,
)

# ----------------------------------------------------------------------
# Job class (now with optional archive flag)
# ----------------------------------------------------------------------
class HeaderExtractionJob:
    def __init__(self, input_folder, company_code, passwords, date_str, file_names,
                 archive_files=True):
        self.input_folder = input_folder
        self.company_code = company_code
        self.passwords = passwords
        self.date_str = date_str
        self.file_names = file_names
        self.archive_files = archive_files      # if False, read files from watch folder directly
        self.output_root = os.path.join(BASE_OUTPUT_FOLDER, f"HD_extract_{date_str}")
        self.archived_folder = os.path.join(self.output_root, "Archived")
        self.unzipped_folder = os.path.join(self.output_root, "Unzipped")
        self.log_file = os.path.join(self.output_root, f"Log_{date_str}.csv")
        self.header_output_file = os.path.join(
            self.output_root, f"{date_str}_headerextract_{company_code}.xlsx"
        )
        for d in [self.output_root, self.archived_folder, self.unzipped_folder]:
            os.makedirs(d, exist_ok=True)

# ----------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------
LOG_HEADERS = [
    "File Path", "File Name", "Sheet Name",
    "Headers Extracted", "Error Msg", "Remarks"
]

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
        print(f"Log write error: {e}")

# ----------------------------------------------------------------------
# Generic helper: is cell empty?
# ----------------------------------------------------------------------
def is_empty_cell(val):
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

# ----------------------------------------------------------------------
# Find contiguous non‑empty column blocks
# ----------------------------------------------------------------------
def find_table_blocks(df):
    non_empty_cols = []
    n_rows = min(df.height, MAX_SEARCH_ROWS)
    for col_idx in range(df.width):
        col_data = df[:, col_idx].to_list()[:n_rows]
        if any(not is_empty_cell(v) for v in col_data):
            non_empty_cols.append(col_idx)
    if not non_empty_cols:
        return []
    blocks = []
    start = non_empty_cols[0]
    prev = non_empty_cols[0]
    for c in non_empty_cols[1:]:
        if c > prev + 1:
            blocks.append((start, prev))
            start = c
        prev = c
    blocks.append((start, prev))
    return blocks

# ----------------------------------------------------------------------
# Improved header row detection (row scoring)
# ----------------------------------------------------------------------
def detect_header_row_for_block(df, col_start, col_end):
    """
    Scan the first MAX_SEARCH_ROWS rows of the block.
    Pick the row with the highest number of non‑empty cells.
    Return (row_index, header_values).
    """
    n_rows = min(df.height, MAX_SEARCH_ROWS)
    best_row = None
    best_headers = []
    best_score = -1

    for row_idx in range(n_rows):
        row_vals = df.row(row_idx)[col_start:col_end+1]
        # Filter out None / empty
        non_empty = [str(v).strip() for v in row_vals if not is_empty_cell(v)]
        score = len(non_empty)
        if score > best_score:
            best_score = score
            best_row = row_idx
            best_headers = non_empty

    if best_row is not None and best_headers:
        return best_row, best_headers
    return None, []

# ----------------------------------------------------------------------
# Content‑based categorisation (extends scanner's infer_column_type_by_content)
# ----------------------------------------------------------------------
def categorise_column(df, col_index, header_row):
    """
    Analyse the data below the header row (up to 200 rows) and return a category string.
    Uses infer_column_type_by_content for DOB/sex/policynum/name,
    then falls back to simple numeric/date/text checks.
    """
    if df.is_empty() or col_index >= df.width:
        return "Unknown"

    # Sample data below header
    data_start = header_row + 1
    if data_start >= df.height:
        return "Empty"

    sample_slice = df[data_start:, col_index]
    non_null = sample_slice.drop_nulls()
    sample = non_null.head(200).to_list()
    if not sample:
        return "Empty"

    str_vals = [str(v).strip() for v in sample if str(v).strip() != '']
    if not str_vals:
        return "Empty"

    # First try the scanner's special detectors
    col_type = infer_column_type_by_content(df, col_index)
    if col_type == 'sex':
        return "Sex"
    elif col_type == 'dob':
        return "DOB"
    elif col_type == 'policynum':
        return "PolicyNumber"
    elif col_type == 'name':
        # Could be full name or first/last – we'll just call it Name
        return "Name"
    elif col_type == 'firstlastname':
        return "FirstLast"

    # Numeric detection
    numeric_count = 0
    for v in str_vals:
        try:
            float(v)
            numeric_count += 1
        except:
            pass
    if numeric_count / len(str_vals) > 0.9:
        return "Numeric"

    # Date detection (using the same date parser from scanner)
    from ofac_scanner_core import parse_date_to_mmddyyyy
    date_count = 0
    for v in str_vals:
        if parse_date_to_mmddyyyy(v) != '':
            date_count += 1
    if date_count / len(str_vals) > 0.7:
        return "Date"

    # Identifier detection (short alphanumeric, 4‑20 chars)
    id_pattern = re.compile(r'^[A-Za-z0-9\-_]{4,20}$')
    id_count = sum(1 for v in str_vals if id_pattern.match(v))
    if id_count / len(str_vals) > 0.7:
        return "Identifier"

    # Text detection (mostly letters and spaces)
    text_pattern = re.compile(r"^[A-Za-zÀ-ÿ'\-\.\s]{2,}$")
    text_count = sum(1 for v in str_vals if text_pattern.match(v))
    if text_count / len(str_vals) > 0.8:
        return "Text"

    return "Unknown"

# ----------------------------------------------------------------------
# Archive processing (isolated temp subfolder)
# ----------------------------------------------------------------------
def _process_archive(job, archive_path, archive_name):
    """Extract archive into a temporary subfolder, return list of extracted file paths."""
    extracted_files = []
    # Create isolated temp dir inside unzipped_folder
    extract_dir = tempfile.mkdtemp(prefix="hd_ext_", dir=job.unzipped_folder)
    try:
        success = False
        for pwd in job.passwords:
            result = subprocess.run(
                [SEVEN_ZIP_PATH, 'x', archive_path, '-aou', f'-o{extract_dir}', f'-p{pwd}'],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                success = True
                break
        if not success:
            result = subprocess.run(
                [SEVEN_ZIP_PATH, 'x', archive_path, '-aou', f'-o{extract_dir}'],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                raise Exception("Archive extraction failed")
        extracted_files = [os.path.join(root, f) for root, _, files in os.walk(extract_dir) for f in files]
    except Exception as e:
        write_log_row(job, {
            'File Path': archive_path,
            'File Name': archive_name,
            'Error Msg': str(e),
            'Remarks': 'Extraction failed'
        })
        # Clean up failed extraction dir
        shutil.rmtree(extract_dir, ignore_errors=True)
        return []

    write_log_row(job, {
        'File Path': archive_path,
        'File Name': archive_name,
        'Remarks': f"Extracted {len(extracted_files)} files"
    })
    return extracted_files

# ----------------------------------------------------------------------
# Process a single file (Excel/CSV)
# ----------------------------------------------------------------------
def _process_single_file(job, file_path, file_name, ext, all_headers):
    temp_path = None
    try:
        if ext in FILE_EXTENSIONS_EXCEL:
            decrypted, _ = unlock_excel(file_path, job.passwords)
            if isinstance(decrypted, io.BytesIO):
                temp_path = os.path.join(job.unzipped_folder, f"temp_{int(time.time())}_{file_name}")
                with open(temp_path, 'wb') as f:
                    f.write(decrypted.getvalue())
                actual_path = temp_path
            else:
                actual_path = decrypted
            sheets = pl.read_excel(actual_path, has_header=False, sheet_id=0, raise_if_empty=False)
        elif ext in FILE_EXTENSIONS_TEXT:
            actual_path = file_path
            df = pl.read_csv(actual_path, has_header=False, truncate_ragged_lines=True)
            sheets = {'': df}
        else:
            return

        for sheet_name, df in sheets.items():
            if df.is_empty():
                continue
            blocks = find_table_blocks(df)
            total_headers_this_sheet = 0
            for table_idx, (col_start, col_end) in enumerate(blocks, start=1):
                hdr_row, headers = detect_header_row_for_block(df, col_start, col_end)
                if headers:
                    for i, h in enumerate(headers):
                        abs_col = col_start + i
                        category = categorise_column(df, abs_col, hdr_row)
                        all_headers.append({
                            'Header': h,
                            'File Path': file_path,
                            'Sheet Name': sheet_name if sheet_name else '',
                            'Table': f"Table {table_idx}",
                            'Column Index': abs_col,
                            'Row Index': hdr_row,
                            'Category': category
                        })
                    total_headers_this_sheet += len(headers)

            if total_headers_this_sheet > 0:
                write_log_row(job, {
                    'File Path': file_path,
                    'File Name': file_name,
                    'Sheet Name': sheet_name if sheet_name else '',
                    'Headers Extracted': total_headers_this_sheet
                })
            else:
                write_log_row(job, {
                    'File Path': file_path,
                    'File Name': file_name,
                    'Sheet Name': sheet_name if sheet_name else '',
                    'Error Msg': 'No header row found'
                })

    except Exception as e:
        write_log_row(job, {
            'File Path': file_path,
            'File Name': file_name,
            'Error Msg': str(e)
        })
    finally:
        if temp_path:
            try:
                os.remove(temp_path)
            except:
                pass

# ----------------------------------------------------------------------
# Main extraction routine
# ----------------------------------------------------------------------
def process_header_extraction(job, progress_callback=None, stop_flag=None):
    if not os.path.exists(job.log_file) or os.path.getsize(job.log_file) == 0:
        with open(job.log_file, 'w', newline='', encoding='utf-8') as f:
            csv.DictWriter(f, fieldnames=LOG_HEADERS).writeheader()

    all_headers = []

    all_file_paths = [
        os.path.join(job.input_folder, f)
        for f in job.file_names
        if os.path.exists(os.path.join(job.input_folder, f))
    ]
    idx = 0
    files_processed = 0
    errors = 0

    while idx < len(all_file_paths):
        if stop_flag and stop_flag():
            write_log_row(job, {'Remarks': 'Extraction stopped by user'})
            break

        src_path = all_file_paths[idx]
        file_name = os.path.basename(src_path)
        if progress_callback:
            progress_callback(f"Processing {file_name}...")

        # Move to Archived (if enabled)
        if job.archive_files:
            try:
                work_path = move_file_to_archived(src_path, job.archived_folder)
            except Exception as e:
                write_log_row(job, {
                    'File Path': src_path, 'File Name': file_name, 'Error Msg': str(e)
                })
                errors += 1
                idx += 1
                continue
        else:
            work_path = src_path   # read directly

        ext = os.path.splitext(file_name)[1].lower().replace('.', '')

        if ext in FILE_EXTENSIONS_ARCHIVE:
            new_files = _process_archive(job, work_path, file_name)
            all_file_paths.extend(new_files)
        else:
            _process_single_file(job, work_path, file_name, ext, all_headers)
            files_processed += 1

        idx += 1

    # Write final output
    if all_headers:
        df = pl.DataFrame(all_headers)
        # Sort for readability
        try:
            df = df.sort(['File Path', 'Sheet Name', 'Table', 'Column Index'])
        except:
            pass
        try:
            df.write_excel(job.header_output_file)
            write_log_row(job, {
                'Remarks': (f"Extraction complete. Total headers: {len(all_headers)}. "
                            f"Files processed: {files_processed}, Errors: {errors}. "
                            f"Output: {job.header_output_file}")
            })
        except Exception as e:
            csv_out = job.header_output_file.replace('.xlsx', '.csv')
            df.write_csv(csv_out)
            write_log_row(job, {
                'Remarks': (f"Excel write failed ({e}). CSV saved to {csv_out}. "
                            f"Total headers: {len(all_headers)}")
            })
    else:
        write_log_row(job, {'Remarks': "No headers found in any file."})

    return True
