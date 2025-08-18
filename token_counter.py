#!/usr/bin/env python3
"""
Token Counter
-------------
Counts tokens in documents across one or more folders and writes a CSV report.

Supported types: .pdf, .docx, .txt, .pptx, .xls, .xlsx, .csv
(Windows-only optional: .doc, .ppt via MS Office COM if available)

Examples:
  python token_counter.py "C:\Data" D:\More -o report.csv
  python token_counter.py ./docs -o out.csv --model gpt-4o --excel-max-rows 10000
  python token_counter.py ./in -o out.csv --encoding cl100k_base

Notes:
- For .doc/.ppt you must be on Windows with MS Office installed.
- CSV schema remains: file_path, token_count
"""

import os
import csv
import sys
import argparse
import logging
import platform
from pathlib import Path
from typing import List, Dict, Tuple, Optional

# Progress bar
from tqdm import tqdm

# We import heavy/optional libs inside functions to keep startup light and avoid platform issues.
import tiktoken  # encoding is selected at runtime

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("token_counter")


# -----------------------------
# Tokenization helpers
# -----------------------------
def get_encoder(encoding: Optional[str], model: Optional[str]):
    """
    Choose a tiktoken encoder either by model or encoding name.
    Defaults to cl100k_base if nothing else works.
    """
    try:
        if model:
            return tiktoken.encoding_for_model(model)
    except Exception as e:
        logger.warning(f"Could not load encoding for model '{model}': {e}. Falling back...")
    try:
        if encoding:
            return tiktoken.get_encoding(encoding)
    except Exception as e:
        logger.warning(f"Could not load encoding '{encoding}': {e}. Falling back...")
    return tiktoken.get_encoding("cl100k_base")


def count_tokens(text: str, encoder) -> int:
    try:
        return len(encoder.encode(text or ""))
    except Exception as e:
        logger.error(f"Error counting tokens: {e}")
        return 0


# -----------------------------
# File processors
# -----------------------------
def process_pdf(file_path: str) -> str:
    try:
        # pypdf is the maintained successor to PyPDF2
        from pypdf import PdfReader
        with open(file_path, "rb") as fh:
            reader = PdfReader(fh)
            parts = []
            for page in reader.pages:
                try:
                    parts.append(page.extract_text() or "")
                except Exception:
                    parts.append("")
            return "\n".join(parts)
    except Exception as e:
        logger.error(f"Error processing PDF file {file_path}: {e}")
        return ""


def process_docx(file_path: str) -> str:
    try:
        from docx import Document
        doc = Document(file_path)
        parts = [p.text for p in doc.paragraphs]
        # include table text as well
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    if cell.text:
                        parts.append(cell.text)
        return "\n".join(filter(None, parts))
    except Exception as e:
        logger.error(f"Error processing DOCX file {file_path}: {e}")
        return ""


def process_txt(file_path: str) -> str:
    # try utf-8; if it fails, fall back to latin-1 rather than silently dropping bytes
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, "r", encoding="latin-1") as f:
                return f.read()
        except Exception as e:
            logger.error(f"Error processing TXT file {file_path}: {e}")
            return ""
    except Exception as e:
        logger.error(f"Error processing TXT file {file_path}: {e}")
        return ""


def process_pptx(file_path: str) -> str:
    try:
        from pptx import Presentation
        prs = Presentation(file_path)
        parts: List[str] = []
        for slide in prs.slides:
            for shape in slide.shapes:
                # python-pptx exposes .text for shapes with text frames/placeholders
                if hasattr(shape, "text") and shape.text:
                    parts.append(shape.text)
            # include slide notes if present
            if hasattr(slide, "has_notes_slide") and slide.has_notes_slide and slide.notes_slide:
                notes_frame = getattr(slide.notes_slide, "notes_text_frame", None)
                if notes_frame and notes_frame.text:
                    parts.append(notes_frame.text)
        return "\n".join(parts)
    except Exception as e:
        logger.error(f"Error processing PPTX file {file_path}: {e}")
        return ""


def process_excel(file_path: str, excel_max_rows: Optional[int]) -> str:
    """
    Reads all sheets and returns a CSV-like text joined across sheets.
    Limits rows per sheet if excel_max_rows is set.
    """
    try:
        import pandas as pd

        suffix = Path(file_path).suffix.lower()
        if suffix == ".xls":
            # xlrd still supports only .xls
            import xlrd  # noqa: F401
            # Read all sheets using pandas with xlrd engine
            xls = pd.read_excel(file_path, sheet_name=None, engine="xlrd", dtype=str, nrows=excel_max_rows)
        else:
            # .xlsx/.xlsm via openpyxl
            xls = pd.read_excel(file_path, sheet_name=None, engine="openpyxl", dtype=str, nrows=excel_max_rows)

        frames = list(xls.values())
        if not frames:
            return ""
        df = pd.concat(frames, ignore_index=True)
        text = df.to_csv(index=False, header=False)  # path_or_buf defaults to None -> returns string
        return text or ""
    except Exception as e:
        logger.error(f"Error processing Excel file {file_path}: {e}")
        return ""


def process_csv(file_path: str, csv_max_rows: Optional[int]) -> str:
    try:
        import pandas as pd
        df = pd.read_csv(file_path, dtype=str, nrows=csv_max_rows)
        text = df.to_csv(index=False, header=False)
        return text or ""
    except Exception as e:
        logger.error(f"Error processing CSV file {file_path}: {e}")
        return ""


def process_doc_ppt_via_com(file_path: str, app: str) -> str:
    """
    Windows-only MS Office COM extractor for legacy .doc and .ppt.
    Returns "" on any error or if platform/app is unavailable.
    """
    if platform.system() != "Windows":
        logger.warning(f"Skipping {file_path}: .{Path(file_path).suffix.lstrip('.')} requires Windows/MS Office.")
        return ""
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        logger.warning(f"Skipping {file_path}: pywin32 not installed or COM unavailable.")
        return ""

    try:
        pythoncom.CoInitialize()
        if app == "Word":
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(str(file_path))
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            return text
        elif app == "PowerPoint":
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt.Presentations.Open(str(file_path), ReadOnly=True)
            parts = []
            for slide in pres.Slides:
                for shape in slide.Shapes:
                    if getattr(shape, "HasTextFrame", False):
                        try:
                            parts.append(shape.TextFrame.TextRange.Text)
                        except Exception:
                            pass
            pres.Close()
            ppt.Quit()
            return "\n".join(parts)
        else:
            return ""
    except Exception as e:
        logger.error(f"Error processing {Path(file_path).suffix} file {file_path} via COM: {e}")
        return ""
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# -----------------------------
# Orchestration
# -----------------------------
def process_file(
    file_path: str,
    encoder,
    excel_max_rows: Optional[int],
    csv_max_rows: Optional[int],
) -> int:
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        text = process_pdf(file_path)
    elif ext == ".docx":
        text = process_docx(file_path)
    elif ext == ".txt":
        text = process_txt(file_path)
    elif ext == ".pptx":
        text = process_pptx(file_path)
    elif ext in (".xls", ".xlsx"):
        text = process_excel(file_path, excel_max_rows=excel_max_rows)
    elif ext == ".csv":
        text = process_csv(file_path, csv_max_rows=csv_max_rows)
    elif ext == ".doc":
        text = process_doc_ppt_via_com(file_path, app="Word")
    elif ext == ".ppt":
        text = process_doc_ppt_via_com(file_path, app="PowerPoint")
    else:
        logger.warning(f"Unsupported file type: {file_path}")
        return 0

    return count_tokens(text, encoder)


def iter_files(root: str) -> List[str]:
    files = []
    for dirpath, _, filenames in os.walk(root):
        for name in filenames:
            files.append(os.path.join(dirpath, name))
    return files


def process_folder(
    folder_path: str,
    encoder,
    excel_max_rows: Optional[int],
    csv_max_rows: Optional[int],
) -> Tuple[List[Dict[str, str]], int, int]:
    results: List[Dict[str, str]] = []
    total_tokens = 0
    total_files = 0

    all_files = iter_files(folder_path)
    for file_path in tqdm(all_files, desc=f"Processing {os.path.basename(folder_path) or folder_path}", unit="file"):
        try:
            token_count = process_file(file_path, encoder, excel_max_rows, csv_max_rows)
            total_tokens += token_count
            total_files += 1
            results.append({"file_path": file_path, "token_count": str(token_count)})
        except KeyboardInterrupt:
            raise
        except Exception as e:
            logger.error(f"Unhandled error processing {file_path}: {e}")
            results.append({"file_path": file_path, "token_count": "0"})

    return results, total_tokens, total_files


def create_csv_report(results: List[Dict[str, str]], total_tokens: int, total_files: int, output_path: str):
    try:
        with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
            fieldnames = ["file_path", "token_count"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in results:
                writer.writerow(row)
            # Footer lines for totals (preserving your original format)
            writer.writerow({"file_path": "Total", "token_count": str(total_tokens)})
            writer.writerow({"file_path": "Total Files", "token_count": str(total_files)})
        logger.info(f"Report saved to: {output_path}")
    except Exception as e:
        logger.error(f"Error creating CSV report: {e}")


def main():
    parser = argparse.ArgumentParser(description="Count tokens in files across multiple folders.")
    parser.add_argument("folders", nargs="+", help="Full path(s) to source folders to process")
    parser.add_argument("-o", "--output", required=True, help="Full path to output CSV file")

    enc = parser.add_argument_group("Tokenization")
    enc.add_argument("--model", default=None, help="Model name to select encoding (e.g., gpt-4o, gpt-4o-mini)")
    enc.add_argument("--encoding", default=None, help="Explicit tiktoken encoding (e.g., cl100k_base, o200k_base)")

    limits = parser.add_argument_group("Row limits")
    limits.add_argument("--excel-max-rows", type=int, default=None, help="Max rows per sheet to read from Excel")
    limits.add_argument("--csv-max-rows", type=int, default=None, help="Max rows to read from CSV")

    args = parser.parse_args()

    output_path = args.output if args.output.lower().endswith(".csv") else f"{args.output}.csv"

    logger.info("Token Counter")
    logger.info("Input folders:")
    for folder in args.folders:
        logger.info(f"- {folder}")
    logger.info(f"Output file: {output_path}")

    # choose encoder
    encoder = get_encoder(args.encoding, args.model)

    all_results: List[Dict[str, str]] = []
    grand_tokens = 0
    grand_files = 0

    for folder in args.folders:
        logger.info(f"Processing folder: {folder}")
        results, tokens, files = process_folder(folder, encoder, args.excel_max_rows, args.csv_max_rows)
        all_results.extend(results)
        grand_tokens += tokens
        grand_files += files

    create_csv_report(all_results, grand_tokens, grand_files, output_path)

    logger.info("Processing complete!")
    logger.info(f"Total tokens: {grand_tokens}")
    logger.info(f"Total files processed: {grand_files}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("Interrupted by user. Exiting...")
        sys.exit(130)
