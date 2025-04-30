import os
import csv
import argparse
from typing import List, Dict, Tuple
import PyPDF2
from docx import Document
from pptx import Presentation
import pandas as pd
import tiktoken
import xlrd
import logging
import win32com.client
import pythoncom
from tqdm import tqdm

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def count_tokens(text: str) -> int:
    try:
        encoding = tiktoken.get_encoding("cl100k_base")
        return len(encoding.encode(text))
    except Exception as e:
        logger.error(f"Error counting tokens: {str(e)}")
        return 0

def process_pdf(file_path: str) -> str:
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            return " ".join(page.extract_text() or "" for page in pdf_reader.pages)
    except Exception as e:
        logger.error(f"Error processing PDF file {file_path}: {str(e)}")
        return ""

def process_doc(file_path: str) -> str:
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        logger.error(f"Error processing DOC file {file_path}: {str(e)}")
        return ""
    finally:
        pythoncom.CoUninitialize()

def process_docx(file_path: str) -> str:
    try:
        doc = Document(file_path)
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)
    except Exception as e:
        logger.error(f"Error processing DOCX file {file_path}: {str(e)}")
        return ""

def process_txt(file_path: str) -> str:
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            return file.read()
    except Exception as e:
        logger.error(f"Error processing TXT file {file_path}: {str(e)}")
        return ""

def process_pptx(file_path: str) -> str:
    try:
        prs = Presentation(file_path)
        return "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, 'text'))
    except Exception as e:
        logger.error(f"Error processing PPTX file {file_path}: {str(e)}")
        return ""

def process_excel(file_path: str) -> str:
    try:
        if file_path.endswith('.xls'):
            workbook = xlrd.open_workbook(file_path)
            sheets = [pd.DataFrame([[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]) 
                      for sheet in workbook.sheets()]
            df = pd.concat(sheets)
        else:
            df = pd.read_excel(file_path)
        return df.to_string()
    except Exception as e:
        logger.error(f"Error processing Excel file {file_path}: {str(e)}")
        return ""

def process_csv(file_path: str) -> str:
    try:
        df = pd.read_csv(file_path)
        return df.to_string()
    except Exception as e:
        logger.error(f"Error processing CSV file {file_path}: {str(e)}")
        return ""

def process_file(file_path: str) -> int:
    processors = {
        '.pdf': process_pdf,
        '.doc': process_doc,
        '.docx': process_docx,
        '.txt': process_txt,
        '.ppt': process_pptx,
        '.pptx': process_pptx,
        '.xls': process_excel,
        '.xlsx': process_excel,
        '.csv': process_csv
    }
    
    try:
        ext = os.path.splitext(file_path)[1].lower()
        processor = processors.get(ext)
        if processor:
            text = processor(file_path)
            return count_tokens(text)
        else:
            logger.warning(f"Unsupported file type: {file_path}")
            return 0
    except Exception as e:
        logger.error(f"Error processing file {file_path}: {str(e)}")
        return 0

def process_folder(folder_path: str) -> Tuple[List[Dict[str, str]], int, int]:
    results = []
    total_tokens = 0
    total_files = 0
    
    for root, _, files in os.walk(folder_path):
        for file in tqdm(files, desc=f"Processing {os.path.basename(root)}", unit="file"):
            file_path = os.path.join(root, file)
            token_count = process_file(file_path)
            if token_count > 0:
                total_tokens += token_count
                total_files += 1
                results.append({
                    'file_path': file_path,
                    'token_count': token_count
                })

    return results, total_tokens, total_files

def create_csv_report(results: List[Dict[str, str]], total_tokens: int, total_files: int, output_path: str):
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['file_path', 'token_count']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                writer.writerow(result)
            
            writer.writerow({'file_path': 'Total', 'token_count': total_tokens})
            writer.writerow({'file_path': 'Total Files', 'token_count': total_files})
        logger.info(f"Report saved to: {output_path}")
    except Exception as e:
        logger.error(f"Error creating CSV report: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description="Count tokens in files across multiple folders.")
    parser.add_argument('folders', nargs='+', help='Full path to source folders to process')
    parser.add_argument('-o', '--output', required=True, help='Full path to output CSV file')
    args = parser.parse_args()

    if not args.output.lower().endswith('.csv'):
        args.output += '.csv'

    logger.info("Token Counter")
    logger.info("Input folders:")
    for folder in args.folders:
        logger.info(f"- {folder}")
    logger.info(f"Output file: {args.output}")

    all_results = []
    total_tokens = 0
    total_files = 0

    for folder in args.folders:
        logger.info(f"Processing folder: {folder}")
        results, tokens, files = process_folder(folder)
        all_results.extend(results)
        total_tokens += tokens
        total_files += files

    create_csv_report(all_results, total_tokens, total_files, args.output)

    logger.info("Processing complete!")
    logger.info(f"Total tokens: {total_tokens}")
    logger.info(f"Total files processed: {total_files}")

if __name__ == "__main__":
    main()