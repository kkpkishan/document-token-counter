import os
import csv
import argparse
from typing import List, Dict, Tuple
import PyPDF2
from docx import Document
from pptx import Presentation
import pandas as pd
import tiktoken

def count_tokens(text: str) -> int:
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(text))

def process_pdf(file_path: str) -> str:
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
    return text

def process_docx(file_path: str) -> str:
    doc = Document(file_path)
    return "\n".join([paragraph.text for paragraph in doc.paragraphs])

def process_txt(file_path: str) -> str:
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        return file.read()

def process_pptx(file_path: str) -> str:
    prs = Presentation(file_path)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                text += shape.text + "\n"
    return text

def process_excel(file_path: str) -> str:
    df = pd.read_excel(file_path)
    return df.to_string()

def process_csv(file_path: str) -> str:
    df = pd.read_csv(file_path)
    return df.to_string()

def process_file(file_path: str) -> int:
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.pdf':
            text = process_pdf(file_path)
        elif ext == '.docx':
            text = process_docx(file_path)
        elif ext == '.txt':
            text = process_txt(file_path)
        elif ext in ['.ppt', '.pptx']:
            text = process_pptx(file_path)
        elif ext in ['.xls', '.xlsx']:
            text = process_excel(file_path)
        elif ext == '.csv':
            text = process_csv(file_path)
        else:
            print(f"Unsupported file type: {file_path}")
            return 0
        return count_tokens(text)
    except Exception as e:
        print(f"Error processing file {file_path}: {str(e)}")
        return 0

def process_folder(folder_path: str) -> Tuple[List[Dict[str, str]], int, int]:
    results = []
    total_tokens = 0
    total_files = 0
    
    for root, _, files in os.walk(folder_path):
        for file in files:
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
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['file_path', 'token_count']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for result in results:
            writer.writerow(result)
        
        writer.writerow({'file_path': 'Total', 'token_count': total_tokens})
        writer.writerow({'file_path': 'Total Files', 'token_count': total_files})

def main():
    parser = argparse.ArgumentParser(description="Count tokens in files across multiple folders.")
    parser.add_argument('folders', nargs='+', help='Full path to source folders to process')
    parser.add_argument('-o', '--output', required=True, help='Full path to output CSV file')
    args = parser.parse_args()

    if not args.output.lower().endswith('.csv'):
        args.output += '.csv'

    print("\n" + "="*50)
    print("Token Counter")
    print("="*50)
    print("\nInput folders:")
    for folder in args.folders:
        print(f"- {folder}")
    print(f"\nOutput file: {args.output}")
    print("="*50 + "\n")

    all_results = []
    total_tokens = 0
    total_files = 0

    for folder in args.folders:
        print(f"Processing folder: {folder}")
        results, tokens, files = process_folder(folder)
        all_results.extend(results)
        total_tokens += tokens
        total_files += files

    create_csv_report(all_results, total_tokens, total_files, args.output)

    print("\n" + "="*50)
    print(f"Processing complete!")
    print(f"Total tokens: {total_tokens}")
    print(f"Total files processed: {total_files}")
    print(f"Report saved to: {args.output}")
    print("="*50 + "\n")

if __name__ == "__main__":
    main()
