# Document Token Counter

This Python script processes various types of documents (PDF, DOCX, TXT, PPTX, XLS/XLSX, CSV) and counts the tokens in each file using the OpenAI tiktoken library. It can process multiple folders and generate a CSV report with token counts for each file and overall statistics.

## Features

- Supports multiple file formats: PDF, DOCX, TXT, PPTX, XLS/XLSX, CSV
- Processes files recursively in specified folders
- Generates a CSV report with token counts for each file
- Provides total token count and total number of files processed

## Requirements

- Python 3.6+
- PyPDF2
- python-docx
- python-pptx
- pandas
- tiktoken
- openpyxl (for Excel file support)

## Installation

1. Clone this repository:
```
git clone https://github.com/yourusername/document-token-counter.git
cd document-token-counter
```

2. Install the required packages:
pip install -r requirements.txt


## Usage

Run the script from the command line:


python token_counter.py "D:\Docs" "D:\More" -o report.csv --workers 8 --quiet
# optional:
#   --model gpt-4o  OR  --encoding cl100k_base
#   --excel-max-rows 50000  --csv-max-rows 100000


- You can specify multiple input folders.
- Use the `-o` or `--output` flag to specify the output CSV file path.

## Output

The script will generate a CSV file with the following columns:
- `file_path`: The path to each processed file
- `token_count`: The number of tokens in the file

At the end of the CSV, there will be two additional rows:
- Total token count across all files
- Total number of files processed
