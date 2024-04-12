import pandas as pd
from tabulate import tabulate
import os
import sys
import logging
import argparse
import requests

# Loggerの設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def excel_range_to_params(range_str):
    start_col, start_row = '', ''
    end_col, end_row = '', ''
    parsing_end_col = False

    # 開始列と終了列の解析
    for char in range_str:
        if char.isdigit():
            if not parsing_end_col:
                start_row += char
            else:
                end_row += char
        elif char == ':':
            parsing_end_col = True
        elif not parsing_end_col:
            start_col += char
        else:
            end_col += char
    
    # 列の範囲を 'A:D' の形式で計算
    col_range = f"{start_col}:{end_col}"

    # skiprowsは開始行-1、nrowsは終了行-開始行+1
    skiprows = int(start_row) - 1
    nrows = int(end_row) - skiprows

    return {
        'usecols': col_range,
        'skiprows': skiprows,
        'nrows': nrows
    }

def range_checker(range_str):
    """Check if the given range is valid."""
    try:
        params = excel_range_to_params(range_str)
        return True
    except:
        return False

def parse_sheet_numbers(sheet_numbers):
    """Parse sheet numbers from command line argument."""
    sheets = []
    for item in sheet_numbers.split(','):
        if '-' in item:
            start, end = map(int, item.split('-'))
            sheets.extend(list(range(start, end+1)))
        else:
            sheets.append(int(item))
    return [{'sheet': sheet, 'range': ''} for sheet in sheets]

def parse_sheet_names(sheet_names):
    """Parse sheet names and ranges from command line argument."""
    sheets = []
    for item in sheet_names.split(','):
        if '!' in item:
            sheet, range_str = item.split('!')
            if not range_checker(range_str):
                raise ValueError(f"Invalid range: {range_str}")
            sheets.append({'sheet': sheet.strip(), 'range': range_str})
        else:
            sheets.append({'sheet': item.strip(), 'range': ''})
    return sheets

def get_secret_from_env():
    return os.environ.get('NOTEPM_SECRET')

def post_markdown_to_api(markdown_data, file_name):
    url = 'https://notepm.jp/docs/api/note/api/v1/notes'
    headers = {
        'Authorization': f'Bearer {get_secret_from_env()}',
        'Content-Type': 'application/json'
    }
    data = {
        'title': file_name,
        'content': markdown_data
    }
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        logging.info(f"Markdown data posted to API successfully for {file_name}")
    except requests.exceptions.RequestException as e:
        logging.error(f"Error posting markdown data to API for {file_name}: {e}")

def escape_newlines(df):
    """セル内の改行をエスケープする"""
    return df.applymap(lambda x: x.replace('\n', '<br>') if isinstance(x, str) else x)

def clean_markdown_lines(markdown_array):
    """Markdownテキストの余分なスペースを削除"""
    cleaned_lines = []
    for line in markdown_array:
        if '|' in line:
            parts = line.split('|')
            cleaned_parts = [part.strip() for part in parts]
            cleaned_line = '|'.join(cleaned_parts)
            cleaned_lines.append(cleaned_line)
        else:
            cleaned_lines.append(line)
    return cleaned_lines

def convert_to_markdown_array(df):
    """DataFrameからMarkdown形式の配列に変換"""
    df_escaped = escape_newlines(df.fillna(''))
    markdown_array = tabulate(df_escaped, tablefmt="pipe", headers="keys", showindex=False).split('\n')
    return clean_markdown_lines(markdown_array)


def save_markdown(markdown_array, output_file, encoding, upload, save_as_file):
    """Markdown配列をファイルに保存し、APIにPOST"""
    markdown_data = '\n'.join(markdown_array)
    if save_as_file or not upload:
        with open(output_file, 'w', encoding=encoding, errors='replace') as f:
            f.write(markdown_data)
        logging.info(f"Markdown file saved as {output_file}, encoding: {encoding}")
    if upload:
        post_markdown_to_api(markdown_data, os.path.basename(output_file))

    
def process_excel_to_markdown(input_file, output_file, sheets, encoding, upload, save_as_file):
    """Excelファイルを読み込み、指定されたシートをMarkdownに変換して保存"""
    try:
        xls = pd.ExcelFile(input_file)
        for sheet in sheets:
            sheet_name = sheet['sheet']
            range_str = sheet['range']
            try:
                if isinstance(sheet_name, int):
                    sheet_name = xls.sheet_names[sheet_name - 1]
                if range_str:
                    params = excel_range_to_params(range_str)
                    df = pd.read_excel(xls, sheet_name=sheet_name, na_filter=False, **params)
                else:
                    df = pd.read_excel(xls, sheet_name=sheet_name, na_filter=False)
                markdown_array = convert_to_markdown_array(df)
                sheet_output_file = f"{output_file}_{sheet_name}"
                save_markdown(markdown_array, sheet_output_file, encoding, upload, save_as_file)
            except pd.errors.XLRDError as e:
                logging.error(f"Error reading sheet {sheet_name}: {e}")
            except MemoryError:
                logging.error(f"Insufficient memory for sheet {sheet_name}. Consider reading the file in smaller chunks.")
            except Exception as e:
                logging.error(f"An unexpected error occurred while processing sheet {sheet_name}: {e}")
    except FileNotFoundError:
        logging.error("File not found. Please provide a valid file path.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")


def get_input_file():
    """対話的に入力ファイルのパスを取得"""
    while True:
        file_path = input("Enter the input file path (file://<file_path>): ")
        if file_path.startswith("file://"):
            file_path = file_path[7:]
            if os.path.isfile(file_path):
                return file_path
            else:
                logging.error("File not found. Please provide a valid file path.")
        else:
            logging.error("Invalid file path format. Please use the format file://<file_path>.")
        

def main():
    parser = argparse.ArgumentParser(description="Convert Excel sheets to Markdown format.",
                                     epilog="Examples:\n"
                                            "  python script.py\n"
                                            "  python script.py input.xlsx -v\n"
                                            "  python script.py input.xlsx -na \"Sheet1!A1:B10,Sheet2\"\n"
                                            "  python script.py input.xlsx -sn \"1-3,5\"\n"
                                            "  python script.py input.xlsx -f output_base_name\n"
                                            "  python script.py input.xlsx --encoding shift-jis\n"
                                            "  python script.py input.xlsx -u\n"
                                            "  python script.py input.xlsx --save-as-file\n"
                                            "  python script.py input.xlsx -u --save-as-file\n"
                                            "  python script.py input.xlsx -na \"Sheet1!A1:B10,Sheet2\" -f output_base -u --save-as-file --encoding utf-16",
                                     formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('input_file', nargs='?', help="Input Excel file.")
    parser.add_argument('-f', '--file-base-name', help="Base name for the output Markdown file.")
    parser.add_argument('-na', '--sheet-name', help="Comma-separated list of sheet names and ranges (e.g., 'Sheet1!A1:B10,Sheet2').")
    parser.add_argument('-sn', '--sheet-number', help="Comma-separated list of sheet numbers and ranges (e.g., '1-3,5').")
    parser.add_argument('--encoding', default='utf-8', help="Character encoding for the output file (default: utf-8).")
    parser.add_argument('-v', '--verbose', action='store_true', help="Enable verbose output.")
    parser.add_argument('-u', '--upload', action='store_true', help="Upload the converted Markdown to the NotePM API.")
    parser.add_argument('--save-as-file', action='store_true', help="Save the converted Markdown as a local file.")

    
    args = parser.parse_args()

    if args.sheet_name and args.sheet_number:
        parser.error("Cannot specify both --sheet-name and --sheet-number.")

    sheets = []
    if args.sheet_name:
        sheets = parse_sheet_names(args.sheet_name)
    elif args.sheet_number:
        sheets = parse_sheet_numbers(args.sheet_number)

    if not args.input_file:
        args.input_file = get_input_file()
    
    if not args.file_base_name:
        args.file_base_name = os.path.splitext(os.path.basename(args.input_file))[0]
    
    # ロギングレベルの設定
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    logging.info("Starting conversion process.")
    process_excel_to_markdown(args.input_file, args.file_base_name, sheets, args.encoding, args.upload, args.save_as_file)
    logging.info("Conversion process completed.")

if __name__ == "__main__":
    main()
