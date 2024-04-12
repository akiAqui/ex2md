import pandas as pd
from tabulate import tabulate
import os
import sys
import logging
import argparse

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


def save_markdown(markdown_array, output_file, encoding):
    """Markdown配列をファイルに保存"""
    with open(output_file, 'w', encoding=encoding, errors='replace') as f:
        for line in markdown_array:
            f.write(line + '\n')
    logging.info(f"Markdown file saved as {output_file}, encoding: {encoding}")

    
def process_excel_to_markdown(input_file, output_file, encoding='utf-8'):
    """Excelファイルを読み込み、すべてのシートをMarkdownに変換して保存"""
    xls = pd.ExcelFile(input_file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name, na_filter=False)
        markdown_array = convert_to_markdown_array(df)
        sheet_output_file = f"{output_file}_{sheet_name}"
        save_markdown(markdown_array, sheet_output_file, encoding)

        
def main():
    parser = argparse.ArgumentParser(description="Convert Excel sheets to Markdown format.")
    parser.add_argument('input_file', help="Input Excel file.")
    parser.add_argument('output_file', help="Output Markdown file base name.")
    parser.add_argument('--encoding', default='utf-8', help="Character encoding for the output file (default: utf-8).")
    parser.add_argument('--verbose', action='store_true', help="Enable verbose output.")
    
    args = parser.parse_args()
    
    # ロギングレベルの設定
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    logging.info("Starting conversion process.")
    process_excel_to_markdown(args.input_file, args.output_file, args.encoding)
    logging.info("Conversion process completed.")

if __name__ == "__main__":
    main()


