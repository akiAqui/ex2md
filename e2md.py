import pandas as pd
from tabulate import tabulate
import os
import sys

def escape_newlines(df):
    """セル内の改行をエスケープする"""
    return df.applymap(lambda x: x.replace('\n', '\\n') if isinstance(x, str) else x)

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
    print(f"Markdown file saved as {output_file}, encoding: {encoding}")

def process_excel_to_markdown(input_file, output_file, encoding='utf-8'):
    """Excelファイルを読み込み、Markdownに変換して保存"""
    df = pd.read_excel(input_file, na_filter=False)  # ExcelデータをDataFrameとして読み込む
    markdown_array = convert_to_markdown_array(df)  # DataFrameをMarkdown形式の配列に変換
    save_markdown(markdown_array, output_file, encoding)  # Markdownデータをファイルに保存

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <inputfile.xlsx> <outputfile.md> [--encoding=utf-8/sjis]")
        sys.exit()
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    encoding = 'utf-8'  # デフォルトのエンコーディングをUTF-8に設定
    if '--encoding=sjis' in sys.argv:
        encoding = 'shift_jis'

    process_excel_to_markdown(input_file, output_file, encoding)

if __name__ == "__main__":
    main()
