#!/usr/bin/env python3
"""
PDFファイルからテキストを抽出し、Markdown形式に整形するスクリプト

テキストベースのPDFを解析し、見出し・テーブル・リストを含む
構造化されたMarkdownファイルを生成します。

使い方:
    python simple_pdf_to_markdown.py input.pdf [output.md]

機能:
    - テキスト抽出とフォントサイズに基づく見出し判定
    - テーブル（表）の自動検出とMarkdown表形式への変換
    - 箇条書き・番号付きリストの検出と変換
    - 段落の整形
"""

import re
import sys
import os
from collections import Counter

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber がインストールされていません。")
    print("以下のコマンドでインストールしてください:")
    print("  pip install pdfplumber")
    sys.exit(1)


# 設定
CONFIG = {
    # 見出し判定のフォントサイズ閾値（相対値）
    'heading_size_ratio': {
        'h1': 1.5,   # 本文の1.5倍以上 → H1
        'h2': 1.3,   # 本文の1.3倍以上 → H2
        'h3': 1.15,  # 本文の1.15倍以上 → H3
    },
    # リスト検出パターン
    'bullet_patterns': [
        r'^[\s]*[•・●○◆◇▪▫■□►▸‣⁃]\s*',  # 記号による箇条書き
        r'^[\s]*[-\*]\s+',                    # - または * による箇条書き
    ],
    'numbered_patterns': [
        r'^[\s]*(\d+)[\.）\)]\s*',           # 1. 2. または 1) 2)
        r'^[\s]*[\(（](\d+)[\)）]\s*',       # (1) (2) または （1）（2）
        r'^[\s]*([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])\s*',  # 丸数字
    ],
}


def extract_font_sizes(page):
    """
    ページ内のフォントサイズを抽出し、本文サイズを推定する
    """
    chars = page.chars
    if not chars:
        return None
    
    sizes = [char.get('size', 0) for char in chars if char.get('size')]
    if not sizes:
        return None
    
    # 最も頻出するフォントサイズを本文サイズとする
    size_counter = Counter(round(s, 1) for s in sizes)
    body_size = size_counter.most_common(1)[0][0]
    
    return body_size


def get_heading_level(font_size, body_size):
    """
    フォントサイズから見出しレベルを判定
    """
    if not body_size or not font_size:
        return None
    
    ratio = font_size / body_size
    
    if ratio >= CONFIG['heading_size_ratio']['h1']:
        return 1
    elif ratio >= CONFIG['heading_size_ratio']['h2']:
        return 2
    elif ratio >= CONFIG['heading_size_ratio']['h3']:
        return 3
    
    return None


def is_bullet_list(text):
    """
    箇条書きリストかどうかを判定
    """
    for pattern in CONFIG['bullet_patterns']:
        if re.match(pattern, text):
            return True
    return False


def is_numbered_list(text):
    """
    番号付きリストかどうかを判定
    """
    for pattern in CONFIG['numbered_patterns']:
        if re.match(pattern, text):
            return True
    return False


def clean_bullet_text(text):
    """
    箇条書きの記号を除去してクリーンなテキストを返す
    """
    for pattern in CONFIG['bullet_patterns']:
        text = re.sub(pattern, '', text)
    return text.strip()


def clean_numbered_text(text):
    """
    番号付きリストの番号を除去してクリーンなテキストを返す
    """
    for pattern in CONFIG['numbered_patterns']:
        text = re.sub(pattern, '', text)
    return text.strip()


def extract_text_with_formatting(page, body_size):
    """
    ページからテキストを抽出し、フォーマット情報を付加
    """
    # 行ごとにテキストと位置情報を取得
    lines = []
    
    # extract_text()で基本テキストを取得
    text = page.extract_text()
    if not text:
        return []
    
    # 文字情報から行ごとのフォントサイズを推定
    chars = page.chars
    if not chars:
        # 文字情報がない場合はプレーンテキストとして返す
        for line in text.split('\n'):
            if line.strip():
                lines.append({
                    'text': line.strip(),
                    'font_size': None,
                    'is_bold': False,
                })
        return lines
    
    # 行ごとにグループ化（Y座標でグループ化）
    char_lines = {}
    for char in chars:
        y = round(char.get('top', 0), 0)
        if y not in char_lines:
            char_lines[y] = []
        char_lines[y].append(char)
    
    # 各行のテキストとフォント情報を抽出
    for y in sorted(char_lines.keys()):
        line_chars = sorted(char_lines[y], key=lambda c: c.get('x0', 0))
        
        # 行のテキストを構築
        line_text = ''.join(c.get('text', '') for c in line_chars)
        line_text = line_text.strip()
        
        if not line_text:
            continue
        
        # 行の主要フォントサイズを取得
        line_sizes = [c.get('size', 0) for c in line_chars if c.get('size')]
        avg_size = sum(line_sizes) / len(line_sizes) if line_sizes else None
        
        # 太字判定（フォント名に'Bold'が含まれるか）
        font_names = [c.get('fontname', '') for c in line_chars]
        is_bold = any('Bold' in fn or 'bold' in fn for fn in font_names)
        
        lines.append({
            'text': line_text,
            'font_size': avg_size,
            'is_bold': is_bold,
        })
    
    return lines


def table_to_markdown(table):
    """
    テーブルをMarkdown形式に変換
    """
    if not table or not table[0]:
        return ""
    
    md_lines = []
    
    # ヘッダー行
    header = table[0]
    header_cells = [str(cell) if cell else '' for cell in header]
    md_lines.append('| ' + ' | '.join(header_cells) + ' |')
    
    # 区切り行
    md_lines.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
    
    # データ行
    for row in table[1:]:
        cells = [str(cell) if cell else '' for cell in row]
        # セル内の改行をスペースに置換
        cells = [cell.replace('\n', ' ') for cell in cells]
        md_lines.append('| ' + ' | '.join(cells) + ' |')
    
    return '\n'.join(md_lines)


def process_page(page, body_size):
    """
    1ページを処理してMarkdownテキストを返す
    """
    md_parts = []
    
    # テーブルを先に抽出（テーブル領域を特定するため）
    tables = page.extract_tables()
    table_bboxes = []
    
    if tables:
        # テーブルの境界ボックスを取得
        table_finder = page.find_tables()
        for t in table_finder:
            table_bboxes.append(t.bbox)
    
    # テーブルをMarkdownに変換
    for table in tables:
        md_table = table_to_markdown(table)
        if md_table:
            md_parts.append(('table', md_table))
    
    # テキストを抽出
    lines = extract_text_with_formatting(page, body_size)
    
    current_list_type = None  # 'bullet', 'numbered', None
    numbered_counter = 0
    paragraph_lines = []
    
    for line_info in lines:
        text = line_info['text']
        font_size = line_info['font_size']
        is_bold = line_info['is_bold']
        
        # 見出し判定
        heading_level = get_heading_level(font_size, body_size)
        
        if heading_level:
            # 段落をフラッシュ
            if paragraph_lines:
                md_parts.append(('paragraph', ' '.join(paragraph_lines)))
                paragraph_lines = []
            current_list_type = None
            
            md_parts.append(('heading', '#' * heading_level + ' ' + text))
            continue
        
        # 箇条書きリスト判定
        if is_bullet_list(text):
            if paragraph_lines:
                md_parts.append(('paragraph', ' '.join(paragraph_lines)))
                paragraph_lines = []
            
            clean_text = clean_bullet_text(text)
            md_parts.append(('bullet', '- ' + clean_text))
            current_list_type = 'bullet'
            continue
        
        # 番号付きリスト判定
        if is_numbered_list(text):
            if paragraph_lines:
                md_parts.append(('paragraph', ' '.join(paragraph_lines)))
                paragraph_lines = []
            
            if current_list_type != 'numbered':
                numbered_counter = 0
            numbered_counter += 1
            
            clean_text = clean_numbered_text(text)
            md_parts.append(('numbered', f'{numbered_counter}. ' + clean_text))
            current_list_type = 'numbered'
            continue
        
        # 通常のテキスト（段落）
        current_list_type = None
        
        # 太字テキストの処理
        if is_bold and len(text) < 100:  # 短い太字は強調として扱う
            if paragraph_lines:
                md_parts.append(('paragraph', ' '.join(paragraph_lines)))
                paragraph_lines = []
            md_parts.append(('bold', f'**{text}**'))
        else:
            paragraph_lines.append(text)
    
    # 残りの段落をフラッシュ
    if paragraph_lines:
        md_parts.append(('paragraph', ' '.join(paragraph_lines)))
    
    return md_parts


def convert_pdf_to_markdown(pdf_path):
    """
    PDFファイルをMarkdownに変換
    """
    md_output = []
    
    with pdfplumber.open(pdf_path) as pdf:
        # 全ページの本文フォントサイズを推定
        all_sizes = []
        for page in pdf.pages:
            chars = page.chars
            if chars:
                sizes = [char.get('size', 0) for char in chars if char.get('size')]
                all_sizes.extend(sizes)
        
        if all_sizes:
            size_counter = Counter(round(s, 1) for s in all_sizes)
            body_size = size_counter.most_common(1)[0][0]
        else:
            body_size = 12  # デフォルト値
        
        # 各ページを処理
        for page_num, page in enumerate(pdf.pages, 1):
            page_parts = process_page(page, body_size)
            
            prev_type = None
            for part_type, content in page_parts:
                # 適切な空行を挿入
                if prev_type and prev_type != part_type:
                    if part_type in ('heading', 'paragraph', 'table'):
                        md_output.append('')
                
                md_output.append(content)
                prev_type = part_type
            
            # ページ間に空行を追加
            if page_num < len(pdf.pages):
                md_output.append('')
    
    return '\n'.join(md_output)


def main():
    """
    メイン関数
    """
    if len(sys.argv) < 2:
        print("使い方: python simple_pdf_to_markdown.py <入力ファイル.pdf> [出力ファイル.md]")
        print()
        print("例:")
        print("  python simple_pdf_to_markdown.py document.pdf")
        print("  python simple_pdf_to_markdown.py document.pdf output.md")
        sys.exit(1)
    
    input_path = sys.argv[1]
    
    # 入力ファイルの存在確認
    if not os.path.exists(input_path):
        print(f"エラー: ファイルが見つかりません: {input_path}")
        sys.exit(1)
    
    # 出力ファイルパスの決定
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        # 入力ファイル名から自動生成
        base_name = os.path.splitext(input_path)[0]
        output_path = base_name + '.md'
    
    print(f"変換中: {input_path} → {output_path}")
    
    try:
        markdown_content = convert_pdf_to_markdown(input_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        print(f"完了: {output_path} を作成しました")
        
    except Exception as e:
        print(f"エラー: 変換中にエラーが発生しました: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
