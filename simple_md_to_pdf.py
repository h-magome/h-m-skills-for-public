#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import markdown
import os
from pathlib import Path

def markdown_to_html(md_file, html_file):
    """マークダウンファイルをHTMLに変換する"""
    
    # マークダウンファイルを読み込み
    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    # マークダウンをHTMLに変換
    md = markdown.Markdown(extensions=['tables', 'fenced_code', 'toc'])
    html_content = md.convert(md_content)
    
    # CSSスタイルを定義（日本語フォント対応）
    css_style = """
    <style>
    @page {
        size: A4;
        margin: 2cm;
    }
    
    body {
        font-family: "Meiryo", "Yu Gothic", "Hiragino Sans", sans-serif;
        font-size: 12pt;
        line-height: 1.6;
        color: #333;
        max-width: 800px;
        margin: 0 auto;
        padding: 20px;
    }
    
    h1 {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 10px;
        margin-top: 30px;
        margin-bottom: 20px;
    }
    
    h2 {
        color: #34495e;
        border-bottom: 1px solid #bdc3c7;
        padding-bottom: 5px;
        margin-top: 25px;
        margin-bottom: 15px;
    }
    
    h3 {
        color: #7f8c8d;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    
    table {
        border-collapse: collapse;
        width: 100%;
        margin: 15px 0;
    }
    
    th, td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
    }
    
    th {
        background-color: #f8f9fa;
        font-weight: bold;
    }
    
    code {
        background-color: #f4f4f4;
        padding: 2px 4px;
        border-radius: 3px;
        font-family: "Consolas", "Monaco", monospace;
    }
    
    pre {
        background-color: #f8f8f8;
        padding: 10px;
        border-radius: 5px;
        overflow-x: auto;
        border: 1px solid #ddd;
    }
    
    ul, ol {
        margin: 10px 0;
        padding-left: 20px;
    }
    
    li {
        margin: 5px 0;
    }
    
    strong {
        color: #2c3e50;
    }
    
    blockquote {
        border-left: 4px solid #3498db;
        padding-left: 15px;
        margin: 15px 0;
        color: #555;
    }
    
    @media print {
        body {
            font-size: 10pt;
        }
        
        h1 {
            page-break-before: always;
        }
        
        h1:first-child {
            page-break-before: avoid;
        }
    }
    </style>
    """
    
    # HTMLテンプレート
    html_template = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>HM スキルシート</title>
        {css_style}
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """
    
    # HTMLファイルに出力
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_template)
    
    print(f"HTMLファイルが作成されました: {html_file}")
    print("このHTMLファイルをブラウザで開き、印刷機能でPDFに保存できます。")
    return html_file

def convert_files():
    """マークダウンファイルをHTMLに変換する"""
    md_files = [
        "HM_スキルシート.md",
        "README.md"
    ]
    
    converted_files = []
    
    for md_file in md_files:
        if os.path.exists(md_file):
            html_file = md_file.replace('.md', '.html')
            try:
                html_path = markdown_to_html(md_file, html_file)
                converted_files.append(html_path)
                print(f"✓ {md_file} → {html_file}")
            except Exception as e:
                print(f"✗ エラー: {md_file} の変換に失敗しました - {e}")
        else:
            print(f"✗ ファイルが見つかりません: {md_file}")
    
    if converted_files:
        print("\n" + "="*50)
        print("変換完了！")
        print("="*50)
        print("作成されたHTMLファイル:")
        for file in converted_files:
            print(f"  - {file}")
        print("\nPDFに変換する方法:")
        print("1. HTMLファイルをブラウザで開く")
        print("2. Ctrl+P (印刷) を押す")
        print("3. 送信先で「PDFに保存」を選択")
        print("4. 保存をクリック")

if __name__ == "__main__":
    convert_files() 