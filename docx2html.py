import sys
import os
import zipfile
import re
from lxml import etree

def convert_docx_to_html(docx_path):
    # Unzip the docx file
    with zipfile.ZipFile(docx_path, "r") as z:
        with z.open("word/document.xml") as f:
            document_xml = f.read()

    # Parse the XML document
    root = etree.fromstring(document_xml)

    # Define namespaces
    namespaces = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }

    # Extract and convert paragraphs
    html_content = []
    for p in root.xpath("//w:p", namespaces=namespaces):
        text_runs = []

        namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        left_chars_key = f'{namespace}leftChars'

        p_start = '<p>'
        ind = p.xpath(".//w:ind", namespaces=namespaces)

        if ind and left_chars_key in ind[0].attrib:
            p_start = '<p class="citation">' if ind[0].attrib[left_chars_key] == '100' else '<p>'

        # フィールド文字を使用したルビを処理するための変数
        in_field = False
        field_ruby_text = ""
        field_ruby_base = ""
        
        # 通常のテキスト実行とルビを含むテキスト実行を別々に処理
        all_runs = p.xpath(".//w:r", namespaces=namespaces)
        
        for r_idx, r in enumerate(all_runs):
            # Check if the run is inside a ruby element
            is_in_ruby = len(r.xpath("ancestor::w:ruby", namespaces=namespaces)) > 0
            
            # Skip runs that are inside ruby elements as they will be processed separately
            if is_in_ruby and not r.xpath("ancestor::w:rubyBase", namespaces=namespaces) and not r.xpath("ancestor::w:rt", namespaces=namespaces):
                continue
            
            # フィールド文字を使用したルビを処理
            fld_char = r.xpath(".//w:fldChar", namespaces=namespaces)
            instr_text = r.xpath(".//w:instrText/text()", namespaces=namespaces)
            
            if fld_char:
                fld_type = fld_char[0].get(f"{{{namespaces['w']}}}fldCharType")
                if fld_type == "begin":
                    in_field = True
                    field_ruby_text = ""
                    field_ruby_base = ""
                elif fld_type == "end" and in_field and field_ruby_text and field_ruby_base:
                    text_runs.append(f'<ruby>{field_ruby_base}<rt>{field_ruby_text}</rt></ruby>')
                    in_field = False
            elif in_field and instr_text:
                for text in instr_text:
                    if "EQ" in text and "\\o\\ad" in text:
                        # フィールドの開始部分は無視
                        continue
                    elif ")," in text:
                        # ベーステキストを抽出
                        match = re.search(r'\),(.*?)\)', text)
                        if match:
                            field_ruby_base = match.group(1)
                    elif text.strip() and not text.startswith("\\"):
                        # ルビテキストを抽出
                        field_ruby_text = text.strip()
                
                # 通常のテキスト処理はスキップ
                continue
                
            # Process ruby elements at the run level
            ruby_elements = r.xpath(".//w:ruby", namespaces=namespaces)
            if ruby_elements:
                for ruby_elem in ruby_elements:
                    ruby_base = ruby_elem.xpath(".//w:rubyBase//w:t/text()", namespaces=namespaces)
                    ruby_text = ruby_elem.xpath(".//w:rt//w:t/text()", namespaces=namespaces)
                    
                    if ruby_base and ruby_text:
                        text = f'<ruby>{ruby_base[0]}<rt>{ruby_text[0]}</rt></ruby>'
                        text_runs.append(text)
            else:
                # Normal text run
                text = "".join(r.xpath(".//w:t/text()", namespaces=namespaces))
                
                # Apply color if it's "00B050"
                color = r.xpath(".//w:rPr/w:color/@w:val", namespaces=namespaces)
                if color and color[0] == "00B050":
                    text = f'<span class="ref">{text}</span>'

                # Apply bold if it's bold
                bold = r.xpath(".//w:b", namespaces=namespaces)
                if bold:
                    text = f'<span class="bold">{text}</span>'

                text_runs.append(text)

        paragraph_text = "".join(text_runs)
        paragraph_text = f"{p_start}{paragraph_text}</p>"

        # Replace special characters and patterns
        paragraph_text = paragraph_text.replace("<p></p>", "<p>&nbsp;</p>")
        paragraph_text = paragraph_text.replace("<p>（了）</p>", '<p class="end">（了）</p>')
        paragraph_text = paragraph_text.replace("<p>＊</p>", '<p class="ast">＊</p>')
        paragraph_text = paragraph_text.replace("――", '<span style="letter-spacing: -1px;">―</span>―')
        paragraph_text = paragraph_text.replace("——", '<span style="letter-spacing: -1px;">―</span>―')

        # Merge consecutive same-styled spans
        paragraph_text = re.sub(r'(<span class="ref">)(.*?)</span>(\s*<span class="ref">)+', r'\1\2', paragraph_text)
        paragraph_text = re.sub(r'</span>(\s*<span class="ref">)+', "", paragraph_text)
        paragraph_text = re.sub(r'(</span>)(\s*<span class="ref">)+', "", paragraph_text)

        # Merge consecutive same-styled spans
        paragraph_text = re.sub(r'(indentback-3)(.*?)</span>(\s*indentback-3)+', r'\1\2', paragraph_text)
        paragraph_text = re.sub(r'</span>(\s*indentback-3)+', "", paragraph_text)
        paragraph_text = re.sub(r'(</span>)(\s*indentback-3)+', "", paragraph_text)
        
        # ルビの重複を修正
        paragraph_text = re.sub(r'<ruby>(.*?)<rt>(.*?)</rt></ruby>\2\1', r'<ruby>\1<rt>\2</rt></ruby>', paragraph_text)
        
        # 同じルビが連続している場合は1つだけにする
        paragraph_text = re.sub(r'(<ruby>([^<]+)<rt>([^<]+)</rt></ruby>)\1+', r'\1', paragraph_text)

        html_content.append(paragraph_text)

    return "\n".join(html_content)

def main():
    if len(sys.argv) != 2:
        print("Usage: python fixed_docx2html_v2.py [docx_file_path]")
        sys.exit(1)

    docx_path = sys.argv[1]
    html_path = os.path.splitext(docx_path)[0] + ".html"

    html_content = convert_docx_to_html(docx_path)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n<title>Converted Document</title>\n</head>\n<body>")
        f.write(html_content)
        f.write("\n</body>\n</html>")
    
    print(f"Converted file saved as {html_path}")

if __name__ == "__main__":
    main()
