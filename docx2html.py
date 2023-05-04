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

        for r in p.xpath(".//w:r[not(ancestor::w:ruby)]", namespaces=namespaces):
            # Check if the run contains ruby
            ruby = r.xpath(".//w:ruby", namespaces=namespaces)
            if ruby:
                ruby_base = ruby[0].xpath(".//w:rubyBase/w:r/w:t/text()", namespaces=namespaces)
                ruby_text = ruby[0].xpath(".//w:rt/w:r/w:t/text()", namespaces=namespaces)
                text = f'<ruby>{ruby_base[0]}<rt>{ruby_text[0]}</rt></ruby>'
            else:
                text = "".join(r.xpath(".//w:t/text()", namespaces=namespaces))
                text = text.lstrip()

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

        # Merge consecutive same-styled spans
        paragraph_text = re.sub(r'(<span class="ref">)(.*?)</span>(\s*<span class="ref">)+', r'\1\2', paragraph_text)
        paragraph_text = re.sub(r'</span>(\s*<span class="ref">)+', "", paragraph_text)
        paragraph_text = re.sub(r'(</span>)(\s*<span class="ref">)+', "", paragraph_text)

        # Merge consecutive same-styled spans
        paragraph_text = re.sub(r'(<span class="bold">)(.*?)</span>(\s*<span class="bold">)+', r'\1\2', paragraph_text)
        paragraph_text = re.sub(r'</span>(\s*<span class="bold">)+', "", paragraph_text)
        paragraph_text = re.sub(r'(</span>)(\s*<span class="bold">)+', "", paragraph_text)


        html_content.append(paragraph_text)

    return "\n".join(html_content)

def main():
    if len(sys.argv) != 2:
        print("Usage: python convert_docx_to_html.py [docx_file_path]")
        sys.exit(1)

    docx_path = sys.argv[1]
    html_path = os.path.splitext(docx_path)[0] + ".html"

    html_content = convert_docx_to_html(docx_path)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n<title>Converted Document</title>\n</head>\n<body>")
        f.write(html_content)
        f.write("\n</body>\n</html>")

if __name__ == "__main__":
    main()