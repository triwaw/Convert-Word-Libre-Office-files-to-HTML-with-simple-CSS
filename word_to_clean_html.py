#!/usr/bin/env python3
"""
Convert MS Word (.doc, .docx) to clean HTML + CSS.
Usage: python word_to_clean_html.py input.docx [output_prefix]
If output_prefix is not given, files are saved as output.html and output.css.
"""

import os
import sys
import subprocess
import re
from pathlib import Path

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

def convert_doc_to_docx(input_doc):
    """Convert .doc to .docx using LibreOffice (headless). Returns path to temporary .docx."""
    output_docx = str(Path(input_doc).with_suffix('.docx'))
    # If output already exists, we'll use a temp name to avoid overwriting
    if os.path.exists(output_docx):
        base, ext = os.path.splitext(input_doc)
        output_docx = f"{base}_converted.docx"
    try:
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'docx', input_doc, '--outdir', os.path.dirname(input_doc)],
            check=True, capture_output=True, text=True
        )
        if not os.path.exists(output_docx):
            # LibreOffice sometimes names the file differently; try to find it
            possible = list(Path(os.path.dirname(input_doc)).glob("*.docx"))
            if possible:
                # take the most recent one (heuristic)
                output_docx = str(max(possible, key=os.path.getmtime))
            else:
                raise RuntimeError("Conversion failed: output file not found")
        return output_docx
    except subprocess.CalledProcessError as e:
        print(f"LibreOffice conversion failed: {e.stderr}")
        sys.exit(1)
    except FileNotFoundError:
        print("LibreOffice (soffice) not found. Please install LibreOffice and ensure it's in your PATH.")
        sys.exit(1)

def get_run_html(run, in_heading=False):
    """
    Convert a docx run to an HTML snippet.
    - Bold → <strong>
    - Italic → <em>
    - Underline → <u>
    - Font containing 'Jameel Noori Nastaleeq' → <span class="urdu">
    These tags are nested appropriately (span inside formatting tags).
    """
    text = run.text
    if not text:
        return ""

    # Determine which formatting applies
    tags = []
    if run.bold:
        tags.append(('strong', None))
    if run.italic:
        tags.append(('em', None))
    if run.underline:
        tags.append(('u', None))

    # Check font name for Jameel Noori Nastaleeq
    font_name = None
    if run.font.name:
        font_name = run.font.name
    # Also check style's font? Not needed, run.font.name already includes direct formatting
    if font_name and "jameel noori nastaleeq" in font_name.lower():
        tags.append(('span', 'urdu'))

    # Build HTML from innermost out
    html = text
    for tag_name, class_name in tags:
        if class_name:
            attrs = f' class="{class_name}"'
        else:
            attrs = ''
        html = f'<{tag_name}{attrs}>{html}</{tag_name}>'

    return html

def process_docx(docx_path):
    """Parse docx and return BeautifulSoup of body content."""
    doc = Document(docx_path)
    soup = BeautifulSoup('<body></body>', 'html.parser')
    body = soup.body

    for para in doc.paragraphs:
        style_name = para.style.name
        # Determine heading level
        heading_level = None
        if style_name and style_name.startswith('Heading'):
            # Extract number from style name (e.g., "Heading 1" -> 1)
            match = re.search(r'(\d+)$', style_name)
            if match:
                level = int(match.group(1))
                if 1 <= level <= 3:
                    heading_level = level

        # Create container tag
        if heading_level:
            tag = soup.new_tag(f'h{heading_level}')
        else:
            tag = soup.new_tag('p')

        # Process runs
        for run in para.runs:
            run_html = get_run_html(run, in_heading=bool(heading_level))
            if run_html:
                # Insert as raw HTML; BeautifulSoup will parse it
                tag.append(BeautifulSoup(run_html, 'html.parser'))
        body.append(tag)

        # Add newline for readability (optional)
        body.append(soup.new_string('\n'))

    return soup

def write_output(soup, css_content, output_prefix):
    """Write HTML and CSS files."""
    html_file = f"{output_prefix}.html"
    css_file = f"{output_prefix}.css"

    # Add link to CSS in head
    html = BeautifulSoup('<html><head></head></html>', 'html.parser')
    html.head.append(html.new_tag('link', rel='stylesheet', href=os.path.basename(css_file)))
    # Replace body with our processed body
    if html.body:
        html.body.replace_with(soup.body)
    else:
        html.append(soup.body)

    # Write files
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(str(html.prettify()))
    with open(css_file, 'w', encoding='utf-8') as f:
        f.write(css_content)

    print(f"Files saved: {html_file}, {css_file}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python word_to_clean_html.py input.docx [output_prefix]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_prefix = sys.argv[2] if len(sys.argv) > 2 else "output"

    # Convert .doc to .docx if needed
    ext = Path(input_path).suffix.lower()
    if ext == '.doc':
        print("Converting .doc to .docx using LibreOffice...")
        docx_path = convert_doc_to_docx(input_path)
        print(f"Converted to: {docx_path}")
    elif ext == '.docx':
        docx_path = input_path
    else:
        print("Unsupported file type. Please provide .doc or .docx.")
        sys.exit(1)

    # Process the .docx
    soup = process_docx(docx_path)

    # Define CSS
    css = f"""/* Clean styles for converted Word document */
body {{
    font-family: Arial, Tahoma, 'Jameel Noori Nastaleeq', sans-serif;
    font-size: 12px;
    line-height: 1.5;
    margin: 1em;
}}

h1 {{
    font-size: 32px;
    font-weight: bold;
    text-decoration: underline;
}}

h2 {{
    font-size: 28px;
    font-weight: bold;
    text-decoration: underline;
}}

h3 {{
    font-size: 18px;
    font-weight: bold;
    text-decoration: underline;
}}

/* Urdu text using Jameel Noori Nastaleeq gets larger size */
.urdu {{
    font-family: 'Jameel Noori Nastaleeq', serif;
    font-size: 22px;
}}

/* Preserve formatting tags */
strong, em, u {{
    /* inherit font-family and size */
}}
"""
    write_output(soup, css, output_prefix)

    # Clean up temporary file if created
    if ext == '.doc' and docx_path != input_path and os.path.exists(docx_path):
        os.remove(docx_path)

if __name__ == "__main__":
    main()