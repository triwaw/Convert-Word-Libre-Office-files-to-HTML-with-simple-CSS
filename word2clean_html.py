#!/usr/bin/env python3
"""
Unified Word/HTML to clean HTML+CSS converter.
Usage: python word_to_clean_html.py [input_file] [output_base]
If no arguments, a file dialog will appear.
"""

import os
import sys
import re
import hashlib
import shutil
import subprocess
from pathlib import Path
from collections import defaultdict

try:
    from bs4 import BeautifulSoup
    import tinycss2
    from tinycss2.ast import QualifiedRule, Declaration
except ImportError:
    print("Missing required Python modules. Install with:")
    print("  pip install beautifulsoup4 tinycss2")
    sys.exit(1)

# ---------- GUI file selection ----------
try:
    import tkinter as tk
    from tkinter import filedialog
    USE_GUI = True
except ImportError:
    USE_GUI = False

def choose_file_gui():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="Select Word or HTML file",
        filetypes=[
            ("Word / HTML files", "*.doc *.docx *.html *.htm"),
            ("All files", "*.*")
        ]
    )
    root.destroy()
    return file_path

# ---------- LibreOffice conversion ----------
def convert_to_html(input_path):
    """Convert .doc/.docx to HTML using LibreOffice. Returns path to generated HTML."""
    print("Converting document to HTML using LibreOffice...")
    out_dir = input_path.parent
    try:
        # Run LibreOffice headless conversion
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'html', str(input_path),
             '--outdir', str(out_dir)],
            check=True, capture_output=True, text=True
        )
    except subprocess.CalledProcessError as e:
        print(f"LibreOffice conversion failed: {e.stderr}")
        sys.exit(1)
    except FileNotFoundError:
        print("LibreOffice (soffice) not found. Please install LibreOffice and add it to PATH.")
        sys.exit(1)

    # Find the generated HTML file (same name but .html)
    html_file = out_dir / (input_path.stem + ".html")
    if not html_file.exists():
        # Sometimes LibreOffice adds a suffix; try to find any .html created recently
        candidates = list(out_dir.glob("*.html"))
        if candidates:
            # take the most recent
            html_file = max(candidates, key=os.path.getmtime)
        else:
            print("Conversion seemed to succeed but output HTML not found.")
            sys.exit(1)
    print(f"Conversion successful: {html_file}")
    return html_file

# ---------- Style extraction (first pass) ----------
def style_to_class_name(style_str):
    """Generate a deterministic class name from a style string."""
    props = [p.strip() for p in style_str.split(';') if p.strip()]
    props.sort()
    normalized = ';'.join(props)
    hash_obj = hashlib.md5(normalized.encode())
    return f"s{hash_obj.hexdigest()[:8]}"

def extract_inline_styles(soup):
    """
    Find all tags with 'style' attribute, generate class names,
    build CSS rules, and replace style with class.
    Returns dict: style_string -> class_name, and list of (class_name, style_string)
    """
    style_map = {}
    rules = []

    for tag in soup.find_all(style=True):
        style_val = tag['style']
        style_val = re.sub(r'\s+', ' ', style_val).strip()
        if not style_val:
            del tag['style']
            continue

        class_name = style_to_class_name(style_val)
        if style_val not in style_map:
            style_map[style_val] = class_name
            rules.append((class_name, style_val))

        # Add class to tag
        if tag.has_attr('class'):
            tag['class'].append(class_name)
        else:
            tag['class'] = [class_name]
        del tag['style']

    return style_map, rules

# ---------- Image handling ----------
def process_images(soup, base_dir, output_dir):
    """Copy local images to images/ subfolder and update src."""
    images_folder = output_dir / "images"
    images_folder.mkdir(exist_ok=True)

    for img in soup.find_all('img'):
        src = img.get('src')
        if not src:
            continue
        src = src.replace('%20', ' ')
        src_path = base_dir / src
        if src_path.exists():
            dest = images_folder / src_path.name
            shutil.copy2(src_path, dest)
            img['src'] = f"images/{src_path.name}"
        else:
            print(f"Warning: Image not found: {src}")

# ---------- CSS parsing ----------
def parse_css(css_text):
    """Parse CSS text and return a dict: classname -> property dict."""
    rules = tinycss2.parse_stylesheet(css_text, skip_whitespace=True)
    class_styles = {}
    for rule in rules:
        if isinstance(rule, QualifiedRule):
            selector = tinycss2.serialize(rule.prelude).strip()
            if selector.startswith('.'):
                class_name = selector[1:].split()[0]  # simple class selector
                decls = tinycss2.parse_declaration_list(rule.content)
                props = {}
                for decl in decls:
                    if isinstance(decl, Declaration):
                        name = decl.name
                        value = tinycss2.serialize(decl.value).strip()
                        props[name] = value
                class_styles[class_name] = props
    return class_styles

# ---------- Style combination ----------
def merge_styles(styles_list):
    merged = {}
    for s in styles_list:
        merged.update(s)
    return merged

def style_dict_to_css(style_dict):
    return '; '.join(f"{k}: {v}" for k, v in style_dict.items())

def get_style_for_classes(class_names, class_styles):
    styles = [class_styles.get(cn, {}) for cn in class_names]
    return merge_styles(styles)

def class_name_from_style(style_dict, existing_map):
    normalized = tuple(sorted(style_dict.items()))
    if normalized in existing_map:
        return existing_map[normalized]
    hash_str = hashlib.md5(str(normalized).encode()).hexdigest()[:8]
    new_name = f"c{hash_str}"
    existing_map[normalized] = new_name
    return new_name

# ---------- Flattening text nodes ----------
def flatten_text_structure(soup, class_styles):
    """
    For each text node, collect all classes and formatting from ancestors,
    combine styles, replace with a single span having a combined class.
    Returns a dict of normalized style -> new class name.
    """
    style_map = {}  # normalized tuple -> class name

    # Process all text nodes
    for text_node in soup.find_all(string=True):
        if not text_node.strip():
            continue

        # Collect classes and formatting properties from ancestors
        classes = []
        formatting_props = []
        current = text_node.parent
        # We'll walk up until body, but stop if we hit an <a> or other structural tag? No, we want all formatting.
        while current and current.name != 'body':
            if current.has_attr('class'):
                classes.extend(current['class'])
            # Handle formatting tags
            if current.name in ('b', 'strong'):
                formatting_props.append(('font-weight', 'bold'))
            elif current.name in ('i', 'em'):
                formatting_props.append(('font-style', 'italic'))
            elif current.name == 'u':
                formatting_props.append(('text-decoration', 'underline'))
            # <font> may have attributes; we don't have direct mapping, but they already contributed via classes.
            # (The first pass already converted <font> style attributes to classes.)
            current = current.parent

        # Base style from classes
        base_style = get_style_for_classes(classes, class_styles)

        # Override with formatting props
        for prop, val in formatting_props:
            base_style[prop] = val

        # Get or create combined class
        combined_class = class_name_from_style(base_style, style_map)

        # Create new span
        new_span = soup.new_tag('span')
        new_span['class'] = combined_class
        new_span.string = text_node

        # Replace the text node
        text_node.replace_with(new_span)

    # Remove empty tags that may have been left
    for tag in soup.find_all():
        if tag.name in ('span', 'font', 'b', 'i', 'u', 'strong', 'em') and not tag.contents:
            tag.decompose()
        # Also remove tags that contain only whitespace and no other content
        if tag.string and not tag.string.strip() and not list(tag.children):
            tag.decompose()

    return style_map

# ---------- Main cleaning function ----------
def clean_html(html_path, output_base=None):
    html_path = Path(html_path).resolve()
    base_dir = html_path.parent

    if output_base is None:
        output_base = html_path.stem + "_clean"
    output_dir = base_dir / output_base
    output_dir.mkdir(exist_ok=True)

    # Read HTML
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    # Remove unwanted meta tags
    unwanted_names = ['generator', 'author', 'created', 'changed', 'AppVersion']
    for meta in soup.find_all('meta'):
        name = meta.get('name', '').lower()
        if name in unwanted_names:
            meta.decompose()

    # Process images
    process_images(soup, base_dir, output_dir)

    # ---- First pass: extract inline styles to classes ----
    inline_style_map, inline_rules = extract_inline_styles(soup)

    # ---- Gather existing <style> tags ----
    style_tags = soup.find_all('style')
    css_parts = []
    for style in style_tags:
        if style.string:
            css_parts.append(style.string.strip())
        style.decompose()
    css_content = "\n".join(css_parts)

    # Add inline style rules
    for class_name, style_str in inline_rules:
        css_content += f"\n.{class_name} {{ {style_str} }}"

    # ---- Parse CSS to get class definitions ----
    class_styles = parse_css(css_content)

    # ---- Second pass: flatten nested structure ----
    new_style_map = flatten_text_structure(soup, class_styles)

    # ---- Generate new CSS for flattened classes ----
    new_css_lines = []
    for normalized, class_name in new_style_map.items():
        style_dict = dict(normalized)
        rule = f".{class_name} {{ " + style_dict_to_css(style_dict) + " }"
        new_css_lines.append(rule)

    # Combine new CSS with any remaining original CSS (like @page, etc.)
    # We'll keep the original CSS as well; unused classes are harmless.
    final_css = "\n".join(new_css_lines) + "\n" + css_content

    # ---- Write CSS file ----
    css_file = output_dir / f"{output_base}.css"
    with open(css_file, 'w', encoding='utf-8') as f:
        f.write(final_css)

    # ---- Add link to CSS in head ----
    head = soup.head
    if not head:
        head = soup.new_tag('head')
        if soup.html:
            soup.html.insert(0, head)
        else:
            soup.insert(0, head)
    link = soup.new_tag('link', rel='stylesheet', href=f"{output_base}.css")
    head.append(link)

    # ---- Write final HTML ----
    html_file = output_dir / f"{output_base}.html"
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(str(soup.prettify()))

    print(f"Cleaned files saved in: {output_dir}")
    print(f"  HTML: {html_file.name}")
    print(f"  CSS:  {css_file.name}")
    print(f"  Images: images/ (if any)")

# ---------- Main ----------
def main():
    input_file = None
    output_base = None

    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
        if len(sys.argv) >= 3:
            output_base = sys.argv[2]
    else:
        if USE_GUI:
            input_file = choose_file_gui()
        else:
            input_file = input("Enter path to Word or HTML file: ").strip()

    if not input_file:
        print("No input file provided. Exiting.")
        sys.exit(1)

    input_path = Path(input_file)
    if not input_path.exists():
        print(f"File not found: {input_file}")
        sys.exit(1)

    # Determine file type and convert if needed
    ext = input_path.suffix.lower()
    if ext in ('.doc', '.docx'):
        # Convert to HTML using LibreOffice
        html_file = convert_to_html(input_path)
        # Now clean that HTML
        clean_html(html_file, output_base)
        # Optionally delete the intermediate HTML? Keep it for now.
    elif ext in ('.html', '.htm'):
        clean_html(input_path, output_base)
    else:
        print("Unsupported file type. Please provide .doc, .docx, or .html")
        sys.exit(1)

if __name__ == "__main__":
    main()