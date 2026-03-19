# Convert-Word-Libre-Office-files-to-HTML-with-simple-CSS
Word/ODT to Clean HTML Converter
Author: Wasif Ali
GitHub: triwaw

This Python script converts Microsoft Word (.doc, .docx) and LibreOffice (.odt) files into a minimal, clean HTML/CSS package. It also accepts existing HTML files (exported from Word/LibreOffice) and performs a deep cleaning to remove clutter, flatten nested formatting, and organise images.
✨ Features
Converts .doc, .docx, .odt to HTML using LibreOffice (headless).
Accepts existing .html / .htm files directly.
Removes unnecessary meta tags (generator, author, dates).
Extracts all inline style attributes and <style> blocks into a single external .css file.
Flattens deeply nested <span>, <font>, <b>, <i>, <u> tags – for each text fragment, it combines all applied styles into one CSS class, dramatically reducing HTML size.
Copies embedded images into an images/ subfolder and updates <img> src attributes.
Preserves hyperlinks (<a> tags) exactly as in the original document.
Produces a portable folder containing .html, .css, and images/.
📋 Requirements
Python 3.6+
Download from python.org.
Python Libraries
Install via pip:
bash
pip install beautifulsoup4 tinycss2
LibreOffice (only needed for .doc, .docx, .odt files)
Download from libreoffice.org.
During installation, ensure the program folder is added to your system PATH so that the soffice command is available from the command line.
Verify by running:

bash
soffice --version
Expected output:

text
LibreOffice 26.2.1.2 620(Build:2)
Note: If you only need to clean existing HTML files, LibreOffice is not required.

🚀 Installation
Clone this repository or download the script word_to_clean_html.py.

Install the required Python packages:

bash
pip install beautifulsoup4 tinycss2
(Optional) If you plan to convert Word/ODT files, install LibreOffice and add it to your PATH.

🖥️ Usage
Option A: GUI Mode (easiest)
Double‑click the script (or run python word_to_clean_html.py without arguments).

A file selection dialog appears – choose your Word/ODT/HTML file.

Enter a base name for the output folder (or press Enter to use default: originalname_clean).

The script creates a new folder with that name, containing the cleaned files.

Option B: Command Line
bash
python word_to_clean_html.py input.docx output_name
input.docx – your source file (.doc, .docx, .odt, .html, .htm).

output_name – (optional) base name for output folder. If omitted, input_stem_clean is used.

Option C: Using a Batch File (Windows)
If you have multiple Python installations, create a .bat file to always run with the correct Python (see the Batch File section).

📁 Output Structure
For an input file mydoc.docx with output name my_clean, you get:

text
my_clean/
├── my_clean.html
├── my_clean.css
└── images/          (if the document contained images)
    ├── pic1.png
    └── pic2.gif
The HTML links to the CSS file (my_clean.css) and images are referenced as images/filename.

🧠 How It Works (Brief)
File Selection / Conversion

If the input is a document (.doc, .docx, .odt), LibreOffice converts it to HTML.

If it’s already HTML, the script proceeds directly.

First Pass – Inline Styles

Every style="..." attribute is replaced with a generated class (e.g., s1a2b3c4), and the CSS rule is collected.

Image Handling

Local images are copied to an images/ folder, and <img src> is updated.

CSS Parsing

All CSS (from <style> blocks and inline styles) is parsed into a dictionary of class → properties.

Second Pass – Flattening

For each text node, the script walks up its ancestors, collects all classes and formatting tags (<b>, <i>, <u>), and merges the CSS properties.

The resulting combined style is given a new class (e.g., c8f3a2d1), and the entire nested structure is replaced by a single <span class="..."> containing the text.

Final CSS & HTML

New CSS rules for the combined classes are written to an external file.

A <link> to the CSS is added to the HTML <head>.

The cleaned HTML and CSS are saved in the output folder.

🪟 Batch File for Windows
If your default python command points to the wrong interpreter (e.g., LibreOffice’s Python), create a batch file in the same folder as the script:

run_clean.bat

batch
@echo off
REM Replace the path below with the full path to YOUR Python 3.6+ interpreter
C:\Users\YourName\AppData\Local\Programs\Python\Python313\python.exe word_to_clean_html.py %*
pause
Replace C:\Users\YourName\... with the actual path to the Python interpreter where beautifulsoup4 and tinycss2 are installed.

Now you can double‑click run_clean.bat or drag‑and‑drop a file onto it to start the conversion.

🔧 Troubleshooting
"Missing required Python modules" even after pip install
This means the Python interpreter you are running does not have the modules installed.
Use the batch file above to explicitly call the correct Python.
Or run the script with the full path, e.g.:
bash
C:\Users\YourName\AppData\Local\Programs\Python\Python313\python.exe word_to_clean_html.py
LibreOffice conversion fails
Ensure soffice is in your PATH. Open a terminal and type soffice --version.
If the command is not found, reinstall LibreOffice and check the option "Add LibreOffice to PATH" (or manually add C:\Program Files\LibreOffice\program to your system’s PATH environment variable).
Images are not copied
The script only copies images that are referenced as local files (e.g., src="image.png"). If the HTML uses absolute URLs (like http://...), they are left untouched.

Hyperlinks are lost?
Hyperlinks (<a> tags) are preserved. The text inside them is flattened, but the <a> tag itself remains.

📜 License
This project is open source and available under the MIT License.

🙏 Acknowledgements
Beautiful Soup
tinycss2
LibreOffice

Enjoy clean HTML!
If you have any questions or suggestions, feel free to open an issue or contribute on GitHub.
