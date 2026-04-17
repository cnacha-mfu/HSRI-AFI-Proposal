# -*- coding: utf-8 -*-
"""
Convert NRIIS_proposal_TH.md → styled HTML → PDF via Chrome headless.
Font: TH Sarabun New 16pt (สวรส. standard)
"""
import sys, os, re, subprocess
sys.stdout.reconfigure(encoding='utf-8')
import markdown

FOLDER = r'G:\My Drive\Research\MORU'
MD_FILE = os.path.join(FOLDER, 'NRIIS_proposal_TH.md')
HTML_FILE = os.path.join(FOLDER, 'NRIIS_proposal_TH.html')
PDF_FILE = os.path.join(FOLDER, 'NRIIS_proposal_TH.pdf')

# Read and clean markdown
with open(MD_FILE, 'r', encoding='utf-8') as f:
    md_text = f.read()

md_text = re.sub(r'^>.*คัดลอก.*\n', '', md_text, flags=re.MULTILINE)
md_text = re.sub(r'^#\s*═+\s*$', '', md_text, flags=re.MULTILINE)
md_text = re.sub(r'^#\s*ส่วนที่.*$', '', md_text, flags=re.MULTILINE)
md_text = re.sub(r'^## ช่อง:\s*', '## ', md_text, flags=re.MULTILINE)
# Remove SVG images (Chrome can handle them but paths may be wrong)
md_text = re.sub(r'!\[.*?\]\(images/.*?\.svg\)', '', md_text)
md_text = re.sub(r'^รูปที่ 3-\d+:.*$', '', md_text, flags=re.MULTILINE)

html_body = markdown.markdown(md_text, extensions=['tables', 'sane_lists'], output_format='html5')

full_html = f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<title>ข้อเสนอโครงการวิจัย สวรส. ปีงบประมาณ 2570</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,400;0,700;1,400;1,700&display=swap');

@page {{
    size: A4;
    margin: 2.54cm 2.54cm 2.54cm 3.17cm;
}}

body {{
    font-family: 'Sarabun', 'TH Sarabun New', sans-serif;
    font-size: 16pt;
    line-height: 1.5;
    color: #000;
}}

h1 {{
    font-size: 20pt;
    font-weight: 700;
    text-align: center;
    margin-top: 0;
    margin-bottom: 12pt;
    color: #1a3c5e;
}}

h2 {{
    font-size: 18pt;
    font-weight: 700;
    margin-top: 20pt;
    margin-bottom: 8pt;
    color: #1a3c5e;
    border-bottom: 2px solid #1a3c5e;
    padding-bottom: 4pt;
    page-break-after: avoid;
}}

h3 {{
    font-size: 16pt;
    font-weight: 700;
    margin-top: 14pt;
    margin-bottom: 6pt;
    color: #2c3e50;
    page-break-after: avoid;
}}

p {{
    margin-top: 0;
    margin-bottom: 6pt;
    text-align: justify;
}}

ul, ol {{
    margin-top: 2pt;
    margin-bottom: 6pt;
}}

li {{
    margin-bottom: 3pt;
}}

table {{
    width: 100%;
    border-collapse: collapse;
    margin: 8pt 0 14pt 0;
    font-size: 13pt;
    page-break-inside: auto;
}}

thead {{ display: table-header-group; }}
tr {{ page-break-inside: avoid; }}

th {{
    background-color: #1a3c5e;
    color: #fff;
    font-weight: 700;
    text-align: center;
    padding: 6pt 4pt;
    border: 1pt solid #1a3c5e;
}}

td {{
    padding: 4pt 4pt;
    border: 1pt solid #ccc;
    vertical-align: top;
}}

tbody tr:nth-child(even) {{ background-color: #f5f7fa; }}

td strong {{ color: #1a3c5e; }}
td em {{ font-style: normal; color: #2c3e50; font-weight: 700; }}

hr {{
    border: none;
    border-top: 2px solid #1a3c5e;
    margin: 20pt 0;
}}

pre, code {{
    font-family: 'Consolas', monospace;
    font-size: 10pt;
    background: #f5f5f5;
}}

pre {{
    padding: 8pt;
    border: 1pt solid #ddd;
    white-space: pre-wrap;
    page-break-inside: avoid;
}}

.cover {{
    text-align: center;
    margin-bottom: 24pt;
    padding-bottom: 16pt;
    border-bottom: 3px solid #1a3c5e;
}}
</style>
</head>
<body>
<div class="cover">
    <h1>ข้อเสนอโครงการวิจัย</h1>
    <p style="font-size:18pt;font-weight:700;color:#1a3c5e;">
        สถาบันวิจัยระบบสาธารณสุข (สวรส.)<br/>ปีงบประมาณ 2570
    </p>
</div>

{html_body}

</body>
</html>"""

# Save HTML
with open(HTML_FILE, 'w', encoding='utf-8') as f:
    f.write(full_html)
print(f'HTML saved: {HTML_FILE}')

# Find Chrome
chrome_paths = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe'),
]
chrome = None
for p in chrome_paths:
    if os.path.exists(p):
        chrome = p
        break

if chrome:
    print('Generating PDF via Chrome headless...')
    result = subprocess.run([
        chrome,
        '--headless',
        '--disable-gpu',
        '--no-sandbox',
        f'--print-to-pdf={PDF_FILE}',
        '--print-to-pdf-no-header',
        f'file:///{HTML_FILE.replace(os.sep, "/")}',
    ], capture_output=True, text=True, timeout=60)
    if os.path.exists(PDF_FILE):
        size_mb = os.path.getsize(PDF_FILE) / 1024 / 1024
        print(f'Done — PDF saved: {PDF_FILE} ({size_mb:.1f} MB)')
    else:
        print(f'Chrome stderr: {result.stderr[:500]}')
        print('PDF generation failed. Use the HTML file and print to PDF from browser.')
else:
    print('Chrome not found. Open the HTML file in browser and print to PDF (Ctrl+P).')
    print(f'HTML file: {HTML_FILE}')
