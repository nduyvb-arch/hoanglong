import docx
import glob
import os

files = glob.glob('bài *.docx')
files.sort()

with open('output.txt', 'w', encoding='utf-8') as out:
    for f in files:
        try:
            doc = docx.Document(f)
            out.write(f"--- {f} ---\n")
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ''])
            out.write(text[:500] + "...\n\n")
        except Exception as e:
            out.write(f"Error reading {f}: {e}\n\n")
