from docx import Document
import re

def remove_zeroes(doc_path):
    doc = Document(doc_path)
    for para in doc.paragraphs:
        for run in para.runs:
            run.text = re.sub(r'(\d+,\d)0\b', r'\1', run.text)
    doc.save('updated_' + doc_path)

remove_zeroes('Документ Microsoft Word.docx')
