from docx import Document
import re

def remove_zeroes(doc_path):
    doc = Document(doc_path)
    for para in doc.paragraphs:
        for run in para.runs:
            for i in range(10, 131):
                num = i / 10
                run.text = re.sub(rf'({num:.2f})м\b', lambda m: f'{m.group(1)[:-1]}м', run.text)
    doc.save('updated_' + doc_path)

remove_zeroes('Документ Microsoft Word.docx')
