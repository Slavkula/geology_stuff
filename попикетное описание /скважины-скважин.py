import re
from docx import Document

def replace_text_in_word_doc(file_path):
    doc = Document(file_path)
    pattern = r"(встречены в пределах) скважин(: \d+[.;])"
    replacement = r"\1 скважины\2"

    for paragraph in doc.paragraphs:
        if re.search(pattern, paragraph.text):
            paragraph.text = re.sub(pattern, replacement, paragraph.text)

    doc.save('updated_' + file_path)

replace_text_in_word_doc('Документ Microsoft Word.docx')
