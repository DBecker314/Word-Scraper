from docx import Document
document = Document('Proxy Review Sheet.docx')
for paragraph in document.paragraphs:
    print(paragraph.text)