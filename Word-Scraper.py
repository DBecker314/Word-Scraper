from docx.api import Document
document = Document("Proxy Review Sheet.docx")
for p in document.paragraphs:
    if p.style.name.startswith("Heading") or p.style.name.startswith("Title"):
        print(p.text)
