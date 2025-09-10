from docx import Document
import pandas as pd
Chunks = []
document = Document('Proxy Review Sheet.docx')
for paragraph in document.paragraphs:
    Chunks.append(paragraph.text)
Data = pd.DataFrame(Chunks)
Data.to_csv('Lesson_Paragraphs')
