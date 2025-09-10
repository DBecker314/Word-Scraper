from pathlib import Path
from docx import Document
import pandas as pd
Folder = input("Input the Folder Path Here:")
Chunks = []

folder_path = Path(Folder)
for item in folder_path.iterdir():
    document = Document(item)
    for paragraph in document.paragraphs:
        Chunks.append(paragraph.text)

Data = pd.DataFrame(Chunks)
Data.to_csv('Lesson_Paragraphs')
