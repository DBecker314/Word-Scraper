from pathlib import Path
from docx import Document
import pandas as pd
input_path = input("Input Source File or Folder Here:")
input_save = input("Output CSV File as:")
chunks = []
folder_path = Path(input_path)
if folder_path.is_dir():
    for item in folder_path.iterdir():
        document = Document(item)
        for paragraph in document.paragraphs:
            chunks.append(paragraph.text)
elif folder_path.is_file():
    document = Document(folder_path)
    for paragraph in document.paragraphs:
        chunks.append(paragraph.text)

Data = pd.DataFrame(chunks)
Data.to_csv(input_save)
