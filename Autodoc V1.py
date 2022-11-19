from docxtpl import DocxTemplate
from pathlib import Path
import pandas as pd

##Path to the word file
base_dir = Path(__file__).parent
word_template_path = base_dir / "Word.docx"
excel_path = base_dir / "Excel.xlsx"
output_dir = base_dir / "Output"

##Create Output folders for word documents
output_dir.mkdir(exist_ok=True)

# Convert Excel sheet into pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")

##Iterate over  each row  in df and render word document
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['NAME of the DOC']}-Doc.docx"
    doc.save(output_path)
