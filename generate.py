#!/usr/bin/env python3

from docx import Document
import pandas as pd
import os

os.makedirs("output", exist_ok=True)

template = Document('input/Personalised report template sent to Nick e centre.docx')
df = pd.read_excel("input/Required document 1 - Data spreadsheet for e-research center.xlsx")
print(df)

for i, row in df.iterrows():
    if i == 0:
        continue
    number = row[0]
    V1 = row["V1"]
    V2 = str(round(row["V2"], 1))
    V3 = row["V3"]
    V4 = str(round(row["V4 (constant)"], 1))
    print(number, V1, V2, V3, V4)
    for para in template.paragraphs:
        if "…V1" in para.text:
            para.text = para.text.replace("…V1", V1)
        if "***V2" in para.text:
            para.text = para.text.replace("***V2", V2)
        if "***V3" in para.text:
            para.text = para.text.replace("***V3", V3)
    filename = f'output/{number}_{V1}.docx'
    template.save(filename)
    print(f"{filename} saved")
    exit(1)