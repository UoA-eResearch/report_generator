#!/usr/bin/env python3

from docx import Document
import pandas as pd
import os
from gauge import gauge

os.makedirs("output", exist_ok=True)

template = Document('input/Personalised report template sent to Nick e centre.docx')
df = pd.read_excel("input/Required document 1 - Data spreadsheet for e-research center.xlsx")
print(df)

LEVELS = ['Low','Low to medium','Medium','Medium to High','High']

for i, row in df.iterrows():
    if i == 0:
        continue
    number = row[0]
    V1 = row["V1"]
    V2 = str(round(row["V2"], 1))
    V3 = row["V3"]
    V4 = str(round(row["V4 (constant)"], 1))
    print(number, V1, V2, V3, V4)
    filename = f'output/{number}_{V1}'
    image_filename = filename + ".png"
    gauge(labels=['Low', '', 'Medium', '', 'High'],
          colors='RdYlGn',
          arrow=row["V2"] / 2,
          fname=image_filename)
    for para in template.paragraphs:
        if "…V1" in para.text:
            para.text = para.text.replace("…V1", V1)
        if "***V2" in para.text:
            para.text = para.text.replace("***V2", V2)
        if "***V3" in para.text:
            para.text = para.text.replace("***V3", V3)
        if para.text == "Diagram to show the score of the company on average (V2) and the score of the industry on average (V4).":
            para.text = ""
            run = para.add_run()
            run.add_picture(image_filename)
        if "Diagram" in para.text:
            print(para.text)
    template.save(filename + ".docx")
    print(f"{filename} saved")
    exit(1)