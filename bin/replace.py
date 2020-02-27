import docx
import pandas as pd
import numpy as np
from docx.shared import Pt
from utils import propietario

document = docx.Document('base.docx')
xlsx = pd.ExcelFile('demo_ccm.xlsx')
df = pd.read_excel(xlsx)
department_propietario = dict(zip(df['depa'], df['propietario']))
department_genre = dict(zip(df['depa'], df['genre']))
one_amount = True
two_amount = False

for department in df['depa']:
    prop = propietario(
        department_propietario[department],
        department,
        department_genre[department],
        0,
        0
        )
    for paragraph in document.paragraphs:

        if 'PROPIETARIO' in paragraph.text:
            if prop.genre == 'female':
                paragraph.text = 'Estimada ' + prop.name +  \
                    'propietaria del ' + str(prop.department)
            else:
                paragraph.text = 'Estimado ' + prop.name + \
                    'propietario del' + str(prop.department)
        elif 'PROPIETARY_ENGLISH' in paragraph.text:
            paragraph.text = 'Dear ' + prop.name + \
                'owner of department ' + str(prop.department)
        
        style = document.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(16)
        document.save(str(prop.department) + 'base.docx')
