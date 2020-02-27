import docx
import pandas as pd
import numpy as np
from docx.shared import Pt
from utils import Propietario

document = docx.Document('base.docx')
xlsx = pd.ExcelFile('demo_ccm.xlsx')
df = pd.read_excel(xlsx)
apartment_owner = dict(zip(df['apartment'], df['propietario']))
owner_genre = dict(zip(df['apartment'], df['genre']))


for apartment in df['apartment']:
    prop = Propietario(
        apartment_owner[apartment],
        apartment,
        owner_genre[apartment],
        0,
        0
        )
    for paragraph in document.paragraphs:

        if 'PROPIETARIO' in paragraph.text:
            if prop.genre == 'female':
                paragraph.text = 'Estimada ' + prop.name +  \
                    'propietaria del ' + str(prop.apartment)
            else:
                paragraph.text = 'Estimado ' + prop.name + \
                    'propietario del' + str(prop.apartment)
        elif 'PROPIETARY_ENGLISH' in paragraph.text:
            paragraph.text = 'Dear ' + prop.name + \
                'owner of apartment ' + str(prop.apartment)
        
        style = document.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(16)
        document.save(str(prop.apartment) + 'base.docx')
