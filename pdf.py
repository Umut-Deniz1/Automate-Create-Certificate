#!/usr/bin/env python
# -*- coding: utf8 -*-

from fpdf import FPDF
import pandas as pd
import os
import shutil

path = "C:/Users/Asus/Desktop/create_certificate/tr_sertifikaEN"
font = os.path.join("C:/Windows/Fonts", "ARIALUNI.TTF")
namelist = ["number-your name"]



class PDF(FPDF):
    def header(self):
        # Logo
        self.image("ENG.jpg",x = -1, y = 0, w = 215)
        self.ln(93)

    # Page footer
    def footer(self):
        self.set_text_color(128)
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'I', 8)
        # Page number
        #self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')


data = pd.read_excel("ser.xls")
df = pd.DataFrame(data)

result = df["NO-ADSOYAD"].values
for i in result:
    namelist.append(i)

#print(namelist)

for name in namelist:
    name = name.split("-")
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()

    pdf.add_font('ARIALUNI', '', font, uni=True)
    pdf.set_font('ARIALUNI', '', 30)

    if "i" in name[1]:
        name[1] = name[1].replace("i","İ").replace("  "," ").upper()
    else:
        name[1] = name[1].replace("  "," ").upper()

    pdf.set_text_color(2, 21, 81)
    #print(len(name[1]))
    if len(name[1]) < 12:
        pdf.cell(60)

    elif len(name[1]) < 15:
        pdf.cell(54)

    elif len(name[1]) < 18:
        pdf.cell(51)

    else:
        pdf.cell(42)

    pdf.cell(0, 0, f'{name[1]}', 0, 1)
    
    if os.path.isdir("tr_sertifikaEN") == False:
        os.mkdir(path)
    pdf.output(f'{name[0]}-EN.pdf', 'F')
    

SOURCE_DIR = 'C:/Users/Asus/Desktop/create_certificate'
DEST_DIR = 'C:/Users/Asus/Desktop/create_certificate/tr_sertifikaEN'

for fname in os.listdir(SOURCE_DIR):
    if fname.lower().endswith('.pdf'):
        shutil.move(os.path.join(SOURCE_DIR, fname), DEST_DIR)
