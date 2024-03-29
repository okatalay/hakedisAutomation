import os
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
import sql_modul as sql
from tkinter import messagebox
from datetime import datetime, timedelta
import json

list = None
list2 = None
denetci_and_kont = None
mimar_or_muh = None
birinci_kisim_insaat = None

save_path = ""

def iskele_kurulum():

    ek_unvan = ""
    if "MK" in list2[0][1] or "mk" in list2[0][1] or "Mk" in list2[0][1]:
        ek_unvan = "101"
    else:
        ek_unvan = "1900"

    try:
        date_string = list[3]
        date_format = "%d.%m.%Y"
        date = datetime.strptime(date_string, date_format)
        three_days_before = date - timedelta(days=3)
        result_string = three_days_before.strftime(date_format)
    except Exception as e:
        messagebox.showerror("HATA !", "Lütfen beton tarihini kontrol ediniz ! ")

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    baslık2 = document.add_heading("KALIP VE TAŞIYICI KALIP İSKELESİ KURULUMU KONTROL TUTANAĞI\n")
    baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    table1 = document.add_table(rows=8, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("", f"  YİBF No: {list2[0][0]}"),
        ("İlgili İdare", f": {list2[0][3]}"),
        ("Yapı Sahibi", f": {list2[0][12]}"),
        ("Yapı Ruhsat tarihi ve nosu", f": {list2[0][4]}"),
        ("Yapının Adresi", f": {list2[0][7]}"),
        ("Pafta/Ada/Parsel No", f": {list2[0][9]}/{list2[0][10]}/{list2[0][11]}"),
        ("Yapı İnşaat Alanı (m²) ve Cinsi", f": {list2[0][6]} m2 {list2[0][5]}"),
        ("Yapı Denetim Kuruluşunun Unvanı/İzin Belge No", f": {list2[0][1]}/{ek_unvan}"),
    ]

    column_widths = [5, 4.8]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(8):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]
            if j == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    document.add_paragraph("\n\n")
    p1 = document.add_paragraph()

    p1.add_run(f"Yukarıda belirtilen yapının ")
    bold_run1 = p1.add_run(f"{list[0]} {list[1]} ")
    bold_run1.bold = True
    p1.add_run("kotunda yapılan denetimde:")
    document.add_paragraph("Kalıp ve taşıyıcı kalıp iskelesi olarak kullanılan malzemenin istenilen nitelikte, kalıp ve taşıyıcı kalıp iskele işçiliğinin ve takviyelerinin yeterli olduğu ölçü, kot, yatay ve düşey düzlemlere uygunluk açısından kalıbın ve taşıyıcı kalıp iskelelerinin ruhsat eki projelerine (çizimler veya malzeme, boyut, mekanik özellikler ve uygulama detaylarını gösteren tablolar veya tip detaylar) uygun olarak yapıldığı, tespit edilmiştir")

    p3 = document.add_paragraph()
    p3.add_run(f"İş bu tutanak ")
    bold_run2 = p3.add_run(f"{result_string} ")
    bold_run2.bold = True
    p3.add_run("tarihinde bir nüshası inşaat mühendisi fenni mesul/ yapı denetçisi tarafından ilgili idareye verilmek üzere üç nüsha olarak düzenlenmiştir."+"\n"*4)

    table2 = document.add_table(rows=3, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("Fenni mesul/Yapı Denetçisi", f"Şantiye Şefi"),
        ("İnşaat Mühendisi", f"{list2[0][13]}"),
        ("İmza", f"İmza")]


    column_widths = [5, 5]
    for row_index in range(3):
        cell = table2.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

    cell = table2.cell(row_index, 1)
    cell.width = Inches(column_widths[1])

    for i in range(3):
        for j in range(2):
            cell = table2.cell(i, j)
            cell.text = cell_contents1[i][j]
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if i == 0 or i == 1:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)

    if "MUHTEŞEM" in list2[0][1] or "muh" in list2[0][1]:

        document.save(
            f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/iskele_kurulum.docx")
        #os.startfile(f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/iskele_kurulum.docx")
    else:
        document.save(
            f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}/iskele_kurulum.docx")
        #os.startfile(f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}")

def kalip_donati():

    ek_unvan = ""
    if "MK" in list2[0][1] or "mk" in list2[0][1] or "Mk" in list2[0][1]:
        ek_unvan = "101"
    else:
        ek_unvan = "1900"

    try:
        date_string = list[3]

        date_format = "%d.%m.%Y"
        date = datetime.strptime(date_string, date_format)
        three_days_before = date - timedelta(days=1)
        result_string = three_days_before.strftime(date_format)
    except Exception as e:
        messagebox.showerror("HATA !", "Lütfen beton tarihini kontrol ediniz ! ")

    print(denetci_and_kont)

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    baslık2 = document.add_heading("KALIP VE DONATI İMALATI KONTROL TUTANAĞI\n")
    baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    table1 = document.add_table(rows=8, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("", f"  YİBF No: {list2[0][0]}"),
        ("İlgili İdare", f": {list2[0][3]}"),
        ("Yapı Sahibi", f": {list2[0][12]}"),
        ("Yapı Ruhsat tarihi ve nosu", f": {list2[0][4]}"),
        ("Yapının Adresi", f": {list2[0][7]}"),
        ("Pafta/Ada/Parsel No", f": {list2[0][9]}/{list2[0][10]}/{list2[0][11]}"),
        ("Yapı İnşaat Alanı (m²) ve Cinsi", f": {list2[0][6]} m2 {list2[0][5]}"),
        ("Yapı Denetim Kuruluşunun Unvanı/İzin Belge No", f": {list2[0][1]}/{ek_unvan}"),
    ]

    column_widths = [5, 4.8]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(8):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]
            if j == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    document.add_paragraph("\n\n")
    p1 = document.add_paragraph()

    p1.add_run(f"Yukarıda belirtilen yapının ")
    bold_run1 = p1.add_run(f"{list[0]} {list[1]} ")
    bold_run1.bold = True
    p1.add_run("kotunda yapılan denetimde:")

    document.add_paragraph("1- Kalıp imalatında kullanılan malzemenin istenilen nitelikte, kalıp işçiliğinin iyi ve takviyelerinin yeterli olduğu, ölçü, kot, yatay ve düşey düzlemlere uygunluk açısından kalıbın projesine uygun olarak yapıldığı, ")
    document.add_paragraph("2- Betonarme demirlerinin projesinde gösterilen adet, çap ve boyda olduğu, projesine uygun olarak döşendiği, ")
    document.add_paragraph("3- Tesisat projelerine uygunluk sağlandığı tespit edilmiştir.")
    document.add_paragraph("Bu durumda beton dökülmesine izin verilmiştir.")


    p3 = document.add_paragraph()
    p3.add_run(f"İş bu tutanak ")
    bold_run2 = p3.add_run(f"{result_string} ")
    bold_run2.bold = True
    p3.add_run("tarihinde bir nüshası inşaat mühendisi fenni mesul/ yapı denetçisi tarafından ilgili idareye verilmek üzere üç nüsha olarak düzenlenmiştir."+"\n"*4)

    if len(denetci_and_kont[0]) == 4:

        table2 = document.add_table(rows=4, cols=4)
        for row in table2.rows:
            row.height = Inches(0)
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.space_after = Inches(0)

        cell_contents1 = [
            (f"{denetci_and_kont[0][0][2]}", f"{denetci_and_kont[0][1][2]}" , f"{denetci_and_kont[0][2][2]}", f"{denetci_and_kont[0][3][2]}"),
            (f"{denetci_and_kont[0][0][3]}", f"{denetci_and_kont[0][1][3]}", f"{denetci_and_kont[0][2][3]}", f"{denetci_and_kont[0][3][3]}"),
            (f"{denetci_and_kont[0][0][1]}", f"{denetci_and_kont[0][1][1]}", f"{denetci_and_kont[0][2][1]}", f"{denetci_and_kont[0][3][1]}"),
            ("İmza", "İmza", "İmza", "İmza",)
        ]


        column_widths = [2, 2, 2, 2]
        for row_index in range(4):
            cell = table2.cell(row_index, 0)
            cell.width = Inches(column_widths[0])

        cell = table2.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

        for i in range(4):
            for j in range(4):
                cell = table2.cell(i, j)
                cell.text = cell_contents1[i][j]
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                if i == 0 or i == 1:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True

    else:

        table2 = document.add_table(rows=4, cols=5)
        for row in table2.rows:
            row.height = Inches(0)
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.space_after = Inches(0)

        cell_contents1 = [
            (f"{denetci_and_kont[0][0][2]}", f"{denetci_and_kont[0][1][2]}", f"{denetci_and_kont[0][2][2]}", f"{denetci_and_kont[0][3][2]}", f"{denetci_and_kont[0][4][2]}"),
            (f"{denetci_and_kont[0][0][3]}", f"{denetci_and_kont[0][1][3]}", f"{denetci_and_kont[0][2][3]}", f"{denetci_and_kont[0][3][3]}", f"{denetci_and_kont[0][4][3]}"),
            (f"{denetci_and_kont[0][0][1]}", f"{denetci_and_kont[0][1][1]}", f"{denetci_and_kont[0][2][1]}", f"{denetci_and_kont[0][3][1]}", f"{denetci_and_kont[0][4][1]}"),
            ("İmza", "İmza", "İmza", "İmza", "İmza")
        ]

        column_widths = [3, 3, 3, 3]
        for row_index in range(4):
            cell = table2.cell(row_index, 0)
            cell.width = Inches(column_widths[0])

        cell = table2.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

        for i in range(4):
            for j in range(5):
                cell = table2.cell(i, j)
                cell.text = cell_contents1[i][j]
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                if i == 0 or i == 1:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True



    document.add_paragraph("\n\n")

    table3 = document.add_table(rows=3, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("Kontrol / Yrd. Kontrol Elemanı", f"Şantiye Şefi"),
        (f"{mimar_or_muh[0][3]}", f"Mimar"),
        (f"{mimar_or_muh[0][1]}", f"{list2[0][13]}")]

    column_widths = [5, 5]
    for row_index in range(3):
        cell = table3.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

    cell = table3.cell(row_index, 1)
    cell.width = Inches(column_widths[1])

    for i in range(3):
        for j in range(2):
            cell = table3.cell(i, j)
            cell.text = cell_contents1[i][j]
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if i == 0 or i == 1:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True


    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)

    if "MUHTEŞEM" in list2[0][1] or "muh" in list2[0][1]:

        document.save(
            f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/kalip_donati.docx")
        # os.startfile(
        #     f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/kalip_donati.docx")
    else:
        document.save(
            f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}/kalip_donati.docx")
        # os.startfile(
        #     f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/kalip_donati.docx")


def beton_dokum():

    ek_unvan = ""
    if "MK" in list2[0][1] or "mk" in list2[0][1] or "Mk" in list2[0][1]:
        ek_unvan = "101"
    else:
        ek_unvan = "1900"


    date_string = list[3]

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    baslık2 = document.add_heading("BETON DÖKÜM TUTANAĞI\n")
    baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    table1 = document.add_table(rows=8, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("", f"  YİBF No: {list2[0][0]}"),
        ("İlgili İdare", f": {list2[0][3]}"),
        ("Yapı Sahibi", f": {list2[0][12]}"),
        ("Yapı Ruhsat tarihi ve nosu", f": {list2[0][4]}"),
        ("Yapının Adresi", f": {list2[0][7]}"),
        ("Pafta/Ada/Parsel No", f": {list2[0][9]}/{list2[0][10]}/{list2[0][11]}"),
        ("Yapı İnşaat Alanı (m²) ve Cinsi", f": {list2[0][6]} m2 {list2[0][5]}"),
        ("Yapı Denetim Kuruluşunun Unvanı/İzin Belge No", f": {list2[0][1]}/{ek_unvan}"),
    ]

    column_widths = [5, 4.8]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(8):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]
            if j == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    document.add_paragraph("\n")
    p1 = document.add_paragraph()

    p1.add_run(f"\tYukarıda belirtilen yapının ")
    bold_run1 = p1.add_run(f"{list[0]} {list[1]} ")
    bold_run1.bold = True
    p1.add_run(f"kotunda ")
    bold_run1 = p1.add_run(f"{date_string} ")
    bold_run1.bold = True
    p1.add_run("tarihinde gerçekleştirilen ")
    bold_run1 = p1.add_run(f"{list[4]} m3 ")
    bold_run1.bold = True
    p1.add_run("beton dökümü, projesine ve standartlarına uygun olarak yapılmıştır. Ayrıca beton ve beton elemanlarının numune alma ve deney metotlarına ilişkin standartlarına uygun olarak ")
    bold_run1 = p1.add_run(f"{list[2]} adet ")
    bold_run1.bold = True
    p1.add_run("beton numunesi alınmıştır. Laboratuvar deney sonuçlarına ilişkin raporlar, olumsuzluk halinde, laboratuvar tarafından düzenlenme tarihinden itibaren üç iş günü içinde, aksi takdirde hakediş eki olarak ilgili idareye iletilecektir. İş bu tutanak, bir nüshası yapı denetim kuruluşunca ilgili idareye verilmek üzere üç nüsha düzenlenmiştir.")



    table2 = document.add_table(rows=4, cols=4)
    for row in table2.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        (f"{birinci_kisim_insaat[0][2]}", f"Yardımcı Kontrol Elemanı", f"Şantiye Şefi", f"Laboratuvar"),
        (f"{birinci_kisim_insaat[0][3]}", f"{mimar_or_muh[0][3]}", f"{list2[0][13]}", f"Teknisyeni"),
        (f"{birinci_kisim_insaat[0][1]}", f"{mimar_or_muh[0][1]}", f"", f""),
        ("İmza", "İmza", "İmza", "İmza",)
    ]

    column_widths = [2, 2, 2, 2]
    for row_index in range(4):
        cell = table2.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

    cell = table2.cell(row_index, 1)
    cell.width = Inches(column_widths[1])

    for i in range(4):
        for j in range(4):
            cell = table2.cell(i, j)
            cell.text = cell_contents1[i][j]
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if i == 0 or i == 1:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True



    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)

    if "MUHTEŞEM" in list2[0][1] or "muh" in list2[0][1]:

        document.save(
            f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/beton_dokum.docx")
        # os.startfile(
        #     f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/beton_dokum.docx")
    else:
        document.save(
            f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}/beton_dokum.docx")
        # os.startfile(
        #     f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/beton_dokum.docx")

def iskele_sokum():

    ek_unvan = ""
    if "MK" in list2[0][1] or "mk" in list2[0][1] or "Mk" in list2[0][1]:
        ek_unvan = "101"
    else:
        ek_unvan = "1900"

    try:
        date_string = list[3]

        date_format = "%d.%m.%Y"
        date = datetime.strptime(date_string, date_format)
        three_days_before = date + timedelta(days=15)
        result_string = three_days_before.strftime(date_format)
    except Exception as e:
        messagebox.showerror("HATA !", "Lütfen beton tarihini kontrol ediniz ! ")

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    baslık2 = document.add_heading("KALIP VE TAŞIYICI KALIP İSKELESİ SÖKÜMÜ KONTROL TUTANAĞI\n")
    baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    table1 = document.add_table(rows=8, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("", f"  YİBF No: {list2[0][0]}"),
        ("İlgili İdare", f": {list2[0][3]}"),
        ("Yapı Sahibi", f": {list2[0][12]}"),
        ("Yapı Ruhsat tarihi ve nosu", f": {list2[0][4]}"),
        ("Yapının Adresi", f": {list2[0][7]}"),
        ("Pafta/Ada/Parsel No", f": {list2[0][9]}/{list2[0][10]}/{list2[0][11]}"),
        ("Yapı İnşaat Alanı (m²) ve Cinsi", f": {list2[0][6]} m2 {list2[0][5]}"),
        ("Yapı Denetim Kuruluşunun Unvanı/İzin Belge No", f": {list2[0][1]}/{ek_unvan}"),
    ]

    column_widths = [5, 4.8]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(8):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]
            if j == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    document.add_paragraph("\n\n")
    p1 = document.add_paragraph()

    p1.add_run(f"Yukarıda belirtilen yapının ")
    bold_run1 = p1.add_run(f"{list[0]} {list[1]} ")
    bold_run1.bold = True
    p1.add_run("kotunda yapılan denetimde:")
    document.add_paragraph("Kalıp ve taşıyıcı kalıp iskelelerinin sökümünde bir mahzur bulunmadığı tespit edilmiştir.")

    p3 = document.add_paragraph()
    p3.add_run(f"İş bu tutanak ")
    bold_run2 = p3.add_run(f"{result_string} ")
    bold_run2.bold = True
    p3.add_run("tarihinde bir nüshası inşaat mühendisi fenni mesul/ yapı denetçisi tarafından ilgili idareye verilmek üzere üç nüsha olarak düzenlenmiştir."+"\n"*4)

    table2 = document.add_table(rows=3, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("Fenni mesul/Yapı Denetçisi", f"Şantiye Şefi"),
        ("İnşaat Mühendisi", f"{list2[0][13]}"),
        ("İmza", f"İmza")]


    column_widths = [5, 5]
    for row_index in range(3):
        cell = table2.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

    cell = table2.cell(row_index, 1)
    cell.width = Inches(column_widths[1])

    for i in range(3):
        for j in range(2):
            cell = table2.cell(i, j)
            cell.text = cell_contents1[i][j]
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if i == 0 or i == 1:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)

    if "MUHTEŞEM" in list2[0][1] or "muh" in list2[0][1]:

        document.save(
            f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/iskele_sokum.docx")
        os.startfile(f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}")
    else:
        document.save(
            f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}/iskele_sokum.docx")
        os.startfile(f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}")

def temel_topraklama():

    ek_unvan = ""
    if "MK" in list2[0][1] or "mk" in list2[0][1] or "Mk" in list2[0][1]:
        ek_unvan = "101"
    else:
        ek_unvan = "1900"

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    baslık2 = document.add_heading("TEMEL TOPRAKLAMA KONTROL TUTANAĞI\n")
    baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    table1 = document.add_table(rows=8, cols=2)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("", f"  YİBF No: {list2[0][0]}"),
        ("İlgili İdare", f": {list2[0][3]}"),
        ("Yapı Sahibi", f": {list2[0][12]}"),
        ("Yapı Ruhsat tarihi ve nosu", f": {list2[0][4]}"),
        ("Yapının Adresi", f": {list2[0][7]}"),
        ("Pafta/Ada/Parsel No", f": {list2[0][9]}/{list2[0][10]}/{list2[0][11]}"),
        ("Yapı İnşaat Alanı (m²) ve Cinsi", f": {list2[0][6]} m2 {list2[0][5]}"),
        ("Yapı Denetim Kuruluşunun Unvanı/İzin Belge No", f": {list2[0][1]}/{ek_unvan}"),
    ]

    column_widths = [5, 4.8]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(8):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]
            if j == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    document.add_paragraph("\n\n")
    p1 = document.add_paragraph()

    p1.add_run(f"\tİş bu tutanak ")
    bold_run1 = p1.add_run(f"{demir_cekme} ")
    bold_run1.bold = True
    p1.add_run("tarihinde yukarıda bilgileri olan binanın temel topraklamasına ilişkin olarak düzenlenmiş ve taraflarca imza altına alınmıştır.")

    document.add_paragraph("-Temel topraklamasının temel aşamasında ve projede belirtilen planına uygun olarak yapıldığı,")
    document.add_paragraph("-Temel donatı deneyine bağlanan garveniz şeridin 30*3,5 mm’lik ve kalitesinin uygun olduğu, klemensle bağlandığı,")
    document.add_paragraph("-Eş potansiyel bara, ana pano ve asansör kuyusuna birer filiz bırakılmış bulunduğu,")
    document.add_paragraph("-Elektrotların usulüne uygun olarak toprağa gömüldüğü,")
    document.add_paragraph("-Temel betonun dökümüne herhangi bir sakınca bulunmadığı,")
    document.add_paragraph("Tespit edilmiştir.\n\n")

    table2 = document.add_table(rows=3, cols=3)
    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("Proje ve Uygulama Denetçisi", f"{list2[0][1]}", "Şantiye Şefi"),
        ("Elektrik Mühendisi", "", "Mimar",),
        ("AHMET PEKER", "", f"{list2[0][13]}"),
    ]


    column_widths = [2, 2, 2]
    for row_index in range(3):
        cell = table2.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

    cell = table2.cell(row_index, 1)
    cell.width = Inches(column_widths[1])

    for i in range(3):
        for j in range(3):
            cell = table2.cell(i, j)
            cell.text = cell_contents1[i][j]
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if i == 0 or i == 1:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True


    document.add_paragraph("\n\n")

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)

    if "TEMEL" in list[0] or "temel" in list[0]:

        if "MUHTEŞEM" in list2[0][1] or "muh" in list2[0][1]:

            document.save(
                f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/{list[0]}/temel_topraklama.docx")
            # os.startfile(
            #     f"{save_path}/MUHTEŞEM/{list2[0][10]}-{list2[0][11]}/temel_topraklama.docx")
        else:
            document.save(
                f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/{list[0]}/temel_topraklama.docx")
            # os.startfile(
            #     f"{save_path}/MK/{list2[0][10]}-{list2[0][11]}/temel_topraklama.docx")



def folder_update():

    home_dir = os.path.expanduser("~")
    desktop_path = os.path.join(home_dir, "Desktop")

    folder_path = os.path.join(desktop_path, "BETON TUTANAKLARI")
    os.makedirs(folder_path, exist_ok=True)

    global save_path

    save_path = f"{desktop_path}/BETON TUTANAKLARI"

    save_folder = "MUHTEŞEM"
    save_folder2 = "MK"

    try:
        result_muh = sql.sql_query("ada, parsel" , "beton_kont", "ydk",
                                         "MUHTEŞEM YAPI DENETİM LİMİTED ŞİRKETİ")
        result_mk = sql.sql_query("ada, parsel" , "beton_kont", "ydk",
                                         "MK YAPI DENETİM LİMİTED ŞİRKETİ")

        desktop_path = os.path.join(os.path.expanduser("~"), save_path)
        parent_folder_name = save_folder
        parent_folder_path = os.path.join(desktop_path, parent_folder_name)
        os.makedirs(parent_folder_path, exist_ok=True)

        def create_folders(folder_structure, parent_path="."):
            for folder_name in folder_structure:
                folder_path = os.path.join(parent_path, folder_name)
                os.makedirs(folder_path, exist_ok=True)
                parent_path = folder_path

        for i in result_muh:
            folder_structure = [f"{i[0]}-{i[1]}", f"{list[0]}"]
            create_folders(folder_structure, parent_path=parent_folder_path)

        desktop_path = os.path.join(os.path.expanduser("~"), save_path)
        parent_folder_name = save_folder2
        parent_folder_path = os.path.join(desktop_path, parent_folder_name)
        os.makedirs(parent_folder_path, exist_ok=True)

        for i in result_mk:
            folder_structure = [f"{i[0]}-{i[1]}", f"{list[0]}"]
            create_folders(folder_structure, parent_path=parent_folder_path)

    except Exception as e:
        print(f"Make Files Error: {e}")


def call_beton(row_values, present_data, donati_cekme):

    print(present_data)
    bet_result = sql.sql_query("denetci_b, kontrol_b", "beton_personel", "ada_b", present_data[0][10], "parsel_b", present_data[0][11])

    global list, list2, demir_cekme, denetci_and_kont, mimar_or_muh, birinci_kisim_insaat

    list = row_values
    list2 = present_data
    demir_cekme = donati_cekme

    denetci_and_kont = []
    for i in bet_result:
        for j in i:
            denetci_and_kont.append(json.loads(j))


    mimar_or_muh = []
    flag = 0
    for i in denetci_and_kont[1]:
        if i[3] == "İnşaat Mühendisi":
            flag = 1
            mimar_or_muh.insert(0, i)
        elif i[3] == "Mimar":
            flag = 1
            mimar_or_muh.append(i)

    birinci_kisim_insaat = []
    for i in denetci_and_kont[0]:
        if i[3] == "İnşaat Mühendisi":
            birinci_kisim_insaat.insert(0, i)



    if flag == 0:
        messagebox.showerror("HATA !", "  Kontrol / Yrd. Kontrol Elemanı İnşaat Mühendisi EKSİK !\n\n  Lütfen YİBF Bilgisini Kontrol edin ! ")
        return

    folder_update()
    iskele_kurulum()
    kalip_donati()
    beton_dokum()
    iskele_sokum()
    temel_topraklama()

