# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
import sql_modul as sql
from tkinter import messagebox

# sonuc_ana = [('22.11.2023', '23.11.2023', '22.11.2023', '', '1', 'MUHTESEM Yapı Denetim Ltd. Şti./ Dosya No: 1900',
# 'MUHTEŞEM Yapı Denetim Ltd. Şti', '23', '17 (A BLOK)', 'P24B16B1A', 'AKSU', '1979672',
# '22.11.2023 / 2023/72A', 'BAĞLIK (GÖKSU) MAHALLESİ KUMLUCA/ANTALYA', '3012', '4A', 'ÜMMÜGÜLSÜM DURAKOĞLU',
# 'OGUZ', '10', '', '', '', '', '', '10', 'ON', '548.327,37', '10', '1.644,98', '1.644,98', '51.542,77',
# '10.966,55', '', '62.509,32',
# 'FENERCİOĞLU SİSTEM BETONARME YAPI KONTROL LABORATUVARI SAN. VE TİC. LTD. ŞTİ.', '2000', '1000')]

sonuc_ana = []
save_path = ""

def call():
    global sonuc_ana
    print(sonuc_ana)
    sonuc_ana = [list(item) for item in sonuc_ana]

    if sonuc_ana[0][10] == "AKSU":
        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])
        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0)
            # section.left_margin = Inches(0.5)
            # section.right_margin = Inches(0.5)

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER\n")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=7, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("ADA / PARSEL", f": {sonuc_ana[0][7]} ADA {sonuc_ana[0][8]} PARSEL"),
                ("MAHALLE", f": {sonuc_ana[0][13]}"),
                ("İL/İLÇE / İLGİLİ İDARE", ": ANTALYA / AKSU / AKSU BELEDİYESİ"),
                ("YİBF NO", f": {sonuc_ana[0][11]}"),
                ("YAPI DENETİM KURULUŞU", f": {sonuc_ana[0][5]}"),
                ("IBAN/VERGİ NUMARASI", f": {iban_vd[0]}"),
                ("KURUMLAR V.D", f": {iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(7):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(7):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz" and sonuc_ana[0][34] != "" :
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :\n")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"{comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run("'nin nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "OSB":

        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])

        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            baslık2 = document.add_heading(
                "ANTALYA\nORGANİZE SANAYİ BÖLGESİ MÜDÜRLÜĞÜ\nİmar ve Fen İşleri Müdürlüğü\nDÖŞEMEALTI MAL MÜDÜRLÜĞÜ’NE")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": {sonuc_ana[0][13]}"),
                ("İlgili İdare", f": ORGANİZE SANAYİ BÖLGESİ"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.")

            if lab_flag == 7:
                document.add_paragraph("\n")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)



            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "ALANYA":

        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])

        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.3)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            document.add_picture('Pictures/alanya.png', width=Inches(6.5), height=Inches(1.2))

            p1 = document.add_paragraph()
            bold_run1 = p1.add_run("Konu :")
            bold_run1.bold = True
            p1.add_run("  Yapı denetim ödemesi hk")

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            for run in baslık2.runs:
                run.font.size = Pt(14)

            for paragraph in document.paragraphs:
                if paragraph.style.name.startswith('Heading'):
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(5)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": Alanya"),
                ("İlgili İdare / Mahalle", f": Alanya / {sonuc_ana[0][13]}"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz" and sonuc_ana[0][34] != "":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run("nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.\n\n\n\n\n")

            if lab_flag == 7:
                p2 = document.add_paragraph("\n" * 4)

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            document.add_picture('Pictures/alanya2.png', width=Inches(6), height=Inches(0.5))

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "ANTALYA BÜYÜKŞEHİR":
        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])

        if query_bel:
            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            document.add_picture('Pictures/buyuksehir.png', width=Inches(6.5), height=Inches(1.2))

            p1 = document.add_paragraph()
            bold_run1 = p1.add_run("Konu :")
            bold_run1.bold = True
            p1.add_run("  Yapı Denetim Hakedişi")
            p1.add_run("\t" * 7 + f"{sonuc_ana[0][1]}")

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": {sonuc_ana[0][10]}"),
                ("İlgili İdare", ": Antalya Büyükşehir Belediyesi"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz" and sonuc_ana[0][34] != "":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.paragraph_format.space_after = Pt(0)
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2
            p1.add_run(f" nolu hesabına aktarılması,")
            p2 = document.add_paragraph()
            p2.paragraph_format.space_after = Pt(0)
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.paragraph_format.space_after = Pt(0)
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run("nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)


            p5 = document.add_paragraph("\tGereğini arz ederim.")

            if lab_flag == 7:
                document.add_paragraph("\n\n\n\n\n")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            isim_1 = document.add_paragraph("\t" * 5 + "     TUNCAY KAYA\t\t    BARIŞ SOYKAM")
            isim_2 = document.add_paragraph("\t" * 5 + "İmar Şube Müdürü V.\t\tKent Estetiği Dai. Bşk")

            isim_1.paragraph_format.space_before = Pt(2)
            isim_1.paragraph_format.space_after = Pt(2)
            isim_2.paragraph_format.space_before = Pt(2)
            isim_2.paragraph_format.space_after = Pt(2)

            document.add_picture('Pictures/buyuksehir2.png', width=Inches(6), height=Inches(1))

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "MURATPAŞA":
        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])

        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            baslık2 = document.add_heading(
                "ANTALYA MURATPAŞA BELELDİYESİ\nİmar Ve Şehircilik Müdürlüğü\nMURATPAŞA MAL MÜDÜRLÜĞÜ’NE")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": Muratpaşa"),
                ("İlgili İdare", f": Muratpaşa Belediyesi"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tGereğini arz ederim.")


            document.add_paragraph("\n\n\n\n")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            isim_1 = document.add_paragraph("\t" * 10 + " ERGİN SARI")
            mev = document.add_paragraph("İmar ve Şehircilik Müdürü")
            mev.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            isim_1.paragraph_format.space_before = Pt(2)
            isim_1.paragraph_format.space_after = Pt(2)
            mev.paragraph_format.space_before = Pt(2)
            mev.paragraph_format.space_after = Pt(2)



            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")


    elif sonuc_ana[0][10] == "DÖŞEMEALTI":

        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])
        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            baslık2 = document.add_heading("T.C.\nDÖŞEMEALTI BELEDİYE BAŞKANLIĞI\nRuhsat ve Denetim Müdürlüğü")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            bold_run1 = p1.add_run("Sayı :")
            bold_run1.bold = True
            p1.add_run("\t" * 11 + f"{sonuc_ana[0][1]}\n")
            p1.add_run(f"Konu  : {sonuc_ana[0][6]}  {sonuc_ana[0][7]}/{sonuc_ana[0][8]} {sonuc_ana[0][4]} Nolu Hak.")

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": Döşemealtı"),
                ("İlgili İdare", f": Döşemealtı Belediyesi"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True  # Make the text bold
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.")

            document.add_paragraph("\n\n\n\n\n")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            isim = document.add_paragraph("\t" * 9 + "       Hayrettin BAHŞİ")
            mev = document.add_paragraph("Ruhsat ve Denetim Müdür V.")
            mev.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            isim.paragraph_format.space_before = Pt(2)
            isim.paragraph_format.space_after = Pt(2)
            mev.paragraph_format.space_before = Pt(2)
            mev.paragraph_format.space_after = Pt(2)

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "MANAVGAT":
        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])

        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            baslık2 = document.add_heading("T.C.\nMANAVGAT BELEDİYESİ\nYapı Kontrol Müdürlüğü\n")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            bold_run1 = p1.add_run("Sayı    : 64053976-310.11.02-")

            p1.add_run("\t" * 8 + f"{sonuc_ana[0][1]}\n")
            p1.add_run(f"Konu  : Denetim Hizmet Bedeli Hakedişi.")

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Mahallesi", f": {sonuc_ana[0][13]}"),
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": ANTALYA"),
                ("İlçesi", f": MANAVGAT"),
                ("İlgili İdare", f": MANAVGAT BELEDİYESİ"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                # ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

                # Populate the first table with data
            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True  # Make the text bold
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2, p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tGereğini arz ederim.")

            if lab_flag == 7:
                document.add_paragraph("\n\n\n\n")
            else:
                document.add_paragraph("\n\n")

            for i in [p1, p2, p3]:
                i.paragraph_format.space_before = Pt(2)
                i.paragraph_format.space_after = Pt(2)



            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            isim = document.add_paragraph("\t" * 10 + "   Mustafa YAVUZ")
            mev = document.add_paragraph("Yapı Kontrol Müd.V.")
            mev.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            isim.paragraph_format.space_before = Pt(2)
            isim.paragraph_format.space_after = Pt(2)
            mev.paragraph_format.space_before = Pt(2)
            mev.paragraph_format.space_after = Pt(2)

            tar = document.add_paragraph("... / 02 / 2024")
            tar.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "FİNİKE":

        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])
        if query_bel:

            muhtesem = [
                "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : ",
                "TR67 0020 5000 0956 6541 9000 01"]
            mk = [
                "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: ",
                "TR05 0020 5000 0958 5145 8000 01"]

            muhtesem_kurulus = ""
            iban_vd = []
            comp = []

            if "MK" in sonuc_ana[0][6] or "mk" in sonuc_ana[0][6]:
                comp = mk
                iban_vd.append("TR05 0020 5000 0958 5145 8000 01")
                iban_vd.append("6220189578")
            else:
                comp = muhtesem
                iban_vd.append("TR67 0020 5000 0956 6541 9000 01")
                iban_vd.append("6230334026")

            print(comp)

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

            document.add_picture('Pictures/finike.png', width=Inches(6.5), height=Inches(1.1))

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER:")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            table1 = document.add_table(rows=8, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", f": Antalya"),
                ("İlçesi", f": Finike"),
                ("İlgili İdare", f": Finike Belediyesi"),
                ("YİBF No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
                ("IBAN/Vergi Numarası", f": {iban_vd[0]}/{iban_vd[1]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(8):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(8):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])

            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp[0]}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.")

            if lab_flag == 7:
                document.add_paragraph("\n\n\n\n")
            else:
                document.add_paragraph("\n\n\n")


            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            isim = document.add_paragraph("\t" * 9 + "    Çiğdem ÇETİNKAYA")
            mev = document.add_paragraph("Ruhsat ve Denetim Müdür Vek.\n")
            mev.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            isim.paragraph_format.space_before = Pt(2)
            isim.paragraph_format.space_after = Pt(2)
            mev.paragraph_format.space_before = Pt(2)
            mev.paragraph_format.space_after = Pt(2)

            document.add_picture('Pictures/finike2.png', width=Inches(6.5), height=Inches(0.8))

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    elif sonuc_ana[0][10] == "KEPEZ":

        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])
        if query_bel:
            muhtesem = "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : TR67 0020 5000 0956 6541 9000 01"
            mk = "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: TR05 0020 5000 0958 5145 8000 01"
            muhtesem_kurulus = ""

            comp = ""
            if "MK" in sonuc_ana[0][6]:
                comp = mk
            else:
                comp = muhtesem

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0)
            # section.left_margin = Inches(0.5)
            # section.right_margin = Inches(0.5)

            baslık = document.add_heading(
                f"T.C.\nANTALYA {sonuc_ana[0][10]} BELEDİYESİ\nİmar ve Şehircilik Müdürlüğü\n{sonuc_ana[0][10]} MAL MÜDÜRLÜĞÜ’NE")
            baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            baslık.paragraph_format.space_after = Pt(0)

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER :\n")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            baslık2.runs[0].font.underline = True
            baslık2.paragraph_format.space_before = Pt(0)

            table1 = document.add_table(rows=6, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", ": Antalya"),
                ("İlgili İdare", f": {sonuc_ana[0][10]}"),
                ("Y.İ.B.F. No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(6):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(6):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz" and sonuc_ana[0][34] != "":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True

            p3.add_run(f"'nin {comp}")
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.\n\n\n\n")

            document.add_paragraph("\t" * 9 + "ERDİ KARA\n" + "\t" * 7 + "              İmar ve Şehircilik Müdür V.")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")

    else:
        query_bel = sql.sql_query("*", "belediye", "adi", sonuc_ana[0][10])
        if query_bel:
            muhtesem = "MUHTEŞEM Yapı Denetim Ltd. Şti.’nin (Kurumlar Vergi Dairesi -6230334026 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN : TR67 0020 5000 0956 6541 9000 01"
            mk = "MK Yapı Denetim Ltd. Şti.'nin (Kurumlar Vergi Dairesi -6220189578 VKN) Kuveyt Türk Bankası Çallı şubesi IBAN: TR05 0020 5000 0958 5145 8000 01"
            muhtesem_kurulus = ""

            comp = ""
            if "MK" in sonuc_ana[0][6]:
                comp = mk
            else:
                comp = muhtesem

            document = Document()
            section = document.sections[0]

            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0)
            # section.left_margin = Inches(0.5)
            # section.right_margin = Inches(0.5)

            baslık = document.add_heading(
                f"T.C.\nANTALYA {sonuc_ana[0][10]} BELEDİYESİ\nİmar ve Şehircilik Müdürlüğü\n{sonuc_ana[0][10]} MAL MÜDÜRLÜĞÜ’NE")
            baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            baslık.paragraph_format.space_after = Pt(0)

            baslık2 = document.add_heading("TAHAKKUKA ESAS BİLGİLER :\n")
            baslık2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslık2.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            baslık2.runs[0].font.underline = True
            baslık2.paragraph_format.space_before = Pt(0)

            table1 = document.add_table(rows=6, cols=2)
            for row in table1.rows:
                row.height = Inches(0)
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.space_after = Inches(0)

            cell_contents1 = [
                ("Ada", f": {sonuc_ana[0][7]}"),
                ("Parsel", f": {sonuc_ana[0][8]}"),
                ("İli", ": Antalya"),
                ("İlgili İdare", f": {sonuc_ana[0][10]}"),
                ("Y.İ.B.F. No", f": {sonuc_ana[0][11]}"),
                ("Yapı Denetim Kuruluşu", f": {sonuc_ana[0][5]}"),
            ]

            column_widths = [2, 4]
            for row_index in range(6):
                cell = table1.cell(row_index, 0)
                cell.width = Inches(column_widths[0])

                cell = table1.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(6):
                for j in range(2):
                    cell = table1.cell(i, j)
                    cell.text = cell_contents1[i][j]

            space = document.add_paragraph("\n")

            if sonuc_ana[0][32] == "" or sonuc_ana[0][32] == " -":
                sonuc_ana[0][32] = " -"
            else:
                sonuc_ana[0][32] = f"{sonuc_ana[0][32]} TL"

            lab_flag = 7

            if sonuc_ana[0][34] != "Lab Seçiniz" and sonuc_ana[0][34] != "":
                lab_flag = 8
                lab_info = sql.sql_query("*", "lab", "adi", sonuc_ana[0][34])
            else:
                lab_flag = 7

            table2 = document.add_table(rows=lab_flag, cols=2)
            cell_contents2 = [
                ("Denetim Hizmet Bedeline esas tutar (KDV hariç)", f": {sonuc_ana[0][26]} TL"),
                ("Gerçekleşen Fiziki Seviye", f": % {sonuc_ana[0][27]}"),
                ("% 3 İlgili İdare Payı (KDV hariç)", f": {sonuc_ana[0][28]} TL"),
                ("% 3 Bakanlık Payı (KDV hariç)", f": {sonuc_ana[0][29]} TL"),
                ("Yapı Denetim Kuruluşu Payı (KDV hariç)", f": {sonuc_ana[0][30]} TL"),
                ("KDV", f": {sonuc_ana[0][31]} TL"),
                ("Tevkifat", f": {sonuc_ana[0][32]}"),
                ("Laboratuvar", f": {sonuc_ana[0][35]} TL"),
            ]

            column_widths = [6, 4]
            for row_index in range(lab_flag):
                cell = table2.cell(row_index, 0)
                cell.width = Inches(column_widths[0])
                cell = table2.cell(row_index, 1)
                cell.width = Inches(column_widths[1])

            for i in range(lab_flag):
                for j in range(2):
                    cell = table2.cell(i, j)
                    cell.text = cell_contents2[i][j]

            baslik2 = document.add_heading("ÖDEME ŞEKLİ :")
            baslik2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            p1 = document.add_paragraph()
            p1.add_run(
                f"1) 333.99.47.1 nolu emanet hesabından, 333.99.47.2 nolu ilgili idare hesabına aktarılan KDV hariç ")
            bold_run1 = p1.add_run(f"{sonuc_ana[0][28]} TL")
            bold_run1.bold = True
            p1.add_run(f"’nin {sonuc_ana[0][10]} Belediyesi ({query_bel[0][1]}) {query_bel[0][2]} ")
            bold_run2 = p1.add_run(str(query_bel[0][3]))
            #bold_run2.bold = True
            p1.add_run(f" nolu hesabına aktarılması,")

            p2 = document.add_paragraph()
            p2.add_run(
                "2) 333.99.47.1 nolu emanet hesabından, 333.99.47.3 nolu Çevre ve Şehircilik Bakanlığı döner sermaye payı hesabına aktarılan KDV hariç ")
            bold_run1 = p2.add_run(f"{sonuc_ana[0][29]} TL")
            bold_run1.bold = True
            p2.add_run(
                "'nin (Ankara Mithatpaşa Vergi Dairesi-1530522399 VKN) Halk Bankası A.Ş. Ankara Kurumsal Şubesi IBAN : ")
            iban = "TR98 0001 2009 4520 0005 0000 22"
            bold_run2 = p2.add_run(iban)
            #bold_run2.bold = True
            p2.add_run(" nolu hesabına aktarılması,")

            p3 = document.add_paragraph()
            p3.add_run("3) 333.38.01 nolu emanet hesabından, KDV dahil ")
            if lab_flag == 8:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]}-{sonuc_ana[0][35]}={sonuc_ana[0][36]} TL")
                bold_run.bold = True
            else:
                bold_run = p3.add_run(f"{sonuc_ana[0][33]} TL")
                bold_run.bold = True
            p3.add_run(f"'nin {comp}")
            bold_run = p3.add_run(f"{comp[1]}")
            #bold_run.bold = True
            p3.add_run(" nolu hesabına aktarılması hususunda,")

            if lab_flag == 8:
                p4 = document.add_paragraph()
                p4.add_run("4) 333.99.47.1 nolu emanet hesabından, KDV dahil ")
                bold_run = p4.add_run(f"{sonuc_ana[0][35]} TL")
                bold_run.bold = True
                p4.add_run(
                    f"’nin {lab_info[0][0]}'nin {lab_info[0][1]} ndeki {lab_info[0][2]} ")
                bold_run = p4.add_run(f" {lab_info[0][3]}")
                #bold_run.bold = True
                p4.add_run(" nolu iban hesabına aktarılması hususunda,")

                for i in [p1, p2 ,p3, p4]:
                    i.paragraph_format.space_before = Pt(2)
                    i.paragraph_format.space_after = Pt(2)

            p5 = document.add_paragraph("\tBilgi ve gereğini rica ederim.")

            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                            paragraph.paragraph_format.space_before = Pt(1)
                            paragraph.paragraph_format.space_after = Pt(1)

            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":
                document.save(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
            else:
                document.save(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")
                os.startfile(
                    f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/tahakkuk.docx")

        else:
            messagebox.showerror("Hata !", "İlgili daireye ait banka bilgileri eksik !")
