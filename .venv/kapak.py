# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

sonuc_ana = []

save_path = ""


def kapak_kepez():
    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.9)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(1.5)

    yibf = document.add_paragraph("Sayı:2024/\nKonu: Hakediş Raporu")
    yibf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    baslık = document.add_heading(f"KEPEZ BELEDİYESİ\nİMAR VE ŞEHİRCİLİK MÜDÜRLÜĞÜ\n"+"\t"*6+"ANTALYA\n\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    if "MK" not in sonuc_ana[0][6]:
        note = document.add_paragraph(f"""
        \tYapı Denetim Hizmeti tarafımızca yürütülmekte olan {sonuc_ana[0][7]} ada {sonuc_ana[0][8]} parselde {sonuc_ana[0][11]} Y.İ.B.F. No'lu yapılan inşaatın {sonuc_ana[0][4]} no’lu Yapı Denetim Hizmet Bedeli hakediş raporu, yazımız ekinde sunulmuştur.
        \tBahse konu inşaatın dosyasında ve mahalinde yapılan incelemeyle ilgili;  Kanun, Yönetmelik, Şartname ve Standartlara uygun, Yapı Ruhsatı eki projelere aykırı bir imalat olmadığı tarafımızdan tespit edildiği için {sonuc_ana[0][4]} no'lu %{sonuc_ana[0][24]} seviye yapı denetim hizmet bedelinin Mal Müdürlüğü'ne ödenmesi hususunda,\n
        \tGerekenin yapılmasını saygılarımla arz ederim.
        """ + "\t" * 9 + f"        {sonuc_ana[0][0]}")

    else:
        note = document.add_paragraph(f"""
        \tYapı Denetim Hizmeti tarafımızca yürütülmekte olan {sonuc_ana[0][7]} ada {sonuc_ana[0][8]} parselde {sonuc_ana[0][11]} Y.İ.B.F. No'lu yapılan inşaatın {sonuc_ana[0][4]} no’lu Yapı Denetim Hizmet Bedeli hakediş raporu, yazımız ekinde sunulmuştur.
        \tBahse konu inşaatın dosyasında ve mahalinde yapılan incelemeyle ilgili;  Kanun, Yönetmelik, Şartname ve Standartlara uygun, Yapı Ruhsatı eki projelere aykırı bir imalat olmadığı tarafımızdan tespit edildiği için {sonuc_ana[0][4]} no'lu %{sonuc_ana[0][24]} seviye yapı denetim hizmet bedelinin Mal Müdürlüğü'ne ödenmesi hususunda,\n
        \tGerekenin yapılmasını saygılarımla arz ederim.\n
        """ + "\t" * 9 + f"        {sonuc_ana[0][0]}")


    ydk = document.add_paragraph(f"{sonuc_ana[0][6]}\n\n\n\n\n")
    ydk.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    if "MK" not in sonuc_ana[0][6]:
        adres_muh = document.add_paragraph(f"""
MUHTEŞEM Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6230334026 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR67 0020 5000 0956 6541 9000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/2 Konyaaltı /ANTALYA
Tel: 0 242 229 35 75
        """)

    else:
        adres_mk = document.add_paragraph(f"""
MK Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6220189578 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR05 0020 5000 0958 5145 8000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/7  Konyaaltı /ANTALYA
Tel: 0 506 136 98 23
        """)

    ekler = document.add_paragraph("    EK: ")
    ekler.add_run("""
    1) 3 Adet Hakediş raporu
    2) 3 Adet Personel Listesi
    3) 3 Adet Makbuz
    4) 3 Adet fatura""")

    # Change the font color of the text in the 'ekler' paragraph to black
    for run in ekler.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)  # Set to black
        if "EK" in run.text:
            run.font.bold = True

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")

def kapak_manavgat():
    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.9)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(1.5)

    yibf = document.add_paragraph("Sayı:2024/\nKonu: Hakediş Raporu")
    yibf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    baslık = document.add_heading(f"MANAVGAT BELEDİYESİ\nYAPI KONTROL MÜDÜRLÜĞÜ’NE\n"+"\t"*6+"ANTALYA\n\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    note = document.add_paragraph(f"""
\t4708 sayılı Yapı Denetim Kanunu uyarınca Denetim görevini üstlendiğimiz Manavgat İlçesi, {sonuc_ana[0][13]} imarın {sonuc_ana[0][7]} ada {sonuc_ana[0][8]} parsel {sonuc_ana[0][11]} YİBF numaralı, {sonuc_ana[0][12]} tarih/ruhsat numaralı, Yapı Sahibi {sonuc_ana[0][16]} olan inşaatın {sonuc_ana[0][4]} no’lu yüzde {sonuc_ana[0][24]} seviye Yapı Denetim Hizmet Bedeli hakediş raporu,yazımız ekinde sunulmuştur.
\tGerekenin yapılmasını saygılarımla arz ederim. {sonuc_ana[0][0]}\n\n""")

    ydk = document.add_paragraph(f"{sonuc_ana[0][6]}\n\n\n\n\n")
    ydk.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    if "MK" not in sonuc_ana[0][6]:
        adres_muh = document.add_paragraph(f"""
MUHTEŞEM Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6230334026 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR67 0020 5000 0956 6541 9000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/2 Konyaaltı /ANTALYA
Tel: 0 242 229 35 75
        """)
    else:
        adres_mk = document.add_paragraph(f"""
MK Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6220189578 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR05 0020 5000 0958 5145 8000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/7  Konyaaltı /ANTALYA
        Tel: 0 506 136 98 23
        """)

    ekler = document.add_paragraph("    EK: ")
    ekler.add_run("""
    1) 3 Adet Hakediş raporu
    2) 3 Adet Personel Listesi
    3) 3 Adet Makbuz
    4) 3 Adet fatura""")

    # Change the font color of the text in the 'ekler' paragraph to black
    for run in ekler.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)  # Set to black
        if "EK" in run.text:
            run.font.bold = True

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")

def kapak_korkuteli():
    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.9)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(1.5)

    baslık = document.add_heading(f"KORKUTELİ BELEDİYESİ\nİmar ve Şehircilik Müdürlüğü’ne\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    details = [
        "Konu: Ulusal Yapı Denetim Sisteminde(UYDS’de) Hakediş Ödeme Talebi",
        f"Yibf No: {sonuc_ana[0][11]}",
        f"Ruhsat Tarih ve Nosu: {sonuc_ana[0][12]}",
        f"İnşaat Mahallesi: {sonuc_ana[0][13]}.",
        f"Ada/Parseli: {sonuc_ana[0][7]}/{sonuc_ana[0][8]}"
    ]

    # Add the details to the document with the label in bold
    for detail in details:
        label, value = detail.split(": ", 1)  # Split the detail into label and value
        p = document.add_paragraph()
        p.add_run(label + ": ").bold = True
        p.add_run(value)


    note = document.add_paragraph(f"\tYukarıda bilgileri bulunan inşaatın firmamızca mahallinde yapılan denetiminde ilgili kanun, yönetmelik, şartname, standartlara, ruhsat ve ek projelere aykırı bir durumun olmadığını taahhüt ederek {sonuc_ana[0][4]} Nolu  %{sonuc_ana[0][24]} seviyedeki denetim hizmet bedeli hakediş raporunun, Ulusal Yapı Denetim Sisteminde(UYDS’de) ödenmesi hususunda;\n\tGereğini arz ederim.")

    ydk = document.add_paragraph(f"{sonuc_ana[0][6]}\n\n\n")
    ydk.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    table_war = document.add_table(rows=1, cols=1, style="Table Grid")
    cell = table_war.cell(0, 0)
    paragraph_content = "Tarafımıza ve Korkuteli Malmüdürlüğü Muhasebe Servis Şefliği’ne iletilecek ilgili onaylı hakkediş evrakları Firmamızda çalışan personellerce elden takipli olacaktı"
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(paragraph_content)
    paragraph.alignment = 1

    note2 = document.add_paragraph()
    note2_run = note2.add_run("\nEKLER:")
    note2_run.bold = True

    document.add_paragraph("""
    1) Denetim Hizmet Bedeline ait Hakediş Raporu(3 Sayfa)
    2) Tahakkuka Esas Bilgiler(3 Sayfa)
    3) Emanet Hesaba Ödenen Dekont ve Alındı Belgeleri(Tamamı)
    4) Personel Bildirgesi(Tarih aralığı belirtilmiş olarak 3 Sayfa)
    5) Yapıya İlişkin Bilgi Formu(YİBF Çıktısı 3 sayfa)
    6) Seviye Tespit Tutanağı(2 Sayfa)
    7) Yapı Denetim Hizmet Sözleşmesi ASLI(%10’luk hakkedişte sadece)
    8) Taahhütname
    9) Fatura
    
       *Borcu yoktur yazısı Malmüdürlüğü Nüshasında güncel tarihli olmalıdır""")

    if "MK" not in sonuc_ana[0][6]:
        adres_muh = document.add_paragraph(f"""
MUHTEŞEM Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6230334026 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR67 0020 5000 0956 6541 9000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/2 Konyaaltı /ANTALYA
Tel: 0 242 229 35 75
        """)
    else:
        adres_mk = document.add_paragraph(f"""
MK Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6220189578 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR05 0020 5000 0958 5145 8000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/7  Konyaaltı /ANTALYA
Tel: 0 506 136 98 23
        """)

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")


def kapak_kent():
    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.9)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(1.5)

    yibf = document.add_paragraph("Sayı:2024/\nKonu: Hakediş Raporu")
    yibf.alignment = WD_ALIGN_PARAGRAPH.LEFT


    baslık = document.add_heading("ANTALYA BÜYÜKŞEHİR BELEDİYESİ\nKENT ESTETİĞİ DAİRE BAŞKANLIĞINA\n"+"\t"*8+"ANTALYA\n\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)



    note = document.add_paragraph(f"\tYapı Denetim Hizmeti tarafımızca yürütülmekte olan {sonuc_ana[0][13]} imarın {sonuc_ana[0][7]} ada {sonuc_ana[0][8]} parsel {sonuc_ana[0][11]} YİBF numaralı, inşaatın {sonuc_ana[0][4]} no’lu yüzde {sonuc_ana[0][24]} seviye Yapı Denetim Hizmet Bedeli hakediş raporu, yazımız ekinde sunulmuştur.\n"+"\tGerekenin yapılmasını saygılarımla arz ederim.\n")

    ydk1 = document.add_paragraph(f"\t"*9+f"      {sonuc_ana[0][0]}")
    ydk2 = document.add_paragraph(f"{sonuc_ana[0][6]}\n\n\n\n\n")
    ydk2.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    if "MK" not in sonuc_ana[0][6]:
        adres_muh = document.add_paragraph(f"""
MUHTEŞEM Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6230334026 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR67 0020 5000 0956 6541 9000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/2 Konyaaltı /ANTALYA
Tel: 0 242 229 35 75
        """)
    else:
        adres_mk = document.add_paragraph(f"""
MK Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6220189578 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR05 0020 5000 0958 5145 8000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/7  Konyaaltı /ANTALYA
        Tel: 0 506 136 98 23
        """)

    ekler = document.add_paragraph("    EK: ")
    ekler.add_run("""
        1) 3 Adet Hakediş raporu
        2) 3 Adet Personel Listesi
        3) 3 Adet Makbuz
        4) 3 Adet fatura""")

    # Change the font color of the text in the 'ekler' paragraph to black
    for run in ekler.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)  # Set to black
        if "EK" in run.text:
            run.font.bold = True

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")

def kapak_diger():

    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.9)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(1.5)

    yibf = document.add_paragraph("Sayı:2024/\nKonu: Hakediş Raporu")
    yibf.alignment = WD_ALIGN_PARAGRAPH.LEFT


    baslık = document.add_heading(f"{sonuc_ana[0][10]} BELEDİYESİ\nİMAR VE ŞEHİRCİLİK MÜDÜRLÜĞÜ\n"+"\t"*8+"ANTALYA\n\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)



    note = document.add_paragraph(f"\tYapı Denetim Hizmeti tarafımızca yürütülmekte olan {sonuc_ana[0][7]} ada {sonuc_ana[0][8]} parselde yapılan inşaatın {sonuc_ana[0][4]} no’lu Yapı Denetim Hizmet Bedeli hakediş raporu, yazımız ekinde sunulmuştur.\n\tGerekenin yapılmasını saygılarımla arz ederim.")

    if "MK" not in sonuc_ana[0][6]:
        ydk1 = document.add_paragraph(f"\t" * 9 + f"      {sonuc_ana[0][0]}")
    else:
        ydk1 = document.add_paragraph(f"\t" * 10 + f"     {sonuc_ana[0][0]}")


    ydk2 = document.add_paragraph(f"{sonuc_ana[0][6]}\n\n\n\n\n")
    ydk2.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    if "MK" not in sonuc_ana[0][6]:
        adres_muh = document.add_paragraph(f"""
MUHTEŞEM Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6230334026 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR67 0020 5000 0956 6541 9000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/2 Konyaaltı /ANTALYA
Tel: 0 242 229 35 75
        """)

    else:
        adres_mk = document.add_paragraph(f"""
MK Yapı Denetim Ltd. Şti. (Kurumlar Vergi Dairesi -6220189578 VKN)
Kuveyt Türk Bankası Çallı şubesi
IBAN: TR05 0020 5000 0958 5145 8000 01
Adres:Pınarbaşı Mah. 708. Sok no:8/7  Konyaaltı /ANTALYA
Tel: 0 506 136 98 23
        """)

    ekler = document.add_paragraph("    EK: ")
    ekler.add_run("""
        1) 3 Adet Hakediş raporu
        2) 3 Adet Personel Listesi
        3) 3 Adet Makbuz
        4) 3 Adet fatura""")

    # Change the font color of the text in the 'ekler' paragraph to black
    for run in ekler.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)  # Set to black
        if "EK" in run.text:
            run.font.bold = True

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/kapak.docx")
