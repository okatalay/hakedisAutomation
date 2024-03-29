import os
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

sonuc_ana = []
save_path = ""

def call():
    document = Document()
    section = document.sections[0]

    section.orientation = WD_ORIENT.PORTRAIT  # Change orientation to portrait
    section.page_width = Inches(8.27)  # Swap width and height for portrait
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

    baslık = document.add_heading("SEVİYE TESPİT TUTANAĞI\n")
    baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER

    baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    yibf = document.add_paragraph(f"YİBF No : {sonuc_ana[0][11]}")
    yibf.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    table1 = document.add_table(rows=7, cols=2)

    for row in table1.rows:
        row.height = Inches(0)
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.space_after = Inches(0)

    cell_contents1 = [
        ("İlgili İdare Belediye/Valilik", f": {sonuc_ana[0][10]}"),
        ("Yapı Ruhsat Tarihi ve No", f": {sonuc_ana[0][12]}"),
        ("Yapının Adresi", f": {sonuc_ana[0][13]}"),
        ("Pafta/Ada/Parsel No", f": {sonuc_ana[0][9]}/{sonuc_ana[0][7]}/{sonuc_ana[0][8]}"),
        ("Yapı İnşaat Alanı(m²) ve Cinsi", f": {sonuc_ana[0][14]}/{sonuc_ana[0][15]}"),
        ("Yapı Sahibi", f": {sonuc_ana[0][16]}"),
        ("Yapı Denetim Kuruluşunun Unvanı", f": {sonuc_ana[0][5]}"),
    ]

    column_widths = [3, 4]
    for row_index in range(7):
        cell = table1.cell(row_index, 0)
        cell.width = Inches(column_widths[0])

        cell = table1.cell(row_index, 1)
        cell.width = Inches(column_widths[1])

    for i in range(7):
        for j in range(2):
            cell = table1.cell(i, j)
            cell.text = cell_contents1[i][j]

    space = document.add_paragraph("\n" * 2)

    table2 = document.add_table(rows=8, cols=3)
    cell_contents2 = [
        ("\nİşin Tanımı (Yapı Bölümü)", "(%)\nTaksit Oranı", "(%)\nGerçekleşme Oranı"),
        ("(a) Ruhsat alımı aşamasında ödenecek olan proje inceleme bedeli", "10", sonuc_ana[0][18]),
        ("(b) Kazı ve temel üst kotuna kadar olan kısım", "10", sonuc_ana[0][19]),
        ("c) Taşıyıcı sistem bölümü", "40", sonuc_ana[0][20]),
        ("(d) Çatı, dolgu duvarları, kapı ve pencere kasaları, tesisat alt yapısı dâhil yapının sıvaya kadar hazır duruma getirilmiş bölümü", "20", sonuc_ana[0][21]),
        ("(e) Mekanik ve elektrik tesisatı ile kalan yapı bölümü", "15", sonuc_ana[0][22]),
        ("(f) İş bitirme tutanağının ilgili idare tarafından onaylanması", "5", sonuc_ana[0][23]),
        ("\t"*7+"         Toplam :", "100", sonuc_ana[0][24]),  # Add empty row for the last row of the second table
    ]

    column_widths = [5, 1, 1]
    for col_index, width in enumerate(column_widths):
        for row_index in range(8):
            cell = table2.cell(row_index, col_index)
            cell.width = Inches(width)

    for row_index in range(8):
        for col_index in range(3):
            cell = table2.cell(row_index, col_index)
            cell.text = cell_contents2[row_index][col_index]
            run = cell.paragraphs[0].runs[0]

            if run.text in cell_contents2[0]:
                run.bold = True
                run.underline = True
                cell.bold = True
            if col_index == 0:
                cell.paragraphs[0].alignment = 0
            else:
                cell.paragraphs[0].alignment = 1

    note = document.add_paragraph(f"\n\t{sonuc_ana[0][2]} tarihi itibariyle yukarıda özellikleri belirtilen yapı yüzde {sonuc_ana[0][25]} ({sonuc_ana[0][24]}) oranında gerçekleşmiştir. İş bu tutanak üç nüsha olarak düzenlenmiştir.\n")

    note2 = document.add_paragraph("\t"*5+"  DÜZENLEYENLER\n")

    table3 = document.add_table(rows=2, cols=3)

    table_3_data = [["Yapı Denetim Kuruluşu Yetkilisi", "Yapı Şantiye Şefi", "Yapı Sahibi"],
                    ["", sonuc_ana[0][17], sonuc_ana[0][16]]]

    table3_column_widths = [10, 1, 1]

    for col_index, width in enumerate(table3_column_widths):
        table3.columns[col_index].width = Inches(width)

    j = 0
    for row in table3.rows:
        for i, cell in enumerate(row.cells):
            cell.text = table_3_data[j][i]
            cell.paragraphs[0].alignment = 1
            run = cell.paragraphs[0].runs[0]

            if run.text in table_3_data[0]:
                run.bold = True
                run.underline = True
                cell.bold = True
        j += 1

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 20 pt
                    paragraph.paragraph_format.space_before = Pt(2)
                    paragraph.paragraph_format.space_after = Pt(1)

    if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

        document.save(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/seviye.docx")
        os.startfile(
            f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/seviye.docx")
    else:
        document.save(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/seviye.docx")
        os.startfile(
            f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/seviye.docx")
