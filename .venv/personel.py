# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
import sql_modul
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Inches, RGBColor, Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import os
import json
from tkinter import messagebox

sonuc_ana = []
save_path = ""

def call():
    root = tk.Tk()
    root.title("MUHTESEM YAPI DENETİM")

    pgen = 1220
    pyuks = 250
    x = (root.winfo_screenwidth() - pgen) // 2
    y = (root.winfo_screenheight() - pyuks) // 2
    root.geometry(f"{pgen}x{pyuks}+{x}+{y - 200}")

    label_text = f"{sonuc_ana[0][3]} - {sonuc_ana[0][2]} TARİH(LER)\'İNDE GERÇEKLEŞEN {sonuc_ana[0][4]} NOLU HAKEDİŞ RAPORUNA AİT PERSONEL BİLDİRGESİ"
    label = tk.Label(root, text=label_text)
    label.place(x=300, y=30)

    frame = ttk.Frame(root)
    frame.grid(column=0, row=0, padx=20, pady=80)

    labels = ["SIRA NO", "ADI SOYADI", "DENETÇİ VASFI ve MESLEĞİ",
              "DENETÇİ NO-ODA SİCİL", "İŞE BAŞLAMA TARİHİ", "AYRILIŞ TARİHİ", "İMZA"]

    for i, label_text in enumerate(labels, start=0):
        label = tk.Label(frame, text=label_text)
        label.grid(row=0, column=i, padx=5, pady=5)

    row_index = 1
    sorgu = sql_modul.sql_query_all("ydk_liste")
    sorgu_isim = []

    for i in sorgu:
        sorgu_isim.append(i[0])

    list_combo = []
    list_label = []

    flag_per = False

    def autocomplete(event, combobox):
        try:
            sonuc2 = sql_modul.sql_query_all("ydk_liste", "adi")
            strings = [t[0].lower() for t in sonuc2]
        except Exception as e:
            print("Line 56  :", e)
        if combobox.get() == "":
            combobox['values'] = strings
        else:
            current_text = combobox.get().lower()
            filtered_options = [option.upper() for option in strings if option.startswith(current_text)]
            combobox['values'] = filtered_options

    def on_combobox_select(event, row, labels_for_row):
        try:
            selected_item = event.widget.get()
            # print(f"Selected item in row {row}: {selected_item}")
            sonuc3 = sql_modul.sql_query("denetci_vasfı, denetci_no, ise_baslama", "ydk_liste", "adi", selected_item)

            i = 0
            for label in labels_for_row:
                label.config(text=f"{sonuc3[0][i]}")
                i += 1
        except Exception as e:
            messagebox.showerror("Hata !", f"{selected_item} kişisine ait bilgiler eksik !")


    def open_combobox_list(event, combo):
        combo.event_generate("<Button-1>")


    def ekle():

        nonlocal row_index, pyuks

        line_name_label = tk.Label(frame, text=f"{row_index}")
        line_name_label.grid(row=row_index, column=0, sticky=tk.EW, padx=5, pady=5)

        combo = ttk.Combobox(frame, values=sorgu_isim, width=30)
        combo.grid(row=row_index, column=1, sticky=tk.W, padx=5, pady=5)
        list_combo.append(combo)

        labels_for_row = []
        for j, label_text in enumerate(["Label 1", "Label 2", "Label 3"], start=2):
            label = tk.Label(frame, text=label_text)
            label.grid(row=row_index, column=j, sticky=tk.W, padx=5, pady=5)
            labels_for_row.append(label)


        combo.bind("<<ComboboxSelected>>", lambda event, row=row_index: on_combobox_select(event, row, labels_for_row))
        combo.bind("<KeyRelease>", lambda event: autocomplete(event, combo))
        combo.bind("<Return>", lambda event: open_combobox_list(event, combo))

        row_index += 1

        pyuks += 30
        root.geometry(f"{pgen}x{pyuks}")

    def delete_last_row():
        nonlocal row_index, pyuks
        if row_index > 1:
            # Destroy widgets of last row
            for widget in frame.grid_slaves(row=row_index - 1):
                widget.grid_forget()
            row_index -= 1
            pyuks -= 30
            root.geometry(f"{pgen}x{pyuks}")
            list_combo.pop()

    delete_button = ttk.Button(root, text="Sil", command=delete_last_row, padding=(10, 3))
    delete_button.grid(column=2, row=1)

    hak_per = sql_modul.sql_query("personel", "hakedis_personel", "ada", sonuc_ana[0][7], "parsel", sonuc_ana[0][8], "daire", sonuc_ana[0][10], "hakedis", sonuc_ana[0][4])
    pers = []

    if len(hak_per) > 0:
        pers = json.loads(hak_per[0][0])

    row_num = 5
    if len(pers) != 0:
        row_num = len(pers)

    for _ in range(row_num):
        ekle()

    for value, combo_box in zip(pers, list_combo):
        try:
            combo_box.set(value)
            combo_box.event_generate("<<ComboboxSelected>>")
        except Exception as e:
            print("Line 143", e)

    ekle_button = ttk.Button(root, text="Ekle", command=ekle, padding=(10, 3))
    ekle_button.grid(column=1, row=1)

    def export_to_docx():
        list_hak = [sonuc_ana[0][7], sonuc_ana[0][8], sonuc_ana[0][10], sonuc_ana[0][4]]
        list_ser = []

        for i in list_combo:
            if i.get() != "":
                list_ser.append(i.get())

        serialized_list = json.dumps(list_ser)
        list_hak.insert(4, serialized_list)
        sql_modul.sql_insert_or_update("hakedis_personel", list_hak)

        document = Document()
        section = document.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE

        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        baslık = document.add_heading(
            f"{sonuc_ana[0][3]} - {sonuc_ana[0][2]} TARİH(LER)\'İNDE GERÇEKLEŞEN {sonuc_ana[0][4]} NOLU HAKEDİŞ RAPORUNA AİT PERSONEL BİLDİRGESİ",
            level=1)
        baslık.alignment = WD_ALIGN_PARAGRAPH.CENTER
        baslık.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        yibf = document.add_paragraph(f"\n\n YİBF No : {sonuc_ana[0][11]}")
        yibf.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        table = document.add_table(rows=1, cols=len(labels), style='Table Grid')
        hdr_cells = table.rows[0].cells

        for i, label_text in enumerate(labels):
            hdr_cells[i].text = label_text

        # j = 0
        # for row in table.rows:
        # for i, cell in enumerate(row.cells):
        # cell.paragraphs[0].alignment = 0

        for i in range(1, row_index):
            row_data = []
            for widget in frame.grid_slaves(row=i):
                if isinstance(widget, tk.Label):
                    row_data.insert(0, widget.cget("text"))
                elif isinstance(widget, ttk.Combobox):
                    row_data.insert(0, widget.get())
            row_cells = table.add_row().cells
            for j, cell_data in enumerate(row_data):
                row_cells[j].text = cell_data

        for ri in table.rows:
            for c in ri.cells:
                c.paragraphs[0].alignment = 1

        for i, label_text in enumerate(labels):
            hdr_cells[i].text = label_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        paragraph.paragraph_format.line_spacing = Pt(10)
                        paragraph.paragraph_format.space_before = Pt(8)
                        paragraph.paragraph_format.space_after = Pt(8)
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        column_widths = [0.5, 4.0, 6.0, 2.0, 1.0, 1.0, 3.0]

        for row in table.rows:
            for i, cell in enumerate(row.cells):
                cell.width = Inches(column_widths[i])

        note = document.add_paragraph(
            "\n\nNOT :" + "\t" * 10 + "          Yukarıda bilgierin kayıtlarımıza uygun olduğunu onaylarım."
                                      "\n-Düzenleme tarihinde işten ayrılanlardan imza şartı aranmayacak," + "\t" * 5 + f"           {sonuc_ana[0][1]}"
                                                                                                                        "\nancak bu kişiler de belirtilecektir.\n-Denetçiler için denetçi no, kontrol elemanları için oda sicil no yazılacaktır." + "\t" * 3 + "      Yapı Denetim Kuruluşu Yetkilisi")
        note.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            if sonuc_ana[0][5] == "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900":

                document_path = f"{save_path}/MUHTEŞEM/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/personel.docx"
            else:
                document_path = f"{save_path}/MK/{sonuc_ana[0][10]}/{sonuc_ana[0][7]}-{sonuc_ana[0][8]}/{sonuc_ana[0][4]} NOLU/personel.docx"

            document.save(document_path)
            os.startfile(document_path)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    export_button = ttk.Button(root, text="Dışa Aktar", command=export_to_docx, padding=(10, 3))
    export_button.grid(column=3, row=1)

    root.mainloop()
