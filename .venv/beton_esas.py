# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog
import sql_modul
import os
import json
from beton_docx import call_beton
from tkinter import messagebox

class MuhtesemYapiDenetimApp:
    def __init__(self):

        self.root = tk.Tk()
        self.root.title("MUHTESEM YAPI DENETİM")
        self.set_window_geometry()

        self.frame = ttk.Frame(self.root)
        self.frame.grid(column=0, row=2, padx=20, pady=10, )

        self.frame2 = ttk.Frame(self.root)
        self.frame2.grid(column=0, row=0, padx=20, pady=20)

        self.frame3 = ttk.Frame(self.root)
        self.frame3.grid(column=0, row=3, padx=10, pady=10)

        self.frame4 = ttk.Frame(self.root)
        self.frame4.grid(column=0, row=1, padx=35, pady=10)

        self.labels = ["KOTLAR", "ELEMAN", "ADEDİ", "TARİHİ", "M3", "7 GÜNLÜK", "28 GÜNLÜK"]
        self.eleman_list = ["KOLON PERDE", "KOLON TABLİYE", "TABLİYE", "KOLON", "PERDE"]
        self.num_say = ["8", "12", "16", "20", "24", "28", "32", "36", "40"]
        self.main_labels = ["YİBF", "YDK", "YDK ADRES", "İDARE", "RUHSAT", "YAPI SINIFI", "ALAN", "MAHALLE", "ADRES", "PAFTA", "ADA", "PARSEL", "YAPI SAHİBİ", "ŞANTİYE ŞEFİ", "MÜTEAHHİT"]

        self.list_entry = []
        self.list_combo = []
        self.list_button = []
        self.denetci_liste = []
        self.kontrol_liste = []

        self.list_all_rows = []
        self.list_all_rows_widgets = []
        self.temp_list_all_row_widgets = []
        self.row_element = [None] * 9

        self.present_data = []
        self.present_data_widgets = []
        self.row_index = 1
        self.i = 0
        self.j = 0

        self.fill_combo()
        self.widgets()
        #self.ekle()

        self.combo_ada.bind("<<ComboboxSelected>>", self.on_ada_selection)
        self.combo_parsel.bind("<<ComboboxSelected>>", self.on_parsel_selection)

    def set_window_geometry(self):

        self.pgen = 880
        self.pyuks = 400
        self.x = (self.root.winfo_screenwidth() - self.pgen) // 2
        self.y = (self.root.winfo_screenheight() - self.pyuks) // 2
        self.root.geometry(f"{self.pgen}x{self.pyuks}+{self.x}+{self.y - 200}")

    def on_ada_selection(self, event):
        selected_item = self.combo_ada.get()
        parsel_sorgu = sql_modul.sql_query("parsel", "beton_kont", "ada", selected_item)


        self.combo_parsel["state"] = "normal"
        self.combo_parsel.set(parsel_sorgu[0][0])
        self.combo_parsel["values"] = sorted(list(set(item[0] for item in parsel_sorgu)))

        # if len(parsel_sorgu) == 1:
        #     self.combo_parsel.current(0)
        #     self.combo_parsel.event_generate("<<ComboboxSelected>>")
        #     self.on_parsel_selection(event)

    def on_parsel_selection(self, event):

        print("CB SELECTED")
        selected_item1 = self.combo_ada.get()
        selected_item2 = self.combo_parsel.get()

        self.row_element = [None] * 9


        self.sonuc_all = sql_modul.sql_query("*", "beton_kont", "ada", selected_item1,
                                             "parsel", selected_item2)

        self.data_as_list = [list(item) for item in self.sonuc_all]
        self.present_data = self.data_as_list

        for ent in self.frm4_wid_entry:
            ent.delete(0, tk.END)

        for entry, value in zip(self.frm4_wid_entry, self.data_as_list[0]):
            entry.insert(0, value)

        self.entry_demir.delete(0, tk.END)

        if self.present_data[0][-1]:
            self.entry_demir.insert(0, self.present_data[0][-1])

        for i in range(self.row_index - 1):

            for widget in self.frame.grid_slaves(row=self.row_index - 1):
                widget.grid_forget()
            self.pyuks -= 30
            self.row_index -= 1

            self.root.geometry(f"{self.pgen}x{self.pyuks}")

        self.parsel_kotlar = sql_modul.sql_query("kot, eleman, adet, tarih, m3, yedi_tar, yedi_day, sekiz_tar, sekiz_day", "dayanim", "ada", self.frm4_wid_entry[10].get(), "parsel", self.frm4_wid_entry[11].get())

        if len(self.parsel_kotlar) == 0:

            self.pgen = 880
            self.root.geometry(f"{self.pgen}x{self.pyuks}")

        for i in range(len(self.parsel_kotlar)):
            self.ekle()

        for vt, current in zip(self.parsel_kotlar, self.list_all_rows_widgets):
            for vt_, current_ in zip(vt, current):
                if isinstance(current_, tk.Entry):
                    current_.delete(0, tk.END)
                    current_.insert(0, vt_)
                elif isinstance(current_, tk.StringVar):
                    current_.set(vt_)

        self.list_all_rows_widgets.clear()
        self.parsel_kotlar.clear()


    def widgets(self):

        self.label_ada = tk.Label(self.frame2, text="ADA")
        self.label_ada.grid(row=0, column=0, padx=5, pady=5)

        self.combo_ada = ttk.Combobox(self.frame2, width=15, values=self.integers)
        self.combo_ada.grid(row=0, column=1, padx=5, pady=5)

        self.label_parsel = tk.Label(self.frame2, text="PARSEL")
        self.label_parsel.grid(row=0, column=2, padx=5, pady=5)

        self.combo_parsel = ttk.Combobox(self.frame2, width=15)
        self.combo_parsel.grid(row=0, column=3, padx=5, pady=5)
        self.combo_parsel["state"] = "disabled"

        for i, label_text in enumerate(self.labels, start=0):
            if label_text == "7 GÜNLÜK":
                label = tk.Label(self.frame, text=label_text)
                label.grid(row=0, column=5, columnspan=2, padx=5, pady=5)
            elif label_text == "28 GÜNLÜK":
                label = tk.Label(self.frame, text=label_text)
                label.grid(row=0, column=7, columnspan=2, padx=5, pady=5)
            else:
                label = tk.Label(self.frame, text=label_text)
                label.grid(row=0, column=i, padx=5, pady=5)

        self.ekle_button = ttk.Button(self.frame3, text="Ekle", command=self.ekle, padding=(10, 3))
        self.ekle_button.grid(column=0, row=0, pady=10)

        self.delete_button = ttk.Button(self.frame3, text="Sil", command=self.delete_last_row, padding=(10, 3))
        self.delete_button.grid(column=1, row=0, sticky=tk.E)
        if self.row_index == 1:
            self.delete_button["state"] = "disabled"

        self.import_button = ttk.Button(self.frame2, text="İçe Aktar", command=self.import_data, padding=(10, 3))
        self.import_button.grid(column=4, row=0, padx=10)

        self.save_button = ttk.Button(self.frame4, text="Kaydet", command=self.save_data, padding=(10, 3))
        self.save_button.grid(column=6, row=0, padx=10)

        #self.update_button = ttk.Button(self.frame4, text="Güncelle", command=self.update_data, padding=(10, 3))
        #self.update_button.grid(column=6, row=1, padx=10)

        self.clear_button = ttk.Button(self.frame4, text="Temizle", command=self.clear_data, padding=(10, 3))
        self.clear_button.grid(column=6, row=1, padx=10)

        self.label_demir = tk.Label(self.frame4, text="DEMİR ÇEKME")
        self.label_demir.grid(row=6, column=0, padx=5, pady=5)
        self.entry_demir = tk.Entry(self.frame4, width=23)
        self.entry_demir.grid(row=6, column=1, padx=5, pady=5)

        self.frm4_wid_entry = []

        for i in range(5):
            label_text = self.main_labels[i]
            tk.Label(self.frame4, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky=tk.E)
            entry = tk.Entry(self.frame4, width=23)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.frm4_wid_entry.append(entry)

        # Second column
        for i in range(5, 10):
            label_text = self.main_labels[i]
            tk.Label(self.frame4, text=label_text).grid(row=i - 5, column=2, padx=5, pady=5, sticky=tk.E)
            entry = tk.Entry(self.frame4, width=23)
            entry.grid(row=i - 5, column=3, padx=5, pady=5)
            self.frm4_wid_entry.append(entry)

        for i in range(10, 15):
            label_text = self.main_labels[i]
            tk.Label(self.frame4, text=label_text).grid(row=i - 10, column=4, padx=5, pady=5, sticky=tk.E)
            entry = tk.Entry(self.frame4, width=23)
            entry.grid(row=i - 10, column=5, padx=5, pady=5)
            self.frm4_wid_entry.append(entry)

    def update_data(self):

        if self.frm4_wid_entry[10].get() != "" and self.frm4_wid_entry[11].get() != "":

            self.update_list = []
            for ent in self.frm4_wid_entry:
                self.update_list.append(ent.get())

            condition_columns = ["ada", "parsel"]
            update_values = self.update_list
            condition_values = [self.present_data[0][10], self.present_data[0][11]]
            sql_modul.sql_update("beton_kont", update_values, condition_columns, condition_values)

            for wid in self.temp_list_all_row_widgets:
                temp_list = []
                for mem in wid:
                    temp_list.append(mem.get())
                self.list_all_rows.append(temp_list)

            for i in self.list_all_rows:
                i.append(self.frm4_wid_entry[10].get())
                i.append(self.frm4_wid_entry[11].get())

            for i in self.list_all_rows:
                if i[0] == "":
                    continue

                sql_modul.sql_into_beton("dayanim", i)

            self.clear_data()
            self.list_all_rows = []

            messagebox.showinfo("Yeni Kayıt", "Güncelleme işlemi tamamlanmıştır.")

        else:
            messagebox.showerror("Eksik Bilgi !", """Lütfen  "Ada/Parsel"  alanlarını doldurunuz !""")

    def clear_data(self):

        for j in self.list_all_rows_widgets:
            for i in j:
                if isinstance(i, tk.Entry):
                    i.delete(0, 'end')
                elif isinstance(i, tk.StringVar):
                    i.set('')

        for ent in self.frm4_wid_entry: ent.delete(0, tk.END)
        self.combo_ada.set("")
        self.combo_parsel.set("")
        self.combo_parsel["state"] = "disabled"
        self.combo_ada.focus_set()
        self.combo_ada["values"] = self.fill_combo()

        for i in range(self.row_index - 1):

            for widget in self.frame.grid_slaves(row=self.row_index - 1):
                widget.grid_forget()
            self.pyuks -= 30
            self.row_index -= 1

            self.root.geometry(f"{self.pgen}x{self.pyuks}")

        self.present_data.clear()
        self.list_all_rows_widgets.clear()
        #self.parsel_kotlar.clear()

    def save_data(self):


        if self.frm4_wid_entry[10].get() != "" and self.frm4_wid_entry[11].get() != "" :

            self.save_list=[]
            for ent in self.frm4_wid_entry:
                self.save_list.append(ent.get())

            self.save_list.append(self.entry_demir.get())

            sql_modul.sql_into_beton("beton_kont", self.save_list)

            for wid in self.temp_list_all_row_widgets:
                temp_list = []
                for mem in wid:
                    temp_list.append(mem.get())
                self.list_all_rows.append(temp_list)

            for i in self.list_all_rows:
                i.append(self.frm4_wid_entry[10].get())
                i.append(self.frm4_wid_entry[11].get())

            for i in self.list_all_rows:
                if i[0] == "":
                    continue

                sql_modul.sql_into_beton("dayanim", i)

            messagebox.showinfo("Yeni Kayıt", "Kayıt işlemi tamamlanmıştır.")

            self.clear_data()
            self.temp_list_all_row_widgets.clear()
            self.list_all_rows = []

        else:
            messagebox.showerror("Eksik Bilgi !", """Lütfen  "Ada/Parsel"  alanlarını doldurunuz !""")
    def import_data(self):
        try:
            self.file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

            if self.file_path != "" and not self.file_path.lower().endswith('.txt'):
                self.messagebox.showerror("Hata !", "Lütfen sadece .txt uzantılı dosya seçiniz !")
                raise ValueError("Invalid file format")

            with open(self.file_path, 'r', encoding='utf-8') as file:
                self.file_content = file.read()

            self.lines = self.file_content.strip().split('\n')

            self.extracted_values = []
            for index, line in enumerate(self.lines):
                try:
                    if index in [1, 2, 3, 5, 11, 19, 22, 25, 26, 29, 30, 31, 33, 37, 38]:
                        self.split_line = line.split('\t', 1)
                        if len(self.split_line) > 1:
                            self.extracted_values.append(self.split_line[1].strip())
                except IndexError as e:
                    print(f"IndexError occurred: {e}")

            for ent in self.frm4_wid_entry:
                ent.delete(0, tk.END)

            for entry, value in zip(self.frm4_wid_entry, self.extracted_values):
                entry.insert(0, value)

            self.present_data.clear()

            temp_list = [[]]
            temp_list[0] = self.extracted_values
            self.present_data = temp_list
            print(self.present_data)

        except UnicodeDecodeError as e:
            print(f"UnicodeDecodeError: {e}")
            print("An error occurred while decoding the file. Please check the file's encoding.")
        except FileNotFoundError:
            print(f"Dosya bulunamadı: {self.file_path}")
        except Exception as e:
            print(f"Lütfen dil kodlamasını 'utf-8' olarak güncelliyiniz:", e)

        start_keyword = 'Denetçi'
        end_keyword = 'Kontrol Elemanı Mimar/Mühendis'
        extracted_data_personel = self.extract_data_between_keywords(self.file_path, start_keyword, end_keyword)

        denetci_list = []
        kontrol = []
        split_index = next(idx for idx, sublist in enumerate(extracted_data_personel) if
                           'Yardımcı Kontrol Elemanı Mimar/Mühendis' in sublist)

        denetci_list = extracted_data_personel[:split_index + 1]
        del denetci_list[0]
        del denetci_list[-1]

        kontrol = extracted_data_personel[split_index + 1:]

        self.denetci_liste = denetci_list
        self.kontrol_liste = kontrol

        serialized_denetci = json.dumps(denetci_list)
        serialized_kontrol = json.dumps(kontrol)


        into_data = [None, None, None, None]
        into_data = self.extracted_values[10:12] + [serialized_denetci, serialized_kontrol]


        sql_modul.sql_into_beton("beton_personel", into_data)


        for i in range(self.row_index - 1):

            for widget in self.frame.grid_slaves(row=self.row_index - 1):
                widget.grid_forget()
            self.pyuks -= 30
            self.row_index -= 1

            self.root.geometry(f"{self.pgen}x{self.pyuks}")

        return denetci_list, kontrol

    def extract_data_between_keywords(self, file_path, start_keyword, end_keyword):
        self.extracted_data = []
        self.is_inside_target_range = False

        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                if start_keyword == line.strip():
                    self.is_inside_target_range = True
                    self.extracted_data.append([line.strip()])
                    continue
                if self.is_inside_target_range:
                    if line.startswith("T.C"):
                        continue
                    columns = [column.strip() for column in line.strip().split('\t') if column.strip()]
                    self.extracted_data.append(columns)
                if end_keyword == line.strip():
                    self.is_inside_target_range = False
                    self.extracted_data.append(line.strip())
                    break
        return self.extracted_data

    def ekle(self):

        self.row_element = [None] * 9

        self.row_widgets_entry = []
        self.row_widgets_combobox = []

        if self.row_index > 1:
            self.delete_button["state"] = "normal"

        self.export = ttk.Button(self.frame, text="Dışa Aktar", command=self.on_export_button_clicked, padding=(10, 3))
        self.export.grid(column=9, row=self.row_index)
        self.export.config(command=lambda btn=self.export: self.on_export_button_clicked(btn))
        self.list_button.append(self.export)

        self.entry = tk.Entry(self.frame, width=15)
        self.entry.grid(row=self.row_index, column=0, sticky=tk.W, padx=5, pady=5)
        self.row_widgets_entry.append(self.entry)

        self.combo2 = ttk.Combobox(self.frame, width=15, values=self.eleman_list)
        self.combo2.grid(row=self.row_index, column=1, sticky=tk.W, padx=5, pady=5)
        self.row_widgets_combobox.append(self.combo2)

        self.combo3 = ttk.Combobox(self.frame, width=10, values=self.num_say)
        self.combo3.grid(row=self.row_index, column=2, sticky=tk.W, padx=5, pady=5)
        self.row_widgets_combobox.append(self.combo3)

        for i in range(3, 9):
            self.entry = tk.Entry(self.frame, width=15)
            self.entry.grid(row=self.row_index, column=i, sticky=tk.W, padx=5, pady=5)
            self.row_widgets_entry.append(self.entry)

        self.row_index += 1

        self.pgen = 1100
        self.pyuks += 30
        self.root.geometry(f"{self.pgen}x{self.pyuks}")

        row_value = []

        for ent, ind in zip(self.row_widgets_entry, [0, 3, 4, 5, 6, 7, 8]):
            self.row_element[ind] = ent

        for comb, ind in zip(self.row_widgets_combobox, [1, 2]):
            self.row_element[ind] = comb


        self.list_all_rows_widgets.append(self.row_element)
        self.temp_list_all_row_widgets.append(self.row_element)
        self.list_entry.append(self.row_widgets_entry)
        self.list_combo.append(self.row_widgets_combobox)

        if self.row_index < 1:
            self.delete_button["state"] = "disabled"

    def on_export_button_clicked(self, button):

        index = self.list_button.index(button)
        entry_widgets = self.list_entry[index]
        combobox_widgets = self.list_combo[index]

        row_element = [None] * 9
        row_value = []

        for ent, ind in zip(entry_widgets, [0, 3, 4, 5, 6, 7, 8]):
            row_element[ind] = ent

        for comb, ind in zip(combobox_widgets, [1, 2]):
            row_element[ind] = comb

        row_values = [val.get() for val in row_element]

        call_beton(row_values, self.present_data, self.entry_demir.get())


    def fill_combo(self):

        try:
            self.ada_list = sql_modul.sql_query_all("beton_kont", "ada")
            self.unique_values = set(t[0] for t in self.ada_list)
            self.integers = [int(value) for value in self.unique_values if value.isdigit()]
            self.integers.sort()
            self.strings = [value for value in self.unique_values if not value.isdigit()]
            self.integers.extend(self.strings)
        except Exception as e:
            print("Line 75  :", e)

        return self.integers

    def delete_last_row(self):

        try:
            if self.row_index < 2:
                self.delete_button["state"] = "disabled"

            if self.row_index > 1:
                for widget in self.frame.grid_slaves(row=self.row_index - 1):
                    widget.grid_forget()
                self.pyuks -= 30
                self.row_index -= 1

                self.root.geometry(f"{self.pgen}x{self.pyuks}")

                self.list_entry.pop()
                self.list_combo.pop()
                self.list_button.pop()
                self.temp_list_all_row_widgets.pop()
                self.parsel_kotlar = sql_modul.sql_query("kot, eleman, adet, tarih, m3, yedi_tar, yedi_day, sekiz_tar, sekiz_day", "dayanim", "ada", self.frm4_wid_entry[10].get(), "parsel", self.frm4_wid_entry[11].get())

                parsel_list = list(self.parsel_kotlar[-1])

                parsel_list.append(self.frm4_wid_entry[10].get())
                parsel_list.append(self.frm4_wid_entry[11].get())

                sql_modul.sql_delete2("dayanim", "kot", parsel_list[0], "eleman", parsel_list[1], "ada", parsel_list[9], "parsel", parsel_list[10])



        except Exception as e:
            print(e)
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MuhtesemYapiDenetimApp()
    app.run()
