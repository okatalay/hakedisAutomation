# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk, filedialog
from beton_esas import MuhtesemYapiDenetimApp


import belediye_liste
import kapak
import personel
import seviye
import sql_modul
import tahakkuk_vars
import ydk_liste
import lab_liste
import os
import json

def folder_update():
    home_dir = os.path.expanduser("~")
    desktop_path = os.path.join(home_dir, "Desktop")

    folder_path = os.path.join(desktop_path, "HAKEDİŞ")
    os.makedirs(folder_path, exist_ok=True)

    save_path = f"{desktop_path}/HAKEDİŞ"

    save_folder = "MUHTEŞEM"
    save_folder2 = "MK"

    for i in [personel, seviye, kapak, tahakkuk_vars]:
        i.save_path = save_path

    try:
        result_muh = sql_modul.sql_query("ada, parsel, ilgili_daire, hakedis_no", "anasayfa", "ydk_unvan", "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900")
        result_mk = sql_modul.sql_query("ada, parsel, ilgili_daire, hakedis_no", "anasayfa", "ydk_unvan", "MK Yapı Denetim Ltd. Şti./ Dosya No: 101")

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
            folder_structure = [f"{i[2]}", f"{i[0]}-{i[1]}", f"{i[3]} NOLU"]
            create_folders(folder_structure, parent_path=parent_folder_path)

        desktop_path = os.path.join(os.path.expanduser("~"), save_path)
        parent_folder_name = save_folder2
        parent_folder_path = os.path.join(desktop_path, parent_folder_name)
        os.makedirs(parent_folder_path, exist_ok=True)

        for i in result_mk:
            folder_structure = [f"{i[2]}", f"{i[0]}-{i[1]}", f"{i[3]} NOLU"]
            create_folders(folder_structure, parent_path=parent_folder_path)

    except Exception as e:
        print(f"Make Files Error: {e}")

folder_update()

root = tk.Tk()
root.title("MK YAPI DENETİM")

root.iconbitmap("company.ico")

ekrangen = root.winfo_screenwidth()
ekranyuks = root.winfo_screenheight()

pgen = 1000
pyuks = 750

x = (ekrangen - pgen) // 2
y = (ekranyuks - pyuks) // 2

root.geometry(f"{pgen}x{pyuks}+{x}+{y - 50}")

root.protocol("WM_DELETE_WINDOW", root.destroy)

frm = tk.Frame(root)
frm.place(relx=0.05, rely=0.1)

labels = [
    "Kapak Tarihi", "YİBF Tarihi", "Hakediş Gerçekleşme Tarihi", "Bir Önceki Hakediş Tarihi",
    "Hakediş No", "Yapı Denetim Kuruluşunun Unvanı", "Yapı Denetim Kuruluşu", "Ada", "Parsel", "Pafta",
    "İlgili İdare", "Y.İ.B.F. No", "Yapı Ruhsat Tarihi ve No", "Yapının Adresi", "Yapı İnşaat Alanı (m²)",
    "Cinsi(SINIFI)", "Yapı Sahibi", "Şantiye Şefi", "10%", "10%", "40%", "20%", "15%", "5%", "100%",
    "Toplam(yazıyla)", "Denetim Hizmet Bedeline esas tutar (KDV hariç)", "Gerçekleşen Fiziki Seviye",
    "% 3 İlgili İdare Payı (KDV hariç)", "% 3 Bakanlık Payı (KDV hariç)", "Yapı Denetim Kuruluşu Payı (KDV hariç)",
    "KDV", "Tevkifat", "Denetim Hizmet Bedeline esas tutar (KDV dahil)", "LAB isim", "LAB Tutar",
    "Denetim Hizmet Bedeline esas tutar (-Lab Tutar)",
]

global_sonuc = []
entries = []
entry_values = []

i = 1
j = 1
for label_text in labels:
    if i < 19:
        label = tk.Label(frm, text=label_text)
        label.grid(row=i, column=0, sticky=tk.E, padx=5, pady=5)
        entry = tk.Entry(frm, width=30)
        entry.grid(row=i, column=1, sticky=tk.W, padx=10, pady=5)
        entries.append(entry)
        i += 1

    else:
        label = tk.Label(frm, text=label_text)
        label.grid(row=j, column=2, sticky=tk.E, padx=5, pady=5)
        entry = tk.Entry(frm, width=30)
        entry.grid(row=j, column=3, sticky=tk.W, padx=10, pady=5)
        entries.append(entry)
        j += 1

entries[34].config(width=22)

try:
    ada_list = sql_modul.sql_query_all("anasayfa", "ada")
    unique_values = set(t[0] for t in ada_list)
    integers = [int(value) for value in unique_values if value.isdigit()]
    integers.sort()
    strings = [value for value in unique_values if not value.isdigit()]
    integers.extend(strings)

except Exception as e:
    print("Line 75  :", e)

sorgu_cb = ttk.Combobox(values=integers)
sorgu_cb.place(x=90, y=20)

sorgu_cb2 = ttk.Combobox(state="disabled")
sorgu_cb2.place(x=320, y=20)

sorgu_cb3 = ttk.Combobox(state="disabled")
sorgu_cb3.place(x=550, y=20)

sorgu_cb4 = ttk.Combobox(state="disabled")
sorgu_cb4.place(x=790, y=20)

sorgu_cb.focus_set()

ext_personel = []
def save_entry_values():

    entry_values.clear()
    if entries[4].get() != "" and entries[7].get() != "" and entries[8].get() != "" and entries[10].get() != "":
        for entry_l in entries:
            entry_values.append(entry_l.get())

        sql_modul.sql_into("anasayfa", entry_values)
        messagebox.showinfo("Yeni Kayıt", "Kayıt işlemi tamamlanmıştır.")
        for ent in entries: ent.delete(0, tk.END)
        lab_combo.set("")
        sorgu_cb.focus_set()

        for combobox in [sorgu_cb2, sorgu_cb3, sorgu_cb4]:
            combobox["state"] = "disabled"
            combobox.set("")
        sorgu_cb.set("")

        try:
            ada_list = sql_modul.sql_query_all("anasayfa", "ada")
            integers = [int(t[0]) for t in ada_list if t[0].isdigit()]
            integers.sort()
            strings = [t[0] for t in ada_list if not t[0].isdigit()]

            integers_set = set(integers)
            integers = sorted(integers_set, key=integers.index)

            integers.extend(strings)
            ada_list = integers
            sorgu_cb['values'] = ada_list
            global_sonuc = []

            folder_update()

        except Exception as e:
            print("Line 121  :", e)
    else:
        messagebox.showerror("Eksik Bilgi !", """Lütfen  "Ada/Parsel/Daire/Hakediş"  alanlarını doldurunuz !""")

    try:
        list_hak = [entry_values[7], entry_values[8], entry_values[10], entry_values[4]]
        del ext_personel[0]
        serialized_list = json.dumps(ext_personel)
        list_hak.insert(4, serialized_list)
        sql_modul.sql_insert_or_update("hakedis_personel", list_hak)
    except Exception as e:
        print("Lines 193", e)

def update_entry_values():
    global global_sonuc

    sorgu_cb.focus_set()

    if entries[4].get() != "" and entries[7].get() != "" and entries[8].get() != "" and entries[10].get() != "":
        entry_values.clear()

        for entry_loop in entries:
            entry_values.append(entry_loop.get())
        condition_columns = ["ada", "parsel", "ilgili_daire", "hakedis_no"]
        update_values = entry_values
        condition_values = [global_sonuc[0][7], global_sonuc[0][8], global_sonuc[0][10],
                            global_sonuc[0][4]]

        print(update_values, condition_columns, condition_values, sep="\n")

        sql_modul.sql_update("anasayfa", update_values, condition_columns, condition_values)
        messagebox.showinfo("Güncelle !", "Verileriniz güncellenmiştir !")

        for ent in entries: ent.delete(0, tk.END)
        lab_combo.set("")
        sorgu_cb.focus_set()

        for combobox in [sorgu_cb2, sorgu_cb3, sorgu_cb4]:
            combobox["state"] = "disabled"
            combobox.set("")
        sorgu_cb.set("")

        try:
            ada_list = sql_modul.sql_query_all("anasayfa", "ada")
            integers = [int(t[0]) for t in ada_list if t[0].isdigit()]
            integers.sort()
            strings = [t[0] for t in ada_list if not t[0].isdigit()]
            integers_set = set(integers)
            integers = sorted(integers_set, key=integers.index)

            integers.extend(strings)
            ada_list = integers

            sorgu_cb['values'] = ada_list
            global_sonuc = []

            folder_update()

        except Exception as e:
            print("Line 121  :", e)
    else:
        messagebox.showerror("Eksik Bilgi !", """Lütfen  "Ada/Parsel/Daire/Hakediş"  alanlarını doldurunuz !""")
def import_file():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

        if file_path != "" and not file_path.lower().endswith('.txt'):
            messagebox.showerror("Hata !", "Lütfen sadece .txt uzantılı dosya seçiniz !")
            raise ValueError("Invalid file format")

        with open(file_path, 'r', encoding='utf-8') as file:
            file_content = file.read()

        lines = file_content.strip().split('\n')

        extracted_values = []
        for index, line in enumerate(lines):
            try:
                if index in [1, 2, 5, 11, 19, 22, 26, 29, 30, 31, 33, 37]:
                    split_line = line.split('\t', 1)
                    if len(split_line) > 1:
                        extracted_values.append(split_line[1].strip())
            except IndexError as e:
                print(f"IndexError occurred: {e}")
                # Handle the error here as needed

        for ent in entries:
            try:
                ent.delete(0, tk.END)
            except Exception as e:
                print(f"Error occurred while deleting entries: {e}")
                # Handle the error here as needed

        index_numbers = [11, 6, 10, 12, 15, 14, 13, 9, 7, 8, 16, 17]

        extracted_members = []
        for index in index_numbers:
            try:
                extracted_members.append(entries[index])
            except IndexError as e:
                print(f"IndexError occurred: {e}")
                # Handle the error here as needed

        for ent, mem in zip(extracted_members, extracted_values):
            try:
                ent.insert(0, mem)
            except Exception as e:
                print(f"Error occurred while inserting into entry: {e}")

        lab_combo.set("Lab Seçiniz")
        lab_combo.event_generate("<<ComboboxSelected>>")

        if "MK" in extracted_values[1]:
            entries[5].insert(0, "MK Yapı Denetim Ltd. Şti./ Dosya No: 101")
        else:
            entries[5].insert(0, "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900")
    except FileNotFoundError as e:
        messagebox.showerror("Dosya !", 'Seçilen dosya bulunamadı !')
        # Handle the file not found error here
    except Exception as e:
        messagebox.showerror("Dil kodu hatası !", 'Lütfen dil kodlamasını "UTF-8" olarak seçiniz !')
        # Handle other exceptions here

    def extract_data_between_keywords(file_path, start_keyword, end_keyword):
        extracted_data = []
        is_inside_target_range = False
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    if start_keyword == line.strip():
                        is_inside_target_range = True
                        extracted_data.append(line.strip())
                        continue
                    if is_inside_target_range:
                        # Skip lines starting with "Denetçi", "['T.C.'", and "['Yardımcı"
                        if line.startswith("T.C") or line.startswith("Yardımcı"):
                            continue
                        columns = [column.strip() for column in line.strip().split('\t') if column.strip()]
                        extracted_data.append(columns[1])
                    if end_keyword == line.strip():
                        is_inside_target_range = False
                        is_inside_target_range = False
                        extracted_data.append(line.strip())
                        break


        except UnicodeDecodeError as e:
            print(f"UnicodeDecodeError: {e}")
            print("An error occurred while decoding the file. Please check the file's encoding.")
        except FileNotFoundError:
            print(f"Dosya bulunamadı: {file_path}")
        except Exception as e:
            print(f"Lütfen dil kodlamasını 'utf-8' olarak güncelliyiniz:", e)
        return extracted_data

    start_keyword = 'Denetçi'
    end_keyword = 'Kontrol Elemanı Mimar/Mühendis'
    extracted_data_personel = extract_data_between_keywords(file_path, start_keyword, end_keyword)


    global ext_personel

    ext_personel = []
    ext_personel = extracted_data_personel


def clear_values():

    for ent in entries: ent.delete(0, tk.END)
    lab_combo.set("")
    sorgu_cb.focus_set()

    for combobox in [sorgu_cb2, sorgu_cb3, sorgu_cb4]:
        combobox["state"] = "disabled"
        combobox.set("")
    sorgu_cb.set("")

lab_sonuc = sql_modul.sql_query_all("lab", "adi")

if lab_sonuc is not None:
    lab_sonuc = [item[0] for item in lab_sonuc]
else:
    print("Line : 189 > Error: lab_sonuc is None, cannot iterate over it.")

kaydet = ttk.Button(root, text="Yeni Kayıt", command=save_entry_values, padding=(15, 5))
kaydet.place(x=810, y=680)

guncelle = ttk.Button(root, text="Güncelle", command=update_entry_values, padding=(15, 5))
guncelle.place(x=700, y=680)

temizle = ttk.Button(root, text="Temizle", command=clear_values, padding=(15, 5))
temizle.place(x=200, y=680)

import_file = ttk.Button(root, text="İçe Aktar", command=import_file, padding=(15, 5))
import_file.place(x=90, y=680)

lab_combo = ttk.Combobox(root, values=lab_sonuc, width=27, height=1)
lab_combo.place(x=730, y=575)

sorgu_l = ttk.Label(text="Ada No :")
sorgu_l.place(x=35, y=20)

sorgu_l2 = ttk.Label(text="Parsel No :")
sorgu_l2.place(x=250, y=20)

sorgu_l3 = ttk.Label(text="İlgili İdare :")
sorgu_l3.place(x=480, y=20)

sorgu_l4 = ttk.Label(text="Hakediş No :")
sorgu_l4.place(x=710, y=20)
def on_select_lab_combo(event, lab_combo_df):
    if lab_combo_df.get():
        entries[34].delete(0, tk.END)
        entries[34].insert(0, lab_combo_df.get())
def on_select_cb4(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4):
    selected_item1 = sorgu_cb.get()
    selected_item2 = sorgu_cb2.get()
    selected_item3 = sorgu_cb3.get()
    selected_item4 = sorgu_cb4.get()

    sonuc = sql_modul.sql_query("*", "anasayfa", "ada", selected_item1,
                                "parsel", selected_item2, "ilgili_daire", selected_item3,
                                "hakedis_no", selected_item4, )
    for entry, sorgu in zip(entries, sonuc[0]):
        entry.delete(0, tk.END)
        entry.insert(0, sorgu)

    lab_combo.set(entries[34].get())

    global global_sonuc
    global_sonuc = sonuc

    personel.sonuc_ana = sonuc
    seviye.sonuc_ana = sonuc
    kapak.sonuc_ana = sonuc
    tahakkuk_vars.sonuc_ana = sonuc

    return sonuc
def on_select_cb3(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4):
    selected_item1 = sorgu_cb.get()
    selected_item2 = sorgu_cb2.get()
    selected_item3 = sorgu_cb3.get()
    sonuc = sql_modul.sql_query("hakedis_no", "anasayfa", "ada", selected_item1,
                                "parsel", selected_item2, "ilgili_daire", selected_item3)

    sorgu_cb4["state"] = "normal"
    sorgu_cb3["state"] = "normal"
    sorgu_cb4["values"] = [item[0] for item in sonuc]
    sorgu_cb4.set(sonuc[0][0])

    sonuc3 = sql_modul.sql_query("*", "anasayfa", "ada", selected_item1, "parsel", selected_item2, "ilgili_daire", selected_item3, "hakedis_no", sonuc[0][0], )

    for entry, sorgu in zip(entries, sonuc3[0]):
        entry.delete(0, tk.END)
        entry.insert(0, sorgu)

    lab_combo.set(entries[34].get())

    global global_sonuc
    global_sonuc = sonuc3

    personel.sonuc_ana = sonuc3
    seviye.sonuc_ana = sonuc3
    kapak.sonuc_ana = sonuc3
    tahakkuk_vars.sonuc_ana = sonuc3

def on_select_cb2(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4):

    selected_item1 = sorgu_cb.get()
    selected_item2 = sorgu_cb2.get()
    selected_item3 = sorgu_cb3.get()
    selected_item = sorgu_cb2.get()
    sonuc = sql_modul.sql_query("ilgili_daire", "anasayfa", "ada", selected_item1,
                                "parsel", selected_item2)

    if len(set(row[0] for row in sonuc)) == 1:
        sorgu_cb4["state"] = "normal"
        sorgu_cb3["state"] = "normal"
        sorgu_cb3["values"] = sonuc[0]
        sorgu_cb3.set(sonuc[0][0])

        selected_item1 = sorgu_cb.get()
        selected_item2 = sorgu_cb2.get()
        selected_item3 = sorgu_cb3.get()
        selected_item4 = sorgu_cb4.get()
        sonuc2 = sql_modul.sql_query("hakedis_no", "anasayfa", "ada", selected_item1, "parsel", selected_item2, "ilgili_daire", selected_item3)

        sorgu_cb4["state"] = "normal"
        sorgu_cb4["values"] = [item[0] for item in sonuc2]
        sorgu_cb4.set(sonuc2[0][0])

        sonuc3 = sql_modul.sql_query("*", "anasayfa", "ada", selected_item1, "parsel", selected_item2, "ilgili_daire", sonuc[0][0], "hakedis_no", sonuc2[0][0], )

        for entry, sorgu in zip(entries, sonuc3[0]):
            entry.delete(0, tk.END)
            entry.insert(0, sorgu)

        lab_combo.set(entries[34].get())

        global global_sonuc
        global_sonuc = sonuc3

        personel.sonuc_ana = sonuc3
        seviye.sonuc_ana = sonuc3
        kapak.sonuc_ana = sonuc3
        tahakkuk_vars.sonuc_ana = sonuc3

    elif len(sonuc) > 0:
        sorgu_cb4["state"] = "disable"
        sorgu_cb3["state"] = "normal"
        sorgu_cb3["values"] = list(set(item[0] for item in sonuc))
        sorgu_cb3.set(sonuc[0][0])

def on_select_cb(event, sorgu_cb2, sorgu_cb3, sorgu_cb4):
    selected_item = sorgu_cb.get()
    idare_sorgu = sql_modul.sql_query("parsel", "anasayfa", "ada", selected_item)

    if len(idare_sorgu) > 0:
        sorgu_cb3["state"] = "disable"
        sorgu_cb4["state"] = "disable"
        sorgu_cb2["state"] = "normal"
        sorgu_cb2.set(idare_sorgu[0][0])
        sorgu_cb2["values"] = list(set(item[0] for item in idare_sorgu))
def belediye_cagir():
    belediye_liste.call()
def ydk_liste_cagir():
    ydk_liste.call()
def personel_cagir():
    try:
        if len(global_sonuc) >= 1:
            personel.call()
        else:
            messagebox.showerror("Hata !", "Lütfen seçim yapınız !")
    except Exception as e:
        print("Line 513", e)

def seviye_cagir():
    if global_sonuc is not None and len(global_sonuc) > 0 :
        seviye.call()
    else:
        messagebox.showerror("Hata !", "Lütfen seçim yapınız !")
def tahakkuk_cagir():

    if global_sonuc is not None and len(global_sonuc) > 0:
        tahakkuk_vars.call()
    else:
        messagebox.showerror("Hata !", "Lütfen seçim yapınız !")
def kapak_cagir():
    try:
        if global_sonuc is not None and len(global_sonuc) > 0:
            if "KEPEZ" in global_sonuc[0][10]:
                kapak.kapak_kepez()
            elif "MANAVGAT" in global_sonuc[0][10]:
                kapak.kapak_manavgat()
            elif "KORK" in global_sonuc[0][10]:
                kapak.kapak_korkuteli()
            elif "ANT" in global_sonuc[0][10]:
                kapak.kapak_kent()
            else:
                kapak.kapak_diger()
        else:
            messagebox.showerror("Hata !", "Lütfen seçim yapınız !")
    except Exception as e:
        print("satır 450", e)
        messagebox.showerror("Hata !", "Lütfen açık kalan bir önceki Word dosyasını kapatın !")

def autocomplete(event, combobox):
    try:
        ada_list = sql_modul.sql_query_all("anasayfa", "ada")
        integers = [int(t[0]) for t in ada_list if t[0].isdigit()]
        integers.sort()
        strings = [t[0] for t in ada_list if not t[0].isdigit()]
        just_int = integers
        integers.extend(strings)
        ada_list = integers

    except Exception as e:
        print("Line 340  :", e)

    if combobox.get() == "":
        combobox['values'] = list(set(ada_list))
    elif combobox.get() == "-":
        combobox['values'] = list(set(strings))
    else:
        current_text = combobox.get()
        filtered_options = [option for option in just_int if str(current_text) in str(option)]
        combobox['values'] = list(set(filtered_options))
def open_combobox_list(event, sorgu_cb):
    sorgu_cb.event_generate("<Button-1>")

sorgu_cb.bind("<KeyRelease>", lambda event: autocomplete(event, sorgu_cb))
sorgu_cb.bind("<Return>", lambda event: open_combobox_list(event, sorgu_cb))

sorgu_cb.bind("<<ComboboxSelected>>", lambda event: on_select_cb(event, sorgu_cb2, sorgu_cb3, sorgu_cb4))
sorgu_cb2.bind("<<ComboboxSelected>>", lambda event: on_select_cb2(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4))
sorgu_cb3.bind("<<ComboboxSelected>>", lambda event: on_select_cb3(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4))
sorgu_cb4.bind("<<ComboboxSelected>>", lambda event: on_select_cb4(event, sorgu_cb, sorgu_cb2, sorgu_cb3, sorgu_cb4))

lab_combo.bind("<<ComboboxSelected>>", lambda event: on_select_lab_combo(event, lab_combo))
def on_key_press(event):

    # entries[5].config(state="normal")
    entry_text = entries[6].get()

    if "MK" in entry_text or "mk" in entry_text or "Mk" in entry_text:

        entries[5].delete(0, tk.END)
        entries[5].insert(0, "MK Yapı Denetim Ltd. Şti./ Dosya No: 101")
        # entries[5].config(state="readonly")
    else:

        entries[5].delete(0, tk.END)
        entries[5].insert(0, "MUHTEŞEM Yapı Denetim Ltd. Şti./ Dosya No: 1900")
        # entries[5].config(state="readonly")

def cald_lab(e33, e35):

    if e33 and e35:

        x = float(e33.replace('.', '').replace(',', '.'))
        y = float(e35.replace('.', '').replace(',', '.'))
        result = x - y
        formatted_result = '{:,.2f}'.format(result).replace(',', 'X').replace('.', ',').replace('X', '.')

        entries[36].delete(0, tk.END)
        entries[36].insert(0, formatted_result)

entries[6].bind("<KeyRelease>", on_key_press)
entries[33].bind("<KeyRelease>", lambda event: cald_lab(entries[33].get(), entries[35].get()))
entries[35].bind("<KeyRelease>", lambda event: cald_lab(entries[33].get(), entries[35].get()))

def beton_eses_start():
    app = MuhtesemYapiDenetimApp()
    app.run()

def lab_cagir():
    lab_liste.call()

menubar = tk.Menu(root)
root.config(menu=menubar)

filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label="Personel", command=personel_cagir)
filemenu.add_command(label="Seviye", command=seviye_cagir)
filemenu.add_command(label="Kapak", command=kapak_cagir)
filemenu.add_command(label="Tahakkuk", command=tahakkuk_cagir)
filemenu.add_separator()
filemenu.add_command(label="Beton Kontrol", command=beton_eses_start)

filemenu2 = tk.Menu(menubar, tearoff=0)
filemenu2.add_command(label="YDK Liste", command=ydk_liste_cagir)
filemenu2.add_command(label="Belediye Liste", command=belediye_cagir)
filemenu2.add_command(label="Lab Liste", command=lab_cagir)

menubar.add_cascade(label="Dosya", menu=filemenu)
menubar.add_cascade(label="Liste", menu=filemenu2)
def close_all_windows():
    for window in root.winfo_children():
        if window != root:
            window.destroy()

    root.destroy()

root.protocol("WM_DELETE_WINDOW", close_all_windows)
root.mainloop()
