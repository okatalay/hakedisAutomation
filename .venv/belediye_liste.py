import tkinter as tk
from tkinter import ttk
import sql_modul

def call():
    def update_member():

        selected_item = tree.selection()
        if selected_item:

            item_data = tree.item(selected_item)
            update_window = tk.Toplevel(root)
            update_window.geometry("420x255")
            update_window.title("Belediye Güncelle")

            frame = tk.Frame(update_window)
            frame.pack(padx=10, pady=25)

            labels = ["Belediye:", "Vergi Dairesi:", "Şube:", "IBAN:"]
            entry_vars = []

            for i, label_text in enumerate(labels):
                tk.Label(frame, text=label_text).grid(row=i, column=0, padx=15, pady=5, sticky="w")
                entry = tk.Entry(frame, width=30)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.delete(0, tk.END)
                entry.insert(0, item_data["values"][i])
                entry_vars.append(entry)

            def update_info():

                new_info = [var.get() for var in entry_vars]
                selected_item = tree.selection()
                item_data = tree.item(selected_item)
                sql_modul.sql_update("belediye", new_info, ["adi", "sube"],
                                     [item_data["values"][0], item_data["values"][2]])
                tree.delete(*tree.get_children())

                data = sql_modul.sql_query_all("belediye")
                for item in data:
                    tree.insert("", "end", values=item)

                update_window.destroy()

            def delete_member():

                new_info = [var.get() for var in entry_vars]

                sql_modul.sql_delete("belediye", "adi", new_info[0], "sube", new_info[2])
                tree.delete(*tree.get_children())

                data = sql_modul.sql_query_all("belediye")
                for item in data:
                    tree.insert("", "end", values=item)

                update_window.destroy()

            update_button = ttk.Button(frame, text="Güncelle", command=update_info, padding=(15, 4))
            update_button.grid(row=len(labels), column=0, columnspan=2, pady=5)

            delete_button = ttk.Button(frame, text="Sil", command=delete_member, padding=(15, 4))
            delete_button.grid(row=len(labels) + 1, column=0, columnspan=2, pady=5)

    root = tk.Tk()
    root.title("Belediye Liste")
    root.geometry("1200x280")

    frm = tk.Frame(root)
    frm.grid(row=0, column=0, padx=30, pady=20)

    frm2 = tk.Frame(root)
    frm2.grid(row=0, column=1)

    fields = ["Belediye", "Vergi Dairesi", "Şube", "IBAN"]
    entries_kisi_ekle = []

    for i, field in enumerate(fields):
        label = tk.Label(frm2, text=field + " :")
        label.grid(column=2, row=i, sticky=tk.E, padx=5, pady=5)
        entry = tk.Entry(frm2, width=25)
        entry.grid(column=3, row=i, sticky=tk.W, padx=10, pady=5)
        entries_kisi_ekle.append(entry)

    def add_member():

        data_entries = []
        for entry in entries_kisi_ekle:
            data_entries.append(entry.get())
            entry.delete(0, tk.END)

        entries_kisi_ekle[0].focus_set()

        sql_modul.sql_into("belediye", [data_entries[0], data_entries[1], data_entries[2], data_entries[3]])
        tree.delete(*tree.get_children())
        data = sql_modul.sql_query_all("belediye")

        for item in data:
            tree.insert("", "end", values=item)

    tree = ttk.Treeview(frm, columns=("Belediye", "Vergi Dairesi", "Şube", "IBAN"),
                        selectmode="browse")
    tree["show"] = "headings"
    tree.heading("Belediye", text="Adı-Soyadı")
    tree.heading("Vergi Dairesi", text="Vergi Dairesi")
    tree.heading("Şube", text="Şube")
    tree.heading("IBAN", text="IBAN")

    for col in ["Belediye", "Vergi Dairesi", "Şube", "IBAN"]:
        tree.column(f"{col}", width=200)

    data = sql_modul.sql_query_all("belediye")

    for item in data:
        tree.insert("", "end", values=item)

    tree.bind("<Double-1>", lambda event: update_member())

    tree.grid(row=0, column=0, sticky='nsew')

    scrollbar = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
    scrollbar.grid(row=0, column=1, sticky='ns')
    tree.configure(yscroll=scrollbar.set)

    add_button = ttk.Button(frm2, text="Belediye Ekle", command=add_member,padding=(15, 4))
    add_button.grid(column=3, row=4, sticky=tk.W, padx=10, pady=15)

    root.mainloop()
