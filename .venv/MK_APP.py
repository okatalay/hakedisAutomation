import tkinter as tk
from tkinter import ttk
import sqlite3 as sql
from tkinter import messagebox as mb
import sql_modul
from beton_esas import MuhtesemYapiDenetimApp

root_giris = tk.Tk()
# root_giris.tk_setPalette("#11114b")
pgen = 380
pyuks = 150

ekrangen = root_giris.winfo_screenwidth()
ekranyuks = root_giris.winfo_screenheight()

x = (ekrangen - pgen) // 2
y = (ekranyuks - pyuks) // 2

root_giris.geometry(f"{pgen}x{pyuks}+{x}+{y}")
root_giris.title("KULLANICI GİRİŞİ")

entry_giris = tk.Entry()
entry_giris.place(x=130, y=30, width=165)
entry_giris.focus_set()

entry_sifre = tk.Entry(show="*")
entry_sifre.place(x=130, y=60, width=165)

# Create labels with right-aligned appearance
label_kullanici = tk.Label(text="Kullanıcı Adı", justify="right", anchor="e")
label_kullanici.place(x=50, y=30)

label_sifre = tk.Label(text="Şifre", justify="right", anchor="e")
label_sifre.place(x=50, y=60)

def giris_yap():
    id = entry_giris.get()
    sifre = entry_sifre.get()
    veri = sql_modul.sql_query("*", "kullanici", "id", id, "sifre", sifre)

    if veri:
        if veri[0][2] == "admin":
            root_giris.destroy()
            import ana_ekran
        elif veri[0][2] == "eng":
            root_giris.destroy()
            app = MuhtesemYapiDenetimApp()
            app.run()
    else:
        mb.showerror("Hatalı Giriş !", "Lütfen geçerli bir kullanıcı adı ve şifre girin.", icon="error")

def enter_key(event):
    giris_yap()

b_giris = ttk.Button(text="Giriş Yap", command=giris_yap, padding=(15, 4))
b_giris.place(x=160, y=90, width=100, height=40)

root_giris.bind('<Return>', enter_key)
root_giris.mainloop()
