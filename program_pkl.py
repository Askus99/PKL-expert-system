import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os
import laporan
from PIL import Image, ImageTk
from docx import Document


LARGE_FONT = ("Verdana", 12)


class LautanFrame(tk.Tk):
    def __init__(gw, *args, **kwargs):
        tk.Tk.__init__(gw, *args, **kwargs)
        lebar = 800
        tinggi = 400

        lebar_layar = gw.winfo_screenwidth()
        tinggi_layar = gw.winfo_screenheight()

        x = int((lebar_layar / 2) - (lebar / 2))
        y = int((tinggi_layar / 2) - (tinggi / 2))

        container = ttk.Frame(gw)
        gw.geometry(f"{lebar}x{tinggi}+{x}+{y}")
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        gw.loginState = tk.IntVar()
        gw.loginState.set(0)
        gw.title("SISTEM PAKAR PERKAYUAN")

        gw.my_menu = tk.Menu(gw)
        gw.config(menu=gw.my_menu)

        gw.new_menu = tk.Menu(gw.my_menu, tearoff=False)
        gw.my_menu.add_cascade(label="New", menu=gw.new_menu)
        gw.my_menu.add_command(label="Logout", command=gw.logout)
        gw.my_menu.add_command(label="Login", command=lambda: gw.show_frame(LoginUser))
        gw.my_menu.add_command(
            label="Laporan",
            command=lambda: gw.show_frame(laporanRekap),
            state="disabled",
        )
        gw.my_menu.add_command(label="About", command=gw.tentang)

        gw.new_menu.add_command(
            label="Ilmu Pakar", command=lambda: gw.show_frame(PageOne), state="disabled"
        )
        gw.new_menu.add_command(
            label="Mencari Data Kayu", command=lambda: gw.show_frame(PageTwo)
        )
        gw.new_menu.add_separator()
        gw.new_menu.add_command(label="Exit", command=gw.destroy)

        gw.frames = {}

        for F in (StartPage, PageOne, PageTwo, LoginUser, laporanRekap):

            frame = F(container, gw)

            gw.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        gw.show_frame(StartPage)

    def show_frame(gw, cont):
        frame = gw.frames[cont]
        frame.tkraise()

    def logout(gw):
        gw.my_menu.entryconfig("Login", state="normal")
        gw.my_menu.entryconfig("Laporan", state="disabled")
        gw.new_menu.entryconfig("Ilmu Pakar", state="disabled")

        gw.loginState.set(0)
        return gw.show_frame(StartPage)

    def tentang(gw):
        pot = tk.Toplevel()
        label1 = tk.Label(pot, text="Alfajri Asnan Kusuma")
        label2 = tk.Label(pot, text="18101152610202")
        label3 = tk.Label(pot, text="alfajriaskus@gmail.com")
        label1.pack()
        label2.pack()
        label3.pack()
        pot.mainloop()


class LoginUser(tk.Frame):
    def __init__(gw, parent, controller):
        tk.Frame.__init__(gw, parent)

        con = sqlite3.connect("kayu.db")
        gw.c = con.cursor()

        gw.us = tk.StringVar()
        gw.pas = tk.StringVar()

        label = tk.Label(gw, text="Silahkan login", font=LARGE_FONT)
        label.place(x=350, y=50)

        labelUs = tk.Label(gw, text="Username")
        gw.entryUs = ttk.Entry(gw, textvariable=gw.us)
        labelUs.place(x=300, y=100)
        gw.entryUs.place(x=400, y=100)

        labelPas = tk.Label(gw, text="Password")
        gw.entryPas = ttk.Entry(gw, textvariable=gw.pas, show="*")
        labelPas.place(x=300, y=150)
        gw.entryPas.place(x=400, y=150)

        btn = ttk.Button(gw, text="Login", command=lambda: gw.login(controller))
        btn.place(x=300, y=200)

    def login(gw, cont):
        global kondisi
        if gw.us.get() == "" or gw.pas.get() == "":
            messagebox.showerror("Error", "Masukkan username dan password")
        else:
            try:
                gw.c.execute(
                    f"""select * from admin
                        where username=? and password=?""",
                    (
                        gw.us.get(),
                        gw.pas.get(),
                    ),
                )

                row = gw.c.fetchone()

                if row == None:
                    messagebox.showerror("Error", "Username dan password salah")

                else:
                    cont.show_frame(StartPage)
                    cont.my_menu.entryconfig("Login", state="disabled")
                    cont.loginState.set(1)

                    if gw.us.get()[0] == "A":
                        messagebox.showinfo("Sukses", "Berhasil login Sebagai Admin")
                        cont.my_menu.entryconfig("Laporan", state="normal")
                        cont.new_menu.entryconfig("Ilmu Pakar", state="normal")
                    else:
                        messagebox.showinfo("Sukses", "Berhasil login Sebagai Karyawan")

            except Exception as e:
                messagebox.showerror("Error", f"Error : {e}", parent=gw)

        gw.entryUs.delete(0, "end")
        gw.entryPas.delete(0, "end")


class StartPage(tk.Frame):
    def __init__(gw, parent, controller):
        tk.Frame.__init__(gw, parent)

        image1 = Image.open("img/final1.jpg")
        image1 = image1.resize((800, 400), Image.ANTIALIAS)
        hutan = ImageTk.PhotoImage(image1)

        label = tk.Label(gw, image=hutan)
        label.image = hutan
        label.pack()


class PageOne(tk.Frame):
    def __init__(gw, parent, controller):
        tk.Frame.__init__(gw, parent)

        # gw.geometry("500x300")

        gw.nama_var = tk.StringVar()
        gw.warna1_var = tk.StringVar()
        gw.warna2_var = tk.StringVar()
        gw.tekstur_var = tk.StringVar()
        gw.arah_var = tk.StringVar()
        gw.kesan_var = tk.StringVar()
        gw.kilap_var = tk.StringVar()
        gw.berat_var = tk.DoubleVar()

        label1 = tk.Label(gw, text="Nama Kayu")
        label1.grid(row=0, column=0, pady=2, padx=10)
        label11 = tk.Label(gw, text=": ")
        label11.grid(row=0, column=1, pady=2)

        label2 = tk.Label(gw, text="Warna Kayu Teras")
        label2.grid(row=1, column=0, pady=2, padx=10)
        label22 = tk.Label(gw, text=": ")
        label22.grid(row=1, column=1, pady=2)

        label3 = tk.Label(gw, text="Warna Kayu Gubal")
        label3.grid(row=2, column=0, pady=2, padx=10)
        label33 = tk.Label(gw, text=": ")
        label33.grid(row=2, column=1, pady=2)

        label4 = tk.Label(gw, text="Tekstur")
        label4.grid(row=3, column=0, pady=2, padx=10)
        label44 = tk.Label(gw, text=": ")
        label44.grid(row=3, column=1, pady=2)

        label5 = tk.Label(gw, text="Arah Serat")
        label5.grid(row=4, column=0, pady=2, padx=10)
        label55 = tk.Label(gw, text=": ")
        label55.grid(row=4, column=1, pady=2)

        label6 = tk.Label(gw, text="Kesan Raba")
        label6.grid(row=5, column=0, pady=2, padx=10)
        label66 = tk.Label(gw, text=": ")
        label66.grid(row=5, column=1, pady=2, padx=10)

        label7 = tk.Label(gw, text="Kilap")
        label7.grid(row=6, column=0, pady=2, padx=10)
        label77 = tk.Label(gw, text=":")
        label77.grid(row=6, column=1, pady=2)

        label8 = tk.Label(gw, text="Berat Jenis")
        label8.grid(row=7, column=0, pady=2, padx=10)
        label88 = tk.Label(gw, text=":")
        label88.grid(row=7, column=1, pady=2)

        gw.entry1 = tk.Entry(gw, textvariable=gw.nama_var)
        gw.entry1.grid(row=0, column=2, pady=2, padx=10)

        gw.entry2 = tk.Entry(gw, textvariable=gw.warna1_var)
        gw.entry2.grid(row=1, column=2, pady=2)

        gw.entry3 = tk.Entry(gw, textvariable=gw.warna2_var)
        gw.entry3.grid(row=2, column=2, pady=2)

        gw.entry4 = tk.Entry(gw, textvariable=gw.tekstur_var)
        gw.entry4.grid(row=3, column=2, pady=2)

        gw.entry5 = tk.Entry(gw, textvariable=gw.arah_var)
        gw.entry5.grid(row=4, column=2, pady=2)

        gw.entry6 = tk.Entry(gw, textvariable=gw.kesan_var)
        gw.entry6.grid(row=5, column=2, pady=2)

        gw.entry7 = tk.Entry(gw, textvariable=gw.kilap_var)
        gw.entry7.grid(row=6, column=2, pady=2)

        gw.entry8 = tk.Entry(gw, textvariable=gw.berat_var)
        gw.entry8.grid(row=7, column=2, pady=2)

        button = ttk.Button(gw, text="submit", command=gw.submit)
        button.grid(row=8, column=0, pady=2)

        buttonq = ttk.Button(
            gw, text="Quit", command=lambda: controller.show_frame(StartPage)
        )
        buttonq.grid(row=8, column=1, pady=2)

    def submit(gw):

        gw.con = sqlite3.connect("kayu.db")
        gw.c = gw.con.cursor()

        name = gw.nama_var.get()
        warna1 = gw.warna1_var.get()
        warna2 = gw.warna2_var.get()
        tekstur = gw.tekstur_var.get()
        arah = gw.arah_var.get()
        kesan = gw.kesan_var.get()
        kilap = gw.kilap_var.get()
        berat = gw.berat_var.get()

        try:
            gw.c.execute(
                """INSERT INTO KAYU
                (NAMA,W_KAYU_TERAS,W_KAYU_GUBAL,TEKSTUR,ARAH_SERAT,
                KESAN_RABA,KILAP,BERAT)
                VALUES(?,?,?,?,?,?,?,?)""",
                (name, warna1, warna2, tekstur, arah, kesan, kilap, berat),
            )
            gw.con.commit()

        except Exception as e:
            print(f"gagal masuk : {e}")

        gw.entry1.delete(0, "end")
        gw.entry2.delete(0, "end")
        gw.entry3.delete(0, "end")
        gw.entry4.delete(0, "end")
        gw.entry5.delete(0, "end")
        gw.entry6.delete(0, "end")
        gw.entry7.delete(0, "end")
        gw.entry8.delete(0, "end")
        gw.entry8.insert(0, 0.0)
        gw.entry1.focus()

        gw.c.close()
        gw.con.close()


class PageTwo(tk.Frame):
    def __init__(gw, parent, controller):
        tk.Frame.__init__(gw, parent)

        query = "SELECT * FROM KAYU"
        gw.con = sqlite3.connect("kayu.db")
        gw.c = gw.con.cursor()

        gw.id_var = tk.IntVar()
        gw.nama_var = tk.StringVar()
        gw.warna1_var = tk.StringVar()
        gw.warna2_var = tk.StringVar()
        gw.tekstur_var = tk.StringVar()
        gw.arah_var = tk.StringVar()
        gw.kesan_var = tk.StringVar()
        gw.kilap_var = tk.StringVar()
        gw.berat_var = tk.DoubleVar()
        gw.berat_var2 = tk.DoubleVar()

        kotak = tk.LabelFrame(gw, text="Identifikasi Kayu")
        kotak.pack(fill="both", expand="yes", padx=10, pady=10)

        kotak2 = tk.LabelFrame(gw)
        kotak2.pack(fill="both", expand="yes", padx=10, pady=10)

        kotak3 = tk.LabelFrame(gw, text="Hasil Pencarian")
        kotak3.pack(fill="both", expand="yes", padx=10, pady=10)

        # percobaan
        # my_canvas = tk.Canvas(kotak3)
        # my_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        # scroll = ttk.Scrollbar(kotak3, orient=tk.HORIZONTAL, command=my_canvas.xview)
        # scroll.pack(side=tk.BOTTOM, fill=tk.X)
        # my_canvas.configure(xscrollcommand=scroll.set)
        # my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
        # sec_frame = tk.Frame(my_canvas)
        # my_canvas.create_window((2,0), window=sec_frame, anchor="sw")

        label1 = tk.Label(kotak, text="Warna kayu teras")
        label1.grid(row=0, column=0, padx=2)
        label11 = tk.Label(kotak, text=":")
        label11.grid(row=0, column=1, padx=2)

        label2 = tk.Label(kotak, text="Warna kayu gubal")
        label2.grid(row=0, column=3, padx=2)
        label22 = tk.Label(kotak, text=":")
        label22.grid(row=0, column=4, padx=2)

        label = tk.Label(kotak, text="Tekstur")
        label.grid(row=0, column=6)
        cb = ttk.Combobox(kotak, width=20, textvariable=gw.tekstur_var)
        cb["values"] = ("kasar", "halus", "")
        cb.grid(row=0, column=7)

        label = tk.Label(kotak, text="Arah serat")
        label.grid(row=1, column=6)
        cb1 = ttk.Combobox(kotak, width=20, textvariable=gw.arah_var)
        cb1["values"] = ("lurus", "berpadu", "bergelombang", "")
        cb1.grid(row=1, column=7)

        label = tk.Label(kotak, text="Kesan Raba")
        label.grid(row=2, column=6)
        cb2 = ttk.Combobox(kotak, width=20, textvariable=gw.kesan_var)
        cb2["values"] = ("licin", "kesat", "")
        cb2.grid(row=2, column=7)

        label = tk.Label(kotak, text="Kilap")
        label.grid(row=3, column=6)
        cb3 = ttk.Combobox(
            kotak,
            width=20,
            textvariable=gw.kilap_var,
        )
        cb3["values"] = ("mengkilap", "kusam", "")
        cb3.grid(row=3, column=7, columnspan=2)

        gw.table = ttk.Treeview(
            kotak3, columns=(1, 2, 3, 4, 5, 6, 7, 8, 9), show="headings", height="6"
        )
        gw.table.pack()

        gw.table.heading(1, text="ID")
        gw.table.heading(2, text="Nama Kayu")
        gw.table.heading(3, text="Warna k.teras")
        gw.table.heading(4, text="Warna k.gubal")
        gw.table.heading(5, text="Tekstur")
        gw.table.heading(6, text="Arah Serat")
        gw.table.heading(7, text="Kesan Raba")
        gw.table.heading(8, text="Kilap")
        gw.table.heading(9, text="Berat")

        gw.table.column(1, anchor="c", width=80)
        gw.table.column(2, anchor="c", width=120)
        gw.table.column(3, anchor="c", width=120)
        gw.table.column(4, anchor="c", width=120)
        gw.table.column(5, anchor="c", width=120)
        gw.table.column(6, anchor="c", width=120)
        gw.table.column(7, anchor="c", width=120)
        gw.table.column(8, anchor="c", width=120)
        gw.table.column(9, anchor="c", width=120)

        gw.table.bind("<Double 1>", lambda event: gw.getrow(controller))

        kurawal = {
            "merah": "merah",
            "coklat": "coklat",
            "kuning": "kuning",
            "bervariasi": "bervariasi",
            "putih": "putih",
            "kelabu": "kelabu",
        }
        z = 0
        for (key, value) in kurawal.items():

            R = tk.Radiobutton(
                kotak,
                text=key,
                variable=gw.warna1_var,
                value=value,
            )
            R.grid(row=z, column=2, padx=2)

            R2 = tk.Radiobutton(
                kotak,
                text=key,
                variable=gw.warna2_var,
                value=value,
            )
            R2.grid(row=z, column=5, padx=2)
            z += 1

        button = ttk.Button(kotak2, text="Cari", command=gw.submit)
        button.grid(row=5, column=0, pady=2)

        buttonq = ttk.Button(
            kotak2, text="Quit", command=lambda: controller.show_frame(StartPage)
        )
        buttonq.grid(row=5, column=1, pady=2)

        gw.c.execute(query)
        rows2 = gw.c.fetchall()
        gw.update(rows2)

        buttonz = ttk.Button(kotak2, text="Refresh", command=lambda: gw.update(rows2))
        buttonz.grid(row=5, column=2, pady=2)

    def submit(gw):

        warna1 = gw.warna1_var.get()
        warna2 = gw.warna2_var.get()
        tek = gw.tekstur_var.get()
        ar = gw.arah_var.get()
        kes = gw.kesan_var.get()
        kil = gw.kilap_var.get()

        try:
            gw.c.execute(
                f"""SELECT * FROM KAYU WHERE
                            W_KAYU_TERAS LIKE ? AND
                            W_KAYU_GUBAL LIKE ? AND
                            TEKSTUR LIKE ? AND
                            ARAH_SERAT LIKE ? AND
                            KESAN_RABA LIKE ? AND
                            KILAP LIKE ?""",
                (
                    "%" + warna1 + "%",
                    "%" + warna2 + "%",
                    "%" + tek + "%",
                    "%" + ar + "%",
                    "%" + kes + "%",
                    "%" + kil + "%",
                ),
            )

        except Exception as e:
            print(f"gagal masuk : {e}")
        rows1 = gw.c.fetchall()

        gw.update(rows1)

    def update(gw, baris):
        gw.table.delete(*gw.table.get_children())
        for i in baris:
            gw.table.insert("", "end", values=i)

    def getrow(gw, cont):
        # rowid = gw.table.identify_row(event.y)
        item = gw.table.item(gw.table.focus())
        gw.id_var.set(item["values"][0])
        gw.nama_var.set(item["values"][1])
        gw.warna1_var.set(item["values"][2])
        gw.warna2_var.set(item["values"][3])
        gw.tekstur_var.set(item["values"][4])
        gw.arah_var.set(item["values"][5])
        gw.kesan_var.set(item["values"][6])
        gw.kilap_var.set(item["values"][7])
        gw.berat_var.set(item["values"][8])

        dsd = tk.IntVar()
        tbc = tk.IntVar()
        vp = tk.IntVar()
        way = tk.IntVar()

        way.set(1)

        top = tk.Toplevel()

        kotak = tk.LabelFrame(top, text="Identitas")
        kotak.pack(fill="both", expand="yes", padx=10, pady=10)

        kot = tk.LabelFrame(top, text="Cetak Laporan")
        kot.pack(fill="both", expand="yes", padx=10, pady=10)

        label = tk.Label(kotak, text="ID\t\t" + str(gw.id_var.get()))
        label.grid(row=0, column=1, sticky="nw")
        label1 = tk.Label(kotak, text="Nama Kayu\t" + gw.nama_var.get())
        label1.grid(row=1, column=1, sticky="nw")
        label2 = tk.Label(kotak, text="Warna Teras\t" + gw.warna1_var.get())
        label2.grid(row=2, column=1, sticky="nw")
        label3 = tk.Label(kotak, text="Warna Gubal\t" + gw.warna2_var.get())
        label3.grid(row=3, column=1, sticky="nw")
        label4 = tk.Label(kotak, text="Tekstur\t\t" + gw.tekstur_var.get())
        label4.grid(row=4, column=1, sticky="nw")
        label5 = tk.Label(kotak, text="Arah Serat\t" + gw.arah_var.get())
        label5.grid(row=5, column=1, sticky="nw")

        labelway = tk.Label(kot, text="Jalur")
        labelway.grid(row=6, column=1, sticky="nw")
        ent1 = tk.Entry(kot, textvariable=way)
        ent1.grid(row=6, column=2)

        labeldsd = tk.Label(kot, text="Diameter Setinggi Dada(CM)")
        labeldsd.grid(row=7, column=1, sticky="nw")
        ent = tk.Entry(kot, textvariable=dsd)
        ent.grid(row=7, column=2)

        labeltbc = tk.Label(kot, text="Tinggi Bebas Cabang(M)")
        labeltbc.grid(row=8, column=1, sticky="nw")
        enti = tk.Entry(kot, textvariable=tbc)
        enti.grid(row=8, column=2)

        labelvp = tk.Label(kot, text="Volume Pohon(M^3)")
        labelvp.grid(row=9, column=1, sticky="nw")
        ento = tk.Entry(kot, textvariable=vp)
        ento.grid(row=9, column=2)

        but = ttk.Button(
            kot,
            text="Cetak Laporan",
            command=lambda: gw.cetak(
                gw.nama_var.get(), way.get(), dsd.get(), tbc.get(), vp.get()
            ),
        )

        if cont.loginState.get() == 1:
            but.grid(row=10, column=1, padx=5, sticky="nsew")

        top.mainloop()

    def cetak(gw, nama_kayu, way, dsd, tbc, vp):
        laporan.func(nama_kayu, way, dsd, tbc, vp)
        messagebox.showinfo("Sukses", "Berhasil Mencetak Laporan")
        


class laporanRekap(tk.Frame):
    def __init__(gw, parent, controller):
        tk.Frame.__init__(gw, parent)

        gw.tbl = ttk.Treeview(gw, columns=(1), show="headings", height=8)
        gw.tbl.pack()
        gw.tbl.heading(1, text="Laporan")

        files = os.listdir("laporan")
        for fille in files:
            gw.tbl.insert("", "end", values=fille)

        gw.tbl.bind("<Double 1>", gw.buka)

        instruksi = tk.Label(gw, text="Klik 2x pada nama file untuk membuka")
        instruksi.pack()

    def buka(gw, event):
        # item_id = event.widget.focus()
        # item = event.widget.item(item_id)
        item_id = gw.tbl.focus()
        item = gw.tbl.item(item_id)
        values = item["values"][0]

        # print("the url is:", values)
        lokasi = os.getcwd()
        os.startfile(lokasi + "/laporan/" + values)


App = LautanFrame()

App.mainloop()
