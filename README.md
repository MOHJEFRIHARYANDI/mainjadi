import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import pandas as pd
import openpyxl
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from collections import Counter
import logging
import os
import re
from datetime import datetime
import matplotlib.pyplot as plt
from collections import Counter
import seaborn as sns

def buat_dashboard_statistik():
    """Tampilkan dashboard statistik dalam jendela terpisah."""
    try:
        file = "data_santri.xlsx"
        if not os.path.exists(file):
            messagebox.showerror("Error", "File data_santri.xlsx tidak ditemukan.")
            return

        # Baca data dari Excel
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        # Ambil data dari Excel
        jenis_kelamin = []
        kelas_tujuan = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > 5 and row[5]:  # Kolom jenis kelamin
                value = row[5].strip().capitalize()  # Hapus spasi tambahan dan kapitalisasi
                if value in ["Laki-laki", "Perempuan"]:  # Hanya data valid
                    jenis_kelamin.append(value)
            if len(row) > 24 and row[24]:  # Kolom kelas tujuan
                kelas_tujuan.append(row[24].strip())
        # Debugging: Cetak semua nilai yang ditemukan di kolom Jenis Kelamin
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > 5:
                print("Nilai Kolom Jenis Kelamin:", row[5])        

        # Debugging: Periksa data
        print("Data Jenis Kelamin (Debug):", jenis_kelamin)
        print("Data Kelas Tujuan (Debug):", kelas_tujuan)

        # Ambil data jenis kelamin dengan validasi
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > 5 and row[5]:  # Kolom jenis kelamin
                value = row[5].strip().capitalize()  # Hapus spasi tambahan dan kapitalisasi
                if value in ["Laki-laki", "Perempuan"]:  # Hanya data valid
                    jenis_kelamin.append(value)
                else:
                    print(f"Nilai tidak valid ditemukan di kolom Jenis Kelamin: {row[5]}")

        # Hitung distribusi
        jk_counter = Counter(jenis_kelamin)
        kelas_counter = Counter(kelas_tujuan)

        # Konfigurasi gaya modern dengan seaborn
        sns.set_theme(style="whitegrid")

        # Buat plot
        fig, axes = plt.subplots(1, 3, figsize=(18, 6))

        # Diagram Donat untuk Jenis Kelamin
        wedges, texts, autotexts = axes[0].pie(
            jk_counter.values(),
            labels=jk_counter.keys(),
            autopct="%.1f%%",
            startangle=90,
            colors=sns.color_palette("pastel"),
            textprops={'fontsize': 12}
        )
        # Tambahkan lubang di tengah untuk membuat grafik donat
        centre_circle = plt.Circle((0, 0), 0.70, fc='white')
        axes[0].add_artist(centre_circle)
        axes[0].set_title("Distribusi Jenis Kelamin (Donut Chart)", fontsize=14)

        # Diagram Batang untuk Kelas Tujuan
        kelas_sorted = dict(sorted(kelas_counter.items(), key=lambda x: x[1], reverse=True))
        sns.barplot(
            x=list(kelas_sorted.keys()),
            y=list(kelas_sorted.values()),
            ax=axes[1]  # Hapus parameter palette
        )
        axes[1].set_title("Jumlah Santri Per Kelas", fontsize=14)
        axes[1].set_xlabel("Kelas", fontsize=12)
        axes[1].set_ylabel("Jumlah Santri", fontsize=12)
        axes[1].tick_params(axis="x", rotation=45, labelsize=10)

        # Grafik Jumlah Santri Menurut Jenis Kelamin (Bar Chart)
        sns.barplot(
            x=list(jk_counter.keys()),
            y=list(jk_counter.values()),
            ax=axes[2]  # Hapus parameter palette
        )
        axes[2].set_title("Jumlah Santri Menurut Jenis Kelamin", fontsize=14)
        axes[2].set_xlabel("Jenis Kelamin", fontsize=12)
        axes[2].set_ylabel("Jumlah Santri", fontsize=12)
        axes[2].tick_params(axis="x", labelsize=10)

        # Grafik Jumlah Santri Menurut Jenis Kelamin (Bar Chart)
        sns.barplot(
            x=list(jk_counter.keys()),
            y=list(jk_counter.values()),
            palette="coolwarm",
            ax=axes[2]
        )
        axes[2].set_title("Jumlah Santri Menurut Jenis Kelamin", fontsize=14)
        axes[2].set_xlabel("Jenis Kelamin", fontsize=12)
        axes[2].set_ylabel("Jumlah Santri", fontsize=12)
        axes[2].tick_params(axis="x", labelsize=10)

        # Atur tata letak grafik
        plt.tight_layout()
        plt.show()

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan saat membuat dashboard: {e}")

# Tambahkan tombol di header
def buat_header(root):
    header_frame = tk.Frame(root, bg="#003366", height=90)
    header_frame.pack(fill="x")

    # Tombol Dashboard
    btn_dashboard = tk.Button(header_frame, text="Dashboard Statistik", command=buat_dashboard_statistik, bg="#17a2b8", fg="white", relief="flat")
    btn_dashboard.pack(side="right", padx=5)

# ========== KONFIGURASI UTAMA ==========
root = tk.Tk()
root.title("APLIKASI DATA SANTRI MIFTAHUL ULUM JALMAK PAMEKASAN")
root.geometry("1300x720")
root.configure(bg="white")

# Logging untuk mencatat error
logging.basicConfig(filename="app.log", level=logging.ERROR, format="%(asctime)s - %(levelname)s - %(message)s")

# ========== TTK STYLE ==========
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
style.map("Treeview", background=[("selected", "#347083")])
style.configure("TButton", font=("Helvetica", 10, "bold"), background="#0052cc", foreground="white", padding=5)
style.map("TButton", background=[("active", "#003d99")])

# ========== HEADER ==========
def buat_header():
    header_frame = tk.Frame(root, bg="#003366", height=90)
    header_frame.pack(fill="x")

    # Logo
    try:
        if os.path.exists("logo_yayasan.png"):
            logo = Image.open("logo_yayasan.png").resize((60, 60))
            logo_photo = ImageTk.PhotoImage(logo)
            logo_label = tk.Label(header_frame, image=logo_photo, bg="#003366")
            logo_label.image = logo_photo
        else:
            raise FileNotFoundError
    except Exception as e:
        logging.error(f"Error memuat logo: {e}")
        logo_label = tk.Label(header_frame, text="[LOGO]", bg="#003366", fg="white", font=("Arial", 14, "bold"))
    logo_label.pack(side="left", padx=20, pady=10)

    # Judul
    judul = tk.Label(header_frame, text="APLIKASI DATA SANTRI\nMIFTAHUL ULUM JALMAK PAMEKASAN",
                     font=("Helvetica", 18, "bold"), fg="white", bg="#003366", justify="center")
    judul.pack(pady=10)

    # Tombol Keluar
    btn_keluar = tk.Button(header_frame, text="Keluar", command=on_closing, bg="#dc3545", fg="white", relief="flat")
    btn_keluar.pack(side="right", padx=5)

    # Tombol Pusat Bantuan
    btn_bantuan = tk.Button(header_frame, text="Pusat Bantuan", command=buka_pusat_bantuan, bg="#17a2b8", fg="white", relief="flat")
    btn_bantuan.pack(side="right", padx=5)

    # Tombol Tentang Aplikasi
    btn_tentang = tk.Button(header_frame, text="Tentang Aplikasi", command=buka_tentang_aplikasi, bg="#28a745", fg="white", relief="flat")
    btn_tentang.pack(side="right", padx=5)

    # Tombol Hubungi Kami
    btn_hubungi = tk.Button(header_frame, text="Hubungi Kami", command=buka_hubungi_kami, bg="#ffc107", fg="black", relief="flat")
    btn_hubungi.pack(side="right", padx=5)
    
    # Tombol Dashboard Statistik
    btn_dashboard = tk.Button(header_frame, text="Dashboard Statistik", command=buat_dashboard_statistik, bg="#17a2b8", fg="white", relief="flat")
    btn_dashboard.pack(side="right", padx=5)    
    
# ========== FUNGSI VALIDASI ==========
def validate_entry(entry, validation_function, error_message):
    """Validasi input pada kolom Entry."""
    def on_validate(event):
        if not validation_function(entry.get()):
            messagebox.showerror("Error", error_message)
            entry.delete(0, tk.END)
    entry.bind("<FocusOut>", on_validate)  # Validasi saat pengguna keluar dari kolom

def enforce_uppercase(event):
    """Mengubah teks input menjadi huruf besar secara otomatis tanpa menghapus teks."""
    widget = event.widget
    current_text = widget.get()
    uppercase_text = current_text.upper()
    if current_text != uppercase_text:  # Ubah teks hanya jika berbeda
        widget.delete(0, tk.END)
        widget.insert(0, uppercase_text)

def validasi_tanggal(tanggal):
    """Validasi format tanggal DD/MM/YYYY."""
    try:
        datetime.strptime(tanggal, "%d/%m/%Y")
        return True
    except ValueError:
        return False

# ========== FORM INPUT ==========
entries = {}

def add_hover_effect(button):
    """Tambahkan efek hover pada tombol."""
    def on_enter(e):
        e.widget['background'] = '#003d99'
    def on_leave(e):
        e.widget['background'] = '#0052cc'
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)

def buat_form_input(parent):
    global entries

    canvas = tk.Canvas(parent, bg="white")
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    form_frame = tk.Frame(canvas, bg="white")
    form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=form_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def add_section(title, fields):
        tk.Label(form_frame, text=title, bg="white", font=("Helvetica", 12, "bold"), pady=5).pack(anchor="w")
        for field in fields:
            row = tk.Frame(form_frame, bg="white")
            row.pack(fill="x", pady=2)
            tk.Label(row, text=field, width=25, anchor="w", bg="white").pack(side="left")
            entry = ttk.Entry(row, width=40)
            entry.pack(side="left", fill="x", expand=True)
            entries[field] = entry
            if field not in ["NIS", "NIK", "NOMOR KK", "NOMOR WHATSAPP"]:
                entry.bind("<KeyRelease>", enforce_uppercase)
            if field == "NIS":
                validate_entry(entry, str.isdigit, "NIS harus berupa angka.")
            elif field == "NIK":
                validate_entry(entry, lambda x: len(x) == 16 and x.isdigit(), "NIK harus terdiri dari 16 angka.")
            elif field == "NOMOR WHATSAPP":
                validate_entry(entry, str.isdigit, "Nomor WhatsApp harus berupa angka.")
            elif field == "TANGGAL LAHIR":
                validate_entry(entry, validasi_tanggal, "Gunakan format DD/MM/YYYY.")

    add_section("A. DATA PRIBADI", ["NO REG", "NIS", "NAMA", "NIK", "NOMOR KK", "JENIS KELAMIN", "TEMPAT LAHIR", "TANGGAL LAHIR", "AGAMA", "KEWARGANEGARAAN", "ANAK KE", "JUMLAH SAUDARA KANDUNG"])
    add_section("B. DATA ALAMAT", ["ALAMAT", "RT", "RW", "DESA", "DUSUN", "KECAMATAN", "KABUPATEN", "PROVINSI"])
    add_section("C. DATA ORANG TUA", ["NAMA AYAH", "NIK AYAH", "NAMA IBU", "NIK IBU"])
    add_section("D. DATA PENDIDIKAN", ["KELAS TUJUAN", "KETERANGAN", "NOMOR WHATSAPP", "TANGGAL MASUK"])

    tombol_frame = tk.Frame(form_frame, bg="white")
    tombol_frame.pack(pady=10)

    btn_simpan = tk.Button(tombol_frame, text="Simpan", bg="#0052cc", fg="white", relief="flat", command=submit_data)
    btn_simpan.pack(side="left", padx=5)
    add_hover_effect(btn_simpan)

    btn_reset = tk.Button(tombol_frame, text="Reset", command=reset_form, relief="flat")
    btn_reset.pack(side="left", padx=5)
    add_hover_effect(btn_reset)

    tk.Button(tombol_frame, text="Ekspor ke Excel", command=export_to_excel, bg="#28a745", fg="white", relief="flat").pack(side="left", padx=5)
    tk.Button(tombol_frame, text="Impor dari Excel", command=import_from_excel, bg="#17a2b8", fg="white", relief="flat").pack(side="left", padx=5)
    tk.Button(tombol_frame, text="Cetak PDF", command=cetak_pdf, bg="#ffc107", fg="black", relief="flat").pack(side="left", padx=5)

    next_no_reg = get_next_no_reg()
    entries["NO REG"].insert(0, str(next_no_reg))
    next_nis = get_next_nis()
    entries["NIS"].insert(0, str(next_nis))

# ========== STATUS BAR ==========
def buat_status_bar():
    status_bar = tk.Label(root, text="Ready", bd=1, relief="sunken", anchor="w", bg="#f0f0f0")
    status_bar.pack(side="bottom", fill="x")
    return status_bar

status_bar = buat_status_bar()

# ========== TABEL ==========
def buat_tabel(parent):
    global tree, search_var, kelas_var

    # Frame untuk Filter dan Pencarian
    filter_frame = tk.Frame(parent, bg="white")
    filter_frame.pack(fill="x")

    tk.Label(filter_frame, text="Cari Nama:", bg="white").pack(side="left", padx=5)
    search_var = tk.StringVar()
    search_entry = ttk.Entry(filter_frame, textvariable=search_var, width=20)
    search_entry.pack(side="left", padx=5)

    tk.Label(filter_frame, text="Filter Kelas:", bg="white").pack(side="left", padx=5)
    kelas_var = tk.StringVar()
    kelas_combo = ttk.Combobox(filter_frame, textvariable=kelas_var, values=["Semua", "1 Awwaliyah", "2 Awwaliyah", "3 Awwaliyah", "4 Wustha", " 5 Wustha", "6 Wustha", "Sifir 0.K", "Sifir 0.B"], state="readonly", width=15)
    kelas_combo.current(0)  # Set pilihan default ke "Semua"
    kelas_combo.pack(side="left", padx=5)

    tk.Button(filter_frame, text="Cari", command=filter_data).pack(side="left", padx=5)
    tk.Button(filter_frame, text="Reset", command=reset_filter).pack(side="left", padx=5)

    # Treeview (Tabel)
    tree_frame = tk.Frame(parent)
    tree_frame.pack(fill="both", expand=True)

    columns = ["NO", "NO REG", "NIS", "NAMA", "ALAMAT", "NAMA AYAH", "NAMA IBU", "TANGGAL MASUK", "KELAS TUJUAN", "KETERANGAN"]
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    tree.pack(fill="both", expand=True)
    
    # Muat data dari Excel
    muat_data_dari_excel(limit=23)

    # Frame untuk Statistik
    statistik_frame = tk.Frame(parent, bg="white")
    statistik_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Buat Tabel Statistik
    buat_tabel_statistik(statistik_frame)

# ========== FUNGSI ==========
def validasi_input(data):
    # Validasi Kolom Wajib
    if not data["NAMA"]:
        return "Nama wajib diisi."
    if not data["KELAS TUJUAN"]:
        return "Kelas tujuan wajib dipilih."

    # Validasi NIS (Harus Angka)
    if not data["NIS"].isdigit():
        return "NIS harus berupa angka."

    # Validasi NIK (Panjang Harus 16)
    if len(data["NIK"]) != 16:
        return "NIK harus terdiri dari 16 angka."

    # Validasi Nomor WhatsApp (Harus Angka)
    if not data["NOMOR WHATSAPP"].isdigit():
        return "Nomor WhatsApp harus berupa angka."

    # Validasi Tanggal Lahir (Format DD/MM/YYYY)
    from datetime import datetime
    try:
        datetime.strptime(data["TANGGAL LAHIR"], "%d/%m/%Y")
    except ValueError:
        return "Format TANGGAL LAHIR tidak valid. Gunakan format DD/MM/YYYY."

    return None  # Jika semua validasi lolos

def submit_data():
    # Ambil semua data dari form
    data = {k: v.get() for k, v in entries.items()}

    # Validasi data
    error_message = validasi_input(data)
    if error_message:
        messagebox.showerror("Error", error_message)
        return

    # Jika semua validasi lolos, simpan ke Excel
    simpan_ke_excel(data)

    # Tambahkan data ke tabel
    tree.insert("", "end", values=[
        len(tree.get_children()) + 1,  # Tambahkan nomor urut di tabel
        data["NO REG"], data["NIS"], data["NAMA"], data["ALAMAT"],
        data["NAMA AYAH"], data["NAMA IBU"], data["TANGGAL MASUK"],
        data["KELAS TUJUAN"], data["KETERANGAN"]
    ])
    reset_form()

def simpan_ke_excel(data):
    file = "data_santri.xlsx"
    kolom = list(entries.keys())
    if not os.path.exists(file):
        wb = Workbook()
        ws = wb.active
        ws.append(list(entries.keys()))  # Tambahkan header kolom
        wb.save(file)

    wb = openpyxl.load_workbook(file)
    ws = wb.active
    ws.append([data.get(k, "") for k in kolom])
    wb.save(file)

def reset_form():
    for e in entries.values():
        e.delete(0, tk.END)

    # Isi otomatis NO REG saat form di-reset
    next_no_reg = get_next_no_reg()
    entries["NO REG"].insert(0, str(next_no_reg))

    # Isi otomatis NIS saat form di-reset
    next_nis = get_next_nis()
    entries["NIS"].insert(0, str(next_nis))

def export_to_excel():
    data = [tree.item(i)['values'] for i in tree.get_children()]
    if data:
        df = pd.DataFrame(data, columns=["NO", "NO REG", "NIS", "NAMA", "ALAMAT", "NAMA AYAH", "NAMA IBU", "TANGGAL MASUK", "KELAS TUJUAN", "KETERANGAN"])
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            df.to_excel(path, index=False)
            messagebox.showinfo("Sukses", "Data berhasil diekspor ke Excel.")

def import_from_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if path:
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                tree.insert("", "end", values=row)
            messagebox.showinfo("Sukses", "Data berhasil diimpor.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengimpor data: {str(e)}")

def filter_data():
    # Ambil nilai dari input pencarian dan filter
    nama_cari = search_var.get().strip().lower()
    kelas_pilih = kelas_var.get()

    # Bersihkan tabel sebelum memuat data baru
    tree.delete(*tree.get_children())

    file = "data_santri.xlsx"
    if os.path.exists(file):
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        # print(f"DEBUG: Filter Kelas = {kelas_pilih}, Pencarian Nama = {nama_cari}")

        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
            # Ambil data dari setiap kolom sesuai dengan urutan di Excel
            no_reg = row[0]
            nis = row[1]
            nama = row[2]
            alamat = row[12]
            nama_ayah = row[20]
            nama_ibu = row[22]
            tanggal_masuk = row[27]
            kelas_tujuan = row[24]
            keterangan = row[25]

            # print(f"DEBUG: Row Data = {row}")

            # Filter berdasarkan nama
            if nama_cari and nama_cari not in nama.lower():
                continue

            # Filter berdasarkan kelas
            if kelas_pilih != "Semua" and (not kelas_tujuan or kelas_tujuan.strip() != kelas_pilih):
                continue

            # Tambahkan data yang sesuai ke tabel
            tree.insert("", "end", values=[
                i, no_reg, nis, nama, alamat, nama_ayah, nama_ibu, tanggal_masuk, kelas_tujuan, keterangan
            ])
            # print(f"DEBUG: Data ditambahkan ke tabel = {row}")

def reset_filter():
    search_var.set("")  # Kosongkan kolom pencarian
    kelas_var.set("Semua")  # Set ComboBox ke pilihan default ("Semua")

    # Muat ulang semua data ke tabel
    muat_data_dari_excel()

def cetak_pdf():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Error", "Pilih data yang akan dicetak.")
        return
    try:
        data = tree.item(selected[0])["values"]
        pdf = canvas.Canvas("biodata_santri.pdf", pagesize=A4)
        pdf.drawString(100, 800, f"NO REG: {data[1]}")
        pdf.drawString(100, 780, f"Nama: {data[3]}")
        pdf.drawString(100, 760, f"Alamat: {data[4]}")
        pdf.drawString(100, 740, f"Kelas: {data[8]}")
        pdf.save()
        messagebox.showinfo("Sukses", "Biodata berhasil dicetak ke PDF.")
    except Exception as e:
        logging.error(f"Gagal mencetak PDF: {str(e)}")
        messagebox.showerror("Error", f"Gagal mencetak PDF: {str(e)}")

def muat_data_dari_excel(limit=23):
    file = "data_santri.xlsx"
    if os.path.exists(file):
        try:
            wb = openpyxl.load_workbook(file)
            ws = wb.active

            # Bersihkan tabel sebelum memuat data baru
            tree.delete(*tree.get_children())  

            # Ambil data dengan batas tertentu
            for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
                if i > limit:
                    break
                if len(row) < 28:  # Validasi jumlah kolom
                    logging.warning(f"Baris {i + 1} memiliki data yang tidak lengkap, dilewati.")
                    continue
                tree.insert("", "end", values=[
                    i,  # NO (nomor urut)
                    row[0],  # NO REG
                    row[1],  # NIS
                    row[2],  # NAMA
                    row[12],  # ALAMAT
                    row[20],  # NAMA AYAH
                    row[22],  # NAMA IBU
                    row[27],  # TANGGAL MASUK
                    row[24],  # KELAS TUJUAN
                    row[25],  # KETERANGAN
                ])
        except Exception as e:
            logging.error(f"Tidak dapat membaca file Excel: {str(e)}")
            messagebox.showerror("Error", f"Tidak dapat membaca file Excel: {str(e)}")
    else:
        logging.info("File data_santri.xlsx tidak ditemukan.")
        messagebox.showinfo("Info", "File data_santri.xlsx tidak ditemukan.")

def get_next_no_reg():
    file = "data_santri.xlsx"
    if os.path.exists(file):
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        no_regs = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0] is not None]
        if no_regs:
            return max(map(int, no_regs)) + 1  # Ambil nomor terakhir dan tambahkan 1
    return 1  # Jika file tidak ada atau kosong, mulai dari 1

def get_next_nis():
    file = "data_santri.xlsx"
    if os.path.exists(file):
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        # Ambil semua nilai di kolom NIS, hanya gunakan yang valid (angka saja)
        nises = [
            row[1] for row in ws.iter_rows(min_row=2, values_only=True)
            if row[1] is not None and str(row[1]).isdigit()
        ]
        if nises:
            return max(map(int, nises)) + 1  # Ambil NIS terbesar dan tambahkan 1
    return 1001  # Jika file tidak ada atau kosong, mulai dari 1001
    
def get_kelas_list():
    file = "data_santri.xlsx"
    kelas_list = []
    if os.path.exists(file):
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        # Ambil semua nilai di kolom kelas (misalnya kolom ke-8)
        for row in ws.iter_rows(min_row=2, values_only=True):
            kelas = row[24]  # Sesuaikan dengan indeks kolom kelas di file Excel Anda (0 = kolom A)
            if kelas and str(kelas).strip() not in kelas_list:  # Pastikan kelas diubah menjadi string
                kelas_list.append(str(kelas).strip())  # Tambahkan kelas sebagai string yang sudah di-trim

    return sorted(kelas_list)  # Kembalikan daftar kelas yang sudah diurutkan
    
def validasi_input(data):
    # Validasi Kolom Wajib
    if not data["NAMA"]:
        return "Nama wajib diisi."
    if not data["KELAS TUJUAN"]:
        return "Kelas tujuan wajib dipilih."

    # Validasi NIS (Harus Angka)
    if not data["NIS"].isdigit():
        return "NIS harus berupa angka."

    # Validasi NIK (Panjang Harus 16)
    if len(data["NIK"]) != 16:
        return "NIK harus terdiri dari 16 angka."

    # Validasi Nomor WhatsApp (Harus Angka)
    if not data["NOMOR WHATSAPP"].isdigit():
        return "Nomor WhatsApp harus berupa angka."

    # Validasi Tanggal Lahir (Format DD/MM/YYYY)
    from datetime import datetime
    try:
        datetime.strptime(data["TANGGAL LAHIR"], "%d/%m/%Y")
    except ValueError:
        return "Format TANGGAL LAHIR tidak valid. Gunakan format DD/MM/YYYY."

    return None  # Jika semua validasi lolos

def submit_data():
    # Ambil semua data dari form
    data = {k: v.get() for k, v in entries.items()}

    # Validasi data
    error_message = validasi_input(data)
    if error_message:
        messagebox.showerror("Error", error_message)
        return

    # Jika semua validasi lolos, simpan ke Excel
    simpan_ke_excel(data)

    # Tambahkan data ke tabel
    tree.insert("", "end", values=[
        len(tree.get_children()) + 1,  # Tambahkan nomor urut di tabel
        data["NO REG"], data["NIS"], data["NAMA"], data["ALAMAT"],
        data["NAMA AYAH"], data["NAMA IBU"], data["TANGGAL MASUK"],
        data["KELAS TUJUAN"], data["KETERANGAN"]
    ])
    reset_form()

    
def on_closing():
    if messagebox.askokcancel("Keluar", "Apakah Anda yakin ingin keluar?"):
        root.destroy()
        
def tentang_aplikasi():
    messagebox.showinfo("Tentang", "Aplikasi Input Data Santri versi 1.0\n© Yayasan Miftahul Ulum Jalmak, 2025")
    
def bantuan():
    messagebox.showinfo("Bantuan", "Hubungi admin untuk bantuan lebih lanjut: admin@yayasan.com\nAbd Muiz Syamsuri\nJalmak Pamekasan Madura Jawa Timur")

def buat_tabel_statistik(parent):
    """Fungsi untuk membuat tabel statistik jumlah santri."""
    # Bersihkan frame sebelumnya (jika ada)
    for widget in parent.winfo_children():
        widget.destroy()

    # Data dari file Excel
    file = "data_santri.xlsx"
    if os.path.exists(file):
        try:
            wb = openpyxl.load_workbook(file)
            ws = wb.active

            # Hitung statistik
            total_santri = 0
            total_perempuan = 0
            total_laki_laki = 0
            kelas_dict = {}

            for row in ws.iter_rows(min_row=2, values_only=True):
                total_santri += 1
                jenis_kelamin = row[23]  # Kolom JENIS KELAMIN
                kelas = row[24]  # Kolom KELAS TUJUAN

                if jenis_kelamin == "Perempuan":
                    total_perempuan += 1
                elif jenis_kelamin == "Laki-Laki":
                    total_laki_laki += 1

                if kelas:
                    kelas_dict[kelas] = kelas_dict.get(kelas, 0) + 1

            # Buat Treeview untuk menampilkan statistik
            columns = ["Keterangan", "Jumlah"]
            statistik_tree = ttk.Treeview(parent, columns=columns, show="headings", height=8)

            # Atur heading
            statistik_tree.heading("Keterangan", text="Keterangan")
            statistik_tree.heading("Jumlah", text="Jumlah")

            # Atur lebar kolom
            statistik_tree.column("Keterangan", width=200, anchor="w")
            statistik_tree.column("Jumlah", width=100, anchor="center")

            # Masukkan data ke tabel
            statistik_tree.insert("", "end", values=["Jumlah Keseluruhan Santri", total_santri])
            statistik_tree.insert("", "end", values=["Jumlah Santri Perempuan", total_perempuan])
            statistik_tree.insert("", "end", values=["Jumlah Santri Laki-Laki", total_laki_laki])

            # Masukkan data jumlah santri setiap kelas
            for kelas, jumlah in kelas_dict.items():
                statistik_tree.insert("", "end", values=[f"Jumlah Santri Kelas {kelas}", jumlah])

            # Tampilkan Treeview
            statistik_tree.pack(fill="both", expand=True, padx=10, pady=10)
        except Exception as e:
            print(f"Error memuat data statistik: {str(e)}")
            messagebox.showerror("Error", f"Gagal memuat data statistik: {str(e)}")
    else:
        messagebox.showinfo("Info", "File data_santri.xlsx tidak ditemukan.")
        
def buka_pusat_bantuan():
    """Tampilkan jendela pusat bantuan."""
    bantuan_window = tk.Toplevel(root)
    bantuan_window.title("Pusat Bantuan")
    bantuan_window.geometry("500x400")
    bantuan_window.configure(bg="white")

    tk.Label(bantuan_window, text="Pusat Bantuan", font=("Helvetica", 16, "bold"), bg="white").pack(pady=10)
    tk.Label(bantuan_window, text="Berikut adalah panduan untuk menggunakan aplikasi ini:", bg="white", anchor="w", justify="left").pack(fill="x", padx=10)
    tk.Label(bantuan_window, text="1. Klik tombol 'Simpan' untuk menyimpan data.\n"
                                  "2. Klik tombol 'Ekspor ke Excel' untuk menyimpan data ke file Excel.\n"
                                  "3. Klik tombol 'Impor dari Excel' untuk memuat data dari file Excel.\n"
                                  "4. Klik tombol 'Cetak PDF' untuk mencetak biodata santri.\n"
                                  "5. Gunakan fitur pencarian dan filter untuk menemukan data dengan cepat.",
             bg="white", justify="left", anchor="w").pack(fill="x", padx=10, pady=10)

    tk.Button(bantuan_window, text="Tutup", command=bantuan_window.destroy, bg="#dc3545", fg="white", relief="flat").pack(pady=10)

def buka_tentang_aplikasi():
    """Tampilkan jendela tentang aplikasi."""
    tentang_window = tk.Toplevel(root)
    tentang_window.title("Tentang Aplikasi")
    tentang_window.geometry("400x300")
    tentang_window.configure(bg="white")

    tk.Label(tentang_window, text="Tentang Aplikasi", font=("Helvetica", 16, "bold"), bg="white").pack(pady=10)
    tk.Label(tentang_window, text="Aplikasi Data Santri\nVersi 1.0\n\n"
                                  "Dikembangkan oleh:\nABD MUIZ SYAMSURI\n\n"
                                  "Hak Cipta © 2025", bg="white", justify="center").pack(pady=10)

    tk.Button(tentang_window, text="Tutup", command=tentang_window.destroy, bg="#dc3545", fg="white", relief="flat").pack(pady=10)

def buka_hubungi_kami():
    """Tampilkan jendela hubungi kami."""
    hubungi_window = tk.Toplevel(root)
    hubungi_window.title("Hubungi Kami")
    hubungi_window.geometry("400x300")
    hubungi_window.configure(bg="white")

    tk.Label(hubungi_window, text="Hubungi Kami", font=("Helvetica", 16, "bold"), bg="white").pack(pady=10)
    tk.Label(hubungi_window, text="Jika Anda membutuhkan bantuan lebih lanjut, hubungi kami melalui:\n\n"
                                  "Email: mdtajalmak@gmail.com\n"
                                  "Telepon: 081999619992\n\n"
                                  "Website: www.yayasanmiftahululumjalmak.com",
             bg="white", justify="center").pack(pady=10)

    tk.Button(hubungi_window, text="Tutup", command=hubungi_window.destroy, bg="#dc3545", fg="white", relief="flat").pack(pady=10)
  
# ========== MAIN ==========
buat_header()

main_frame = tk.Frame(root, bg="white")
main_frame.pack(fill="both", expand=True)

frame_kiri = tk.Frame(main_frame, width=650, bg="white")
frame_kanan = tk.Frame(main_frame, bg="white")

frame_kiri.pack(side="left", fill="both", expand=True, padx=10, pady=10)
frame_kanan.pack(side="right", fill="both", expand=True, padx=10, pady=10)

buat_form_input(frame_kiri)
buat_tabel(frame_kanan)

# ========== FOOTER ==========
footer = tk.Frame(root, bg="#f0f0f0", height=30)
footer.pack(fill="x")
tk.Label(footer, text="© 2025 - Yayasan Miftahul Ulum Jalmak | Created by : Abd. Muiz Syamsuri", bg="#f0f0f0", font=("Arial", 9)).pack()

# Muat data dari Excel ke tabel saat aplikasi dimulai
muat_data_dari_excel()

# Jalankan aplikasi
root.mainloop()

