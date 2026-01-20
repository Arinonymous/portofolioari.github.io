import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import messagebox, filedialog
import csv
import time

def ambil_data(npsn):
    url = f"https://dapo.kemendikdasmen.go.id/sekolah/{npsn}"
    headers = {"User-Agent": "Mozilla/5.0"}

    r = requests.get(url, headers=headers, timeout=15)
    if r.status_code != 200:
        return None

    soup = BeautifulSoup(r.text, "html.parser")

    def get_text(label):
        item = soup.find("td", string=label)
        return item.find_next_sibling("td").text.strip() if item else "-"

    return {
        "NPSN": npsn,
        "Nama Sekolah": get_text("Nama Sekolah"),
        "Alamat": get_text("Alamat"),
        "Desa/Kelurahan": get_text("Desa/Kelurahan"),
        "Kecamatan": get_text("Kecamatan"),
        "Kabupaten/Kota": get_text("Kabupaten/Kota"),
        "Provinsi": get_text("Provinsi"),
        "Status Sekolah": get_text("Status Sekolah"),

        # 🔹 FIELD TAMBAHAN
        "Akses Internet": get_text("Akses Internet"),
        "Nama Kepala Sekolah": get_text("Kepala Sekolah"),
        "Nama Operator": get_text("Operator"),
        "Jumlah Ruang Kelas": get_text("Ruang Kelas")
    }

def export_csv():
    npsn_list = text.get("1.0", tk.END).strip().splitlines()
    npsn_list = [n.strip() for n in npsn_list if n.strip().isdigit()]

    if not npsn_list:
        messagebox.showwarning("Kosong", "Masukkan NPSN terlebih dahulu")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV File", "*.csv")]
    )

    if not file_path:
        return

    hasil = []

    for idx, npsn in enumerate(npsn_list, start=1):
        status_label.config(text=f"Proses {idx}/{len(npsn_list)} : {npsn}")
        root.update()

        data = ambil_data(npsn)
        if data:
            hasil.append(data)

        time.sleep(1)  # rate limit aman

    if not hasil:
        messagebox.showerror("Gagal", "Tidak ada data yang berhasil diambil")
        return

    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=hasil[0].keys())
        writer.writeheader()
        writer.writerows(hasil)

    messagebox.showinfo("Selesai", "Data berhasil diexport ke CSV")
    status_label.config(text="Selesai")

# ================= GUI =================

root = tk.Tk()
root.title("Export Data Sekolah dari NPSN")
root.geometry("680x460")

tk.Label(
    root,
    text="Masukkan NPSN (1 baris = 1 NPSN)",
    font=("Arial", 11)
).pack(pady=5)

text = tk.Text(root, height=12, width=80)
text.pack()

tk.Button(
    root,
    text="Export ke CSV",
    command=export_csv,
    height=2
).pack(pady=10)

status_label = tk.Label(root, text="", fg="blue")
status_label.pack()

root.mainloop()
