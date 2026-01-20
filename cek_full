import requests
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import time

def ambil_data(npsn):
    api = f"https://dapo.kemendikdasmen.go.id/api/getSekolah/{npsn}"
    headers = {"User-Agent": "Mozilla/5.0"}

    r = requests.get(api, headers=headers, timeout=10)
    if r.status_code != 200:
        return None

    j = r.json()

    return {
        "NPSN": npsn,
        "Nama Sekolah": j.get("nama", "-"),
        "Alamat": j.get("alamat", "-"),
        "Desa/Kelurahan": j.get("desa_kelurahan", "-"),
        "Kecamatan": j.get("kecamatan", "-"),
        "Kabupaten/Kota": j.get("kabupaten_kota", "-"),
        "Provinsi": j.get("provinsi", "-"),
        "Status Sekolah": j.get("status_sekolah", "-"),
        "Akses Internet": j.get("akses_internet", "-"),
        "Nama Kepala Sekolah": j.get("kepala_sekolah", "-"),
        "Nama Operator": j.get("operator", "-"),
        "Jumlah Ruang Kelas": j.get("ruang_kelas", "-")
    }

def export_csv():
    npsn_list = text.get("1.0", tk.END).strip().splitlines()
    npsn_list = [n.strip() for n in npsn_list if n.strip().isdigit()]

    if not npsn_list:
        messagebox.showwarning("Kosong", "Masukkan NPSN")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")]
    )
    if not path:
        return

    hasil = []

    for i, npsn in enumerate(npsn_list, 1):
        status.config(text=f"Proses {i}/{len(npsn_list)} : {npsn}")
        root.update()

        data = ambil_data(npsn)
        if data:
            hasil.append(data)

        time.sleep(0.5)

    if not hasil:
        messagebox.showerror("Gagal", "Semua data kosong")
        return

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=hasil[0].keys())
        writer.writeheader()
        writer.writerows(hasil)

    messagebox.showinfo("Selesai", "CSV berhasil dibuat")
    status.config(text="Selesai")

# GUI
root = tk.Tk()
root.title("Export Data Sekolah dari NPSN (API)")
root.geometry("680x420")

tk.Label(root, text="Masukkan NPSN (1 baris 1 NPSN)").pack(pady=5)

text = tk.Text(root, height=12, width=80)
text.pack()

tk.Button(root, text="Export CSV", command=export_csv).pack(pady=10)

status = tk.Label(root, text="", fg="blue")
status.pack()

root.mainloop()
