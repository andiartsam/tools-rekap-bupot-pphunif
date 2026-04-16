print("Proses dimulai...")
print("Import libraries...")

import pdfplumber
import pandas as pd
import re
from pathlib import Path
from tkinter import Tk, filedialog
from tqdm import tqdm

print("Membuka Jendela...")
# ===== PATH =====
Tk().withdraw()  # biar jendela utama tkinter nggak muncul

folder_path = filedialog.askdirectory(
    title="Pilih folder berisi PDF Bukti Potong"
)

if not folder_path:
    raise SystemExit("Folder tidak dipilih. Proses dibatalkan.")

FOLDER_PDF = Path(folder_path)
OUTPUT_CSV = "rekap_bupot_2025.csv"
LOG_ERROR = "error_bupot.log"

print(f"Memproses file PDF di folder: {FOLDER_PDF}")

# ===== EXTRACT DATA =====
data = []
error_files = []

def clean_number(text):
    return int(text.replace(".", "").strip())

# Hitung total PDF terlebih dahulu untuk progress bar yang akurat
# rglob mencari file PDF di folder dan semua subfoldernya
pdf_files = sorted(FOLDER_PDF.rglob("*.pdf"))
total_files = len(pdf_files)

for pdf_file in tqdm(pdf_files, desc="Memproses PDF", total=total_files):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = pdf.pages[0].extract_text()

        # Extract nomor BUPOT dari baris ke-6 (misal: "2502ALC3U")
        lines = text.split('\n')
        nomor_bupot = lines[6].split()[0] if len(lines) > 6 else None
        
        # Extract masa pajak (misal: "06-2025")
        masa_pajak_match = re.search(r"(\d{2}-\d{4})", lines[6]) if len(lines) > 6 else None
        masa_pajak = masa_pajak_match.group(1) if masa_pajak_match else None

        # Cari baris yang mengandung kode objek (format XX-XXX-XX)
        kode_objek = None
        dpp = None
        pph = None
        
        for line in lines:
            kode_match = re.search(r"(\d{2}-\d{3}-\d{2})\s+(.+?)\s+(\d+(?:\.?\d+)*)\s+(\d+(?:\.?\d+)*)\s+(\d+(?:\.?\d+)*)", line)
            if kode_match:
                kode_objek = kode_match.group(1)
                dpp_str = kode_match.group(3)  # Kolom B.5 (DPP)
                tarif_str = kode_match.group(4)  # Kolom B.6 (Tarif)
                pph_str = kode_match.group(5)  # Kolom B.7 (PPh)
                
                # Konversi ke angka (hilangkan titik pemisah ribuan)
                dpp = clean_number(dpp_str)
                pph = clean_number(pph_str)
                break

        # Extract NPWP Pemotong (misal: "0959157934512000")
        npwp_match = re.search(r"C\.1 NPWP.*?:\s*(\d+)", text)
        npwp_pemotong = npwp_match.group(1) if npwp_match else None
        
        # Extract Nama Pemotong (misal: "BADAN PENDAPATAN DAERAH KOTA SEMARANG")
        nama_match = re.search(r"C\.3 NAMA PEMOTONG.*?:\s*(.+?)(?:\n|$)", text)
        nama_pemotong = nama_match.group(1).strip() if nama_match else None

        # Extract Tanggal BUPOT dan Dokumen (misal: "04 Juni 2025")
        tanggal_match = re.search(r"Jenis Dokumen.*?Tanggal\s*:\s*(.+?)(?:\n|$)", text)
        tanggal_dokumen = tanggal_match.group(1).strip() if tanggal_match else None
        
        tanggal_dok_match = re.search(r"C\.4 TANGGAL\s*:\s*(.+?)(?:\n|$)", text)
        tanggal_bupot = tanggal_dok_match.group(1).strip() if tanggal_dok_match else None

        data.append({
            "nomor_bupot": nomor_bupot,
            "Masa_Pajak": masa_pajak,
            "Kode_objek_pajak": kode_objek,
            "DPP": dpp,
            "Pajak_Penghasilan": pph,
            "NPWP_Pemotong": npwp_pemotong,
            "Nama_Pemotong": nama_pemotong,
            "tanggal_dokumen": tanggal_dokumen,
            "Tanggal_bupot": tanggal_bupot
        })

    except Exception as e:
        error_files.append((pdf_file.name, str(e)))
        print(f"ERROR di {pdf_file.name}: {e}")



print(f"\nJumlah file BERHASIL: {len(data)}")
print(f"Jumlah file ERROR: {len(error_files)}")

# ===== EXPORT CSV =====
if len(data) > 0:
    df = pd.DataFrame(data)
    df.to_csv(OUTPUT_CSV, sep=";", index=False, encoding="utf-8")
    print(f"CSV berhasil dibuat: {OUTPUT_CSV}")
    print(f"Total baris: {len(df)}")
else:
    print("ERROR: Tidak ada data yang berhasil diextract!")

# ===== LOG ERROR =====
if len(error_files) > 0:
    with open(LOG_ERROR, "w") as f:
        for filename, error_msg in error_files:
            f.write(f"{filename}\n")
            f.write(f"  Error: {error_msg}\n\n")
    print(f"Error log disimpan: {LOG_ERROR}")

print("\nSelesai. CSV terbentuk & error dicatat.")
