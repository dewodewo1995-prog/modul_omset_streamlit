import os
import shutil as sh
import tempfile
import zipfile
import pandas as pd
import streamlit as st

# Konfigurasi halaman web Streamlit
st.set_page_config(
    page_title="Rekap Omset Web", page_icon="📊", layout="centered"
)

st.title("📊 Aplikasi Rekap Omset (Versi Web Cloud)")
st.write(
    "Upload file ZIP atau file-file Excel harian langsung melalui browser!"
)


# --- TATA CARA CEPAT PENGGUNAAN (EXPANDER) ---
with st.expander("📖 Tata Cara Cepat Penggunaan (Klik untuk Membaca)", expanded=False):
  st.markdown("""
    1. **Format Penamaan File Harian**: Pastikan file Excel harian menggunakan format penamaan `tanggal_3hurufbulan` (Contoh: `1_okt.xlsx`, `2_okt.xlsx`, dst).
    2. **Metode Upload Data**: 
       * Pilih metode **Upload Banyak File Excel** jika ingin memilih dan meng-upload file Excel harian satu per satu sekaligus.
       * Atau pilih metode **Upload File ZIP** jika file harian sudah dikompres menjadi satu file ZIP/RAR.
    3. **Menu Proses Baru**: Pilih tab *Proses Baru (Upload Web)* jika Anda ingin membuat rekap omset dari awal menggunakan file master template dari sistem.
    4. **Menu Update Excel Existing**: Pilih tab *Update Excel Existing (Web)* jika Anda ingin memperbarui file rekap omset yang sudah pernah dibuat sebelumnya. Jangan lupa lakukan upload data harian atau file ZIP/RAR pendukungnya.
    5. **Eksekusi Proses**: Klik tombol **Jalankan Proses Penggabungan** pada menu *Proses Baru*, atau klik **Jalankan Update Excel Target** pada menu *Update Existing* untuk memproses dan mendownload hasilnya.
    """)

st.markdown("---")


# --- SIDEBAR / MENU UTAMA ---
st.sidebar.header("⚙️ Parameter Pengaturan")

# 1. Pilihan Jumlah Tanggal (1 - 31)
jumlah_tanggal = st.sidebar.selectbox(
    "Pilih Jumlah Tanggal (1 - 31):", options=list(range(1, 32)), index=0
)

# 2. Pilihan Bulan
bulan_dict = {
    "Januari": "jan",
    "Februari": "feb",
    "Maret": "mar",
    "April": "apr",
    "Mei": "mei",
    "Juni": "jun",
    "Juli": "jul",
    "Agustus": "agu",
    "September": "sep",
    "Oktober": "okt",
    "November": "nov",
    "Desember": "des",
}
pilihan_bulan_nama = st.sidebar.selectbox(
    "Pilih Nama Bulan:", options=list(bulan_dict.keys())
)
nama_bulan = bulan_dict[pilihan_bulan_nama]


# --- FUNGSI PROSES DATA DARI DIREKTORI SEMENTARA ---
def proses_gabung_data_dari_folder(source_dir, jml_tgl, bln):
  df_list = []
  log_messages = []

  for i in range(1, jml_tgl + 1):
    file_base = os.path.join(source_dir, f"{i}_{bln}")
    file_path_xlsx = file_base + ".xlsx"
    file_path_xls = file_base + ".xls"

    target_file = None
    if os.path.exists(file_path_xlsx):
      target_file = file_path_xlsx
    elif os.path.exists(file_path_xls):
      target_file = file_path_xls

    if target_file:
      try:
        df = pd.read_excel(target_file, sheet_name="Sheet1", header=0)
        df["tanggal"] = i
        df_list.append(df)
        log_messages.append(f"✅ Berhasil memproses tanggal: {i}")
      except Exception as e:
        log_messages.append(f"❌ Error membaca file tanggal {i}: {e}")
    else:
      log_messages.append(f"❌ File Excel tanggal {i}_{bln} tidak ditemukan.")

  if not df_list:
    return None, log_messages

  df_gabungan = pd.concat(df_list, axis=0, ignore_index=True)
  return df_gabungan, log_messages


# --- AREA TAMPILAN UT