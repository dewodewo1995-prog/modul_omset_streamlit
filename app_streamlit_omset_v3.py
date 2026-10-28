import os
import shutil as sh
import pandas as pd
import streamlit as st

# --- COBA IMPORT TKINTER (Hanya jalan di komputer lokal) ---
try:
  import tkinter as tk
  from tkinter import filedialog

  HAS_TKINTER = True
except ImportError:
  HAS_TKINTER = False

# Konfigurasi halaman web Streamlit
st.set_page_config(
    page_title="Rekap Omset Web", page_icon="📊", layout="centered"
)

st.title("📊 Aplikasi Rekap Omset (Versi Web)")
st.write(
    "Upload file-file Excel harian atau pilih direktori/file langsung dari"
    " sistem kamu!"
)


# --- FUNGSI HELPER UNTUK MEMANGGIL WINDOWS DIALOG ---
def pilih_folder_lokal():
  if not HAS_TKINTER:
    return ""
  root = tk.Tk()
  root.withdraw()
  root.attributes("-topmost", True)
  folder_path = filedialog.askdirectory(title="Pilih Folder Data Omset")
  root.destroy()
  return folder_path


def pilih_file_target():
  if not HAS_TKINTER:
    return ""
  root = tk.Tk()
  root.withdraw()
  root.attributes("-topmost", True)
  file_path = filedialog.askopenfilename(
      title="Pilih File Excel Target",
      filetypes=[("Excel Files", "*.xlsx *.xls")],
  )
  root.destroy()
  return file_path


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

st.sidebar.markdown("---")

# 3. Bagian Pilih Folder dengan Tombol Dialog
st.sidebar.subheader("📁 Direktori Folder Sumber")

if "path_folder" not in st.session_state:
  st.session_state.path_folder = ""

if HAS_TKINTER:
  if st.sidebar.button("📂 Pilih Folder via Dialog"):
    selected_dir = pilih_folder_lokal()
    if selected_dir:
      st.session_state.path_folder = selected_dir
else:
  st.sidebar.info(
      "💡 Mode Web Cloud: Ketik path folder sumber data harian di bawah."
  )

path_folder = st.sidebar.text_input(
    "Path Folder Lokal:",
    value=st.session_state.path_folder,
    placeholder="Contoh: D:/DataOmset",
)

st.session_state.path_folder = path_folder

st.sidebar.caption(
    "Pastikan file harian (1_jan.xlsx, dst) berada di folder sumber ini."
)


# --- FUNGSI UTAMA PENGGABUNGAN ---
def proses_gabung_data(path, jml_tgl, bln):
  df_list = []
  log_messages = []

  for i in range(1, jml_tgl + 1):
    file_base = os.path.join(path, f"{i}_{bln}")
    file_path_xlsx = file_base + ".xlsx"
    file_path_xls = file_base + ".xls"

    if os.path.exists(file_path_xlsx):
      df = pd.read_excel(file_path_xlsx, sheet_name="Sheet1", header=0)
    elif os.path.exists(file_path_xls):
      df = pd.read_excel(file_path_xls, sheet_name="Sheet1", header=0)
    else:
      log_messages.append(f"❌ File Excel tanggal {i}_{bln} tidak ditemukan.")
      continue

    df["tanggal"] = i
    df_list.append(df)
    log_messages.append(f"✅ Berhasil memproses tanggal: {i}")

  if not df_list:
    return None, log_messages

  df_gabungan = pd.concat(df_list, axis=0, ignore_index=True)
  return df_gabungan, log_messages


# --- AREA TAMPILAN UTAMA (TABS) ---
tab1, tab2 = st.tabs(["🚀 Proses Baru (Template Repo)", "🔄 Update Excel Existing"])

with tab1:
  st.subheader("Proses Penggabungan & Template dari Repository")
  st.write(
      "Sistem akan membaca file harian dari folder lokal, memasukkannya ke"
      " `template_omset.xlsx` bawaan repo, lalu menyediakan tombol download"
      " hasilnya."
  )

  if st.button("🚀 Jalankan Proses Baru", type="primary"):
    if not path_folder or not os.path.exists(path_folder):
      st.error("❌ ERROR: Path folder sumber data tidak valid atau kosong!")
    else:
      with st.spinner("Sedang memproses data..."):
        df_gabungan, logs = proses_gabung_data(
            path_folder, jumlah_tanggal, nama_bulan
        )

        for log in logs:
          st.text(log)

        if df_gabungan is not None:
          try:
            # Mengambil template_omset.xlsx langsung dari direktori repo (tempat file python berada)
            template_path = os.path.join(
                os.path.dirname(__file__), "template_omset.xlsx"
            )

            if not os.path.exists(template_path):
              st.error(
                  "❌ ERROR: File `template_omset.xlsx` tidak ditemukan di"
                  " dalam repository GitHub!"
              )
            else:
              # Buat salinan template di memory/temp supaya file asli di repo aman
              output_file_path = os.path.join(
                  path_folder, "hasil_template_omset.xlsx"
              )
              sh.copy(template_path, output_file_path)

              # Masukkan data gabungan ke file salinan tersebut
              with pd.ExcelWriter(
                  output_file_path,
                  engine="openpyxl",
                  mode="a",
                  if_sheet_exists="replace",
              ) as writer:
                df_gabungan.to_excel(writer, sheet_name="statement", index=False)

              st.success(
                  "🎉 Proses Selesai! File berhasil digabungkan ke template"
                  " repository."
              )

              with open(output_file_path, "rb") as f:
                st.download_button(
                    label="📥 Download Hasil Template Excel",
                    data=f,
                    file_name="hasil_template_omset.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )

          except PermissionError:
            st.error(
                "❌ ERROR: File Excel sedang dibuka di aplikasi lain (Excel),"
                " tutup dulu filenya!"
            )


with tab2:
  st.subheader("Update File Excel Target Tertentu")
  st.write(
      "Gunakan fitur ini jika kamu ingin langsung memperbarui file Excel rekap"
      " tujuan yang berbeda."
  )

  if "target_file_path" not in st.session_state:
    st.session_state.target_file_path = ""

  if HAS_TKINTER:
    if st.button("📂 Pilih File Excel Target via Dialog"):
      selected_file = pilih_file_target()
      if selected_file:
        st.session_state.target_file_path = selected_file

  target_file_path = st.text_input(
      "Masukkan Path Lengkap File Excel Target untuk Di-update:",
      value=st.session_state.target_file_path,
      placeholder="Contoh: C:/Users/Namamu/Documents/rekap_akhir.xlsx",
  )

  st.session_state.target_file_path = target_file_path

  if st.button("🔄 Jalankan Update Excel"):
    if not path_folder or not os.path.exists(path_folder):
      st.error("❌ ERROR: Path folder sumber data tidak valid!")
    elif not target_file_path or not os.path.exists(target_file_path):
      st.error("❌ ERROR: File Excel target tidak ditemukan!")
    else:
      with st.spinner("Sedang memperbarui file Excel..."):
        df_gabungan, logs = proses_gabung_data(
            path_folder, jumlah_tanggal, nama_bulan
        )

        for log in logs:
          st.text(log)

        if df_gabungan is not None:
          try:
            with pd.ExcelWriter(
                target_file_path,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
              df_gabungan.to_excel(writer, sheet_name="statement", index=False)

            st.success(
                "🎉 Berhasil memperbarui file Excel target di:"
                f" `{target_file_path}`"
            )

            with open(target_file_path, "rb") as f:
              st.download_button(
                  label="📥 Download File yang Telah Di-update",
                  data=f,
                  file_name=os.path.basename(target_file_path),
                  mime=(
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  ),
              )

          except PermissionError:
            st.error(
                "❌ ERROR: File target sedang dibuka di Excel. Tutup filenya"
                " terlebih dahulu!"
            )