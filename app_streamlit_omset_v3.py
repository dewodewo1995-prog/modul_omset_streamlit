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


# --- AREA TAMPILAN UTAMA (TABS) ---
tab1, tab2 = st.tabs(
    ["🚀 Proses Baru (Upload Web)", "🔄 Update Excel Existing (Web)"]
)

with tab1:
  st.subheader("Proses Penggabungan & Template via Web")
  st.write(
      "Pilih metode upload data harian kamu di bawah ini, lalu sistem akan"
      " memasukkannya ke `template_omset.xlsx` dari repository."
  )

  metode_upload = st.radio(
      "Pilih Metode Upload Data Harian (Tab 1):",
      ["📁 Upload File ZIP", "📄 Upload Banyak File Excel (.xlsx/.xls)"],
      key="metode_tab1",
  )

  uploaded_zip = None
  uploaded_files = None

  if metode_upload == "📁 Upload File ZIP":
    uploaded_zip = st.file_uploader(
        "Upload file .zip yang berisi file-file harian:",
        type=["zip"],
        key="zip_tab1",
    )
  else:
    uploaded_files = st.file_uploader(
        "Upload file-file Excel harian (bisa pilih banyak sekaligus):",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="files_tab1",
    )

  if st.button("🚀 Jalankan Proses Penggabungan", type="primary"):
    with tempfile.TemporaryDirectory() as temp_dir:
      sumber_data_siap = False

      if metode_upload == "📁 Upload File ZIP" and uploaded_zip is not None:
        with st.spinner("Mengekstrak file ZIP..."):
          with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
          sumber_data_siap = True

      elif (
          metode_upload == "📄 Upload Banyak File Excel (.xlsx/.xls)"
          and uploaded_files
      ):
        with st.spinner("Menyiapkan file yang di-upload..."):
          for uf in uploaded_files:
            temp_file_path = os.path.join(temp_dir, uf.name)
            with open(temp_file_path, "wb") as f:
              f.write(uf.getbuffer())
          sumber_data_siap = True
      else:
        st.error("❌ Harap upload file ZIP atau file Excel terlebih dahulu!")

      if sumber_data_siap:
        with st.spinner("Sedang memproses data omset..."):
          df_gabungan, logs = proses_gabung_data_dari_folder(
              temp_dir, jumlah_tanggal, nama_bulan
          )

          for log in logs:
            st.text(log)

          if df_gabungan is not None:
            try:
              template_path = os.path.join(
                  os.path.dirname(__file__), "template_omset.xlsx"
              )

              if not os.path.exists(template_path):
                st.error(
                    "❌ ERROR: File `template_omset.xlsx` belum di-upload ke"
                    " repository GitHub!"
                )
              else:
                output_file_path = os.path.join(
                    temp_dir, "hasil_template_omset.xlsx"
                )
                sh.copy(template_path, output_file_path)

                with pd.ExcelWriter(
                    output_file_path,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                  df_gabungan.to_excel(
                      writer, sheet_name="statement", index=False
                  )

                st.success(
                    "🎉 Proses Selesai! File berhasil digabungkan ke template."
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
            except Exception as e:
              st.error(f"❌ Terjadi kesalahan saat memproses Excel: {e}")


with tab2:
  st.subheader("Update File Excel Target Tertentu via Web")
  st.write(
      "Upload file Excel target yang ingin kamu *update* isinya secara"
      " langsung beserta data harian sumbernya."
  )

  target_file_upload = st.file_uploader(
      "Upload File Excel Target (.xlsx/.xls):",
      type=["xlsx", "xls"],
      key="target_excel",
  )

  metode_upload_tab2 = st.radio(
      "Pilih Metode Upload Data Harian (Tab 2):",
      ["📁 Upload File ZIP", "📄 Upload Banyak File Excel (.xlsx/.xls)"],
      key="metode_tab2",
  )

  uploaded_zip_tab2 = None
  uploaded_files_tab2 = None

  if metode_upload_tab2 == "📁 Upload File"
