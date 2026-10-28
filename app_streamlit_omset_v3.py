import os
import shutil as sh
import tempfile
import zipfile
import openpyxl
import pandas as pd
import streamlit as st

# Konfigurasi halaman web Streamlit
st.set_page_config(
    page_title="Rekap Omset Web", page_icon="📊", layout="wide"
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
    3. **Menu Proses Baru**: Pilih tab *Proses Baru (Upload Web)* jika Anda ingin membuat rekap omset dari awal menggunakan file master template dari sistem. Kolom tanggal (`DATE` mulai dari sel C8) akan ter-generate otomatis sesuai bulan dan tahun pilihan Anda.
    4. **Menu Update Excel Existing**: Pilih tab *Update Excel Existing (Web)* jika Anda ingin memperbarui file rekap omset yang sudah pernah dibuat sebelumnya.
    5. **Eksekusi Proses**: Klik tombol **Jalankan Proses Penggabungan** pada menu *Proses Baru*, atau klik **Jalankan Update Excel Target** pada menu *Update Existing* untuk memproses dan mendownload hasilnya.
    """)

st.markdown("---")


# --- PARAMETER PENGATURAN UTAMA ---
with st.container(border=True):
  st.subheader("⚙️ Parameter Pengaturan Data")
  col1, col2, col3 = st.columns(3)

  with col1:
    jumlah_tanggal = st.selectbox(
        "Pilih Jumlah Tanggal (1 - 31):", options=list(range(1, 32)), index=0
    )

  with col2:
    bulan_dict = {
        "Januari": ("jan", 1),
        "Februari": ("feb", 2),
        "Maret": ("mar", 3),
        "April": ("apr", 4),
        "Mei": ("mei", 5),
        "Juni": ("jun", 6),
        "Juli": ("jul", 7),
        "Agustus": ("agu", 8),
        "September": ("sep", 9),
        "Oktober": ("okt", 10),
        "November": ("nov", 11),
        "Desember": ("des", 12),
    }
    pilihan_bulan_nama = st.selectbox(
        "Pilih Nama Bulan:", options=list(bulan_dict.keys())
    )
    nama_bulan, angka_bulan = bulan_dict[pilihan_bulan_nama]

  with col3:
    pilihan_tahun = st.selectbox(
        "Pilih Tahun:", options=[2026, 2027, 2028, 2029, 2030], index=0
    )

st.markdown("---")


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
      " memasukkannya ke `template_omset.xlsx` dan otomatis mengisi kolom DATE"
      " dari sel C8."
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
        with st.spinner("Sedang memproses data omset & update kolom DATE..."):
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

                # --- LOAD OPENPYXL & TARGETKAN SHEET ANALISA (KOLOM C, BARIS 8) ---
                wb = openpyxl.load_workbook(output_file_path)
                if "ANALISA" in wb.sheetnames:
                  ws = wb["ANALISA"]
                else:
                  ws = wb.active  # Fallback ke sheet aktif jika nama ANALISA berbeda

                # List singkatan bulan bahasa Inggris
                eng_months = [
                    "",
                    "Jan",
                    "Feb",
                    "Mar",
                    "Apr",
                    "May",
                    "Jun",
                    "Jul",
                    "Aug",
                    "Sep",
                    "Oct",
                    "Nov",
                    "Dec",
                ]
                m_str = eng_months[angka_bulan]

                start_row = 8  # Mulai dari baris 8
                col_index = 3  # Kolom C adalah kolom ke-3

                for d in range(1, jumlah_tanggal + 1):
                  date_str = f"{d:02d}-{m_str}-{pilihan_tahun}"
                  # Menulis langsung ke Kolom C pada baris ke-(8 + d - 1)
                  ws.cell(row=start_row + d - 1, column=col_index, value=date_str)

                # Simpan perubahan openpyxl ke file output
                wb.save(output_file_path)

                # Masukkan sheet statement menggunakan pandas
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
                    "🎉 Proses Selesai! Kolom DATE (C8 ke bawah) berhasil"
                    f" di-overwrite ke bulan {pilihan_bulan_nama}"
                    f" {pilihan_tahun}."
                )

                with open(output_file_path, "rb") as f:
                  st.download_button(
                      label="📥 Download Hasil Template Excel",
                      data=f,
                      file_name=f"rekap_omset_{pilihan_bulan_nama}_{pilihan_tahun}.xlsx",
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

  if metode_upload_tab2 == "📁 Upload File ZIP":
    uploaded_zip_tab2 = st.file_uploader(
        "Upload file .zip data harian untuk update target:",
        type=["zip"],
        key="zip_tab2",
    )
  else:
    uploaded_files_tab2 = st.file_uploader(
        "Upload file-file Excel harian untuk update target:",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="files_tab2",
    )

  if st.button("🔄 Jalankan Update Excel Target", type="primary"):
    if not target_file_upload:
      st.error("❌ Harap upload file Excel target terlebih dahulu!")
    else:
      with tempfile.TemporaryDirectory() as temp_dir:
        sumber_data_siap = False

        if (
            metode_upload_tab2 == "📁 Upload File ZIP"
            and uploaded_zip_tab2 is not None
        ):
          with st.spinner("Mengekstrak file ZIP..."):
            with zipfile.ZipFile(uploaded_zip_tab2, "r") as zip_ref:
              zip_ref.extractall(temp_dir)
            sumber_data_siap = True

        elif (
            metode_upload_tab2 == "📄 Upload Banyak File Excel (.xlsx/.xls)"
            and uploaded_files_tab2
        ):
          with st.spinner("Menyiapkan file yang di-upload..."):
            for uf in uploaded_files_tab2:
              temp_file_path = os.path.join(temp_dir, uf.name)
              with open(temp_file_path, "wb") as f:
                f.write(uf.getbuffer())
            sumber_data_siap = True
        else:
          st.error(
              "❌ Harap upload file ZIP atau file Excel sumber harian terlebih"
              " dahulu!"
          )

        if sumber_data_siap:
          with st.spinner("Sedang memperbarui file Excel target..."):
            df_gabungan, logs = proses_gabung_data_dari_folder(
                temp_dir, jumlah_tanggal, nama_bulan
            )

            for log in logs:
              st.text(log)

            if df_gabungan is not None:
              try:
                target_path_local = os.path.join(
                    temp_dir, target_file_upload.name
                )
                with open(target_path_local, "wb") as f:
                  f.write(target_file_upload.getbuffer())

                with pd.ExcelWriter(
                    target_path_local,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                  df_gabungan.to_excel(
                      writer, sheet_name="statement", index=False
                  )

                st.success("🎉 Berhasil memperbarui file Excel target!")

                with open(target_path_local, "rb") as f:
                  st.download_button(
                      label="📥 Download File Target yang Telah Di-update",
                      data=f,
                      file_name=target_file_upload.name,
                      mime=(
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      ),
                  )
              except Exception as e:
                st.error(
                    f"❌ Terjadi kesalahan saat memperbarui file target: {e}"
                )