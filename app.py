import io
import re
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Nasmoco Tax Automation & Coretax XML",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Nasmoco Tax Automation & Coretax XML Generator")
st.markdown(
    "Aplikasi otomatisasi rekapitulasi akun hutang pajak kepala `2145` dari GL"
    " Nasmoco, pencocokan database vendor unifikasi, dan konversi ke XML"
    " Coretax DJP."
)

# Sidebar untuk Upload File
st.sidebar.header("1. Upload Data Sumber")
gl_file = st.sidebar.file_uploader(
    "Upload File GL ALL (.xlsx)", type=["xlsx"]
)
db_file = st.sidebar.file_uploader(
    "Upload Database Kamus Vendor (.xlsx)", type=["xlsx"]
)


def load_database(db_uploaded):
  try:
    df_db = pd.read_excel(db_uploaded)
    return df_db
  except Exception as e:
    st.error(f"Gagal membaca file database: {e}")
    return None


def parse_gl_data(gl_uploaded):
  try:
    df_gl = pd.read_excel(gl_uploaded, sheet_name="ORI")
    return df_gl
  except Exception as e:
    st.error(f"Gagal membaca file GL: {e}")
    return None


# Fungsi Generator XML Coretax Unifikasi
def generate_coretax_xml(df_tax):
  root = ET.Element("BupotUnifikasi")
  root.set("xmlns", "http://pajak.go.id/coretax/bupotunifikasi")

  header = ET.SubElement(root, "Header")
  ET.SubElement(header, "TahunPajak").text = "2026"
  ET.SubElement(header, "MasaPajak").text = "09"
  ET.SubElement(header, "NPWPnPemberi").text = (
      "010000000000000"  # NPWP Cabang Nasmoco
  )

  list_bupot = ET.SubElement(root, "DaftarBuktiPotong")

  for idx, row in df_tax.iterrows():
    bupot = ET.SubElement(list_bupot, "BuktiPotong")
    ET.SubElement(bupot, "NomorBukti").text = str(row["No Bukti"])
    ET.SubElement(bupot, "Tanggal").text = (
        str(row["Tanggal"])[:10] if pd.notna(row["Tanggal"]) else ""
    )
    ET.SubElement(bupot, "NPWP_NIK").text = str(row["NPWP / NIK"])
    ET.SubElement(bupot, "NamaWp").text = str(row["Vendor Kamus"])
    ET.SubElement(bupot, "JenisPajak").text = str(row["Jenis Pajak"])
    ET.SubElement(bupot, "Tarif").text = str(row["Tarif"])
    ET.SubElement(bupot, "JumlahDPP").text = str(row["Kredit (DPP/Potongan)"])

  xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
  return xml_str


if gl_file and db_file:
  df_db = load_database(db_file)
  df_gl = parse_gl_data(gl_file)

  if df_db is not None and df_gl is not None:
    st.success(
        f"File GL berhasil dimuat & Database Kamus siap ({len(df_db)} vendor"
        " terdaftar)!"
    )

    if st.button("🚀 Proses Rekapitulasi & Matching Pajak"):
      with st.spinner(
          "Sedang memproses transaksi dan mencocokkan dengan kamus vendor..."
      ):
        trans_list = []

        # Deteksi transaksi baris per baris di GL
        for idx, row in df_gl.iterrows():
          val_no = row.iloc[0]
          val_date = row.iloc[1]
          col2_val = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""

          # Jika kolom 0 adalah nomor urut transaksi dan kolom 2 berawalan 2145
          if (
              pd.notna(val_no)
              and pd.notna(val_date)
              and col2_val.startswith("2145")
          ):
            akun_val = col2_val
            keterangan = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""

            # Matching vendor dengan database kamus
            matched_vendor = "TIDAK DITEMUKAN"
            matched_npwp = "-"
            matched_jenis = "-"
            matched_tarif = 0.0

            for _, db_row in df_db.iterrows():
              v_name = str(db_row["Nama Vendor"]).upper()
              if v_name in keterangan.upper():
                matched_vendor = db_row["Nama Vendor"]
                matched_npwp = str(db_row["NPWP"])
                matched_jenis = str(db_row["Jenis Pajak"])
                matched_tarif = float(db_row["Tarif Pajak"])
                break

            # Ekstraksi khusus NIK untuk PPh 21 Pihak Ketiga (21450200)
            extracted_nik = "-"
            if akun_val == "21450200":
              nik_match = re.search(r"==(\d{15,16})", keterangan)
              if nik_match:
                extracted_nik = nik_match.group(1)

            trans_list.append({
                "Akun": akun_val,
                "Tanggal": val_date,
                "No Bukti": row.iloc[3],
                "Keterangan GL": keterangan,
                "Vendor Kamus": matched_vendor,
                "NPWP / NIK": (
                    extracted_nik if extracted_nik != "-" else matched_npwp
                ),
                "Jenis Pajak": matched_jenis,
                "Tarif": matched_tarif,
                "Kredit (DPP/Potongan)": row.iloc[8],
            })

        df_tax = pd.DataFrame(trans_list)

        if not df_tax.empty:
          st.success(
              f"Berhasil menemukan {len(df_tax)} transaksi hutang pajak"
              " berawalan 2145!"
          )
          st.subheader("📋 Hasil Ekstraksi & Matching Pajak")
          st.dataframe(df_tax.head(50))

          # Ringkasan per Jenis Pajak / Akun
          summary = (
              df_tax.groupby(["Akun", "Jenis Pajak"])["Kredit (DPP/Potongan)"]
              .sum()
              .reset_index()
          )
          st.subheader("📊 Ringkasan Total per Jenis Pajak")
          st.dataframe(summary)

          # Tombol Download Excel Rekap
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_tax.to_excel(writer, sheet_name="Detail_Rekap_Pajak", index=False)
            summary.to_excel(writer, sheet_name="Summary", index=False)
          excel_data = output.getvalue()

          col1, col2 = st.columns(2)
          with col1:
            st.download_button(
                label="📥 Download File Rekap Excel Lengkap (.xlsx)",
                data=excel_data,
                file_name="Rekap_Pajak_Nasmoco_Matched.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

          # Tombol Download XML Coretax
          with col2:
            xml_data = generate_coretax_xml(df_tax)
            st.download_button(
                label="📄 Download File XML Coretax DJP (.xml)",
                data=xml_data,
                file_name="Coretax_Unifikasi_Nasmoco.xml",
                mime="application/xml",
            )
        else:
          st.warning(
              "Tidak ditemukan transaksi dengan akun berawalan 2145 pada baris"
              " transaksi."
          )
else:
  st.info(
      "👈 Silakan upload file **GL ALL.xlsx** dan **impor database"
      " 2610.xlsx** di panel sebelah kiri untuk memulai."
  )
