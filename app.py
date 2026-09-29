from datetime import datetime, timedelta
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Deteksi Pelayanan Lab Tanpa Kunjungan FKTP",
    page_icon="🔬",
    layout="wide",
)

st.title("🔬 Deteksi Otomatis Kunjungan Lab Tanpa Pengantar Klinik/Puskesmas")
st.caption(
    "Analisis otomatis dari 1 file rekap: Mendeteksi peserta yang langsung ke Faskes Lab tanpa ada kunjungan FKTP sebelumnya."
)

# --- SIDEBAR PENGATURAN ---
with st.sidebar:
    st.header("⚙️ Pengaturan Parameter")

    # Kata kunci untuk mengenali faskes laboratorium
    keyword_lab_input = st.text_area(
        "Kata Kunci Nama Faskes Laboratorium (pisahkan dengan koma):",
        value="lab, laboratorium, prodia, kimia farma, pramita, cito",
        help="Sistem akan mendeteksi faskes sebagai Laboratorium jika namanya mengandung salah satu kata kunci ini.",
    )
    keywords_lab = [k.strip().lower() for k in keyword_lab_input.split(",")]

    toleransi_hari = st.number_input(
        "Maksimal Rentang Kunjungan FKTP Sebelumnya (Hari):",
        min_value=0,
        max_value=90,
        value=30,
        help="Misal 30 hari: Sistem akan mencari apakah ada kunjungan ke Klinik/Puskesmas dalam 30 hari mundur sebelum tanggal periksa lab.",
    )

    st.markdown("---")
    st.info(
        "📁 **Format Kolom Wajib di Excel:**\n- `no_kartu`\n- `tgl_kunjungan`\n- `nama_faskes`\n- `no_kunjungan`"
    )

# --- AREA UPLOAD FILE ---
file_upload = st.file_uploader(
    "Unggah File Data Kunjungan (.xlsx / .csv):", type=["xlsx", "xls", "csv"]
)


def is_laboratorium(nama_faskes, keywords):
    nama_str = str(nama_faskes).lower()
    return any(k in nama_str for k in keywords if k)


if file_upload is not None:
    try:
        # Load Data
        with st.spinner("Membaca file data..."):
            if file_upload.name.endswith(".csv"):
                df = pd.read_csv(
                    file_upload, dtype={"no_kartu": str, "no_kunjungan": str}
                )
            else:
                df = pd.read_excel(
                    file_upload, dtype={"no_kartu": str, "no_kunjungan": str}
                )

            # Standarisasi kolom
            df.columns = [c.strip().lower() for c in df.columns]

            kolom_wajib = {"no_kartu", "tgl_kunjungan", "nama_faskes", "no_kunjungan"}
            if not kolom_wajib.issubset(set(df.columns)):
                st.error(
                    f"Kolom file tidak sesuai. Wajib memiliki kolom: {kolom_wajib}"
                )
                st.stop()

            # Bersihkan tipe data
            df["no_kartu"] = df["no_kartu"].astype(str).str.strip()
            df["tgl_kunjungan_dt"] = pd.to_datetime(
                df["tgl_kunjungan"], errors="coerce"
            ).dt.date
            df = df.dropna(subset=["tgl_kunjungan_dt", "no_kartu"])

            # Tandai tipe faskes (Lab vs Non-Lab/FKTP)
            df["is_lab"] = df["nama_faskes"].apply(
                lambda x: is_laboratorium(x, keywords_lab)
            )

        # Pisahkan dataset
        df_lab = df[df["is_lab"]].copy()
        df_fktp = df[~df["is_lab"]].copy()

        st.success(
            f"✅ File berhasil dimuat! Total kunjungan: **{len(df):,}** baris (Kunjungan Lab: **{len(df_lab):,}**, Kunjungan FKTP/Klinik/Puskesmas: **{len(df_fktp):,}**)"
        )

        if df_lab.empty:
            st.warning(
                "Tidak ditemukan kunjungan dengan nama faskes Laboratorium berdasarkan kata kunci di sidebar. Coba sesuaikan kata kuncinya."
            )
            st.stop()

        # --- PROSES DETEKSI OTOMATIS ---
        with st.spinner("Menganalisis riwayat kunjungan lab ke FKTP..."):
            hasil_analisis = []

            for _, row_lab in df_lab.iterrows():
                no_kartu = row_lab["no_kartu"]
                tgl_lab = row_lab["tgl_kunjungan_dt"]
                batas_awal = tgl_lab - timedelta(days=toleransi_hari)

                # Cari kunjungan FKTP peserta ini pada rentang [Tgl_Lab - N hari s/d Tgl_Lab]
                fktp_peserta = df_fktp[
                    (df_fktp["no_kartu"] == no_kartu)
                    & (df_fktp["tgl_kunjungan_dt"] >= batas_awal)
                    & (df_fktp["tgl_kunjungan_dt"] <= tgl_lab)
                ]

                if fktp_peserta.empty:
                    # Cek apakah pernah ada kunjungan FKTP di luar rentang itu (riwayat lama)
                    fktp_lama = df_fktp[
                        (df_fktp["no_kartu"] == no_kartu)
                        & (df_fktp["tgl_kunjungan_dt"] < batas_awal)
                    ]
                    if not fktp_lama.empty:
                        last_fktp = fktp_lama.sort_values(
                            by="tgl_kunjungan_dt", ascending=False
                        ).iloc[0]
                        ket = f"Ada kunjungan lama ({last_fktp['tgl_kunjungan_dt']} di {last_fktp['nama_faskes']}), tapi melebihi {toleransi_hari} hari."
                    else:
                        ket = "Sama sekali tidak ada riwayat kunjungan ke Klinik/Puskesmas."

                    hasil_analisis.append({
                        "No Kartu": no_kartu,
                        "No Kunjungan Lab": row_lab["no_kunjungan"],
                        "Nama Faskes Lab": row_lab["nama_faskes"],
                        "Tanggal Lab": tgl_lab,
                        "Status": "⚠️ ANOMALI (Tanpa Pengantar)",
                        "Keterangan": ket,
                        "FKTP Pengantar": "-",
                        "Tanggal Kunjungan FKTP": "-",
                        "No Kunjungan FKTP": "-",
                    })
                else:
                    # Ambil kunjungan FKTP paling mendekati hari lab
                    fktp_terpilih = fktp_peserta.sort_values(
                        by="tgl_kunjungan_dt", ascending=False
                    ).iloc[0]
                    hasil_analisis.append({
                        "No Kartu": no_kartu,
                        "No Kunjungan Lab": row_lab["no_kunjungan"],
                        "Nama Faskes Lab": row_lab["nama_faskes"],
                        "Tanggal Lab": tgl_lab,
                        "Status": "✅ VALID (Ada FKTP)",
                        "Keterangan": f"Ditemukan kunjungan pengantar di {fktp_terpilih['nama_faskes']}.",
                        "FKTP Pengantar": fktp_terpilih["nama_faskes"],
                        "Tanggal Kunjungan FKTP": fktp_terpilih[
                            "tgl_kunjungan_dt"
                        ],
                        "No Kunjungan FKTP": fktp_terpilih["no_kunjungan"],
                    })

            df_hasil = pd.DataFrame(hasil_analisis)

        # --- RINGKASAN METRIK ---
        total_lab = len(df_hasil)
        anomali_cnt = (
            df_hasil["Status"] == "⚠️ ANOMALI (Tanpa Pengantar)"
        ).sum()
        valid_cnt = total_lab - anomali_cnt

        st.markdown("### 📊 Ringkasan Hasil Deteksi")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Pemeriksaan Lab", f"{total_lab:,}")
        c2.metric("Valid (Ada FKTP Pengantar)", f"{valid_cnt:,}")
        c3.metric(
            "Anomali (Tanpa Kunjungan FKTP)",
            f"{anomali_cnt:,}",
            delta=(
                f"{(anomali_cnt/total_lab)*100:.1f}% Temuan"
                if total_lab > 0
                else "0%"
            ),
            delta_color="inverse",
        )

        st.markdown("---")

        # --- TABEL HASIL TEMUAN ---
        tab_anomali, tab_semua = st.tabs(
            [f"🚨 Daftar Temuan Anomali ({anomali_cnt})", "📋 Semua Data Lab"]
        )

        with tab_anomali:
            df_temuan_only = df_hasil[
                df_hasil["Status"] == "⚠️ ANOMALI (Tanpa Pengantar)"
            ]
            if not df_temuan_only.empty:
                st.dataframe(df_temuan_only, use_container_width=True)
            else:
                st.success(
                    "🎉 Semua kunjungan Lab memiliki riwayat kunjungan ke Klinik/Puskesmas!"
                )

        with tab_semua:
            st.dataframe(df_hasil, use_container_width=True)

        # --- DOWNLOAD EXCEL HASIL AUDIT ---
        output_file = "Laporan_Anomali_Lab_Tanpa_FKTP.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_hasil.to_excel(writer, index=False, sheet_name="Hasil_Deteksi")

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Unduh Laporan Hasil Deteksi Lengkap (.xlsx)",
                data=f,
                file_name=f"Laporan_Audit_Lab_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
