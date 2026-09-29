import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Deteksi Anomali Pelayanan Laboratorium",
    page_icon="🩺",
    layout="wide"
)

# Header Aplikasi
st.title("🩺 Detektor Anomali Pelayanan Laboratorium Tanpa Kunjungan")
st.caption("Sistem Verifikasi & Deteksi Kepatuhan Pelayanan Lab Faskes")

# --- SIDEBAR: PENGATURAN DATA & PARAMETER ---
with st.sidebar:
    st.header("⚙️ 1. Unggah Master Data Kunjungan")
    file_kunjungan = st.file_uploader(
        "Upload Rekap Kunjungan Faskes (.xlsx / .csv)", 
        type=["xlsx", "xls", "csv"],
        help="Wajib berisi kolom: no_kartu, tgl_kunjungan, kode_faskes, nama_faskes, no_kunjungan"
    )
    
    st.markdown("---")
    st.header("⏱️ 2. Parameter Toleransi")
    toleransi_hari = st.number_input(
        "Toleransi Mundur Kunjungan (Hari)",
        min_value=0,
        max_value=30,
        value=1,
        help="0 = Harus di hari yang sama. 1 = Boleh ada kunjungan H-1 sebelum lab."
    )
    
    st.info("💡 **Format Kolom Wajib (Kunjungan):**\n- `no_kartu`\n- `tgl_kunjungan`\n- `nama_faskes`\n- `no_kunjungan`")

# Fungsi load & normalisasi data kunjungan (Cache agar cepat)
@st.cache_data
def load_data_kunjungan(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, dtype={'no_kartu': str, 'no_kunjungan': str})
    else:
        df = pd.read_excel(file, dtype={'no_kartu': str, 'no_kunjungan': str})
    
    # Standarisasi nama kolom ke huruf kecil
    df.columns = [c.strip().lower() for c in df.columns]
    
    # Konversi kolom tanggal ke format Date
    if 'tgl_kunjungan' in df.columns:
        df['tgl_kunjungan_dt'] = pd.to_datetime(df['tgl_kunjungan'], errors='coerce').dt.date
    return df

# Validasi apakah file master sudah diunggah
if file_kunjungan is None:
    st.warning("⚠️ Silakan unggah file master data kunjungan faskes di sidebar kiri terlebih dahulu.")
    st.stop()

# Membaca data kunjungan
try:
    df_kunjungan = load_data_kunjungan(file_kunjungan)
    st.sidebar.success(f"✅ Data kunjungan dimuat: {len(df_kunjungan):,} baris")
except Exception as e:
    st.error(f"Gagal membaca file data kunjungan: {e}")
    st.stop()

# Validasi kolom wajib
kolom_wajib = {'no_kartu', 'tgl_kunjungan', 'nama_faskes', 'no_kunjungan'}
if not kolom_wajib.issubset(set(df_kunjungan.columns)):
    st.error(f"File kunjungan tidak memiliki kolom yang sesuai! Kolom wajib: {kolom_wajib}")
    st.stop()

# --- FUNGSI DETEKSI LOGIKA UTAMA ---
def evaluasi_lab(no_kartu, no_lab, tgl_lab, df_master, toleransi):
    no_kartu_str = str(no_kartu).strip()
    tgl_lab_dt = pd.to_datetime(tgl_lab).date()
    batas_awal = tgl_lab_dt - timedelta(days=toleransi)

    # 1. Filter riwayat peserta
    peserta_kunjungan = df_master[df_master['no_kartu'].astype(str).str.strip() == no_kartu_str]

    if peserta_kunjungan.empty:
        return {
            "status": "ANOMALI TINGGI",
            "keterangan": "Tidak ada riwayat kontak faskes sama sekali dalam database.",
            "faskes_asal": "-",
            "tgl_kunjungan": "-",
            "no_kunjungan": "-"
        }

    # 2. Filter kunjungan dalam jendela toleransi [T_lab - N s/d T_lab]
    kunjungan_valid = peserta_kunjungan[
        (peserta_kunjungan['tgl_kunjungan_dt'] >= batas_awal) & 
        (peserta_kunjungan['tgl_kunjungan_dt'] <= tgl_lab_dt)
    ]

    if kunjungan_valid.empty:
        # Ambil riwayat kunjungan paling akhir untuk referensi audit
        riwayat_terakhir = peserta_kunjungan.sort_values(by='tgl_kunjungan_dt', ascending=False).iloc[0]
        return {
            "status": "ANOMALI",
            "keterangan": f"Tidak ada kunjungan pada rentang {batas_awal} s/d {tgl_lab_dt}. Terakhir berkunjung {riwayat_terakhir['tgl_kunjungan_dt']} di {riwayat_terakhir['nama_faskes']}.",
            "faskes_asal": riwayat_terakhir['nama_faskes'],
            "tgl_kunjungan": riwayat_terakhir['tgl_kunjungan_dt'],
            "no_kunjungan": riwayat_terakhir['no_kunjungan']
        }
    else:
        kunjungan_cocok = kunjungan_valid.sort_values(by='tgl_kunjungan_dt', ascending=False).iloc[0]
        return {
            "status": "VALID",
            "keterangan": f"Ditemukan kunjungan sesuai pada tanggal {kunjungan_cocok['tgl_kunjungan_dt']}.",
            "faskes_asal": kunjungan_cocok['nama_faskes'],
            "tgl_kunjungan": kunjungan_cocok['tgl_kunjungan_dt'],
            "no_kunjungan": kunjungan_cocok['no_kunjungan']
        }

# --- TAMPILAN UTAMA (TABS) ---
tab1, tab2 = st.tabs(["🔍 Verifikasi Satuan", "📁 Audit Massal (Batch File)"])

# === TAB 1: FORM INPUT SATUAN ===
with tab1:
    st.subheader("Verifikasi Single Pelayanan Lab")
    with st.form("form_cek_satuan"):
        col1, col2, col3 = st.columns(3)
        with col1:
            input_no_kartu = st.text_input("Nomor Kartu Peserta", placeholder="Contoh: 0001234567890")
        with col2:
            input_no_lab = st.text_input("Nomor Pelayanan / Klaim Lab", placeholder="Contoh: LAB-2026-0091")
        with col3:
            input_tgl_lab = st.date_input("Tanggal Pelayanan Lab", value=datetime.today())
        
        btn_submit = st.form_submit_button("🔍 Cek Kunjungan Faskes", use_container_width=True)

    if btn_submit:
        if not input_no_kartu:
            st.error("Silakan masukkan Nomor Kartu Peserta.")
        else:
            hasil = evaluasi_lab(input_no_kartu, input_no_lab, input_tgl_lab, df_kunjungan, toleransi_hari)
            
            if hasil['status'] == "VALID":
                st.success(f"### ✅ HASIL: VALID")
                st.write(f"**Keterangan:** {hasil['keterangan']}")
                st.json({
                    "Nomor Kartu": input_no_kartu,
                    "Nomor Lab": input_no_lab,
                    "Tanggal Lab": str(input_tgl_lab),
                    "Faskes Terkait": hasil['faskes_asal'],
                    "Tanggal Kunjungan Faskes": str(hasil['tgl_kunjungan']),
                    "Nomor Kunjungan Faskes": str(hasil['no_kunjungan'])
                })
            elif hasil['status'] == "ANOMALI":
                st.warning(f"### ⚠️ HASIL: POTENSI ANOMALI")
                st.write(f"**Peringatan:** Pelayanan Lab **{input_no_lab}** tidak memiliki kunjungan faskes pengantar pada rentang tanggal yang dipersyaratkan.")
                st.write(f"**Keterangan:** {hasil['keterangan']}")
            else:
                st.error(f"### 🚨 HASIL: ANOMALI TINGGI (TANPA HISTORI)")
                st.write(f"**Peringatan:** Peserta **{input_no_kartu}** sama sekali tidak terdaftar memiliki kunjungan di faskes manapun dalam database.")

# === TAB 2: AUDIT MASSAL (BATCH) ===
with tab2:
    st.subheader("Audit Massal Klaim Pelayanan Laboratorium")
    st.caption("Unggah file daftar klaim lab untuk diperiksa secara otomatis terhadap master kunjungan.")
    
    file_batch_lab = st.file_uploader(
        "Upload File Rekap Lab (.xlsx / .csv)", 
        type=["xlsx", "xls", "csv"], 
        key="upload_batch_lab",
        help="Wajib berisi kolom: no_kartu, no_lab, tgl_lab"
    )
    
    if file_batch_lab is not None:
        try:
            if file_batch_lab.name.endswith('.csv'):
                df_lab = pd.read_csv(file_batch_lab, dtype={'no_kartu': str, 'no_lab': str})
            else:
                df_lab = pd.read_excel(file_batch_lab, dtype={'no_kartu': str, 'no_lab': str})
            
            df_lab.columns = [c.strip().lower() for c in df_lab.columns]
            
            kolom_lab_wajib = {'no_kartu', 'no_lab', 'tgl_lab'}
            if not kolom_lab_wajib.issubset(set(df_lab.columns)):
                st.error(f"File Lab harus memiliki kolom: {kolom_lab_wajib}")
            else:
                st.write(f"Total baris klaim lab: **{len(df_lab):,}** baris")
                
                if st.button("🚀 Mulai Audit Massal", type="primary"):
                    with st.spinner("Memproses pencocokan data..."):
                        hasil_audit = []
                        for _, row in df_lab.iterrows():
                            res = evaluasi_lab(
                                row['no_kartu'], 
                                row['no_lab'], 
                                row['tgl_lab'], 
                                df_kunjungan, 
                                toleransi_hari
                            )
                            hasil_audit.append({
                                "no_kartu": row['no_kartu'],
                                "no_lab": row['no_lab'],
                                "tgl_lab": row['tgl_lab'],
                                "status_audit": res['status'],
                                "keterangan": res['keterangan'],
                                "faskes_terkait": res['faskes_asal'],
                                "tgl_kunjungan_terkait": res['tgl_kunjungan'],
                                "no_kunjungan_terkait": res['no_kunjungan']
                            })
                        
                        df_hasil = pd.DataFrame(hasil_audit)
                        
                        # Ringkasan Metrik
                        total = len(df_hasil)
                        valid_cnt = (df_hasil['status_audit'] == 'VALID').sum()
                        anomali_cnt = total - valid_cnt
                        
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Total Pemeriksaan Lab", f"{total:,}")
                        c2.metric("Lolos / Valid", f"{valid_cnt:,}")
                        c3.metric("Potensi Anomali", f"{anomali_cnt:,}", delta=f"{(anomali_cnt/total)*100:.1f}%", delta_color="inverse")
                        
                        # Tampilkan Data Anomali Teratas
                        st.markdown("#### 🚨 Daftar Temuan Anomali")
                        df_temuan = df_hasil[df_hasil['status_audit'] != 'VALID']
                        st.dataframe(df_temuan, use_container_width=True)
                        
                        # Tombol Download Excel Hasil Temuan
                        output_file = "hasil_audit_lab.xlsx"
                        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                            df_hasil.to_excel(writer, index=False, sheet_name="Hasil_Audit")
                        
                        with open(output_file, "rb") as f:
                            st.download_button(
                                label="📥 Download Laporan Audit Lengkap (.xlsx)",
                                data=f,
                                file_name=f"Hasil_Audit_Lab_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses audit batch: {e}")
