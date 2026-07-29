import io  # Tambahkan ini di bagian atas bersama streamlit & pandas
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
# ==========================================
# FUNGSI HELPER CONVERT TO EXCEL
# ==========================================
def convert_df_to_excel(df_to_download):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_download.to_excel(writer, index=False, sheet_name='Hasil_WLA')
    processed_data = output.getvalue()
    return processed_data
# ==========================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="HWAS - Hybrid Workload Assessment System",
    page_icon="🧠",
    layout="wide"
)

st.title("🧠 Hybrid Workload Assessment System (HWAS)")
st.caption("Integrasi Beban Kerja Kuantitatif (FTE) dan Stres Psikologis (NASA-TLX)")

# ==========================================
# 2. SIDEBAR: PARAMETER & IMPORT DATA
# ==========================================
st.sidebar.header("📁 Pengaturan & Import Data")

uploaded_file = st.sidebar.file_uploader(
    "Upload File Excel WLA", 
    type=["xlsx", "xls"],
    help="Gunakan template resmi HWAS (Kalkulasi FTE & NASA-TLX)"
)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Parameter Jam Kerja Efektif")

hari_kalender = st.sidebar.number_input("Hari Kalender / Tahun", value=365, step=1)
weekend = st.sidebar.number_input("Weekend (Sabtu & Minggu)", value=104, step=1)
cuti_tahunan = st.sidebar.number_input("Cuti Tahunan", value=12, step=1)
libur_nasional = st.sidebar.number_input("Libur Nasional / Cuti Bersama", value=17, step=1)

allowance = st.sidebar.slider(
    "Faktor Efisiensi (Allowance)", 
    min_value=0.50, 
    max_value=1.00, 
    value=0.85, 
    step=0.01,
    help="Standar Indoor=0.875, Outdoor/Gudang=0.85"
)

# Kalkulasi Jam Kerja Efektif (JKE)
hari_kerja_efektif = hari_kalender - (weekend + cuti_tahunan + libur_nasional)
jam_formal_tahun = hari_kerja_efektif * 8
jke_jam_tahun = jam_formal_tahun * allowance
jke_menit_tahun = jke_jam_tahun * 60

st.sidebar.info(
    f"**Hari Kerja Efektif:** {hari_kerja_efektif} Hari/Thn\n\n"
    f"**JKE Tahun:** {jke_jam_tahun:.1f} Jam ({jke_menit_tahun:,.0f} Menit)"
)

# ==========================================
# 3. DEFINISI DUMMY DATA DEFAULT (FALLBACK)
# ==========================================
default_tasks = [
    {"Task": "Memeriksa jadwal penerimaan barang", "Durasi": 4, "Freq": 1, "Basis": "per kedatangan (104x/thn)", "MD": 50, "PD": 20, "TD": 40, "OP": 80, "EF": 45, "FR": 25},
    {"Task": "Memastikan kesesuaian barang vs dokumen", "Durasi": 10, "Freq": 1, "Basis": "per kedatangan (104x/thn)", "MD": 75, "PD": 20, "TD": 65, "OP": 70, "EF": 70, "FR": 60},
    {"Task": "Memantau penempatan barang sesuai kode", "Durasi": 5, "Freq": 1, "Basis": "per kedatangan (104x/thn)", "MD": 40, "PD": 40, "TD": 30, "OP": 85, "EF": 45, "FR": 20},
    {"Task": "Memastikan kesiapan picking & loading", "Durasi": 30, "Freq": 2, "Basis": "per hari", "MD": 60, "PD": 65, "TD": 70, "OP": 75, "EF": 65, "FR": 45},
    {"Task": "Memantau kebersihan dan keamanan gudang", "Durasi": 20, "Freq": 1, "Basis": "per hari", "MD": 30, "PD": 50, "TD": 30, "OP": 80, "EF": 35, "FR": 15},
    {"Task": "Berkoordinasi dengan Admin Furniture", "Durasi": 5, "Freq": 25, "Basis": "per hari", "MD": 65, "PD": 30, "TD": 75, "OP": 70, "EF": 60, "FR": 50},
    {"Task": "Berkoordinasi dengan Kurir Furniture", "Durasi": 10, "Freq": 2, "Basis": "per hari", "MD": 45, "PD": 30, "TD": 50, "OP": 80, "EF": 45, "FR": 30},
    {"Task": "Menyiapkan barang display lantai 2", "Durasi": 15, "Freq": 1, "Basis": "per hari", "MD": 35, "PD": 60, "TD": 40, "OP": 85, "EF": 50, "FR": 20},
    {"Task": "Menyiapkan barang pesanan Gosend/Instant", "Durasi": 3, "Freq": 10, "Basis": "per hari", "MD": 40, "PD": 50, "TD": 85, "OP": 75, "EF": 60, "FR": 40},
    {"Task": "Membantu merakit lemari ekspedisi (Lalamove)", "Durasi": 5, "Freq": 10, "Basis": "per hari", "MD": 50, "PD": 80, "TD": 60, "OP": 75, "EF": 75, "FR": 45}
]

# Setel aktif awal menggunakan default data
active_tasks = default_tasks

# ==========================================
# 4. PARSING DOKUMEN EXCEL DINAMIS (JIKA UPLOAD)
# ==========================================
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. Cari baris header secara otomatis
        df_fte_raw_temp = pd.read_excel(xls, sheet_name="Kalkulasi FTE", header=None)
        fte_header_row = df_fte_raw_temp[df_fte_raw_temp.apply(lambda row: row.astype(str).str.contains("Deskripsi Tugas").any(), axis=1)].index[0]
        
        df_tlx_raw_temp = pd.read_excel(xls, sheet_name="Analisis Psikologis TLX", header=None)
        tlx_header_row = df_tlx_raw_temp[df_tlx_raw_temp.apply(lambda row: row.astype(str).str.contains("Mental Demand").any(), axis=1)].index[0]
        
        # 2. Baca sheet sesuai header
        df_fte = pd.read_excel(xls, sheet_name="Kalkulasi FTE", header=fte_header_row)
        df_tlx = pd.read_excel(xls, sheet_name="Analisis Psikologis TLX", header=tlx_header_row)
        
        df_fte.columns = df_fte.columns.str.strip()
        df_tlx.columns = df_tlx.columns.str.strip()
        
        # 3. Filter ketat: buang NaN, teks #REF!, dan baris Ringkasan Total
        df_fte = df_fte[
            df_fte["Deskripsi Tugas / Aktivitas Pekerjaan"].notna() & 
            ~df_fte["Deskripsi Tugas / Aktivitas Pekerjaan"].astype(str).str.contains("#REF!|TOTAL", case=False, na=False)
        ].copy()
        
        df_tlx = df_tlx[
            df_tlx["Aktivitas / Tugas Pekerjaan"].notna() & 
            ~df_tlx["Aktivitas / Tugas Pekerjaan"].astype(str).str.contains("#REF!|RATA-RATA", case=False, na=False)
        ].copy()
        
        # 4. Susun task ter-upload
        uploaded_tasks = []
        min_len = min(len(df_fte), len(df_tlx))
        
        for i in range(min_len):
            rf = df_fte.iloc[i]
            rt = df_tlx.iloc[i]
            
            uploaded_tasks.append({
                "Task": str(rf["Deskripsi Tugas / Aktivitas Pekerjaan"]),
                "Durasi": float(rf["Waktu / Durasi (Menit)"]),
                "Freq": float(rf["Frekuensi / Volume"]),
                "Basis": str(rf["Basis Frekuensi"]),
                "MD": float(rt["Mental Demand (MD)"]),
                "PD": float(rt["Physical Demand (PD)"]),
                "TD": float(rt["Temporal Demand (TD)"]),
                "OP": float(rt["Performance (OP)"]),
                "EF": float(rt["Effort (EF)"]),
                "FR": float(rt["Frustration (FR)"])
            })
            
        if uploaded_tasks:
            active_tasks = uploaded_tasks  # Timpa data aktif
            st.sidebar.success("✅ Data Excel Berhasil Di-load!")
            
    except Exception as e:
        st.sidebar.error(f"Gagal memproses data Excel: {e}")

# ==========================================
# 5. KALKULASI DINAMIS REALTIME (AKURAT)
# ==========================================
total_menit_kerja = 0
for task in active_tasks:
    d = task["Durasi"]
    f = task["Freq"]
    b = str(task["Basis"]).lower()
    
    if "kedatangan" in b:
        mult = 104
    elif "minggu" in b:
        mult = 52
    elif "bulan" in b:
        mult = 12
    else:  # 'per hari' atau default lainnya
        mult = hari_kerja_efektif
        
    total_menit_kerja += (d * f * mult)

# Indeks FTE Total
total_fte = total_menit_kerja / jke_menit_tahun

# Rata-rata Skor Raw TLX
scores_tlx = []
for task in active_tasks:
    s = (task["MD"] + task["PD"] + task["TD"] + (100 - task["OP"]) + task["EF"] + task["FR"]) / 6.0
    scores_tlx.append(s)

skor_tlx_avg = sum(scores_tlx) / len(scores_tlx) if scores_tlx else 0

# Penentuan Kuadran
if total_fte > 1.28 and skor_tlx_avg > 66.66:
    kuadran_text = "KUADRAN I: BURNOUT / EKSTREM"
elif total_fte <= 1.28 and skor_tlx_avg > 66.66:
    kuadran_text = "KUADRAN II: ANOMALI PSIKOLOGIS"
elif total_fte < 1.00 and skor_tlx_avg <= 66.66:
    kuadran_text = "KUADRAN III: UNDER-UTILIZED / KAPASITAS LONGGAR"
else:
    kuadran_text = "KUADRAN IV: OPERASIONAL RUTIN / BALANCE"

# ==========================================
# 6. STRUCTURING TABS UNTUK INTERFACE
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Executive Dashboard", 
    "📝 Input & Data Editor", 
    "🧠 Analisis Psikologis TLX", 
    "📋 Laporan Manajerial"
])

# ------------------------------------------
# TAB 1: EXECUTIVE DASHBOARD
# ------------------------------------------
with tab1:
    st.subheader("Ringkasan Hasil Evaluasi Beban Kerja Integratif")
    
    col1, col2, col3 = st.columns([1, 1, 1.2])
    
    with col1:
        st.metric(
            label="Total Indeks FTE (Kapasitas Fisik)",
            value=f"{total_fte:.4f}",
            delta="OVERLOAD" if total_fte > 1.28 else ("UNDERLOAD" if total_fte < 1.00 else "IDEAL")
        )
        st.caption("Target FTE Ideal: 1.00 - 1.28")
        
    with col2:
        st.metric(
            label="Skor Raw TLX (Stres Psikologis)",
            value=f"{skor_tlx_avg:.2f} / 100",
            delta="BEBAN TINGGI" if skor_tlx_avg > 66.66 else "BEBAN SEDANG/RENDAH"
        )
        st.caption("Batas Kritis Beban Mental: > 66.66")
        
    with col3:
        st.markdown("### Diagnosis Utama:")
        st.info(f"**{kuadran_text}**")

    st.markdown("---")
    st.subheader("📍 Matriks Kuadran HWAS (FTE vs NASA-TLX)")
    
    # Visualisasi Scatter Plot Kuadran dengan Plotly
    fig = go.Figure()
    
    # Area Kuadran (Background Shapes)
    fig.add_shape(type="rect", x0=0.5, y0=0, x1=1.28, y1=66.66, fillcolor="rgba(144, 238, 144, 0.3)", line_width=0) # III
    fig.add_shape(type="rect", x0=1.28, y0=0, x1=2.0, y1=66.66, fillcolor="rgba(173, 216, 230, 0.3)", line_width=0) # IV
    fig.add_shape(type="rect", x0=0.5, y0=66.66, x1=1.28, y1=100, fillcolor="rgba(255, 255, 224, 0.4)", line_width=0) # II
    fig.add_shape(type="rect", x0=1.28, y0=66.66, x1=2.0, y1=100, fillcolor="rgba(255, 182, 193, 0.4)", line_width=0) # I
    
    # Titik Posisi Evaluasi
    fig.add_trace(go.Scatter(
        x=[total_fte],
        y=[skor_tlx_avg],
        mode="markers+text",
        marker=dict(size=22, color="#1f77b4"),
        text=["Posisi Jabatan Saat Ini"],
        textposition="top center"
    ))
    
    fig.update_layout(
        xaxis_title="Indeks FTE (Beban Fisik)",
        yaxis_title="Skor Raw TLX (Beban Mental)",
        xaxis=dict(range=[0.5, max(2.0, total_fte + 0.2)]),
        yaxis=dict(range=[0, 100]),
        height=450,
        margin=dict(l=20, r=20, t=20, b=20)
    )
    
    st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------
# TAB 2: INPUT & DATA EDITOR
# ------------------------------------------
with tab2:
    st.subheader("Detail Daftar Aktivitas & Beban Kuantitatif")
    df_editor = pd.DataFrame(active_tasks)
    st.dataframe(df_editor, use_container_width=True)
    
    # Pasang tombol download di bawah tabel
    excel_bytes = convert_df_to_excel(df_editor)
    st.download_button(
        label="📥 Unduh Tabel Ini (Excel .xlsx)",
        data=excel_bytes,
        file_name="Data_Aktivitas_HWAS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------
# TAB 3: ANALISIS PSIKOLOGIS TLX
# ------------------------------------------
with tab3:
    st.subheader("🧠 Rincian Evaluasi 6 Dimensi NASA-TLX")
    st.caption("Evaluasi tingkat beban mental, fisik, temporal, performa, usaha, dan frustrasi per aktivitas tugas.")
    
    if active_tasks:
        # 1. Buat DataFrame rincian TLX
        df_tlx_detail = pd.DataFrame(active_tasks)[["Task", "MD", "PD", "TD", "OP", "EF", "FR"]].copy()
        
        # Hitung Raw TLX per baris tugas untuk melengkapi laporan
        df_tlx_detail["Skor Raw TLX"] = (
            df_tlx_detail["MD"] + 
            df_tlx_detail["PD"] + 
            df_tlx_detail["TD"] + 
            (100 - df_tlx_detail["OP"]) + 
            df_tlx_detail["EF"] + 
            df_tlx_detail["FR"]
        ) / 6.0
        
        # Tampilkan Tabel
        st.dataframe(df_tlx_detail, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📥 Ekspor Data Analisis Psikologis TLX")
        
        # 2. Konversi ke biner Excel menggunakan fungsi helper convert_df_to_excel
        excel_tlx_bytes = convert_df_to_excel(df_tlx_detail)
        
        # 3. Tombol Unduh Excel
        st.download_button(
            label="📊 Unduh Rincian NASA-TLX (Excel .xlsx)",
            data=excel_tlx_bytes,
            file_name="Analisis_Psikologis_NASA_TLX.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------------------
# TAB 4: LAPORAN MANAJERIAL
# ------------------------------------------
with tab4:
    st.subheader("📋 Rekomendasi & Intervensi Manajerial")
    st.markdown(f"**Diagnosis Posisi:** {kuadran_text}")
    
    # Narasi Evaluasi Beban Fisik
    if total_fte > 1.28:
        st.error("⚠️ **Rekomendasi Beban Fisik:** Posisi ini mengalami *Overload* (FTE > 1.28). Diperlukan rekrutmen SDM tambahan atau redistribusi alur kerja.")
    elif total_fte < 1.00:
        st.warning("ℹ️ **Rekomendasi Beban Fisik:** Posisi ini *Under-utilized* (FTE < 1.00). Lakukan *job enrichment* atau penambahan cakupan tugas.")
    else:
        st.success("✅ **Rekomendasi Beban Fisik:** Kapasitas fisik berada pada rentang ideal (1.00 - 1.28).")
        
    # Narasi Evaluasi Beban Mental
    if skor_tlx_avg > 66.66:
        st.error("⚠️ **Rekomendasi Beban Mental:** Stres psikologis tinggi (> 66.66). Lakukan evaluasi birokrasi, otomatisasi sistem, atau perbaikan sarana kerja.")
    else:
        st.success("✅ **Rekomendasi Beban Mental:** Tingkat stres psikologis dalam kondisi wajar/terkendali (<= 66.66).")

    # --------------------------------------
    # BAGIAN TOMBOL DOWNLOAD EXCEL (DI SINI TEMPATNYA)
    # --------------------------------------
    st.markdown("---")
    st.subheader("📥 Ekspor Rekap Data WLA")
    st.caption("Unduh berkas Excel berisi seluruh rincian kalkulasi aktivitas, FTE, dan parameter NASA-TLX.")
    
    # 1. Konversi active_tasks ke DataFrame lalu ke biner Excel
    df_download = pd.DataFrame(active_tasks)
    excel_bytes = convert_df_to_excel(df_download)
    
    # 2. Tombol Unduh
    st.download_button(
        label="📊 Download Hasil Analisis (Excel .xlsx)",
        data=excel_bytes,
        file_name="Hasil_Analisis_HWAS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
