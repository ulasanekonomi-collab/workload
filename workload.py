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
# Pembagian Tab (Tambahkan Tab 5)
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Dashboard", 
    "📝 Input & Data Editor", 
    "🧠 Analisis Psikologis TLX", 
    "📋 Laporan Manajerial",
    "📅 Beban Kerja Harian & Insidental"  # Tab Baru
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
# Tombol Download Ringkasan Integratif untuk Tab 1
    st.markdown("---")
    summary_data = {
        "Indikator Utama": [
            "Jumlah Total Aktivitas (Task)",
            "Nilai Full-Time Equivalent (FTE)",
            "Status Beban Kerja Fisik",
            "Rata-rata Skor NASA-TLX",
            "Status Beban Mental",
            "Kuadran Diagnosis HWAS"
        ],
        "Hasil Evaluasi": [
            len(active_tasks),
            f"{total_fte:.2f}",
            "Overload" if total_fte > 1.28 else ("Underload" if total_fte < 1.00 else "Ideal"),
            f"{skor_tlx_avg:.2f}",
            "Tinggi (Stres)" if skor_tlx_avg > 66.66 else "Normal/Terkendali",
            kuadran_text
        ]
    }
    df_summary_integratif = pd.DataFrame(summary_data)
    
    st.download_button(
        label="📊 Unduh Ringkasan Integratif (.xlsx)",
        data=convert_df_to_excel(df_summary_integratif),
        file_name="Ringkasan_Evaluasi_Integratif_HWAS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ------------------------------------------
# TAB 3: ANALISIS PSIKOLOGIS TLX
# ------------------------------------------
with tab3:
    st.subheader("🧠 Rincian Evaluasi 6 Dimensi NASA-TLX")
    st.caption("Evaluasi tingkat beban mental, fisik, temporal, performa, usaha, dan frustrasi per aktivitas tugas.")
    
    if active_tasks:
        # 1. Buat dataframe rincian TLX
        df_tlx_detail = pd.DataFrame(active_tasks)[["Task", "MD", "PD", "TD", "OP", "EF", "FR"]].copy()

        # Hitung Raw TLX per baris tugas
        df_tlx_detail["Skor Raw TLX"] = (
            df_tlx_detail["MD"] + 
            df_tlx_detail["PD"] + 
            df_tlx_detail["TD"] + 
            (100 - df_tlx_detail["OP"]) + 
            df_tlx_detail["EF"] + 
            df_tlx_detail["FR"]
        ) / 6.0

        # Tambahkan Kategori Beban Mental
        df_tlx_detail["Kategori Mental"] = df_tlx_detail["Skor Raw TLX"].apply(
            lambda x: "Sangat Tinggi" if x > 80 else ("Tinggi" if x > 60 else "Sedang")
        )

        # Rename kolom agar rapi saat ditampilkan & diunduh
        df_tlx_detail.columns = [
            "Deskripsi Tugas", "Mental (MD)", "Fisik (PD)", "Waktu (TD)", 
            "Performa (OP)", "Usaha (EF)", "Frustrasi (FR)", "Skor Raw TLX", "Kategori Beban Mental"
        ]

        # Tampilkan tabel di Streamlit
        st.dataframe(df_tlx_detail, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Ekspor Data Analisis Psikologis TLX")

        # Conversi ke biner Excel
        excel_tlx_bytes = convert_df_to_excel(df_tlx_detail)

        # Tombol Unduh Excel Hasil Analisis TLX
        st.download_button(
            label="📊 Unduh Rincian NASA-TLX (.xlsx)",
            data=excel_tlx_bytes,
            file_name="Hasil_Analisis_Psychological_NASA_TLX.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# TAB 4: LAPORAN MANAJERIAL
with tab4:
    st.subheader("📋 Laporan Eksekutif Integratif HWAS")
    st.caption("Ringkasan komprehensif hasil evaluasi beban fisik (FTE) dan psikologis (NASA-TLX).")

    if active_tasks:
        # 1. Menyiapkan DataFrame Ringkasan Eksekutif
        df_summary_mgr = pd.DataFrame({
            "Parameter Evaluasi": [
                "Total Aktivitas (Task)", 
                "Indeks Total FTE", 
                "Status Beban Fisik",
                "Rata-rata Skor NASA-TLX", 
                "Status Beban Mental",
                "Diagnosis Kuadran HWAS"
            ],
            "Nilai / Status": [
                len(active_tasks),
                f"{total_fte:.2f}",
                "Overload" if total_fte > 1.28 else ("Underload" if total_fte < 1.00 else "Ideal"),
                f"{skor_tlx_avg:.2f}",
                "Tinggi (Stres)" if skor_tlx_avg > 66.66 else "Normal/Terkendali",
                kuadran_text
            ]
        })

        # Display Ringkasan di Tab
        st.dataframe(df_summary_mgr, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Unduh Laporan Manajerial Lengkap")

        # 2. Siapkan data FTE untuk sheet pendukung (Kalkulasi realtime)
        fte_list = []
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
            else:
                mult = hari_kerja_efektif
                
            tot_menit = d * f * mult
            indeks_fte = tot_menit / jke_menit_tahun
            
            fte_list.append({
                "Deskripsi Tugas": task["Task"],
                "Durasi (Menit)": d,
                "Frekuensi": f,
                "Basis Waktu": task["Basis"],
                "Total Menit/Tahun": tot_menit,
                "Indeks FTE": round(indeks_fte, 4)
            })
            
        df_fte_export = pd.DataFrame(fte_list)

        # Siapkan data TLX untuk sheet pendukung
        df_tlx_export = pd.DataFrame(active_tasks)[["Task", "MD", "PD", "TD", "OP", "EF", "FR"]].copy()
        df_tlx_export["Skor Raw TLX"] = (
            df_tlx_export["MD"] + df_tlx_export["PD"] + df_tlx_export["TD"] + 
            (100 - df_tlx_export["OP"]) + df_tlx_export["EF"] + df_tlx_export["FR"]
        ) / 6.0
        df_tlx_export["Kategori Mental"] = df_tlx_export["Skor Raw TLX"].apply(
            lambda x: "Sangat Tinggi" if x > 80 else ("Tinggi" if x > 60 else "Sedang")
        )
        df_tlx_export.columns = ["Deskripsi Tugas", "Mental (MD)", "Fisik (PD)", "Waktu (TD)", "Performa (OP)", "Usaha (EF)", "Frustrasi (FR)", "Skor Raw TLX", "Kategori Beban Mental"]
        # 3. Fungsi Helper Multi-Sheet Khusus Tab Manajerial
        def export_multisheet_managerial(df_summary, df_fte, df_tlx):
            import io
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='Executive Summary')
                df_fte.to_excel(writer, index=False, sheet_name='Rincian FTE')
                df_tlx.to_excel(writer, index=False, sheet_name='Rincian NASA-TLX')
            return output.getvalue()

        # 4. Tombol Download Laporan Manajerial Multi-Sheet
        st.download_button(
            label="💼 Unduh Laporan Manajerial Komprehensif (.xlsx)",
            data=export_multisheet_managerial(df_summary_mgr, df_fte_export, df_tlx_export),
            file_name="Laporan_Manajerial_Komprehensif_HWAS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# ------------------------------------------
# TAB 5: BEBAN KERJA HARIAN & INSIDENTAL
# ------------------------------------------
with tab5:
    st.subheader("📅 Analisis & Simulasi Beban Kerja Harian Interaktif")
    st.caption("Evaluasi kecukupan waktu kerja harian pekerja berdasarkan tugas rutin, tugas insidental terdaftar, dan input tugas ad-hoc baru.")

    # 1. Parameter Batas Harian
    kapasitas_harian_efektif = round(jke_menit_tahun / hari_kerja_efektif, 1) if hari_kerja_efektif > 0 else 408.0

    st.info(f"💡 **Target Efektif Harian Pekerja:** **{kapasitas_harian_efektif:.0f} Menit/Hari** ({kapasitas_harian_efektif/60:.2f} Jam/Hari)")

    # Session State untuk menyimpan Tugas Tambahan / Ad-Hoc buatan user
    if "custom_incidental_tasks" not in st.session_state:
        st.session_state.custom_incidental_tasks = []

    # 2. SEKSI A: Checklist Tugas Insidental dari Master Data
    st.markdown("---")
    st.markdown("### 1️⃣ Pilih Tugas Non-Harian dari Master Data")
    
    rutin_harian = []
    insidental_periodik = []

    for task in active_tasks:
        b = str(task.get("Basis", "")).lower()
        if "hari" in b:
            rutin_harian.append(task)
        else:
            insidental_periodik.append(task)

    menit_rutin_harian = sum(t.get("Durasi", 0) * t.get("Freq", 0) for t in rutin_harian)
    menit_insidental_master = 0

    if insidental_periodik:
        cols = st.columns(2)
        for idx, task in enumerate(insidental_periodik):
            col_target = cols[idx % 2]
            with col_target:
                durasi_kejadian = task.get("Durasi", 0) * task.get("Freq", 0)
                is_active = st.checkbox(
                    f"📌 **{task.get('Task', 'Tugas')}** (+{durasi_kejadian} Menit) [{task.get('Basis', '-')}]",
                    key=f"incidental_chk_{idx}"
                )
                if is_active:
                    menit_insidental_master += durasi_kejadian
    else:
        st.write("*(Tidak ada tugas non-harian dalam Master Data)*")

    # 3. SEKSI B: Form Input Tugas Baru / Ad-Hoc di Luar Daftar
    st.markdown("---")
    st.markdown("### 2️⃣ Input Tugas Tambahan Baru (Di Luar Daftar Master)")
    st.caption("Gunakan form ini jika pekerja menerima instruksi tugas mendadak/insidental baru hari ini.")

    with st.form("form_tugas_baru", clear_on_submit=True):
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        with col_f1:
            nama_tugas_baru = st.text_input("Deskripsi / Nama Tugas Baru", placeholder="Misal: Unloading Kontainer Darurat")
        with col_f2:
            durasi_baru = st.number_input("Durasi (Menit)", min_value=1, value=30, step=5)
        with col_f3:
            freq_baru = st.number_input("Frekuensi (Kali)", min_value=1, value=1, step=1)

        col_f4, col_f5 = st.columns(2)
        with col_f4:
            stres_mental = st.slider("Estimasi Stres Mental (NASA-TLX)", min_value=0, max_value=100, value=60, help="0 = Sangat Ringan, 100 = Ekstrem Sangat Berat")
        with col_f5:
            st.write("") # Spacing
            btn_tambah = st.form_submit_button("➕ Tambahkan Tugas Insidental Ini", use_container_width=True)

        if btn_tambah and nama_tugas_baru.strip():
            st.session_state.custom_incidental_tasks.append({
                "Task": nama_tugas_baru.strip(),
                "Durasi": durasi_baru,
                "Freq": freq_baru,
                "TotalMenit": durasi_baru * freq_baru,
                "MentalScore": stres_mental
            })
            st.success(f"✅ Tugas '{nama_tugas_baru}' berhasil ditambahkan ke simulasi harian!")

    # 3. Tampilkan Daftar Tugas Ad-Hoc & Hitung Menitnya
    menit_custom_insidental = 0
    if st.session_state.custom_incidental_tasks:
        st.write("**Daftar Tugas Tambahan Baru Hari Ini:**")
        for i, c_task in enumerate(st.session_state.custom_incidental_tasks):
            tot = c_task["TotalMenit"]
            menit_custom_insidental += tot
            
            c_col1, c_col2 = st.columns([4, 1])
            with c_col1:
                st.write(f"• **{c_task['Task']}** — {c_task['Durasi']} mnt x {c_task['Freq']}x = **{tot} Menit** (Beban Mental: {c_task['MentalScore']}/100)")
            with c_col2:
                if st.button("❌ Hapus", key=f"del_custom_{i}"):
                    st.session_state.custom_incidental_tasks.pop(i)
                    st.rerun()

    # 4. Kalkulasi Total & Dampak Kecepatan Harian
    total_insidental_keseluruhan = menit_insidental_master + menit_custom_insidental
    total_menit_hari_ini = menit_rutin_harian + total_insidental_keseluruhan
    total_jam_hari_ini = total_menit_hari_ini / 60.0
    rasio_beban_harian = (total_menit_hari_ini / kapasitas_harian_efektif) * 100 if kapasitas_harian_efektif > 0 else 0

    st.markdown("---")
    st.markdown("### 📊 Rekapitulasi Beban Kerja Harian (Real-Time)")

    col_h1, col_h2, col_h3, col_h4 = st.columns(4)
    with col_h1:
        st.metric(
            label="Beban Rutin Harian",
            value=f"{menit_rutin_harian} Mnt",
            delta=f"{(menit_rutin_harian/60):.2f} Jam"
        )
    with col_h2:
        st.metric(
            label="Insidental Master Data",
            value=f"{menit_insidental_master} Mnt",
            delta=f"{(menit_insidental_master/60):.2f} Jam" if menit_insidental_master > 0 else "0 Jam"
        )
    with col_h3:
        # Menggunakan nama variabel yang sudah disesuaikan (menit_custom_insidental)
        st.metric(
            label="Tugas Ad-Hoc Baru",
            value=f"{menit_custom_insidental} Mnt",
            delta=f"{(menit_custom_insidental/60):.2f} Jam" if menit_custom_insidental > 0 else "0 Jam"
        )
    with col_h4:
        st.metric(
            label="TOTAL WAKTU KERJA",
            value=f"{total_menit_hari_ini} Mnt",
            delta=f"{total_jam_hari_ini:.2f} Jam / Target {kapasitas_harian_efektif/60:.1f} Jam",
            delta_color="inverse" if total_menit_hari_ini > kapasitas_harian_efektif else "normal"
        )

    # Indikator Status Kebijakan Manajerial
    st.write(f"**Tingkat Penggunaan Kapasitas Harian:** {rasio_beban_harian:.1f}%")

    if total_menit_hari_ini > kapasitas_harian_efektif:
        kelebihan = total_menit_hari_ini - kapasitas_harian_efektif
        st.error(
            f"🚨 **OVERLOAD HARIAN!** Total pengerjaan tugas melebihi kapasitas kerja efektif sebesar **{kelebihan:.0f} Menit ({kelebihan/60:.2f} Jam)**.\n\n"
            f"💡 **Rekomendasi:** Disarankan membayar **Lembur Insidental ({kelebihan/60:.2f} Jam)** atau mengalihkan tugas ad-hoc ke stasiun kerja lain."
        )
    elif total_menit_hari_ini >= (kapasitas_harian_efektif * 0.85):
        st.warning("⚠️ **KAPASITAS PADAT.** Pekerja beroperasi mendekati ambang batas kelelahan harian.")
    else:
        sisa = kapasitas_harian_efektif - total_menit_hari_ini
        st.success(f"✅ **KAPASITAS NORMAL.** Masih tersisa alokasi waktu efektif **{sisa:.0f} Menit ({sisa/60:.2f} Jam)**.")

    # ---------------------------------------------------------
    # 5. EKSPOR SIMULASI HARIAN LENGKAP BERSAMA RINCIAN TUGAS (.XLSX)
    # ---------------------------------------------------------
    st.markdown("---")
    st.markdown("### 📥 Unduh Hasil Simulasi & Rincian Tugas Harian")

    # A. Tabel Summary Beban Kerja Harian
    df_harian_summary = pd.DataFrame([
        {"Indikator Harian": "Kapasitas Efektif Harian (Standard)", "Nilai (Menit)": kapasitas_harian_efektif, "Nilai (Jam)": round(kapasitas_harian_efektif/60, 2)},
        {"Indikator Harian": "Beban Rutin Harian", "Nilai (Menit)": menit_rutin_harian, "Nilai (Jam)": round(menit_rutin_harian/60, 2)},
        {"Indikator Harian": "Tambahan Insidental (Master Data)", "Nilai (Menit)": menit_insidental_master, "Nilai (Jam)": round(menit_insidental_master/60, 2)},
        {"Indikator Harian": "Tugas Ad-Hoc Baru (Input User)", "Nilai (Menit)": menit_custom_insidental, "Nilai (Jam)": round(menit_custom_insidental/60, 2)},
        {"Indikator Harian": "TOTAL BEBAN KERJA HARI INI", "Nilai (Menit)": total_menit_hari_ini, "Nilai (Jam)": round(total_jam_hari_ini, 2)},
        {"Indikator Harian": "Status Utilisasi Kapasitas", "Nilai (Menit)": f"{rasio_beban_harian:.1f}%", "Nilai (Jam)": "OVERLOAD" if total_menit_hari_ini > kapasitas_harian_efektif else "NORMAL"}
    ])

    # B. Tabel Rincian Tugas yang Aktif Hari Ini (Rutin + Insidental Terpilih + Ad-Hoc Baru)
    rincian_tugas_hari_ini = []

    # 1. Masukkan Tugas Rutin Harian
    for r_task in rutin_harian:
        tot_m = r_task.get("Durasi", 0) * r_task.get("Freq", 0)
        rincian_tugas_hari_ini.append({
            "Kategori Tugas": "Rutin Harian",
            "Deskripsi Tugas": r_task.get("Task", "-"),
            "Durasi (Menit)": r_task.get("Durasi", 0),
            "Frekuensi": r_task.get("Freq", 0),
            "Total Menit Hari Ini": tot_m,
            "Estimasi Stres Mental": r_task.get("MD", 50)
        })

    # 2. Masukkan Tugas Insidental Master Data yang Dicentang (dari checklist)
    # Catatan: Variabel 'tugas_insidental_diaktifkan' menampung tugas yang dicentang user
    for task_name in tugas_insidental_diaktifkan:
        # Cari detail tugas dari insidental_periodik
        match_task = next((t for t in insidental_periodik if t.get("Task") == task_name), None)
        if match_task:
            tot_m = match_task.get("Durasi", 0) * match_task.get("Freq", 0)
            rincian_tugas_hari_ini.append({
                "Kategori Tugas": "Insidental (Master Data)",
                "Deskripsi Tugas": match_task.get("Task", "-"),
                "Durasi (Menit)": match_task.get("Durasi", 0),
                "Frekuensi": match_task.get("Freq", 0),
                "Total Menit Hari Ini": tot_m,
                "Estimasi Stres Mental": match_task.get("MD", 50)
            })

    # 3. Masukkan Tugas Ad-Hoc Baru Hasil Input User
    if st.session_state.custom_incidental_tasks:
        for c_task in st.session_state.custom_incidental_tasks:
            rincian_tugas_hari_ini.append({
                "Kategori Tugas": "Ad-Hoc Baru (Input Manual)",
                "Deskripsi Tugas": c_task["Task"],
                "Durasi (Menit)": c_task["Durasi"],
                "Frekuensi": c_task["Freq"],
                "Total Menit Hari Ini": c_task["TotalMenit"],
                "Estimasi Stres Mental": c_task["MentalScore"]
            })

    df_rincian_harian_export = pd.DataFrame(rincian_tugas_hari_ini)

    # C. Helper Multi-Sheet Khusus Simulasi Harian
    def export_harian_multisheet(df_summary, df_detail):
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_summary.to_excel(writer, index=False, sheet_name='Ringkasan Harian')
            df_detail.to_excel(writer, index=False, sheet_name='Daftar Tugas Aktif Hari Ini')
        return output.getvalue()

    # D. Tombol Download Excel Multi-Sheet
    st.download_button(
        label="📥 Unduh Simulasi & Rincian Tugas Harian (.xlsx)",
        data=export_harian_multisheet(df_harian_summary, df_rincian_harian_export),
        file_name="Simulasi_Beban_Kerja_dan_Rincian_Tugas_Harian.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
# ------------------------------------------
# FOOTER SIDEBAR: FOTO EKSTRA KECIL SEJAJAR TEKS
# ------------------------------------------
st.sidebar.markdown("---")

col_foto, col_teks = st.sidebar.columns([1, 4])

with col_foto:
    try:
        # Menentukan lebar pasti gambar dalam piksel (width=45)
        st.image("harapku-sarah-saputri.jpg", width=45)
    except Exception:
        st.write("📷")

with col_teks:
    st.markdown(
        "<div style='font-size: 0.8rem; color: #555; line-height: 1.3; padding-top: 4px;'>"
        "Dikembangkan oleh:<br>"
        "<b>Sarah Saputri & Tim</b>, 2026"
        "</div>", 
        unsafe_allow_html=True
    ) 
