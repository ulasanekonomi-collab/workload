import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io

# ==========================================
# 1. KONFIGURASI HALAMAN & TEMA
# ==========================================
st.set_page_config(
    page_title="HWAS - Hybrid Workload Assessment",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Styling
st.markdown("""
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #555555;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 10px;
        padding: 15px;
        border-left: 5px solid #1B365D;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. HEADER & SIDEBAR
# ==========================================
st.markdown("<p class='main-title'>🧠 Hybrid Workload Assessment System (HWAS)</p>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>Integrasi Beban Kerja Kuantitatif (FTE) dan Stres Psikologis (NASA-TLX)</p>", unsafe_allow_html=True)

st.sidebar.header("⚙️ Pengaturan & Import Data")
uploaded_file = st.sidebar.file_path = st.sidebar.file_uploader("Upload File Excel WLA", type=["xlsx", "xls"])

# Sidebar Parameters JKE
st.sidebar.subheader("📅 Parameter Jam Kerja Efektif")
hari_kalender = st.sidebar.number_input("Hari Kalender / Tahun", value=365)
weekend = st.sidebar.number_input("Weekend (Sabtu & Minggu)", value=104)
cuti = st.sidebar.number_input("Cuti Tahunan", value=12)
libur_nasional = st.sidebar.number_input("Libur Nasional / Cuti Bersama", value=17)
allowance = st.sidebar.slider("Faktor Efisiensi (Allowance)", 0.50, 1.00, 0.85, step=0.05)

# Kalkulasi JKE
hke = hari_kalender - (weekend + cuti + libur_nasional)
jam_formal = 8
jke_jam_tahun = hke * jam_formal * allowance
jke_menit_tahun = jke_jam_tahun * 60

st.sidebar.info(f"**Hari Kerja Efektif:** {hke} Hari\n\n**JKE Tahun:** {jke_jam_tahun:.1f} Jam ({jke_menit_tahun:,.0f} Menit)")

# ==========================================
# 3. DUMMY DATA DEFAULT (GUDANG FURNITURE)
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
    {"Task": "Menyiapkan barang pesanan Gosend/Instant", "Durasi": 1, "Freq": 10, "Basis": "per hari", "MD": 40, "PD": 50, "TD": 85, "OP": 75, "EF": 60, "FR": 40},
    {"Task": "Membantu merakit lemari ekspedisi (Lalamove)", "Durasi": 5, "Freq": 10, "Basis": "per hari", "MD": 50, "PD": 80, "TD": 60, "OP": 75, "EF": 75, "FR": 45},
    {"Task": "Merencanakan operasional saat keterlambatan", "Durasi": 10, "Freq": 1, "Basis": "per minggu", "MD": 80, "PD": 20, "TD": 75, "OP": 65, "EF": 70, "FR": 65},
    {"Task": "Stock Opname (Sabtu 08.00 - 14.00)", "Durasi": 360, "Freq": 1, "Basis": "per minggu", "MD": 85, "PD": 75, "TD": 80, "OP": 70, "EF": 85, "FR": 70},
    {"Task": "Loading/unloading insidental malam hari", "Durasi": 120, "Freq": 2, "Basis": "per bulan", "MD": 60, "PD": 85, "TD": 70, "OP": 70, "EF": 80, "FR": 75},
]

df_tasks = pd.DataFrame(default_tasks)

# ==========================================
# 4. PEMROSESAN DATA & KALKULASI
# ==========================================
def hitung_menit_tahun(row, hke):
    d = row['Durasi']
    f = row['Freq']
    b = row['Basis']
    if 'kedatangan' in b:
        return d * f * 104
    elif 'hari' in b:
        return d * f * hke
    elif 'minggu' in b:
        return d * f * 52
    elif 'bulan' in b:
        return d * f * 12
    return 0

df_tasks['Total_Menit_Tahun'] = df_tasks.apply(lambda r: hitung_menit_tahun(r, hke), axis=1)
df_tasks['FTE'] = df_tasks['Total_Menit_Tahun'] / jke_menit_tahun
df_tasks['RTLX'] = (df_tasks['MD'] + df_tasks['PD'] + df_tasks['TD'] + (100 - df_tasks['OP']) + df_tasks['EF'] + df_tasks['FR']) / 6

total_fte = df_tasks['FTE'].sum()
avg_rtl = df_tasks['RTLX'].mean()

# Penentuan Kuadran
if total_fte > 1.28 and avg_rtl > 66.66:
    kuadran = "KUADRAN I: EXTREME OVERLOAD & BURNOUT RISK"
    color_q = "red"
elif total_fte <= 1.28 and avg_rtl > 66.66:
    kuadran = "KUADRAN II: ANOMALI PSIKOLOGIS (Stres Kognitif/Sistem)"
    color_q = "orange"
elif total_fte < 1.00 and avg_rtl <= 66.66:
    kuadran = "KUADRAN III: UNDER-UTILIZED / KAPASITAS LONGGAR"
    color_q = "blue"
else:
    kuadran = "KUADRAN IV: OPERASIONAL RUTIN / BALANCE"
    color_q = "green"

# ==========================================
# 5. TAB TAMPILAN APLIKASI
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Executive Dashboard", "📋 Input & Data Editor", "🧠 Analisis Psikologis TLX", "📝 Laporan Managerial"])

# --- TAB 1: EXECUTIVE DASHBOARD ---
with tab1:
    st.subheader("Ringkasan Hasil Evaluasi Beban Kerja Integratif")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Indeks FTE (Kapasitas Fisik)", f"{total_fte:.4f}", 
                  delta="FIT / OPTIMAL" if 1.0 <= total_fte <= 1.28 else ("OVERLOAD" if total_fte > 1.28 else "UNDERLOAD"),
                  delta_color="normal" if 1.0 <= total_fte <= 1.28 else "inverse")
        st.caption("Target FTE Ideal: 1.00 - 1.28")

    with col2:
        st.metric("Skor Raw TLX (Stres Psikologis)", f"{avg_rtl:.2f} / 100", 
                  delta="BEBAN MENTAL TINGGI" if avg_rtl > 66.66 else "BEBAN SEDANG/RENDAH",
                  delta_color="inverse" if avg_rtl > 66.66 else "normal")
        st.caption("Batas Kritis Beban Mental: > 66.66")

    with col3:
        st.subheader("Diagnosis Utama:")
        st.markdown(f"**:style[color:{color_q};]{{ {kuadran} }}**")

    st.markdown("---")
    
    # Visualisasi Matriks Scatter Plot
    st.subheader("📍 Matriks Kuadran HWAS (FTE vs NASA-TLX)")
    
    fig = px.scatter(
        x=[total_fte], y=[avg_rtl],
        labels={'x': 'Indeks FTE (Beban Fisik)', 'y': 'Skor Raw TLX (Beban Mental)'},
        text=["Gudang Furniture (Indra Lesmana)"],
        size=[20], color_discrete_sequence=['navy']
    )
    
    # Tambah Garis Batas Kuadran
    fig.add_shape(type="rect", x0=0.5, y0=0, x1=1.28, y1=66.66, fillcolor="lightgreen", opacity=0.3, line_width=0)
    fig.add_shape(type="rect", x0=0.5, y0=66.66, x1=1.28, y1=100, fillcolor="orange", opacity=0.3, line_width=0)
    fig.add_shape(type="rect", x0=1.28, y0=66.66, x1=2.0, y1=100, fillcolor="red", opacity=0.3, line_width=0)
    fig.add_shape(type="rect", x0=1.28, y0=0, x1=2.0, y1=66.66, fillcolor="lightblue", opacity=0.3, line_width=0)
    
    fig.update_layout(xaxis_range=[0.5, 2.0], yaxis_range=[0, 100], height=450)
    st.plotly_chart(fig, use_container_width=True)

# --- TAB 2: DATA EDITOR ---
with tab2:
    st.subheader("Edit Data Tugas & Rating Beban Kerja")
    st.caption("Anda dapat mengubah angka durasi, frekuensi, maupun rating 6 dimensi psikologis secara langsung pada tabel di bawah ini:")
    
    edited_df = st.data_editor(df_tasks, num_rows="dynamic", use_container_width=True)

# --- TAB 3: ANALISIS TLX ---
with tab3:
    st.subheader("Breakdown Profil 6 Dimensi Psikologis (NASA-TLX)")
    
    avg_md = edited_df['MD'].mean()
    avg_pd = edited_df['PD'].mean()
    avg_td = edited_df['TD'].mean()
    avg_op = 100 - edited_df['OP'].mean()
    avg_ef = edited_df['EF'].mean()
    avg_fr = edited_df['FR'].mean()
    
    df_radar = pd.DataFrame({
        'Dimensi': ['Mental Demand', 'Physical Demand', 'Temporal Demand', 'Performance (Inversed)', 'Effort', 'Frustration'],
        'Skor': [avg_md, avg_pd, avg_td, avg_op, avg_ef, avg_fr]
    })
    
    fig_radar = px.line_polar(df_radar, r='Skor', theta='Dimensi', line_close=True)
    fig_radar.update_traces(fill='toself', fillcolor='rgba(0, 102, 102, 0.4)', line_color='#006666')
    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])))
    
    st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 4: LAPORAN MANAGERIAL ---
with tab4:
    st.subheader("📄 Laporan Rekomendasi Managerial Otomatis")
    
    st.markdown(f"""
    ### **HASIL ANALISIS UNTUK JABATAN: GUDANG FURNITURE**
    * **Status FTE Kuantitatif:** {total_fte:.4f} (FIT / OPTIMAL)
    * **Status Stres Psikologis:** {avg_rtl:.2f} (BEBAN MENTAL TINGGI)
    * **Diagnosis:** **{kuadran}**
    
    ---
    
    #### **REKOMENDASI INTERVENSI STRATEGIS:**
    1. **Digitalisasi Verifikasi Dokumen:** Terapkan *Barcode Scanner* untuk menekan ketidaksesuaian Surat Jalan vs Fisik Barang (*Menurunkan Mental Demand & Frustration*).
    2. **Otomasi Alat Kerja:** Sediakan perkakas listrik (*power tools*) untuk rakitan lemari ekspedisi Lalamove (*Menurunkan Physical Demand & Effort*).
    3. **Penataan Shift Malam:** Berikan kompensasi *off-shift* pagi setelah tugas insidental bongkar muat malam hari ($22.00 - 24.00$).
    """)

st.sidebar.markdown("---")
st.sidebar.success("Aplikasi HWAS Siap Digunakan!")
