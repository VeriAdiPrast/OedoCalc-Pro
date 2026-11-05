import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import io
import math

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="OedoCalc Pro & Buku Panduan",
    page_icon="🧪",
    layout="wide"
)

# 2. Fungsi Pembuat Template Excel Otomatis (In-Memory)
def buat_template_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        waktu_menit = [0, 0, 0, 1, 2, 4, 8, 15, 30, 60, 120, 240, 480, 1440]
        waktu_detik = [6, 15, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        
        data = {
            'Waktu (Menit)': waktu_menit,
            'Waktu (Detik)': waktu_detik,
            'p=0.25 kg/cm2': [0.012, 0.018, 0.025, 0.034, 0.045, 0.058, 0.072, 0.086, 0.098, 0.108, 0.115, 0.120, 0.124, 0.128],
            'p=0.50 kg/cm2': [0.045, 0.062, 0.080, 0.102, 0.130, 0.162, 0.198, 0.235, 0.268, 0.295, 0.315, 0.330, 0.340, 0.350],
            'p=1.00 kg/cm2': [0.110, 0.142, 0.178, 0.215, 0.260, 0.312, 0.370, 0.432, 0.485, 0.530, 0.562, 0.585, 0.600, 0.615],
            'p=2.00 kg/cm2': [0.240, 0.295, 0.355, 0.420, 0.495, 0.580, 0.672, 0.765, 0.840, 0.905, 0.955, 0.990, 1.015, 1.035],
            'p=4.00 kg/cm2': [0.480, 0.560, 0.650, 0.745, 0.855, 0.980, 1.115, 1.240, 1.345, 1.430, 1.500, 1.550, 1.585, 1.615],
            'p=8.00 kg/cm2': [0.890, 1.010, 1.150, 1.290, 1.450, 1.620, 1.800, 1.960, 2.080, 2.180, 2.250, 2.300, 2.335, 2.365]
        }
        df_template = pd.DataFrame(data)
        df_template.to_excel(writer, index=False, sheet_name='Data_Oedometer')
    return output.getvalue()

# 3. Fungsi Logika Analisis Oedometer
def proses_oedometer(df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan):
    waktu_menit = df.iloc[:, 0].astype(float) + (df.iloc[:, 1].astype(float) / 60.0)
    kolom_beban = df.columns[2:]
    dial_awal = float(df.iloc[0, 2])
    
    beban_list = []
    penurunan_akhir_list = []
    dict_pembacaan = {}

    for idx, col in enumerate(kolom_beban):
        try:
            p_val = float(str(col).replace('kg/cm2','').replace('p=','').strip())
        except:
            p_val = (idx + 1) * 0.25
        beban_list.append(p_val)
        
        pembacaan = df[col].astype(float).tolist()
        dict_pembacaan[col] = pembacaan
        penurunan_akhir_list.append(abs(pembacaan[-1] - dial_awal) / 100.0)

    akar_waktu = np.sqrt(waktu_menit)

    # Grafik Penurunan vs Akar Waktu
    total_grafik = len(kolom_beban)
    cols_grid = 2
    rows_grid = math.ceil(total_grafik / cols_grid)
    
    fig_time, axes = plt.subplots(nrows=rows_grid, ncols=cols_grid, figsize=(10, 3 * rows_grid))
    axes = axes.flatten() if total_grafik > 1 else [axes]

    for idx, col in enumerate(kolom_beban):
        pembacaan_dial = dict_pembacaan[col]
        axes[idx].plot(akar_waktu[:len(pembacaan_dial)], pembacaan_dial, marker='o', color='navy', linewidth=1.5, markersize=3)
        axes[idx].set_title(f'Beban: {col}', fontsize=10, fontweight='bold')
        axes[idx].set_xlabel('√Waktu (√menit)', fontsize=8)
        axes[idx].set_ylabel('Pembacaan Arloji (mm)', fontsize=8)
        axes[idx].grid(True, linestyle='--', alpha=0.6)
        axes[idx].invert_yaxis()

    for i in range(idx + 1, len(axes)):
        fig_time.delaxes(axes[i])

    plt.tight_layout()

    # Grafik e - log p
    p_kpa = [b * 98.0665 for b in beban_list]
    e_list = [(e0 - ((dh / H0) * (1 + e0))) for dh in penurunan_akhir_list]
    
    if len(p_kpa) >= 2 and p_kpa[-1] > p_kpa[-2]:
        Cc = (e_list[-2] - e_list[-1]) / np.log10(p_kpa[-1] / p_kpa[-2])
    else:
        Cc = 0.3
    
    Tv = 0.848
    H_dr = tebal_lempung / 2.0 if "2 Arah" in drainase else tebal_lempung
    t90_tahun = (Tv * (H_dr ** 2)) / cv if cv > 0 else 0
    t90_bulan = t90_tahun * 12.0

    fig_elogp, ax2 = plt.subplots(figsize=(6, 4))
    ax2.plot(p_kpa, e_list, marker='s', linestyle='-', color='teal', linewidth=2)
    ax2.set_xscale('log')
    ax2.set_xlabel('Tegangan Efektif, p (kPa) - Skala Log')
    ax2.set_ylabel('Angka Pori, e')
    ax2.set_title('Kurva Konsolidasi Utama (e - log p)')
    ax2.grid(True, which="both", ls="--", linewidth=0.5)
    ax2.invert_yaxis()
    plt.tight_layout()

    # Klasifikasi & Rekomendasi Technical
    if Cc < 0.2:
        kemampatan = "RENDAH (Low Compressibility) - Tanah Padat"
        rekomendasi = "Pondasi Dangkal (Footplat / Batu Kali)."
        pekerjaan = "- Galian tanah & urugan pasir\n- Lean concrete\n- Pembesian & pengecoran"
        total_bawah = luas_bangunan * 500000
        total_atas = luas_bangunan * 600000
    elif 0.2 <= Cc <= 0.4:
        kemampatan = "SEDANG (Medium Compressibility) - Lempung Sedang"
        rekomendasi = "Pondasi Minipile (Tiang Pancang Skala Kecil)."
        pekerjaan = "- Mobilisasi alat pancang\n- Pengadaan & pemancangan tiang\n- Pile Cap"
        total_bawah = (luas_bangunan * 1150000) + 25000000
        total_atas = (luas_bangunan * 1350000) + 25000000
    else:
        kemampatan = "TINGGI (High Compressibility) - Tanah Sangat Lunak"
        rekomendasi = "Perbaikan Tanah Masif (PVD + Preloading)."
        pekerjaan = "- Geotextile & Sand Blanket\n- Penetrasi PVD dengan Rig\n- Preloading & Monitoring"
        total_bawah = (luas_bangunan * 2600000) + 45000000
        total_atas = (luas_bangunan * 2900000) + 45000000

    laporan_teknis = (
        f"--- HASIL ANALISIS DATA MENTAH LAB ---\n"
        f"Beban Diuji (kg/cm²)  : {beban_list}\n"
        f"Indeks Kompresi (Cc)  : {Cc:.4f}\n"
        f"Klasifikasi Tanah     : {kemampatan}\n\n"
        f"--- ESTIMASI WAKTU PENURUNAN LAPANGAN ---\n"
        f"Tebal Lapisan Lempung : {tebal_lempung} meter\n"
        f"Waktu Penurunan t90   : {t90_tahun:.2f} Tahun ({t90_bulan:.0f} Bulan)\n\n"
        f"--- REKOMENDASI TEKNIS ---\n"
        f"Metode Penanganan     : {rekomendasi}\n"
        f"Tahapan Pekerjaan     :\n{pekerjaan}\n"
    )

    rab_text = f"ESTIMASI TOTAL BIAYA PROYEK ({luas_bangunan} m²):\nRp {total_bawah:,.0f} s/d Rp {total_atas:,.0f}"

    df_export = pd.DataFrame({
        'Beban (kg/cm2)': beban_list,
        'Beban (kPa)': p_kpa,
        'Penurunan Total (mm)': penurunan_akhir_list,
        'Angka Pori (e)': e_list
    })

    return laporan_teknis, rab_text, fig_elogp, fig_time, df_export

# 4. TAB NAVIGASI UTAMA
tab_kalkulator, tab_panduan = st.tabs(["🚀 Aplikasi Kalkulator", "📖 Buku Panduan Penggunaan"])

# --- TAB 1: KALKULATOR GEOTEKNIK ---
with tab_kalkulator:
    st.title("🧪 OedoCalc Pro: Analisis Konsolidasi Tanah")
    
    col_dl, col_space = st.columns([1, 2])
    with col_dl:
        st.download_button(
            label="📥 Unduh Template Format Excel (.xlsx)",
            data=buat_template_excel(),
            file_name="Template_OedoCalc_Pro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    col1, col2 = st.columns([1, 1])

    with col1:
        uploaded_file = st.file_uploader("Unggah File Data Mentah (.xlsx / .csv)", type=["csv", "xlsx"])
        D_sample = st.number_input("Diameter Cincin (mm)", value=63.5)
        H0 = st.number_input("Tinggi Sampel H0 (mm)", value=20.0)
        e0 = st.number_input("Angka Pori Awal (e0)", value=1.15)
        tebal_lempung = st.number_input("Tebal Lempung (m)", value=10.0)
        cv = st.number_input("Cv (m²/thn)", value=1.5)
        drainase = st.radio("Kondisi Drainase", ["1 Arah", "2 Arah (Atas & Bawah)"])
        luas_bangunan = st.number_input("Luas Area Proyek (m²)", value=500.0)
        btn = st.button("🚀 Olah Data & Buat Grafik", type="primary")

    with col2:
        if uploaded_file and btn:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            laporan, rab, fig_elogp, fig_time, df_export = proses_oedometer(df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan)
            
            subtab1, subtab2, subtab3, subtab4 = st.tabs(["📄 Laporan Teknis", "💰 Estimasi RAB", "📈 Kurva e - log p", "📉 Grafik Penurunan"])
            with subtab1:
                st.code(laporan)
            with subtab2:
                st.success(rab)
            with subtab3:
                st.pyplot(fig_elogp)
            with subtab4:
                st.pyplot(fig_time)
                
            csv_bytes = df_export.to_csv(index=False).encode('utf-8')
            st.download_button("📊 Unduh Ringkasan Hasil (.csv)", data=csv_bytes, file_name="ringkasan_oedometer.csv", mime="text/csv")

# --- TAB 2: BUKU PANDUAN PENGGUNAAN ---
with tab_panduan:
    st.header("📖 Buku Panduan Penggunaan OedoCalc Pro")
    
    st.markdown("""
    ### 1. Pendahuluan & Parameter Masukan
    Aplikasi **OedoCalc Pro** dirancang untuk mengolah data mentah pembacaan arloji dial konsolidasi oedometer (*ASTM D2435 / SNI 2812:2011*) secara otomatis.
    
    **Parameter Input Utama:**
    * **Diameter Cincin ($D$):** Ukuran sampel tanah laboratorium (Standar: 63.5 mm).
    * **Tinggi Sampel ($H_0$):** Tinggi awal benda uji (Standar: 20.0 mm).
    * **Angka Pori Awal ($e_0$):** Rasio volume pori terhadap volume padatan tanah awal.
    * **Koefisien Konsolidasi ($C_v$):** Kecepatan penurunan tanah ($m^2/tahun$).
    * **Drainase:** Pilihan 1 arah (lapisan kedap di bawah) atau 2 arah (porous stone di atas & bawah).

    ---

    ### 2. Formulas Penurunan & Konversi Tegangan
    * **Konversi Tegangan:**
      $$p_{kPa} = p_{kg/cm^2} \times 98.0665$$
    * **Perubahan Angka Pori ($\Delta e$):**
      $$e = e_0 - \left( \frac{\Delta H}{H_0} \times (1 + e_0) \right)$$
    * **Waktu Penurunan $t_{90}$ Lapangan:**
      $$t_{90} = \frac{T_v \cdot H_{dr}^2}{C_v} \quad (T_v = 0.848)$$

    ---

    ### 3. Template Standar Tabel Data Excel
    Berikut adalah struktur tabel yang harus digunakan pada file Excel (`.xlsx`) atau CSV sebelum diunggah ke aplikasi:
    """)
    
    # Template Data Sample Table untuk Display di Panduan
    template_display = pd.DataFrame({
        'Waktu (Menit)': [0, 0, 0, 1, 2, 4, 8, 15, 30, 60, 120, 240, 480, 1440],
        'Waktu (Detik)': [6, 15, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        'p=0.25 kg/cm2': [0.012, 0.018, 0.025, 0.034, 0.045, 0.058, 0.072, 0.086, 0.098, 0.108, 0.115, 0.120, 0.124, 0.128],
        'p=0.50 kg/cm2': [0.045, 0.062, 0.080, 0.102, 0.130, 0.162, 0.198, 0.235, 0.268, 0.295, 0.315, 0.330, 0.340, 0.350],
        'p=1.00 kg/cm2': [0.110, 0.142, 0.178, 0.215, 0.260, 0.312, 0.370, 0.432, 0.485, 0.530, 0.562, 0.585, 0.600, 0.615],
        'p=2.00 kg/cm2': [0.240, 0.295, 0.355, 0.420, 0.495, 0.580, 0.672, 0.765, 0.840, 0.905, 0.955, 0.990, 1.015, 1.035],
        'p=4.00 kg/cm2': [0.480, 0.560, 0.650, 0.745, 0.855, 0.980, 1.115, 1.240, 1.345, 1.430, 1.500, 1.550, 1.585, 1.615],
        'p=8.00 kg/cm2': [0.890, 1.010, 1.150, 1.290, 1.450, 1.620, 1.800, 1.960, 2.080, 2.180, 2.250, 2.300, 2.335, 2.365]
    })
    
    st.dataframe(template_display, use_container_width=True)
    
    st.info("💡 **Catatan:** Kolom 1 adalah Waktu dalam Menit, Kolom 2 adalah Waktu sisa dalam Detik, dan Kolom 3 dst. adalah Pembacaan Dial Penurunan (mm) untuk setiap tahap beban.")
