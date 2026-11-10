import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import io
import math

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="OedoCalc Pro & Geolocation",
    page_icon="🧪",
    layout="wide"
)

# 2. Fungsi Pembuat Template Excel Kosong
def buat_template_excel_kosong():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        waktu_menit = [0, 0, 0, 1, 2, 4, 8, 15, 30, 60, 120, 240, 480, 1440]
        waktu_detik = [6, 15, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        
        data = {
            'Waktu (Menit)': waktu_menit,
            'Waktu (Detik)': waktu_detik,
            'p=0.25 kg/cm2': [None] * len(waktu_menit),
            'p=0.50 kg/cm2': [None] * len(waktu_menit),
            'p=1.00 kg/cm2': [None] * len(waktu_menit),
            'p=2.00 kg/cm2': [None] * len(waktu_menit),
            'p=4.00 kg/cm2': [None] * len(waktu_menit),
            'p=8.00 kg/cm2': [None] * len(waktu_menit)
        }
        df_template = pd.DataFrame(data)
        df_template.to_excel(writer, index=False, sheet_name='Data_Oedometer')
    return output.getvalue()

# 3. Fungsi Pembuat Single File Excel Gabungan Output
def buat_excel_output_gabungan(df_export, laporan_teknis, rab_text, Cc, t90_tahun, t90_bulan, total_bawah, total_atas):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Ringkasan Hasil Uji & Angka Pori
        df_export.to_excel(writer, index=False, sheet_name='Ringkasan_Konsolidasi')
        
        # Sheet 2: Laporan Teknis & Parameter
        df_parameter = pd.DataFrame({
            'Parameter / Indikator': [
                'Indeks Kompresi (Cc)',
                'Waktu Penurunan t90 (Tahun)',
                'Waktu Penurunan t90 (Bulan)',
                'Estimasi RAB Batas Bawah (Rp)',
                'Estimasi RAB Batas Atas (Rp)'
            ],
            'Nilai Output': [
                round(Cc, 4),
                round(t90_tahun, 2),
                round(t90_bulan, 0),
                total_bawah,
                total_atas
            ]
        })
        df_parameter.to_excel(writer, index=False, sheet_name='Laporan_Teknis_&_RAB')
        
    return output.getvalue()

# 4. Logika Analisis Oedometer
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

    # Klasifikasi & Rekomendasi
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

    return laporan_teknis, rab_text, fig_elogp, fig_time, df_export, Cc, t90_tahun, t90_bulan, total_bawah, total_atas

# 5. TAB NAVIGASI UTAMA
tab_kalkulator, tab_peta = st.tabs(["🚀 Aplikasi Kalkulator", "🗺️ Peta Lokasi Sampel Tanah"])

# --- TAB 1: KALKULATOR GEOTEKNIK ---
with tab_kalkulator:
    st.title("🧪 OedoCalc Pro: Analisis Konsolidasi Tanah")
    
    col_dl, col_space = st.columns([1, 2])
    with col_dl:
        st.download_button(
            label="📥 Unduh Template Format Excel Kosong (.xlsx)",
            data=buat_template_excel_kosong(),
            file_name="Template_Kosong_OedoCalc_Pro.xlsx",
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
            laporan, rab, fig_elogp, fig_time, df_export, Cc, t90_t, t90_b, rb_b, rb_a = proses_oedometer(
                df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan
            )
            
            subtab1, subtab2, subtab3, subtab4 = st.tabs(["📄 Laporan Teknis", "💰 Estimasi RAB", "📈 Kurva e - log p", "📉 Grafik Penurunan"])
            with subtab1:
                st.code(laporan)
            with subtab2:
                st.success(rab)
            with subtab3:
                st.pyplot(fig_elogp)
            with subtab4:
                st.pyplot(fig_time)
                
            # Tombol unduh single file gabungan (.xlsx)
            excel_gabungan_bytes = buat_excel_output_gabungan(df_export, laporan, rab, Cc, t90_t, t90_b, rb_b, rb_a)
            st.download_button(
                label="📦 Unduh Seluruh Output Lengkap (.xlsx)",
                data=excel_gabungan_bytes,
                file_name="Laporan_Lengkap_OedoCalc_Pro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- TAB 2: PETA LOKASI SAMPEL TANAH ---
with tab_peta:
    st.header("🗺️ Peta Pemetaan Lokasi Pengambilan Sampel Tanah (Borehole)")
    st.write("Gunakan fitur ini untuk melacak dan memetakan koordinat titik bor tempat sampel tanah diambil di lapangan.")

    col_map_in, col_map_view = st.columns([1, 2])

    with col_map_in:
        st.subheader("📍 Data Koordinat Sampel")
        nama_proyek = st.text_input("Nama Proyek / Titik BH", value="BH-01 Proyek Gedung A")
        lat = st.number_input("Latitude (Lintang)", value=-6.2088, format="%.6f")
        lon = st.number_input("Longitude (Bujur)", value=106.8456, format="%.6f")
        kedalaman = st.number_input("Kedalaman Sampel (meter)", value=4.5)
        
        gmaps_url = f"https://www.google.com/maps?q={lat},{lon}"
        st.markdown(f"🔗 **[Buka di Google Maps Direct]({gmaps_url})**")

    with col_map_view:
        st.subheader("🌍 Visualisasi Peta Poin Lokasi")
        df_lokasi = pd.DataFrame({'lat': [lat], 'lon': [lon]})
        st.map(df_lokasi, zoom=14)

    st.divider()
    st.subheader("📋 Ringkasan Geolocation")
    st.info(f"**Titik Sampel:** {nama_proyek} | **Kedalaman:** {kedalaman} m | **Koordinat:** Lat `{lat}`, Lon `{lon}`")
