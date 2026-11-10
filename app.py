import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import io
import math
import folium
from streamlit_folium import st_folium

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

# 3. Pembuat Single File Excel Gabungan Output
def buat_excel_output_gabungan(df_tabel_10_5, df_export, laporan_teknis, pembahasan_text, kesimpulan_text, Cc, t90_tahun, t90_bulan, total_bawah, total_atas):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tabel_10_5.to_excel(writer, index=False, sheet_name='Tabel_10_5_Analisis_Konsolidasi')
        df_export.to_excel(writer, index=False, sheet_name='Ringkasan_Konsolidasi')
        
        df_parameter = pd.DataFrame({
            'Parameter / Indikator': [
                'Indeks Kompresi (Cc)',
                'Waktu Penurunan t90 (Tahun)',
                'Waktu Penurunan t90 (Bulan)',
                'Estimasi RAB Batas Bawah (Rp)',
                'Estimasi RAB Batas Atas (Rp)',
                'Pembahasan Geoteknik',
                'Kesimpulan Geoteknik'
            ],
            'Nilai Output': [
                round(Cc, 4),
                round(t90_tahun, 2),
                round(t90_bulan, 0),
                total_bawah,
                total_atas,
                pembahasan_text.replace("\n", " "),
                kesimpulan_text.replace("\n", " ")
            ]
        })
        df_parameter.to_excel(writer, index=False, sheet_name='Laporan_Teknis_&_RAB')
        
    return output.getvalue()

# 4. Logika Analisis Oedometer
def proses_oedometer(df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs):
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

    H0_cm = H0 / 10.0
    Hs_cm = H0_cm / (1 + e0)

    stress_knm2 = [b * 100.0 for b in beban_list]
    log_stress = [round(np.log10(s), 2) if s > 0 else 0.0 for s in stress_knm2]
    
    vdr_div = [dh * 10.0 for dh in penurunan_akhir_list]
    vdr_cm = [dh / 10.0 for dh in penurunan_akhir_list]
    H_cm_list = [H0_cm - v_cm for v_cm in vdr_cm]
    Hv_cm_list = [h - Hs_cm for h in H_cm_list]
    e_calc_list = [hv / Hs_cm for hv in Hv_cm_list]

    df_tabel_10_5 = pd.DataFrame({
        'Stress (kN/m²)': stress_knm2,
        'Log-Stress (kN/m²)': log_stress,
        'VDR (div)': [round(v, 2) for v in vdr_div],
        'VDR (cm)': [round(v, 5) for v in vdr_cm],
        'H (cm)': [round(h, 2) for h in H_cm_list],
        'Hv (cm)': [round(hv, 2) for hv in Hv_cm_list],
        'e': [round(e_val, 2) for e_val in e_calc_list]
    })

    p_kpa = [b * 98.0665 for b in beban_list]
    e_list = e_calc_list
    
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

    cv_cm2_min = cv * (10000.0 / 525600.0)

    pembahasan_text = (
        f"10.7 Pembahasan\n\n"
        f"Berdasarkan data yang diperoleh, pengujian konsolidasi menunjukkan adanya penurunan tebal sampel "
        f"tanah seiring dengan penambahan beban vertikal secara bertahap. Hal ini membuktikan bahwa tanah yang "
        f"diuji bersifat kompresibel, di mana proses pemampatan terjadi akibat keluarnya air pori dari rongga tanah, "
        f"sesuai dengan prinsip konsolidasi Terzaghi. Berdasarkan hasil perhitungan, didapatkan nilai koefisien "
        f"konsolidasi (Cv) sebesar {cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit) dengan Indeks Kompresi (Cc) sebesar {Cc:.4f}. "
        f"Nilai Cv ini merupakan parameter krusial dalam menentukan kecepatan laju penurunan tanah di lapangan; "
        f"semakin kecil nilai Cv, maka semakin lambat proses konsolidasi yang terjadi. Selain itu, hubungan antara "
        f"angka pori (e) dan logaritma tekanan (log P) menunjukkan tren penurunan yang konsisten, yang mengonfirmasi "
        f"bahwa penambahan beban menyebabkan pengurangan volume rongga pori tanah. Hasil ini sangat penting untuk "
        f"memprediksi besarnya settlement atau penurunan yang mungkin terjadi pada struktur bangunan yang akan didirikan "
        f"di atas lapisan tanah tersebut di masa depan."
    )

    kesimpulan_text = (
        f"10.8 Kesimpulan\n\n"
        f"Berdasarkan hasil analisis uji konsolidasi yang dilakukan di laboratorium mekanika tanah, didapatkan "
        f"nilai koefisien konsolidasi (Cv) sebesar {cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit) dan waktu penurunan (t90) "
        f"sebesar {t90_tahun:.2f} tahun ({t90_bulan:.0f} bulan). Nilai ini merupakan parameter fundamental yang mendeskripsikan "
        f"laju pemampatan tanah dan keluarnya air pori saat menerima beban. Semakin kecil nilai Cv, semakin lama waktu "
        f"yang dibutuhkan oleh lapisan tanah tersebut untuk mencapai penurunan maksimumnya. Informasi ini sangat esensial "
        f"dalam tahap rekayasa geoteknik untuk merancang fondasi yang aman dari risiko penurunan diferensial di masa mendatang.\n\n"
        f"10.9 Daftar Pustaka\n"
        f"Das, Braja M. 2002. Mekanika Tanah (Prinsip-prinsip Rekayasa Geoteknis). Jilid 1. Jakarta: Erlangga.\n"
        f"Badan Standardisasi Nasional. 2011. Cara Uji Konsolidasi Tanah Satu Dimensi (SNI 2812:2011). Jakarta."
    )

    rab_text = f"ESTIMASI TOTAL BIAYA PROYEK ({luas_bangunan} m²):\nRp {total_bawah:,.0f} s/d Rp {total_atas:,.0f}"

    df_export = pd.DataFrame({
        'Beban (kg/cm2)': beban_list,
        'Beban (kPa)': p_kpa,
        'Penurunan Total (mm)': penurunan_akhir_list,
        'Angka Pori (e)': e_list
    })

    return laporan_teknis, pembahasan_text, kesimpulan_text, rab_text, fig_elogp, fig_time, df_export, df_tabel_10_5, Cc, t90_tahun, t90_bulan, total_bawah, total_atas

# 5. PENDEKLARASIAN TAB NAVIGASI UTAMA (Dibuat dulu di sini!)
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
        Gs = st.number_input("Berat Jenis Tanah (Gs)", value=2.65)
        tebal_lempung = st.number_input("Tebal Lempung (m)", value=10.0)
        cv = st.number_input("Cv (m²/thn)", value=1.5)
        drainase = st.radio("Kondisi Drainase", ["1 Arah", "2 Arah (Atas & Bawah)"])
        luas_bangunan = st.number_input("Luas Area Proyek (m²)", value=500.0)
        btn = st.button("🚀 Olah Data & Buat Grafik", type="primary")

    with col2:
        if uploaded_file and btn:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            laporan, pembahasan, kesimpulan, rab, fig_elogp, fig_time, df_export, df_tabel_10_5, Cc, t90_t, t90_b, rb_b, rb_a = proses_oedometer(
                df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs
            )
            
            subtab1, subtab2, subtab3, subtab4, subtab5, subtab6 = st.tabs([
                "📋 Tabel 10.5 Analisis", 
                "📝 Pembahasan & Kesimpulan", 
                "📄 Laporan Teknis", 
                "💰 Estimasi RAB", 
                "📈 Kurva e - log p", 
                "📉 Grafik Penurunan"
            ])
            
            with subtab1:
                st.subheader("Tabel 10.5 Analisis Data Konsolidasi")
                st.dataframe(df_tabel_10_5, use_container_width=True)
            with subtab2:
                st.markdown(f"### {pembahasan}")
                st.divider()
                st.markdown(f"### {kesimpulan}")
            with subtab3:
                st.code(laporan)
            with subtab4:
                st.success(rab)
            with subtab5:
                st.pyplot(fig_elogp)
            with subtab6:
                st.pyplot(fig_time)
                
            excel_gabungan_bytes = buat_excel_output_gabungan(df_tabel_10_5, df_export, laporan, pembahasan, kesimpulan, Cc, t90_t, t90_b, rb_b, rb_a)
            st.download_button(
                label="📦 Unduh Seluruh Output Lengkap (.xlsx)",
                data=excel_gabungan_bytes,
                file_name="Laporan_Lengkap_OedoCalc_Pro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- TAB 2: PETA LOKASI SAMPEL TANAH INTERAKTIF ---
with tab_peta:
    st.header("🗺️ Peta Pemetaan Lokasi Pengambilan Sampel Tanah (Borehole)")
    st.write("Anda dapat mengubah lokasi sampel dengan **memasukkan koordinat manual** atau **mengklik langsung pada peta**.")

    if 'latitude' not in st.session_state:
        st.session_state.latitude = -6.208800
    if 'longitude' not in st.session_state:
        st.session_state.longitude = 106.845600

    col_map_in, col_map_view = st.columns([1, 2])

    with col_map_in:
        st.subheader("📍 Input / Ubah Lokasi Sampel")
        nama_proyek = st.text_input("Nama Proyek / Titik BH", value="BH-01 Proyek Gedung A")
        
        lat_input = st.number_input("Latitude (Lintang)", value=float(st.session_state.latitude), format="%.6f", key="input_lat")
        lon_input = st.number_input("Longitude (Bujur)", value=float(st.session_state.longitude), format="%.6f", key="input_lon")
        
        st.session_state.latitude = lat_input
        st.session_state.longitude = lon_input

        kedalaman = st.number_input("Kedalaman Sampel (meter)", value=4.5)
        
        gmaps_url = f"https://www.google.com/maps?q={st.session_state.latitude},{st.session_state.longitude}"
        st.markdown(f"🔗 **[Buka Titik Ini di Google Maps Direct]({gmaps_url})**")

    with col_map_view:
        st.subheader("🌍 Klik pada Peta untuk Mengubah Lokasi")
        
        m = folium.Map(
            location=[st.session_state.latitude, st.session_state.longitude], 
            zoom_start=13,
            tiles="OpenStreetMap"
        )
        
        folium.Marker(
            [st.session_state.latitude, st.session_state.longitude],
            popup=f"<b>{nama_proyek}</b><br>Kedalaman: {kedalaman} m",
            tooltip=nama_proyek,
            icon=folium.Icon(color="red", icon="info-sign")
        ).add_to(m)

        map_data = st_folium(m, height=400, width=700)

        if map_data and map_data.get("last_clicked"):
            new_lat = map_data["last_clicked"]["lat"]
            new_lng = map_data["last_clicked"]["lng"]
            
            if new_lat != st.session_state.latitude or new_lng != st.session_state.longitude:
                st.session_state.latitude = new_lat
                st.session_state.longitude = new_lng
                st.rerun()

    st.divider()
    st.subheader("📋 Ringkasan Geolocation Terkini")
    st.info(f"**Titik Sampel:** {nama_proyek} | **Kedalaman:** {kedalaman} m | **Koordinat Aktif:** Lat `{st.session_state.latitude:.6f}`, Lon `{st.session_state.longitude:.6f}`")
