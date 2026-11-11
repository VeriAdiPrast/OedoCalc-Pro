import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import io
import math
import base64
import folium
from streamlit_folium import st_folium

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="OedoCalc Pro & Geolocation",
    page_icon="🧪",
    layout="wide"
)

# 2. Database Regional Harga Sewa Alat & Material Geoteknik Indonesia
DATA_REGIONAL = {
    "Jawa (Jabodetabek / Jawa Barat)": {
        "pvd_m": 45000, "preloading_m3": 120000, "rig_mobilization": 35000000, 
        "instrumentation": 15000000, "minipile_m": 220000, "hspd_mobilization": 25000000,
        "shallow_m2": 350000
    },
    "Jawa Tengah / Jawa Timur": {
        "pvd_m": 42000, "preloading_m3": 110000, "rig_mobilization": 30000000, 
        "instrumentation": 12000000, "minipile_m": 200000, "hspd_mobilization": 22000000,
        "shallow_m2": 320000
    },
    "Sumatera": {
        "pvd_m": 55000, "preloading_m3": 150000, "rig_mobilization": 50000000, 
        "instrumentation": 20000000, "minipile_m": 270000, "hspd_mobilization": 35000000,
        "shallow_m2": 420000
    },
    "Kalimantan": {
        "pvd_m": 65000, "preloading_m3": 175000, "rig_mobilization": 65000000, 
        "instrumentation": 25000000, "minipile_m": 310000, "hspd_mobilization": 45000000,
        "shallow_m2": 480000
    },
    "Sulawesi": {
        "pvd_m": 60000, "preloading_m3": 160000, "rig_mobilization": 55000000, 
        "instrumentation": 22000000, "minipile_m": 290000, "hspd_mobilization": 40000000,
        "shallow_m2": 450000
    },
    "Papua & Maluku": {
        "pvd_m": 85000, "preloading_m3": 230000, "rig_mobilization": 95000000, 
        "instrumentation": 35000000, "minipile_m": 420000, "hspd_mobilization": 70000000,
        "shallow_m2": 650000
    }
}

# 3. Fungsi Pembuat Template Excel Kosong
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

# 4. Konversi Matplotlib Figure ke Base64
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

# 5. Pembuat Laporan HTML/PDF Formal Siap Cetak
def buat_laporan_html_lengkap(df_tabel_10_5, fig_elogp, fig_time, pembahasan_text, kesimpulan_text, rab_html, nama_proyek, lat, lon, kedalaman, tebal_lempung, cv, Cc, t90_tahun, t90_bulan, wilayah):
    img_elogp = fig_to_base64(fig_elogp)
    img_time = fig_to_base64(fig_time)
    cv_cm2_min = cv * (10000.0 / 525600.0)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Laporan Hasil Uji Konsolidasi Oedometer</title>
        <style>
            @page {{ size: A4; margin: 20mm; }}
            body {{ font-family: 'Times New Roman', Times, serif; color: #111827; line-height: 1.6; font-size: 11pt; padding: 10px; }}
            .header-kop {{ text-align: center; border-bottom: 3px double #1e3a8a; padding-bottom: 10px; margin-bottom: 20px; }}
            .header-kop h1 {{ font-size: 16pt; margin: 0; color: #1e3a8a; text-transform: uppercase; letter-spacing: 1px; }}
            .header-kop p {{ font-size: 9pt; margin: 3px 0 0 0; color: #4b5563; font-style: italic; }}
            h2 {{ font-size: 12pt; color: #1e3a8a; border-bottom: 1px solid #94a3b8; padding-bottom: 4px; margin-top: 25px; text-transform: uppercase; }}
            h3 {{ font-size: 11pt; color: #0f172a; margin-top: 15px; margin-bottom: 5px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10pt; }}
            th, td {{ border: 1px solid #64748b; padding: 6px 10px; text-align: left; }}
            th {{ background-color: #f1f5f9; font-weight: bold; color: #1e293b; }}
            p {{ text-align: justify; text-indent: 30px; margin-top: 5px; margin-bottom: 10px; }}
            .img-container {{ display: flex; gap: 15px; margin-top: 15px; justify-content: center; }}
            .img-box {{ flex: 1; text-align: center; font-size: 9pt; font-weight: bold; }}
            .img-box img {{ width: 100%; max-width: 350px; height: auto; border: 1px solid #cbd5e1; margin-top: 5px; border-radius: 4px; }}
            .box-rab {{ background-color: #ffffff; border: 1px solid #cbd5e1; padding: 10px; font-size: 10pt; margin-top: 10px; }}
        </style>
    </head>
    <body>
        <div class="header-kop">
            <h1>Laporan Analisis Geoteknik & Uji Konsolidasi</h1>
            <p>Laboratorium Mekanika Tanah — Standard ASTM D2435 / SNI 2812:2011</p>
        </div>

        <h2>BAB 1. INFORMASI PROYEK & SAMPLING</h2>
        <table>
            <tr><th width="25%">Nama Proyek / Titik BH</th><td width="25%">{nama_proyek}</td><th width="25%">Kedalaman Sampel</th><td width="25%">{kedalaman} m</td></tr>
            <tr><th>Koordinat Geolocation</th><td>Lat: {lat:.6f}, Lon: {lon:.6f}</td><th>Wilayah Acuan Harga</th><td>{wilayah}</td></tr>
            <tr><th>Tebal Lapisan Lempung</th><td>{tebal_lempung} m</td><th>Koefisien Konsolidasi (Cv)</th><td>{cv:.4f} m²/thn</td></tr>
            <tr><th>Indeks Kompresi (Cc)</th><td>{Cc:.4f}</td><th>Waktu Penurunan (t90)</th><td>{t90_tahun:.2f} Thn ({t90_bulan:.0f} Bln)</td></tr>
        </table>

        <h2>BAB 2. GRAFIK & KURVA KONSOLIDASI</h2>
        <div class="img-container">
            <div class="img-box">
                <div>Gambar 2.1 Kurva Konsolidasi Utama (e - log p)</div>
                <img src="data:image/png;base64,{img_elogp}" />
            </div>
            <div class="img-box">
                <div>Gambar 2.2 Grafik Penurunan vs √Waktu</div>
                <img src="data:image/png;base64,{img_time}" />
            </div>
        </div>

        <h2>BAB 3. PEMBAHASAN & KESIMPULAN</h2>
        <h3>3.1 Pembahasan Geoteknik</h3>
        <p>Berdasarkan data yang diperoleh, pengujian konsolidasi menunjukkan adanya penurunan tebal sampel tanah seiring dengan penambahan beban vertikal secara bertahap. Hal ini membuktikan bahwa tanah yang diuji bersifat kompresibel, di mana proses pemampatan terjadi akibat keluarnya air pori dari rongga tanah, sesuai dengan prinsip konsolidasi Terzaghi.</p>
        <p>Berdasarkan hasil perhitungan, didapatkan nilai koefisien konsolidasi (Cv) sebesar <strong>{cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit)</strong> dengan Indeks Kompresi (Cc) sebesar <strong>{Cc:.4f}</strong>. Nilai Cv ini merupakan parameter krusial dalam menentukan kecepatan laju penurunan tanah di lapangan; semakin kecil nilai Cv, maka semakin lambat proses konsolidasi yang terjadi. Hubungan antara angka pori (e) dan logaritma tekanan (log P) menunjukkan tren penurunan yang konsisten, mengonfirmasi pengurangan volume rongga pori tanah.</p>

        <h3>3.2 Kesimpulan</h3>
        <p>Berdasarkan hasil analisis uji konsolidasi yang dilakukan di laboratorium mekanika tanah, didapatkan nilai koefisien konsolidasi (Cv) sebesar <strong>{cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit)</strong> dan waktu penurunan (t90) sebesar <strong>{t90_tahun:.2f} tahun ({t90_bulan:.0f} bulan)</strong>. Nilai ini merupakan parameter fundamental yang mendeskripsikan laju pemampatan tanah dan keluarnya air pori saat menerima beban.</p>

        <h3>3.3 Daftar Pustaka</h3>
        <ol style="font-size: 10pt; margin-top: 5px; padding-left: 20px;">
            <li>Das, Braja M. 2002. <em>Mekanika Tanah (Prinsip-prinsip Rekayasa Geoteknis)</em>. Jilid 1. Jakarta: Erlangga.</li>
            <li>Badan Standardisasi Nasional. 2011. <em>Cara Uji Konsolidasi Tanah Satu Dimensi (SNI 2812:2011)</em>. Jakarta.</li>
        </ol>

        <h2>BAB 4. REKOMENDASI TEKNIS & ESTIMASI RAB REGIONAL</h2>
        <div class="box-rab">
            {rab_html}
        </div>
    </body>
    </html>
    """
    return html_content.encode('utf-8')

# 6. Pembuat Single File Excel Gabungan Output
def buat_excel_output_gabungan(df_tabel_10_5, df_export, laporan_teknis, pembahasan_text, kesimpulan_text, Cc, t90_tahun, t90_bulan, df_rab):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tabel_10_5.to_excel(writer, index=False, sheet_name='Tabel_10_5_Analisis_Konsolidasi')
        df_export.to_excel(writer, index=False, sheet_name='Ringkasan_Konsolidasi')
        df_rab.to_excel(writer, index=False, sheet_name='Detail_RAB_Regional')
        
        df_parameter = pd.DataFrame({
            'Parameter / Indikator': [
                'Indeks Kompresi (Cc)', 'Waktu Penurunan t90 (Tahun)', 'Waktu Penurunan t90 (Bulan)',
                'Pembahasan Geoteknik', 'Kesimpulan Geoteknik'
            ],
            'Nilai Output': [
                round(Cc, 4), round(t90_tahun, 2), round(t90_bulan, 0),
                pembahasan_text.replace("\n", " "), kesimpulan_text.replace("\n", " ")
            ]
        })
        df_parameter.to_excel(writer, index=False, sheet_name='Laporan_Teknis')
        
    return output.getvalue()

# 7. Logika Analisis Oedometer & Pembuat RAB Detil per Wilayah
def proses_oedometer(df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs, wilayah_terpilih):
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

    # --- PERHITUNGAN RAB DETIL DAN REKOMENDASI PERBAIKAN ---
    reg_price = DATA_REGIONAL.get(wilayah_terpilih, DATA_REGIONAL["Jawa (Jabodetabek / Jawa Barat)"])
    
    if Cc < 0.2:
        kemampatan = "RENDAH (Low Compressibility) - Tanah Padat"
        rekomendasi_nama = "Pondasi Dangkal (Footplat / Batu Kali) & Pemadatan Standard"
        alat_butuh = "Bulldozer 100 HP, Vibro Roller 10 Ton, Excavator Backhoe 0.8 m³"
        
        pekerjaan_items = [
            ("Galian & Pembersihan Lahan", luas_bangunan * 0.5, "m³", reg_price["shallow_m2"] * 0.2),
            ("Urugan Pasir & Lapisan Subbase", luas_bangunan * 0.3, "m³", reg_price["shallow_m2"] * 0.4),
            ("Pekerjaan Lean Concrete & Pembesian", luas_bangunan, "m²", reg_price["shallow_m2"] * 0.4),
            ("Mobilisasi / Demobilisasi Alat Berat", 1, "ls", 15000000)
        ]
    elif 0.2 <= Cc <= 0.4:
        kemampatan = "SEDANG (Medium Compressibility) - Lempung Sedang"
        rekomendasi_nama = "Pondasi Minipile (Tiang Pancang Presisi / HSPD)"
        alat_butuh = "Hydraulic Static Pile Driver (HSPD) / Mobile Rig Pancang, Mobile Crane 25 Ton, Dumptruck 12 m³"
        
        jumlah_tiang = math.ceil((luas_bangunan / 12.0))
        panjang_tiang_total = jumlah_tiang * tebal_lempung
        
        pekerjaan_items = [
            ("Pengadaan Tiang Pancang Minipile 28x28", panjang_tiang_total, "meter", reg_price["minipile_m"] * 0.6),
            ("Jasa Pemancangan dengan Rig HSPD", panjang_tiang_total, "meter", reg_price["minipile_m"] * 0.4),
            ("Pekerjaan Galian & Pile Cap Structural", luas_bangunan, "m²", reg_price["shallow_m2"] * 0.8),
            ("Mobilisasi Rig Pemancang & Crane", 1, "ls", reg_price["hspd_mobilization"])
        ]
    else:
        kemampatan = "TINGGI (High Compressibility) - Tanah Sangat Lunak"
        rekomendasi_nama = "Perbaikan Tanah Masif (PVD + Preloading Timbunan Tanah)"
        alat_butuh = "PVD Rig Inserter (Mandrel Rig), Excavator 20 Ton, Vibro Roller 12 Ton, Dumptruck 20 m³, Instrument Settlement Plate & Inclinometer"
        
        # Grid PVD 1.2m x 1.2m
        titik_pvd = math.ceil(luas_bangunan / 1.44)
        panjang_pvd_total = titik_pvd * (tebal_lempung + 1.0)
        volume_preloading = luas_bangunan * 3.0 # Timbunan tinggi 3m
        
        pekerjaan_items = [
            ("Pengadaan & Penetrasi PVD (Prefabricated Vertical Drain)", panjang_pvd_total, "meter", reg_price["pvd_m"]),
            ("Pengadaan & Pemadatan Timbunan Preloading (Tanah Pilihan)", volume_preloading, "m³", reg_price["preloading_m3"]),
            ("Mobilisasi & Demobilisasi PVD Rig Inserter", 1, "ls", reg_price["rig_mobilization"]),
            ("Pemasangan Instrumentasi Geoteknik (Settlement Plate)", 1, "ls", reg_price["instrumentation"])
        ]

    # Buat Dataframe RAB
    list_uraian, list_vol, list_sat, list_harga, list_total = [], [], [], [], []
    total_rab_keseluruhan = 0

    for item in pekerjaan_items:
        uraian, vol, sat, hg = item
        subtotal = vol * hg
        list_uraian.append(uraian)
        list_vol.append(vol)
        list_sat.append(sat)
        list_harga.append(hg)
        list_total.append(subtotal)
        total_rab_keseluruhan += subtotal

    df_rab = pd.DataFrame({
        'Uraian Pekerjaan / Material / Alat': list_uraian,
        'Volume': list_vol,
        'Satuan': list_sat,
        'Harga Satuan Regional (Rp)': list_harga,
        'Total Biaya (Rp)': list_total
    })

    laporan_teknis = (
        f"--- HASIL ANALISIS DATA MENTAH LAB ---\n"
        f"Beban Diuji (kg/cm²)  : {beban_list}\n"
        f"Indeks Kompresi (Cc)  : {Cc:.4f}\n"
        f"Klasifikasi Tanah     : {kemampatan}\n\n"
        f"--- ESTIMASI WAKTU PENURUNAN LAPANGAN ---\n"
        f"Tebal Lapisan Lempung : {tebal_lempung} meter\n"
        f"Waktu Penurunan t90   : {t90_tahun:.2f} Tahun ({t90_bulan:.0f} Bulan)\n\n"
        f"--- REKOMENDASI TEKNIS ---\n"
        f"Metode Penanganan     : {rekomendasi_nama}\n"
        f"Daftar Alat Dibutuhkan: {alat_butuh}\n"
    )

    cv_cm2_min = cv * (10000.0 / 525600.0)

    pembahasan_text = (
        f"Berdasarkan data yang diperoleh, pengujian konsolidasi menunjukkan adanya penurunan tebal sampel tanah seiring dengan penambahan beban vertikal secara bertahap. Hal ini membuktikan bahwa tanah yang diuji bersifat kompresibel, di mana proses pemampatan terjadi akibat keluarnya air pori dari rongga tanah, sesuai dengan prinsip konsolidasi Terzaghi. Berdasarkan hasil perhitungan, didapatkan nilai koefisien konsolidasi (Cv) sebesar {cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit) dengan Indeks Kompresi (Cc) sebesar {Cc:.4f}."
    )

    kesimpulan_text = (
        f"Berdasarkan hasil analisis uji konsolidasi laboratorium mekanika tanah, didapatkan nilai koefisien konsolidasi (Cv) sebesar {cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit) dan waktu penurunan (t90) sebesar {t90_tahun:.2f} tahun ({t90_bulan:.0f} bulan). Informasi ini sangat esensial dalam tahap rekayasa geoteknik untuk merancang fondasi yang aman."
    )

    return laporan_teknis, pembahasan_text, kesimpulan_text, df_rab, total_rab_keseluruhan, rekomendasi_nama, alat_butuh, fig_elogp, fig_time, df_export, df_tabel_10_5, Cc, t90_tahun, t90_bulan

# 8. PENDEKLARASIAN TAB NAVIGASI UTAMA
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
        wilayah_terpilih = st.selectbox(
            "📍 Pilih Wilayah Acuan Harga Sewa Alat & Bahan",
            options=list(DATA_REGIONAL.keys())
        )
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
            laporan, pembahasan, kesimpulan, df_rab, total_rab, rek_nama, alat_butuh, fig_elogp, fig_time, df_export, df_tabel_10_5, Cc, t90_t, t90_b = proses_oedometer(
                df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs, wilayah_terpilih
            )
            
            subtab1, subtab2, subtab3, subtab4, subtab5, subtab6 = st.tabs([
                "📋 Tabel 10.5 Analisis", 
                "📝 Pembahasan & Kesimpulan", 
                "📄 Laporan Teknis", 
                "💰 Estimasi RAB Regional", 
                "📈 Kurva e - log p", 
                "📉 Grafik Penurunan"
            ])
            
            with subtab1:
                st.subheader("Tabel 10.5 Analisis Data Konsolidasi")
                st.dataframe(df_tabel_10_5, use_container_width=True)
            with subtab2:
                st.markdown(f"### 3.1 Pembahasan Geoteknik\n{pembahasan}")
                st.divider()
                st.markdown(f"### 3.2 Kesimpulan\n{kesimpulan}")
            with subtab3:
                st.code(laporan)
            with subtab4:
                st.subheader(f"🏗️ Rekomendasi Teknis: {rek_nama}")
                st.info(f"**Tipe Alat Berat Dibutuhkan:** {alat_butuh}")
                st.subheader(f"📊 Rincian RAB Regional ({wilayah_terpilih})")
                st.dataframe(df_rab.style.format({'Harga Satuan Regional (Rp)': 'Rp {:,.0f}', 'Total Biaya (Rp)': 'Rp {:,.0f}'}), use_container_width=True)
                st.success(f"### ESTIMASI TOTAL BIAYA PROYEK: Rp {total_rab:,.0f}")
            with subtab5:
                st.pyplot(fig_elogp)
            with subtab6:
                st.pyplot(fig_time)
                
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                excel_gabungan_bytes = buat_excel_output_gabungan(df_tabel_10_5, df_export, laporan, pembahasan, kesimpulan, Cc, t90_t, t90_b, df_rab)
                st.download_button(
                    label="📦 Unduh Seluruh Output Lengkap (.xlsx)",
                    data=excel_gabungan_bytes,
                    file_name="Laporan_Lengkap_OedoCalc_Pro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_btn2:
                rab_html_table = f"<h3>Rekomendasi Metode: {rek_nama}</h3><p><strong>Alat Berat Dibutuhkan:</strong> {alat_butuh}</p>" + df_rab.to_html(index=False) + f"<h4>TOTAL RAB ESTIMASI REGIONAL ({wilayah_terpilih}): Rp {total_rab:,.0f}</h4>"
                html_bytes = buat_laporan_html_lengkap(
                    df_tabel_10_5, fig_elogp, fig_time, pembahasan, kesimpulan, rab_html_table, 
                    "BH-01 Proyek Gedung", st.session_state.curr_lat, st.session_state.curr_lon, 
                    4.5, tebal_lempung, cv, Cc, t90_t, t90_b, wilayah_terpilih
                )
                st.download_button(
                    label="📄 Unduh Rekap Laporan PDF / HTML (.html)",
                    data=html_bytes,
                    file_name="Laporan_Rekap_OedoCalc_Pro.html",
                    mime="text/html"
                )

# --- TAB 2: PETA LOKASI SAMPEL TANAH INTERAKTIF ---
with tab_peta:
    st.header("🗺️ Peta Pemetaan Lokasi Pengambilan Sampel Tanah (Borehole)")
    st.write("Lokasi dapat dipindahkan dengan **mengklik titik baru di atas peta** atau **mengubah angka koordinat** di sebelah kiri.")

    col_map_in, col_map_view = st.columns([1, 2])

    with col_map_in:
        st.subheader("📍 Input / Ubah Lokasi Sampel")
        nama_proyek = st.text_input("Nama Proyek / Titik BH", value="BH-01 Proyek Gedung A", key="peta_nama")
        
        lat_input = st.number_input(
            "Latitude (Lintang)", 
            value=float(st.session_state.curr_lat), 
            format="%.6f", 
            key="peta_lat_input"
        )
        lon_input = st.number_input(
            "Longitude (Bujur)", 
            value=float(st.session_state.curr_lon), 
            format="%.6f", 
            key="peta_lon_input"
        )
        
        st.session_state.curr_lat = lat_input
        st.session_state.curr_lon = lon_input

        kedalaman = st.number_input("Kedalaman Sampel (meter)", value=4.5, key="peta_depth")
        
        gmaps_url = f"https://www.google.com/maps?q={st.session_state.curr_lat},{st.session_state.curr_lon}"
        st.markdown(f"🔗 **[Buka Titik Ini di Google Maps Direct]({gmaps_url})**")

    with col_map_view:
        st.subheader("🌍 Klik pada Peta untuk Menggeser Posisi")
        
        m = folium.Map(
            location=[st.session_state.curr_lat, st.session_state.curr_lon], 
            zoom_start=13,
            tiles="OpenStreetMap"
        )
        
        folium.Marker(
            [st.session_state.curr_lat, st.session_state.curr_lon],
            popup=f"<b>{nama_proyek}</b><br>Kedalaman: {kedalaman} m",
            tooltip=nama_proyek,
            icon=folium.Icon(color="red", icon="info-sign")
        ).add_to(m)

        map_data = st_folium(m, height=400, width=700, key="folium_interactive_map")

        if map_data and map_data.get("last_clicked"):
            click_lat = map_data["last_clicked"]["lat"]
            click_lng = map_data["last_clicked"]["lng"]
            
            if abs(click_lat - st.session_state.curr_lat) > 0.000001 or abs(click_lng - st.session_state.curr_lon) > 0.000001:
                st.session_state.curr_lat = click_lat
                st.session_state.curr_lon = click_lng
                st.rerun()

    st.divider()
    st.subheader("📋 Ringkasan Geolocation Terkini")
    st.info(f"**Titik Sampel:** {nama_proyek} | **Kedalaman:** {kedalaman} m | **Koordinat Aktif:** Lat `{st.session_state.curr_lat:.6f}`, Lon `{st.session_state.curr_lon:.6f}`")
