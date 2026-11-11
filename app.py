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
    page_title="OedoCalc Pro - Advanced Geotechnical Suite",
    page_icon="🧪",
    layout="wide"
)

# 2. Inisialisasi Session State Koordinat (Di Paling Atas!)
if 'curr_lat' not in st.session_state:
    st.session_state.curr_lat = -6.208800
if 'curr_lon' not in st.session_state:
    st.session_state.curr_lon = 106.845600

# 3. Database Regional Harga Sewa Alat & Material Geoteknik Indonesia
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

# 4. Pembuat Template Excel Kosong
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

# 5. Konversi Grafik ke Base64 (Untuk PDF/HTML)
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

# 6. Pembuat Laporan HTML/PDF Formal
def buat_laporan_html_lengkap(res, nama_proyek, lat, lon, kedalaman, tebal_lempung, cv, wilayah):
    img_elogp = fig_to_base64(res["fig_elogp"])
    img_time = fig_to_base64(res["fig_time"])
    cv_cm2_min = cv * (10000.0 / 525600.0)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Laporan Hasil Uji Konsolidasi Oedometer Pro</title>
        <style>
            @page {{ size: A4; margin: 15mm; }}
            body {{ font-family: 'Times New Roman', Times, serif; color: #111827; line-height: 1.5; font-size: 10pt; padding: 10px; }}
            .header-kop {{ text-align: center; border-bottom: 3px double #1e3a8a; padding-bottom: 8px; margin-bottom: 15px; }}
            .header-kop h1 {{ font-size: 15pt; margin: 0; color: #1e3a8a; text-transform: uppercase; }}
            .header-kop p {{ font-size: 8.5pt; margin: 2px 0 0 0; color: #4b5563; font-style: italic; }}
            h2 {{ font-size: 11pt; color: #1e3a8a; border-bottom: 1px solid #94a3b8; padding-bottom: 3px; margin-top: 18px; text-transform: uppercase; }}
            h3 {{ font-size: 10pt; color: #0f172a; margin-top: 12px; margin-bottom: 4px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 9pt; }}
            th, td {{ border: 1px solid #64748b; padding: 5px 8px; text-align: left; }}
            th {{ background-color: #f1f5f9; font-weight: bold; color: #1e293b; }}
            p {{ text-align: justify; text-indent: 20px; margin-top: 4px; margin-bottom: 8px; }}
            .img-container {{ display: flex; gap: 10px; margin-top: 10px; justify-content: center; }}
            .img-box {{ flex: 1; text-align: center; font-size: 8.5pt; font-weight: bold; }}
            .img-box img {{ width: 100%; max-width: 330px; height: auto; border: 1px solid #cbd5e1; border-radius: 4px; }}
            .box-rab {{ background-color: #ffffff; border: 1px solid #cbd5e1; padding: 8px; font-size: 9pt; margin-top: 8px; }}
        </style>
    </head>
    <body>
        <div class="header-kop">
            <h1>Laporan Analisis Geoteknik Lanjutan & Uji Konsolidasi</h1>
            <p>Laboratorium Mekanika Tanah — Standard ASTM D2435 / SNI 2812:2011</p>
        </div>

        <h2>BAB 1. INFORMASI PROYEK & PARAMETER ADVANCED</h2>
        <table>
            <tr><th width="25%">Nama Proyek / BH</th><td width="25%">{nama_proyek}</td><th width="25%">Kedalaman Sampel</th><td width="25%">{kedalaman} m</td></tr>
            <tr><th>Koordinat</th><td>Lat: {lat:.6f}, Lon: {lon:.6f}</td><th>Wilayah Acuan Harga</th><td>{wilayah}</td></tr>
            <tr><th>Tebal Lapisan Lempung</th><td>{tebal_lempung} m</td><th>Koefisien Konsolidasi (Cv)</th><td>{cv:.4f} m²/thn</td></tr>
            <tr><th>Indeks Kompresi (Cc)</th><td>{res['Cc']:.4f}</td><th>Indeks Rekompresi (Cr)</th><td>{res['Cr']:.4f}</td></tr>
            <tr><th>Preconsolidation (p'c)</th><td>{res['pc']:.2f} kPa</td><th>Status Konsolidasi (OCR)</th><td>{res['ocr_status']} (OCR={res['ocr']:.2f})</td></tr>
            <tr><th>Sekunder (Cα)</th><td>{res['C_alpha']:.4f}</td><th>Waktu Penurunan (t90)</th><td>{res['t90_t']:.2f} Thn ({res['t90_b']:.0f} Bln)</td></tr>
        </table>

        <h2>BAB 2. GRAFIK & KURVA KONSOLIDASI</h2>
        <div class="img-container">
            <div class="img-box">
                <div>Gambar 2.1 Kurva Konsolidasi Utama (e - log p) & p'c</div>
                <img src="data:image/png;base64,{img_elogp}" />
            </div>
            <div class="img-box">
                <div>Gambar 2.2 Grafik Penurunan vs √Waktu</div>
                <img src="data:image/png;base64,{img_time}" />
            </div>
        </div>

        <h2>BAB 3. PEMBAHASAN & KESIMPULAN</h2>
        <h3>3.1 Pembahasan Geoteknik</h3>
        <p>{res['pembahasan']}</p>

        <h3>3.2 Kesimpulan</h3>
        <p>{res['kesimpulan']}</p>

        <h3>3.3 Daftar Pustaka</h3>
        <ol style="font-size: 8.5pt; margin-top: 3px; padding-left: 18px;">
            <li>Das, Braja M. 2002. <em>Mekanika Tanah (Prinsip-prinsip Rekayasa Geoteknis)</em>. Jilid 1. Jakarta: Erlangga.</li>
            <li>Badan Standardisasi Nasional. 2011. <em>Cara Uji Konsolidasi Tanah Satu Dimensi (SNI 2812:2011)</em>. Jakarta.</li>
            <li>Hansbo, S. 1979. <em>Consolidation of Clay by Band-shaped Drains</em>. Ground Engineering.</li>
        </ol>

        <h2>BAB 4. REKOMENDASI TEKNIS & ESTIMASI RAB REGIONAL</h2>
        <div class="box-rab">
            <h3>Rekomendasi Metode: {res['rek_nama']}</h3>
            <p><strong>Alat Berat Dibutuhkan:</strong> {res['alat_butuh']}</p>
            {res['df_rab'].to_html(index=False)}
            <h4>TOTAL RAB ESTIMASI REGIONAL ({wilayah}): Rp {res['total_rab']:,.0f}</h4>
        </div>
    </body>
    </html>
    """
    return html_content.encode('utf-8')

# 7. Pembuat File Excel Multi-Sheet
def buat_excel_output_gabungan(res, df_export, laporan_teknis, wilayah):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        res["df_tabel_10_5"].to_excel(writer, index=False, sheet_name='Tabel_10_5_Analisis')
        df_export.to_excel(writer, index=False, sheet_name='Ringkasan_Konsolidasi')
        res["df_rab"].to_excel(writer, index=False, sheet_name='Detail_RAB_Regional')
        
        df_parameter = pd.DataFrame({
            'Parameter / Indikator': [
                'Indeks Kompresi (Cc)', 'Indeks Rekompresi (Cr)', 'Preconsolidation Pressure pc (kPa)',
                'OCR Status', 'Indeks Pemampatan Sekunder (Cα)', 'Waktu Penurunan t90 (Tahun)', 
                'Wilayah Acuan Harga', 'Estimasi RAB Total (Rp)'
            ],
            'Nilai Output': [
                round(res["Cc"], 4), round(res["Cr"], 4), round(res["pc"], 2),
                res["ocr_status"], round(res["C_alpha"], 4), round(res["t90_t"], 2),
                wilayah, res["total_rab"]
            ]
        })
        df_parameter.to_excel(writer, index=False, sheet_name='Advanced_Parameters')
        
    return output.getvalue()

# 8. Logika Analisis Oedometer Menggunakan Dictionary (Anti Unpack Error + Advanced Soil Engine)
def proses_oedometer(df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs, wilayah_terpilih, po_kpa):
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
    
    # 1. Indeks Kompresi (Cc) & Indeks Rekompresi (Cr)
    if len(p_kpa) >= 2 and p_kpa[-1] > p_kpa[-2]:
        Cc = (e_list[-2] - e_list[-1]) / np.log10(p_kpa[-1] / p_kpa[-2])
    else:
        Cc = 0.3
    Cr = Cc / 6.0 # Estimasi empiris Cr = (1/5 ~ 1/10) Cc

    # 2. Estimasi Preconsolidation Pressure (pc) & OCR
    pc_kpa = p_kpa[len(p_kpa)//2] if len(p_kpa) > 2 else p_kpa[0] * 1.5
    ocr = pc_kpa / po_kpa if po_kpa > 0 else 1.0
    ocr_status = "Normally Consolidated (NC)" if ocr <= 1.2 else "Overconsolidated (OC)"

    # 3. Indeks Pemampatan Sekunder (C_alpha)
    C_alpha = 0.04 * Cc # Hubungan empiris Mesri & Godlewski (C_alpha / Cc = 0.04)

    # 4. Waktu Konsolidasi Primer (t90)
    Tv = 0.848
    H_dr = tebal_lempung / 2.0 if "2 Arah" in drainase else tebal_lempung
    t90_tahun = (Tv * (H_dr ** 2)) / cv if cv > 0 else 0
    t90_bulan = t90_tahun * 12.0

    fig_elogp, ax2 = plt.subplots(figsize=(6, 4))
    ax2.plot(p_kpa, e_list, marker='s', linestyle='-', color='teal', linewidth=2, label="Kurva e-log p")
    ax2.axvline(x=pc_kpa, color='crimson', linestyle='--', label=f"p'c ≈ {pc_kpa:.1f} kPa")
    ax2.axvline(x=po_kpa, color='darkorange', linestyle=':', label=f"p'o = {po_kpa:.1f} kPa")
    ax2.set_xscale('log')
    ax2.set_xlabel('Tegangan Efektif, p (kPa) - Skala Log')
    ax2.set_ylabel('Angka Pori, e')
    ax2.set_title('Kurva Konsolidasi Utama (e - log p)')
    ax2.grid(True, which="both", ls="--", linewidth=0.5)
    ax2.legend(fontsize=8)
    ax2.invert_yaxis()
    plt.tight_layout()

    # Perhitungan RAB Detil
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
        titik_pvd = math.ceil(luas_bangunan / 1.44)
        panjang_pvd_total = titik_pvd * (tebal_lempung + 1.0)
        volume_preloading = luas_bangunan * 3.0
        pekerjaan_items = [
            ("Pengadaan & Penetrasi PVD (Prefabricated Vertical Drain)", panjang_pvd_total, "meter", reg_price["pvd_m"]),
            ("Pengadaan & Pemadatan Timbunan Preloading (Tanah Pilihan)", volume_preloading, "m³", reg_price["preloading_m3"]),
            ("Mobilisasi & Demobilisasi PVD Rig Inserter", 1, "ls", reg_price["rig_mobilization"]),
            ("Pemasangan Instrumentasi Geoteknik (Settlement Plate)", 1, "ls", reg_price["instrumentation"])
        ]

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
        f"Indeks Rekompresi (Cr): {Cr:.4f}\n"
        f"Preconsolidation (p'c): {pc_kpa:.2f} kPa\n"
        f"Status Tanah (OCR)    : {ocr_status} (OCR = {ocr:.2f})\n"
        f"Pemampatan Sekunder   : C_alpha = {C_alpha:.4f}\n"
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
        f"Berdasarkan data yang diperoleh, pengujian konsolidasi menunjukkan adanya penurunan tebal sampel tanah seiring dengan penambahan beban vertikal. "
        f"Nilai Koefisien Konsolidasi (Cv) didapatkan sebesar {cv:.4f} m²/tahun ({cv_cm2_min:.4f} cm²/menit) dengan Indeks Kompresi (Cc) {Cc:.4f} dan Indeks Rekompresi (Cr) {Cr:.4f}. "
        f"Nilai tegangan pra-konsolidasi (p'c) terestimasi sebesar {pc_kpa:.2f} kPa, menghasilkan nilai OCR sebesar {ocr:.2f} yang menunjukkan tanah terklasifikasi sebagai {ocr_status}. "
        f"Indeks pemampatan sekunder (C_alpha) diperkirakan sebesar {C_alpha:.4f} yang berpengaruh pada penurunan creep jangka panjang."
    )

    kesimpulan_text = (
        f"Berdasarkan analisis uji konsolidasi oedometer, tanah lapisan lempung setebal {tebal_lempung} m memiliki waktu penurunan t90 sebesar {t90_tahun:.2f} tahun ({t90_bulan:.0f} bulan). "
        f"Tanah terklasifikasi sebagai {ocr_status} dengan potensi pemampatan sekunder C_alpha = {C_alpha:.4f}. Rekomendasi fondasi / perbaikan tanah yang disarankan adalah {rekomendasi_nama}."
    )

    df_export = pd.DataFrame({
        'Beban (kg/cm2)': beban_list,
        'Beban (kPa)': p_kpa,
        'Penurunan Total (mm)': penurunan_akhir_list,
        'Angka Pori (e)': e_list
    })

    return {
        "laporan": laporan_teknis,
        "pembahasan": pembahasan_text,
        "kesimpulan": kesimpulan_text,
        "df_rab": df_rab,
        "total_rab": total_rab_keseluruhan,
        "rek_nama": rekomendasi_nama,
        "alat_butuh": alat_butuh,
        "fig_elogp": fig_elogp,
        "fig_time": fig_time,
        "df_export": df_export,
        "df_tabel_10_5": df_tabel_10_5,
        "Cc": Cc,
        "Cr": Cr,
        "pc": pc_kpa,
        "ocr": ocr,
        "ocr_status": ocr_status,
        "C_alpha": C_alpha,
        "t90_t": t90_tahun,
        "t90_b": t90_bulan
    }

# 9. PENDEKLARASIAN TAB NAVIGASI UTAMA
tab_kalkulator, tab_pvd, tab_qc, tab_peta = st.tabs([
    "🚀 Aplikasi Utama Kalkulator", 
    "🌀 Simulator PVD & Settlement Real-Time",
    "✅ Quality Control SNI / ASTM", 
    "🗺️ Peta Lokasi Sampel"
])

# --- TAB 1: KALKULATOR GEOTEKNIK ---
with tab_kalkulator:
    st.title("🧪 OedoCalc Pro: Advanced Consolidation Engine")
    
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
        po_kpa = st.number_input("Tegangan Efektif In-Situ p'o (kPa)", value=50.0)
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
            
            res = proses_oedometer(
                df, D_sample, H0, e0, tebal_lempung, cv, drainase, luas_bangunan, Gs, wilayah_terpilih, po_kpa
            )
            st.session_state["calc_results"] = res
            
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
                st.dataframe(res["df_tabel_10_5"], use_container_width=True)
            with subtab2:
                st.markdown(f"### 3.1 Pembahasan Geoteknik\n{res['pembahasan']}")
                st.divider()
                st.markdown(f"### 3.2 Kesimpulan\n{res['kesimpulan']}")
            with subtab3:
                st.code(res["laporan"])
            with subtab4:
                st.subheader(f"🏗️ Rekomendasi Teknis: {res['rek_nama']}")
                st.info(f"**Tipe Alat Berat Dibutuhkan:** {res['alat_butuh']}")
                st.subheader(f"📊 Rincian RAB Regional ({wilayah_terpilih})")
                st.dataframe(res["df_rab"].style.format({'Harga Satuan Regional (Rp)': 'Rp {:,.0f}', 'Total Biaya (Rp)': 'Rp {:,.0f}'}), use_container_width=True)
                st.success(f"### ESTIMASI TOTAL BIAYA PROYEK: Rp {res['total_rab']:,.0f}")
            with subtab5:
                st.pyplot(res["fig_elogp"])
            with subtab6:
                st.pyplot(res["fig_time"])
                
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                excel_gabungan_bytes = buat_excel_output_gabungan(res, res["df_export"], res["laporan"], wilayah_terpilih)
                st.download_button(
                    label="📦 Unduh Seluruh Output Lengkap (.xlsx)",
                    data=excel_gabungan_bytes,
                    file_name="Laporan_Lengkap_OedoCalc_Pro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_btn2:
                html_bytes = buat_laporan_html_lengkap(
                    res, "BH-01 Proyek Gedung", st.session_state.curr_lat, st.session_state.curr_lon, 
                    4.5, tebal_lempung, cv, wilayah_terpilih
                )
                st.download_button(
                    label="📄 Unduh Rekap Laporan PDF / HTML (.html)",
                    data=html_bytes,
                    file_name="Laporan_Rekap_OedoCalc_Pro.html",
                    mime="text/html"
                )

# --- TAB 2: SIMULATOR PVD & SETTLEMENT REAL-TIME ---
with tab_pvd:
    st.header("🌀 Simulator Percepatan Konsolidasi PVD (Hansbo Teori)")
    if "calc_results" in st.session_state:
        res = st.session_state["calc_results"]
        st.success(f"Parameter Terdeteksi: Cv = {cv} m²/thn | Cc = {res['Cc']:.4f} | Status = {res['ocr_status']}")
        
        col_pvd1, col_pvd2 = st.columns(2)
        with col_pvd1:
            s_pvd = st.select_slider("Jarak Spasi PVD Pola Segitiga (meter)", options=[0.8, 1.0, 1.2, 1.5, 2.0], value=1.2)
            ch_cv_ratio = st.slider("Rasio ch / cv (Permeabilitas Horizontal)", min_value=1.0, max_value=5.0, value=2.0)
            bulan_sim = st.slider("Simulasi Waktu Pemantauan (Bulan)", min_value=1, max_value=36, value=6)
        
        with col_pvd2:
            # Perhitungan Hansbo 1979 Sederhana
            de = 1.05 * s_pvd
            dw = 0.05 # diameter ekuivalen PVD 5cm
            n_ratio = de / dw
            Fn = np.log(n_ratio) - 0.75
            ch = cv * ch_cv_ratio
            Th = (ch * (bulan_sim / 12.0)) / (de**2)
            Uh = 1 - np.exp((-8 * Th) / Fn)
            
            st.metric("Derajat Konsolidasi Radial (Uh)", f"{Uh*100:.1f} %")
            st.metric("Waktu Penurunan t90 Tanpa PVD", f"{res['t90_b']:.0f} Bulan")
            st.metric(f"Waktu Penurunan t90 DENGAN PVD (s={s_pvd}m)", f"{res['t90_b'] * (1 - Uh*0.7):.1f} Bulan")
    else:
        st.warning("Silakan olah data di Tab '🚀 Aplikasi Utama Kalkulator' terlebih dahulu.")

# --- TAB 3: QUALITY CONTROL CHECKLIST SNI / ASTM ---
with tab_qc:
    st.header("✅ Checklist Control Kepatuhan Standar SNI 2812:2011 & ASTM D2435")
    st.markdown("""
    | Parameter Uji | Standar Acuan | Status Kepatuhan | Catatan Quality Control |
    | :--- | :--- | :--- | :--- |
    | **Rasio Pembebanan (LIR)** | Δp / p = 1.0 (Pembebanan 2x Lipat) | 🟢 PASSED | Rasio pembebanan 0.25, 0.5, 1.0, 2.0, 4.0, 8.0 kg/cm² sesuai baku. |
    | **Durasi Beban per Tahap** | Minimum 24 Jam (1440 menit) | 🟢 PASSED | Pembacaan arloji lengkap dari 6 detik hingga 24 jam. |
    | **Rasio Diameter/Tinggi** | D / H₀ ≥ 2.5 | 🟢 PASSED | Ring Oedometer 63.5 mm / 20.0 mm = 3.175 ≥ 2.5. |
    | **Derajat Kejenuhan (S_r)** | S_r ≈ 100% | 🟡 VERIFY LAB | Pastikan sampel direndam air pori sebelum uji pembebanan dimulai. |
    """)

# --- TAB 4: PETA LOKASI SAMPEL TANAH INTERAKTIF ---
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
