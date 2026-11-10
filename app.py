import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium

# --- TAB 2: PETA LOKASI SAMPEL TANAH INTERAKTIF ---
with tab_peta:
    st.header("🗺️ Peta Pemetaan Lokasi Pengambilan Sampel Tanah (Borehole)")
    st.write("Anda dapat mengubah lokasi sampel dengan **memasukkan koordinat manual**, **mengisi nama lokasi**, atau **mengklik langsung pada peta** di bawah ini.")

    # Inisialisasi State Koordinat jika belum ada
    if 'latitude' not in st.session_state:
        st.session_state.latitude = -6.208800  # Default Jakarta
    if 'longitude' not in st.session_state:
        st.session_state.longitude = 106.845600

    col_map_in, col_map_view = st.columns([1, 2])

    with col_map_in:
        st.subheader("📍 Input / Ubah Lokasi Sampel")
        
        nama_proyek = st.text_input("Nama Proyek / Titik BH", value="BH-01 Proyek Gedung A")
        
        # Form Input Koordinat (Otomatis terhubung dengan klik di peta)
        lat_input = st.number_input(
            "Latitude (Lintang)", 
            value=float(st.session_state.latitude), 
            format="%.6f", 
            key="input_lat"
        )
        lon_input = st.number_input(
            "Longitude (Bujur)", 
            value=float(st.session_state.longitude), 
            format="%.6f", 
            key="input_lon"
        )
        
        # Update session state dari input manual
        st.session_state.latitude = lat_input
        st.session_state.longitude = lon_input

        kedalaman = st.number_input("Kedalaman Sampel (meter)", value=4.5)
        
        # Link Navigasi ke Google Maps Eksternal
        gmaps_url = f"https://www.google.com/maps?q={st.session_state.latitude},{st.session_state.longitude}"
        st.markdown(f"🔗 **[Buka Titik Ini di Google Maps Direct]({gmaps_url})**")

    with col_map_view:
        st.subheader("🌍 Klik pada Peta untuk Mengubah Lokasi")
        
        # Buat objek peta folium
        m = folium.Map(
            location=[st.session_state.latitude, st.session_state.longitude], 
            zoom_start=13,
            tiles="OpenStreetMap"
        )
        
        # Tambahkan marker pada posisi sampel saat ini
        folium.Marker(
            [st.session_state.latitude, st.session_state.longitude],
            popup=f"<b>{nama_proyek}</b><br>Kedalaman: {kedalaman} m",
            tooltip=nama_proyek,
            icon=folium.Icon(color="red", icon="info-sign")
        ).add_to(m)

        # Tampilkan peta interaktif yang dapat diklik
        map_data = st_folium(m, height=400, width=700)

        # Jika pengguna mengklik titik baru di peta, perbarui koordinat secara otomatis
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
