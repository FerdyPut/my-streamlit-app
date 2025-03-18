import streamlit as st
import pandas as pd
from datetime import datetime, time
import os
from io import BytesIO
import openpyxl
from datetime import datetime
from zoneinfo import ZoneInfo

st.title("Form Survey Promo HCO Chain Bulan Maret 2025")
background_url = "https://upload.wikimedia.org/wikipedia/commons/c/cc/Logo_Siantar_Top.svg"

st.markdown(
    f"""
    <style>
    .stApp::before {{
        content: "";
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: url("{background_url}") no-repeat center center fixed;
        background-size: cover;
        filter: blur(10px);
        opacity: 0.6; /* transparansi background */
        z-index: -1;
    }}

    /* Biar konten tetap tajam */
    .stApp {{
        background: transparent;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

file_path = "data_survey.xlsx"
tab0, tab1, tab2 = st.tabs(["Survey Promo", "Form Survey Inputan", "Hasil Inputan Survey"])

with tab0:
        st.header("Survey Promo - By Business Support Infromation (BSI)")
        logo_url = "https://upload.wikimedia.org/wikipedia/commons/c/cc/Logo_Siantar_Top.svg" 
        st.markdown(
                f"""
                <div style="text-align: center;">
                    <marquee behavior="alternate" scrollamount="5" width="400">
                        <img src="{logo_url}" width="150">
                    </marquee>
                </div>
                """,
                unsafe_allow_html=True
            )
        st.markdown(
            """
        - Website ini Hak Cipta/Milik BSI di bawah Bu Nuril (Head Division)
        - Akan diupdate secara berkala mengenai produk, jenis promo, periode promo, dan lainnya.
        - Mohon dicermati dengan baik.
        """
        )

with tab1:
        st.header("Form Input")
        nama_surveyor = st.selectbox("Nama Surveyor:", [
            "Joko Lestiono", "Achsin", "Asep Rio", "Widi Setiawan", "Hadi Syamsi G.", 
            "Andri", "Bambang", "Sigit Ari W.", "Iswandi", "Zazat Sudrajat", 
            "Wahyu", "D. Muh Sidiq", "Abdur R Slamet", "Yoga Bagus Permana", 
            "M. Dedek Kurniawan", "Ferdian", "Adang Syafaat", "Syamsul", 
            "Alom", "Mulyono", "Fernando"
        ],key="nama_surveyor")

        tahun = st.selectbox("Tahun:", ["2025", "2026", "2027"],key="tahun")
        bulan = st.selectbox("Bulan: ", [str(i) for i in range(1, 13)],key="bulan")

#-------------------------------------------HARUS DIUPDATE NAMA PRODUK, PERIODE, DAN JENIS PROMO
        today = datetime.now(ZoneInfo("Asia/Jakarta")).date()
        sheet_url = "https://docs.google.com/spreadsheets/d/1GIfUGSMLfCMiDMy1aFHm_05F1IJXzY3kY89QCceFDOA/export?format=csv"
        df = pd.read_csv(sheet_url)
        df['Tanggal Survey'] = pd.to_datetime(df['Tanggal Survey']).dt.date
        
        df_today = df[df['Tanggal Survey'] == today]
        if df_today.empty:
            st.warning(f"Tidak ada outlet yang disurvey hari ini ({today})")
        else:
            outlet_data = {}
            for _, row in df_today.iterrows():
                outlet = row['Tipe Outlet']
                if outlet not in outlet_data:
                    outlet_data[outlet] = []
                outlet_data[outlet].append({
                    "nama_produk": row['Nama Produk'],
                    "jenis_promo": row['Jenis Promo'],
                    "periode_promo": row['Periode Promo']
                })
        
            tipe_outlet = st.selectbox("Pilih Outlet (Hari Ini):", list(outlet_data.keys()))
        
            if tipe_outlet:
                st.write(f"Promo di {tipe_outlet}:")
                for promo in outlet_data[tipe_outlet]:
                    st.write(f"- **{promo['nama_produk']}**: {promo['jenis_promo']} (Periode: {promo['periode_promo']})")
                
                produk_list = outlet_data.get(tipe_outlet, [])
                produk_names = [p["nama_produk"] for p in produk_list]
        
                if produk_list:
                    nama_produk = st.selectbox("Nama Produk:", produk_names, key="nama_produk")
                    produk_terpilih = next((p for p in produk_list if p["nama_produk"] == nama_produk), {})
                    periode_promo = produk_terpilih.get("periode_promo", "")
        
                    st.text_input("Jenis Promo:", value=produk_terpilih.get("jenis_promo", ""), disabled=True)
                    st.text_input("Periode Promo:", value=periode_promo, disabled=True)
                    
                    # Cek apakah jenis_promo mengandung kata 'gratis'
                    jenis_promo = produk_terpilih.get("jenis_promo", "").lower()
                    if "gratis" in jenis_promo:
                        st.text("Promo mengandung kata 'gratis'")
                    else:
                        st.text("Promo tidak mengandung kata 'gratis'")
                else:
                    st.info("Belum ada produk untuk outlet ini.")


        tipe_account = st.selectbox("Tipe Account:", ["Chain"], index=0, disabled=True)
        kode_outlet = st.text_input("Kode Outlet:", key="kode_outlet")
        st.caption("Isikan - jika tidak tau")
        
        tanggal = st.date_input("Tanggal", value=datetime.today(),key="tanggal")

        col1, col2 = st.columns(2)
        with col1:
            jam = st.selectbox("Jam Kunjungan:", list(range(0, 24)), format_func=lambda x: f"{x:02d}",key="jam")
        with col2:
            menit = st.selectbox("Menit:", list(range(0, 60)), format_func=lambda x: f"{x:02d}",key="menit")

        jam_final = time(hour=jam, minute=menit)
        kota = st.selectbox("Kota:", [
            "Surabaya", "Bandung", "Bekasi", "Semarang", "Jember", "Medan", "Pamekasan", 
            "Solo", "Malang", "Kediri", "Tangerang", "Makassar", "Palembang", 
            "Pekanbaru", "Pontianak"
        ],key="kota")
        alamat_outlet = st.text_input("Alamat Outlet:",key="alamat_outlet")
    #---------------------------------------------------------------DISPLAY PRODUK
        # Sebelum pertanyaan display produk
        st.subheader(f"Detail Produk")

        produk_display = st.selectbox(f"Apakah produk {nama_produk} terdisplay di toko?", [" ", "Iya", "Stock Kosong", "Tidak Jual"], key="produk_display")

        harga_produk = None
        expired_date = None
        sisa_stock = None

        if produk_display == "Iya":
            harga_produk = st.number_input(
                f"Berapa harga produk {nama_produk} per pcs yang tertera di rak / server kasir?",
                min_value=0
            , key="harga_produk")
            st.caption("Note: Harga Asli, sebelum potongan promo (angka nya saja)")
            expired_date = st.date_input(f"Tanggal Expired Date produk {nama_produk}", key="expired_date")

            if "gratis" in jenis_promo.lower():
                sisa_stock = st.number_input(
                    f"Berapa sisa produk {nama_produk} per pcs yang tertera di display?",
                    min_value=0
                , key="sisa_stock")
                st.caption("Note: Kalau tidak ada/kosong/habis isi 0")


    #-----------------------------------------------------------KATALOG PRODUK
        if produk_display in ["Iya", "Stock Kosong", "Tidak Jual"]:
            st.subheader(f"Informasi Katalog Produk")

            promo_mailer = st.selectbox(
                "Apakah promo tersebut tercantum di mailer / katalog promo? (jika mailer/katalog promo habis/tidak ada, tanyakan & minta di kasir)", 
                ["", "Iya", "Tidak", "Tidak Tahu"]
            , key="promo_mailer")

            if promo_mailer == "Tidak Tahu":
                keterangan = st.text_input("Keterangan:", key="keterangan")
            else:
                keterangan = st.text_input("Keterangan:", value="-", disabled=True)



    #----------------------------------------------------------------ALASAN MATERIAL PROMO
            if promo_mailer in ["Iya", "Tidak", "Tidak Tahu"]:
                st.subheader(f"Alasan Material Promo Produk")
                material_promo = st.selectbox(
                    "Apakah material promo (seperti : wobler, price tag atau lainnya) terpasang di rak / pricetag produk yang di promosikan?", 
                    ["", "Ya", "Tidak"]
                ,key="material_promo")

                alasan_material = "-"
                keterangan_alasan_material = "-"

                if material_promo == "Tidak":
                    alasan_material_list = [
                        "Distribusi/pengiriman material promo (seperti : wobler, atau lainnya) belum sampai ke toko",
                        "Material promo (seperti : wobler, atau lainnya) sudah sampai di toko, namun belum terpasang oleh pihak toko",
                        "Outlet tidak menjual produk tersebut",
                        "Lainnya (Isi sendiri)"
                    ]
                    alasan_select = st.selectbox(
                        "Kenapa material promo (seperti : wobler, price tag atau lainnya) tidak terpasang di rak / pricetag produk yang di promosikan?",
                        alasan_material_list
                    ,key="alasan_select")

                    if alasan_select == "Lainnya (Isi sendiri)":
                        alasan_material = st.text_input("Silakan isi alasan lainnya:",key="alasan_material")
                    else:
                        alasan_material = alasan_select


        #--------------------------------------------------------------------SETTING KASIR
                if material_promo in ["Ya", "Tidak"]:
                    st.subheader(f"Informasi Tersetting Produk dan Harga di Struk Produk {nama_produk}")
                    promo_di_kasir = st.selectbox(
                        "Apakah promo tersetting di sistem server kasir?", 
                        ["", "Ya", "Tidak"]
                    ,key="promo_di_kasir")

                    info_struk = st.text_input("Informasi potongan harga yang tertera di struk:",key="info_struk")

                if st.button("Submit"):
                    errors = []
                
                    # Validasi field wajib utama
                    if not nama_surveyor or not kode_outlet or not kota:
                        errors.append("Nama Surveyor, Kode Outlet, dan Kota wajib diisi!")
                
                    # Validasi tambahan jika produk_display sudah dipilih
                    if produk_display in ["Iya", "Stock Kosong", "Tidak Jual"]:
                        if promo_mailer.strip() == "":
                            errors.append("Promo Mailer wajib diisi.")
                        if promo_mailer == "Tidak Tahu" and keterangan.strip() == "":
                            errors.append("Keterangan wajib diisi jika memilih 'Tidak Tahu' pada promo mailer.")
                        if material_promo.strip() == "":
                            errors.append("Material Promo wajib diisi.")
                        if material_promo == "Tidak" and alasan_material.strip() == "":
                            errors.append("Alasan Material wajib diisi jika material tidak terpasang.")
                        if promo_di_kasir.strip() == "":
                            errors.append("Promo di server kasir wajib diisi.")
                        if info_struk.strip() == "":
                            errors.append("Informasi potongan harga di struk wajib diisi.")
                        if produk_display == "Iya":
                            if harga_produk is None or harga_produk == 0:
                                errors.append("Harga produk wajib diisi jika produk terdisplay.")
                            if expired_date is None:
                                errors.append("Tanggal expired wajib diisi jika produk terdisplay.")
                            if "gratis" in jenis_promo and (sisa_stock is None):
                                errors.append("Sisa stock wajib diisi untuk promo yang mengandung 'gratis'.")
                
                    # Jika ada error, tampilkan semua
                    if errors:
                        for err in errors:
                            st.error(err)
                    else:
                        # Logic untuk Kode Stock
                        if sisa_stock is not None:
                            if sisa_stock <= 3:
                                kode_stock = 1
                            else:
                                kode_stock = 2
                        else:
                            kode_stock = "-"
                        new_data = {
                            "Timestamp Pengisian" : datetime.now(ZoneInfo('Asia/Jakarta')).strftime('%Y-%m-%d %H:%M:%S'),
                            "Tipe Outlet": tipe_outlet,
                            "Tipe Account": tipe_account,
                            "Bulan": bulan,
                            "Tahun": tahun,
                            "Nama Produk": nama_produk,
                            "Periode Promo": periode_promo,
                            "Jenis Promo": produk_terpilih.get("jenis_promo", ""),
                            "Periode Survey": "-",  # Kosong dulu
                            "Kota": kota,
                            "Kode Outlet": kode_outlet,
                            "Alamat": alamat_outlet,
                            "Nama Surveyor": nama_surveyor,
                            "Tanggal Kunjugan": tanggal.strftime('%Y-%m-%d'),
                            "Jam Kunjungan": jam_final.strftime('%H:%M'),
                            "Produk Display": produk_display,
                            "Promo Mailer": promo_mailer,
                            "Keterangan Mailer": keterangan,
                            "Material Promo": material_promo,
                            "Kode Material": "-",  # Kosong dulu
                            "Alasan Material Tidak Terpasang": alasan_material,
                            "Promo Tersetting di Server Kasir": promo_di_kasir,
                            "Sisa Stock": sisa_stock,
                            "Kode Stock": kode_stock,  # Kosong dulu
                            "Informasi Potongan Harga di Struk": info_struk,
                            "Expired Date": expired_date.strftime('%Y-%m-%d') if expired_date else "-",
                            "Harga Produk": harga_produk
                        }

                        if os.path.exists(file_path):
                            df_existing = pd.read_excel(file_path, engine="openpyxl")
                        else:
                            df_existing = pd.DataFrame()

                        order = [
                            "Timestamp Pengisian","Tipe Outlet", "Tipe Account", "Bulan", "Tahun", "Nama Produk", "Periode Promo", "Jenis Promo", 
                            "Periode Survey", "Kota", "Kode Outlet", "Alamat", "Nama Surveyor", "Tanggal Kunjugan", 
                            "Jam Kunjungan", "Produk Display", "Promo Mailer", "Keterangan Mailer", "Material Promo", 
                            "Kode Material", "Alasan Material Tidak Terpasang", "Promo Tersetting di Server Kasir", 
                            "Sisa Stock", "Kode Stock", "Informasi Potongan Harga di Struk", "Expired Date", "Harga Produk"
                        ]

                        df_new = pd.DataFrame([new_data])
                        df_existing = pd.concat([df_existing, df_new], ignore_index=True)
                        df_existing = df_existing.fillna("-")
                        df_existing = df_existing[order]  # <-- Reorder kolom di sini
                        df_existing.to_excel(file_path, index=False, engine='openpyxl')
                        import time

                        st.success("Data Sudah Tersimpan!")
                        st.info("Jikalau mau menginput data lagi silahkan refresh website!")   
with tab2:
    if "admin_login" not in st.session_state:
        st.session_state.admin_login = False

    if not st.session_state.admin_login:
        st.subheader("🔒 Akses Overview Data")
        password = st.text_input("Masukkan Password:", type="password")
        if st.button("Login"):
            if password == "mis01":  # Ganti sesuai kebutuhanmu
                st.session_state.admin_login = True
                st.success("Login berhasil! Silakan akses data.")
                st.rerun()
            else:
                st.error("Password salah, coba lagi.")
    else:
            st.header("Hasil Inputan Data Survey Promo HCO Chain")
            st.markdown("""
            **Petunjuk Penggunaan:**
            - Data survey yang sudah diinput akan tampil di bawah ini dan dapat langsung diedit.
            - Klik **Simpan Perubahan** untuk menyimpan perubahan data setelah merubah data langsung pada tabel di bawah.
            - Untuk Menghapus satu baris data: **Checkbox** pada bagian kiri sendiri di suatu baris yang ingin dihapus kemudian klik gambar **icon hapus** pada pojok kanan tabel (kecil), setelah itu klik **simpan perubahan**
            - Klik **Hapus Semua Data** jika ingin menghapus seluruh data survey.
            - Klik **Download Excel** untuk mengunduh data dalam format Excel.
            - **Surveyor** tidak perlu mendownload excel, dan jikalau revisi bisa dilakukan edit data atau delete data.
            """)
            
            if os.path.exists(file_path):
                df_existing = pd.read_excel(file_path, engine='openpyxl')
        
                edited_df = st.data_editor(
                    df_existing,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="editor",
                    disabled=False
                )
        
                buffer = BytesIO()
                edited_df.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
        
                buffer = BytesIO()
                edited_df.fillna("-").to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                colA, colB, colC = st.columns(3)
        
                with colA:
                    if st.button("Simpan Perubahan"):
                        edited_df.to_excel(file_path, index=False, engine='openpyxl')
                        st.success("Perubahan berhasil disimpan.")
                        st.rerun()
        
                with colB:
                    # Tombol awal untuk trigger konfirmasi
                    if "show_confirm" not in st.session_state:
                        st.session_state.show_confirm = False
                
                    if st.button("Hapus Semua Data"):
                        st.session_state.show_confirm = True
                
                    # Jika tombol sudah ditekan, baru munculkan konfirmasi
                    if st.session_state.show_confirm:
                        st.warning("⚠️ Apakah Anda yakin ingin menghapus semua data? Tindakan ini tidak dapat dibatalkan!")
                
                        confirm = st.checkbox("Saya yakin ingin menghapus semua data")
                
                        col_confirm, col_cancel = st.columns(2)
                        with col_confirm:
                            if confirm:
                                if st.button("Konfirmasi Hapus", key="confirm_hapus"):
                                    try:
                                        os.remove(file_path)
                                        st.success("✅ Semua data berhasil dihapus.")
                                    except FileNotFoundError:
                                        st.error("❌ File tidak ditemukan.")
                                    st.session_state.show_confirm = False
                                    st.rerun()
                        with col_cancel:
                            if st.button("Kembali", key="cancel_hapus"):
                                st.session_state.show_confirm = False
                                st.info("Penghapusan data dibatalkan.")
                                st.rerun()
                with colC:
                    st.download_button(
                        label="Download Excel",
                        data=buffer,
                        file_name="data_survey.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.markdown("---")
                if st.button("Logout ❌"):
                        st.session_state.admin_login = False
                        st.success("Berhasil logout.")
                        st.rerun()
            else:
                st.info("Belum ada data yang tersimpan.")
