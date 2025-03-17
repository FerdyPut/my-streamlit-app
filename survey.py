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
        # Data produk per outlet
        outlet_data = {
            "Indomaret": [
                {
                    "nama_produk": "Goriorio Vanilla 23 Gr",
                    "jenis_promo": "Beli 1 pcs, gratis 1 pcs",
                    "periode_promo": "1 - 5 April 2025"
                }
            ],
            "Alfamart": [
                {
                    "nama_produk": "French Fries 192 Gr",
                    "jenis_promo": "Beli 1 pcs, potongan Rp. 200",
                    "periode_promo": "5 - 9 April 2025"
                },
                {
                    "nama_produk": "Goriorio Vanilla 23 Gr",
                    "jenis_promo": "Beli 2 pcs, gratis 1 pcs",
                    "periode_promo": "6 - 10 April 2025"
                }
            ],
            # Tambah outlet dan produk lain di sini
        }
        
        # Pilih outlet
        tipe_outlet = st.selectbox("Pilih Outlet:", list(outlet_data.keys()))
        
        # Tampilkan produk yang sesuai outlet
        produk_list = outlet_data.get(tipe_outlet, [])
        produk_names = [p["nama_produk"] for p in produk_list]
        
        if produk_list:
            # Pilih produk berdasarkan outlet
            nama_produk = st.selectbox("Nama Produk:", produk_names, key="nama_produk")
            
            # Ambil detail promo & periode
            produk_terpilih = next((p for p in produk_list if p["nama_produk"] == nama_produk), {})
            
            # Ambil periode promo
            periode_promo = produk_terpilih.get("periode_promo", "")
            
            # Tampilkan informasi promo dan periode
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
                f"Berapa harga produk {nama_produk} yang tertera di rak / server kasir?",
                min_value=0
            , key="harga_produk")
            st.caption("Note: Harga Asli, sebelum potongan promo (angka nya saja)")
            expired_date = st.date_input(f"Tanggal Expired Date produk {nama_produk}", key="expired_date")

            if "gratis" in jenis_promo.lower():
                sisa_stock = st.number_input(
                    f"Berapa sisa produk {nama_produk} yang tertera di display?",
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
                    if not nama_surveyor or not kode_outlet or not kota:
                        st.error("Nama Surveyor, Kode Outlet, dan Kota wajib diisi!")
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

                        st.success("Data Sudah Tersimpan di Overview Excel!")
                        st.info("Jikalau mau menginput data lagi silahkan refresh website!")   
with tab2:
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
            if st.button("Hapus Semua Data"):
                os.remove(file_path)
                st.success("Semua data berhasil dihapus.")
                st.rerun()

        with colC:
            st.download_button(
                label="Download Excel",
                data=buffer,
                file_name="data_survey.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.info("Belum ada data yang tersimpan.")
