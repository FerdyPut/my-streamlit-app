import streamlit as st
import pandas as pd
from datetime import datetime, time
import os
from io import BytesIO
import openpyxl
from datetime import datetime, date
from zoneinfo import ZoneInfo

st.set_page_config(page_title="Form Survey Promo HCO Chain dan Lokal - Siantar Top - BSI ", page_icon=":bookmark_tabs:", layout="centered")
st.title("Form Survey Promo HCO Chain dan Lokal Tahun 2025")
st.markdown(
    """
    <style>
/* Menyesuaikan ukuran teks di dalam container widget */
    section.main div[role="radiogroup"] label, /* Untuk radio button */
    section.main label, /* Untuk semua widget */
    div[data-testid="stWidgetLabel"] {
        font-size: 22px !important;
        font-weight: bold;
        color: #FFFFF;
    }

    /* Mengubah ukuran label dalam widget */
    div[data-testid="stWidgetLabel"] {
        font-size: 20px !important;
        font-weight: bold;
        color: #FFFFFF;
    }
    /* Perbesar teks secara umum */
    body, div, input, select, button, label {
        font-size: 20px !important;
    }

    /* Perbesar judul */
    .stApp h1 {
        font-size: 30px !important;
    }

    /* Perbesar teks tabel */
    .dataframe th, .dataframe td, .tabs {
        font-size: 18px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)




# Inisialisasi session state untuk overview
if "overview" not in st.session_state:
    st.session_state["overview"] = pd.DataFrame(columns=["Timestamp Pengisian", "Nama Surveyor", "Nama Produk", "Status"])

# File untuk menyimpan data
EXCEL_FILE = "overview_data.xlsx"
file_path = "data_survey.xlsx"

def load_from_excel():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE, engine="openpyxl")
    return pd.DataFrame(columns=["Timestamp Pengisian", "Nama Surveyor", "Nama Produk"])

def save_to_excel(df, file_path):
    df.to_excel(file_path, index=False, engine="openpyxl")

tab0, tab1, tab2, tab3 = st.tabs(["Survey Promo", "Form Survey Inputan","Cek Produk yang sudah Input", "Hasil Inputan Survey"])

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
        # Inisialisasi session_state jika belum ada
        if "prev_nama_surveyor" not in st.session_state:
            st.session_state["prev_nama_surveyor"] = None
        if "prev_kota" not in st.session_state:
            st.session_state["prev_kota"] = None
        if "prev_tipe_outlet" not in st.session_state:
            st.session_state["prev_tipe_outlet"] = None
        if "prev_nama_produk" not in st.session_state:
            st.session_state["prev_nama_produk"] = None

        if "alamat_outlet" not in st.session_state:
            st.session_state["alamat_outlet"] = ""
        if "jam" not in st.session_state:
            st.session_state["jam"] = 0
        if "menit" not in st.session_state:
            st.session_state["menit"] = 0
        if "tanggal" not in st.session_state:
            st.session_state["tanggal"] = datetime.today()
        if "kode_outlet" not in st.session_state:
            st.session_state["kode_outlet"] = ""

        pertanyaan_produk = {
            "harga_produk": 0.0,  # Gunakan angka bukan string
            "expired_date": None,
            "sisa_stock": 0,
            "promo_mailer": "",
            "keterangan": "",
            "material_promo": "",
            "alasan_material": "",
            "promo_di_kasir": "",
            "info_struk": ""
        }

        # Pastikan setiap key dalam session_state sudah diinisialisasi dengan benar
        for key, value in pertanyaan_produk.items():
            if key not in st.session_state:
                st.session_state[key] = value
#------------------------------------------------------------
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
        kota = st.selectbox("Kota:", [
        "Surabaya", "Bandung", "Bekasi", "Semarang", "Jember", "Medan", "Pamekasan", 
        "Solo", "Malang", "Kediri", "Tangerang", "Makassar", "Palembang", 
        "Pekanbaru", "Pontianak"
        ],key="kota")


        today = datetime.now().date()
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
                    for i, promo in enumerate(outlet_data[tipe_outlet], start=1):
                        st.markdown(f"""
                        <div style='margin-bottom: 12px; font-size: 16px; font-weight: bold;'>
                            {i}. {promo['nama_produk']} - {promo['jenis_promo']} (Periode: {promo['periode_promo']})
                        </div>
                        """, unsafe_allow_html=True)
                        
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

    # Daftar outlet yang termasuk Chain
        chain_outlet = [
            "Indomaret", "Indogrosir", "Alfamart", 
            "Alfamidi", "Lion Superindo", "Clandys", "Family Mart"
        ]

#-------------------------------------------------
        # **Cek perubahan Nama Surveyor, Kota, atau Tipe Outlet → Reset Semua**
        # Cek apakah harus reset semua atau hanya pertanyaan produk
        reset_all = False
        reset_produk = False

        if (
            nama_surveyor != st.session_state["prev_nama_surveyor"] or 
            kota != st.session_state["prev_kota"] or 
            tipe_outlet != st.session_state["prev_tipe_outlet"]
        ):
            reset_all = True

        if nama_produk != st.session_state["prev_nama_produk"]:
            reset_produk = True

        # **Update prev_ values lebih awal untuk menghindari bug**
        st.session_state.update({
            "prev_nama_surveyor": nama_surveyor,
            "prev_kota": kota,
            "prev_tipe_outlet": tipe_outlet,
            "prev_nama_produk": nama_produk
        })

        # **Eksekusi reset sesuai kondisi**
        if reset_all:
            st.session_state.update({
                "alamat_outlet": "",
                "jam": 0,
                "menit": 0,
                "kode_outlet": "",
                "tanggal": datetime.today(),
            })

        if reset_produk:  # Ini harus tetap dijalankan meskipun reset_all sudah berjalan
            for key in pertanyaan_produk:
                st.session_state[key] = ""




        # Tentukan tipe account berdasarkan outlet yang dipilih
        if tipe_outlet in chain_outlet:
            tipe_account_value = "Chain"
        else:
            tipe_account_value = "Lokal"

        # Tampilkan tipe account (disabled)
        tipe_account = st.selectbox("Tipe Account (HCO):", ["Chain", "Lokal"], index=0 if tipe_account_value == "Chain" else 1, disabled=True)
        kode_outlet = st.text_input("Kode Outlet:", key="kode_outlet")
        st.caption("Isikan - jika tidak tau")


        tanggal = st.date_input("Tanggal Kunjungan", value=datetime.today(),key="tanggal")

        col1, col2 = st.columns(2)
        with col1:
            jam = st.selectbox("Jam Kunjungan:", list(range(0, 24)), format_func=lambda x: f"{x:02d}",key="jam")
        with col2:
            menit = st.selectbox("Menit:", list(range(0, 60)), format_func=lambda x: f"{x:02d}",key="menit")

        jam_final = time(hour=jam, minute=menit)

        alamat_outlet = st.text_input("Alamat Outlet:",key="alamat_outlet")
        


#---------------------------------------------------------------DISPLAY PRODUK
        st.subheader(f"Detail Produk")
        produk_display = st.selectbox(
            f"Apakah produk {nama_produk} terdisplay di toko?", 
            [" ", "Iya", "Stock Kosong", "Tidak Jual"], 
            key="produk_display"
        )

        # Default None untuk validasi yang lebih aman
    # Reset Pertanyaan Produk dengan tipe data yang benar
       # Default None untuk validasi yang lebih aman
        harga_produk = None
        expired_date = None
        sisa_stock = None
        promo_mailer = ""
        keterangan = ""
        material_promo = ""
        alasan_material = ""
        promo_di_kasir = ""
        info_struk = ""

        # Step 1 - Jika sudah pilih display
        if produk_display.strip() != "":
            # --- Chain & Lokal: Harga & Stock --- (hanya jika display = "Iya")
            if produk_display == "Iya" and tipe_account_value in ["Chain", "Lokal"]:
                harga_produk_value = st.session_state.get("harga_produk", 0.0)
                if harga_produk_value is None or isinstance(harga_produk_value, str):
                    harga_produk_value = 0.0  # Default jika terjadi kesalahan tipe data

                # Gunakan nilai yang sudah diperbaiki dalam number_input
                harga_produk = st.number_input(
                    f"Berapa harga produk {nama_produk} per pcs yang tertera di rak / server kasir?",
                    min_value=0.0,
                    key="harga_produk",
                    value=harga_produk_value
                )
                st.caption("Note: Harga Asli, sebelum potongan promo (angka nya saja)")

                if "gratis" in jenis_promo.lower():
                    sisa_stock = st.number_input(
                        f"Berapa sisa produk {nama_produk} per pcs yang tertera di display?",
                        min_value=0, key="sisa_stock"
                    )
                    st.caption("Note: Kalau tidak ada/kosong/habis isi 0")

            # --- Chain khusus: expired date --- (hanya jika display = "Iya")
            if produk_display == "Iya" and tipe_account_value == "Chain":
                import datetime
                expired_date = st.date_input(
                    f"Tanggal Expired Date produk {nama_produk}",
                    value=datetime.date.today(),  # Pastikan ini valid
                    key="expired_date"
                )
            # Step 2 - Katalog MUNCUL untuk semua display selain kosong
            if tipe_account_value == "Chain":
                st.subheader(f"Informasi Katalog Produk")
                promo_mailer = st.selectbox(
                    "Apakah promo tersebut tercantum di mailer / katalog promo? (jika mailer/katalog promo habis/tidak ada, tanyakan & minta di kasir)", 
                    ["", "Iya", "Tidak", "Tidak Tahu"],
                    key="promo_mailer"
                )
                if promo_mailer == "Tidak Tahu":
                    keterangan = st.text_input("Keterangan:", key="keterangan")
                elif promo_mailer in ["Iya", "Tidak"]:
                    keterangan = "-"

            # Step 3 - Material Promo (tetap muncul untuk semua tipe account)
            if tipe_account_value != "Chain" or promo_mailer.strip() != "":
                st.subheader(f"Alasan Material Promo Produk")
                material_promo = st.selectbox(
                    "Apakah material promo (seperti : wobler, price tag atau lainnya) terpasang di rak / pricetag produk yang di promosikan?", 
                    ["", "Ya", "Tidak"], 
                    key="material_promo"
                )

                # Step 4 - Alasan material jika jawab "Tidak"
                if material_promo == "Tidak":
                    alasan_material_list = [
                        "Distribusi/pengiriman material promo (seperti : wobler, atau lainnya) belum sampai ke toko",
                        "Material promo (seperti : wobler, atau lainnya) sudah sampai di toko, namun belum terpasang oleh pihak toko",
                        "Outlet tidak menjual produk tersebut",
                        "Lainnya (Isi sendiri)"
                    ]
                    alasan_select = st.selectbox(
                        "Kenapa material promo tidak terpasang di rak / pricetag produk yang di promosikan?",
                        alasan_material_list, key="alasan_select"
                    )
                    if alasan_select == "Lainnya (Isi sendiri)":
                        alasan_material = st.text_input("Silakan isi alasan lainnya:", key="alasan_material")
                    else:
                        alasan_material = alasan_select
                elif material_promo == "Ya":
                    alasan_material = "-"

                # Step 5 - Struk baru muncul setelah material promo terisi
                if material_promo != "":
                    st.subheader(f"Informasi Tersetting Produk dan Harga di Struk Produk {nama_produk}")
                    promo_di_kasir = st.selectbox(
                        "Apakah promo tersetting di sistem server kasir?", 
                        ["", "Ya", "Tidak", "Tidak Tahu"], 
                        key="promo_di_kasir"
                    )
                    info_struk = st.text_input("Informasi potongan harga yang tertera di struk:", key="info_struk")

            # --- SUBMIT ---
             #--------------- TOMBOL SUBMIT DATA   
            if st.button("Submit"):
                errors = []

                # Validasi field wajib utama
                if not nama_surveyor or not kode_outlet or not kota:
                    errors.append("Nama Surveyor, Kode Outlet, dan Kota wajib diisi!")
                if "overview" not in st.session_state or st.session_state["overview"].empty:
                    st.session_state["overview"] = pd.DataFrame(columns=["Timestamp Pengisian", "Nama Surveyor", "Nama Produk", "Status"])

#-------------------------SESSION NTUK OVERVIEW SURVEYOR               
                
                if errors:
                    for err in errors:
                        st.error(err)
                else:
                    # Format data baru
                    from datetime import datetime
                    new_data = pd.DataFrame([{
                        "Timestamp Pengisian": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        "Nama Surveyor": nama_surveyor,
                        "Nama Produk": nama_produk,
                        "Kode Outlet": kode_outlet,
                        "Status": "Sudah Terinput"
                    }])

                    # 🔹 Tambahkan ke session_state["overview"]
                    st.session_state["overview"] = pd.concat([st.session_state["overview"], new_data], ignore_index=True)

                    # 🔹 Simpan ke Excel
                    file_path = "data_survey.xlsx"
                    if os.path.exists(file_path):
                        df_existing = pd.read_excel(file_path, engine="openpyxl")
                    else:
                        df_existing = pd.DataFrame()  # Buat DataFrame kosong jika file belum ada

                    # 🔹 Pastikan hanya menambahkan data utama (bukan overview)
                    df_existing = pd.concat([df_existing], ignore_index=True)

                    # 🔹 Simpan ke file Excel
                    df_existing.to_excel(file_path, index=False, engine="openpyxl")

#-------------------------------------------------------

                # --- Validasi Step Lainnya
                if produk_display in ["Iya", "Stock Kosong", "Tidak Jual"]:
                    # Field yang hanya wajib untuk Chain
                    if tipe_account_value == "Chain":
                        if promo_mailer.strip() == "":
                            errors.append("Promo Mailer wajib diisi untuk account Chain.")

                        # Hanya wajib isi expired date jika produk display = "Iya"
                        if produk_display == "Iya" and expired_date is None:
                            errors.append("Tanggal expired wajib diisi untuk account Chain jika produk terdisplay.")

                    # Field yang wajib untuk Lokal dan Chain
                    if material_promo.strip() == "":
                        errors.append("Material Promo wajib diisi.")
                    if material_promo == "Tidak" and alasan_material.strip() == "":
                        errors.append("Alasan Material wajib diisi jika material tidak terpasang.")
                    if promo_di_kasir.strip() == "":
                        errors.append("Promo di server kasir wajib diisi.")
                    if info_struk.strip() == "":
                        errors.append("Informasi potongan harga di struk wajib diisi.")

                    # Validasi tambahan hanya jika produk_display == "Iya"
                    if produk_display == "Iya":
                        if harga_produk is None or harga_produk == 0:
                            errors.append("Harga produk wajib diisi jika produk terdisplay.")
                        # Sisa stock hanya untuk promo dengan 'gratis'
                        if "gratis" in jenis_promo.lower():
                            if sisa_stock is None:
                                errors.append("Sisa stock wajib diisi untuk promo yang mengandung 'gratis'.")
                # Jika ada error
                if errors:
                    for err in errors:
                        st.error(err)
                else:
                    # Logic kode_stock
                    if isinstance(sisa_stock, int):
                        kode_stock = 1 if sisa_stock <= 3 else 2
                    else:
                        kode_stock = "-"

                    if isinstance(expired_date, (datetime, date)):
                        export_expired = expired_date.strftime('%Y-%m-%d')
                    else:
                        export_expired = "-"

                    export_expired = expired_date.strftime('%Y-%m-%d') if isinstance(expired_date, (datetime, date)) else "-"
                    export_sisa_stock = sisa_stock if isinstance(sisa_stock, int) else "-"
                    export_harga = harga_produk if isinstance(harga_produk, int) else "-"
                    # Data dict
                    new_data = {
                        "Timestamp Pengisian": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        "Tipe Outlet": tipe_outlet,
                        "Tipe Account": tipe_account,
                        "Bulan": bulan,
                        "Tahun": tahun,
                        "Nama Produk": nama_produk,
                        "Periode Promo": periode_promo,
                        "Jenis Promo": produk_terpilih.get("jenis_promo", ""),
                        "Periode Survey": "-",
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
                        "Kode Material": "-",
                        "Alasan Material Tidak Terpasang": alasan_material,
                        "Promo Tersetting di Server Kasir": promo_di_kasir,
                        "Sisa Stock": export_sisa_stock,
                        "Kode Stock": kode_stock,
                        "Informasi Potongan Harga di Struk": info_struk,
                        "Expired Date": export_expired,
                        "Harga Produk": export_harga
                    }

                    if os.path.exists(file_path):
                        df_existing = pd.read_excel(file_path, engine="openpyxl")
                    else:
                        df_existing = pd.DataFrame()

                    order = [
                        "Timestamp Pengisian","Tipe Outlet", "Tipe Account",  "Tahun", "Bulan", "Periode Promo",  "Nama Produk", "Jenis Promo", 
                        "Periode Survey", "Kota","Nama Surveyor", "Tanggal Kunjugan", 
                        "Jam Kunjungan",  "Kode Outlet", "Alamat", "Produk Display", "Promo Mailer", "Keterangan Mailer", "Material Promo", 
                        "Kode Material", "Alasan Material Tidak Terpasang", "Promo Tersetting di Server Kasir", 
                        "Sisa Stock", "Kode Stock", "Informasi Potongan Harga di Struk", "Expired Date", "Harga Produk"]


                    df_new = pd.DataFrame([new_data])
                    df_existing = pd.concat([df_existing, df_new], ignore_index=True)
                    df_existing = df_existing.fillna("-")
                    df_existing = df_existing[order]
                    df_existing.to_excel(file_path, index=False, engine="openpyxl")
                    st.success("Data berhasil disimpan!")
                    
                    # Simpan ke overview_data.xlsx
                    save_to_excel(st.session_state["overview"], EXCEL_FILE)
                    st.info("Jika sudah menginput silahkan input kembali, dengan memilih kembali produk lainnya yang sesuai (jika masih dalam satu outlet yang sama). Jika sudah berbeda outlet/kota, maka silahkan refresh website dan input kembali dari awal.")   
with tab2:
    st.subheader("Produk yang sudah diinput")

    # Load data saat aplikasi dijalankan
    df_overview = load_from_excel()
    if not df_overview.empty:
        list_surveyor = df_overview["Nama Surveyor"].dropna().unique().tolist()
        if list_surveyor:
            nama_surveyor_selected = st.selectbox("Nama Anda", list_surveyor)
            df_filtered = df_overview[df_overview["Nama Surveyor"] == nama_surveyor_selected]
            
            if not df_filtered.empty:
                df_filtered["Status"] = "Sudah terinput semua"
                st.dataframe(df_filtered[["Timestamp Pengisian", "Nama Produk", "Kode Outlet", "Status"]], hide_index=True)
            else:
                st.info(f"Tidak ada produk yang diinput oleh {nama_surveyor_selected}.")
    else:
        st.warning("Belum ada data yang tersimpan.")

    # **Tombol Hapus dengan Konfirmasi**
    st.info("Hapus jika semua sudah lengkap dalam 1 komponen produk!")
    
    with st.expander("⚠️ Konfirmasi Penghapusan Data"):
        konfirmasi = st.checkbox("Saya yakin ingin menghapus semua data overview")

        if konfirmasi:
            if st.button("Hapus Semua Data Overview", type="primary"):
                st.session_state["overview"] = pd.DataFrame(columns=["Timestamp Pengisian", "Nama Surveyor", "Nama Produk", "Kode Outlet", "Status"])
                save_to_excel(st.session_state["overview"], EXCEL_FILE)
                st.success("Semua data overview telah dihapus!")
                st.rerun()
        else:
            st.warning("Centang konfirmasi terlebih dahulu untuk menghapus data.")



with tab3:
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
            st.markdown("""
            **Petunjuk Penggunaan:**
            - Data survey yang sudah diinput akan tampil di bawah ini dan dapat langsung diedit.
            - Klik **Simpan Perubahan** untuk menyimpan perubahan data setelah merubah data langsung pada tabel di bawah.
            - Untuk Menghapus satu baris data: **Checkbox** pada bagian kiri sendiri di suatu baris yang ingin dihapus kemudian klik gambar **icon hapus** pada pojok kanan tabel (kecil), setelah itu klik **simpan perubahan**
            - Klik **Hapus Semua Data** jika ingin menghapus seluruh data survey.
            - Klik **Download Excel** untuk mengunduh data dalam format Excel.
            - **Surveyor** tidak perlu mendownload excel, dan jikalau revisi bisa dilakukan edit data atau delete data.
            """)
            st.markdown(f"""
            <div style='margin-bottom: 12px; text-align: center;'>
                Link Spreadsheet untuk Update Data Produk/Promo/Survey: <br>
                <a href='https://intip.in/UpdateProdukPromoSurveyHCO' target='_blank' 
                   style='background-color: #626F47; color: #ffffff; padding: 4px 8px; 
                          border-radius: 5px; font-size: 90%; font-weight: bold; text-decoration: none; display: inline-block; margin-top: 5px;'>
                    https://intip.in/UpdateProdukPromoSurveyHCO
                </a>
            </div>
            """, unsafe_allow_html=True)

            st.header("Hasil Inputan Data Survey Promo HCO Chain")
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
