import streamlit as st
import pandas as pd
from datetime import datetime, time
import os
from io import BytesIO
import openpyxl
from datetime import datetime
from zoneinfo import ZoneInfo
from datetime import datetime, date
import pytz
st.set_page_config(page_title="Form Survey Promo HCO Chain dan Lokal - Siantar Top - BSI ", page_icon=":bookmark_tabs:", layout="centered")
st.title("Form Survey Promo HCO Chain dan Lokal Tahun 2025")
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
import zipfile
import os

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
                    "periode_promo": row['Periode Promo'],
st.set_page_config(page_title="BSI - Support Information", layout="wide")

FOLDER_PATH = "saved_files"

if not os.path.exists(FOLDER_PATH):
    os.makedirs(FOLDER_PATH)

def load_files():
    files = []
    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith(".xlsx"):
            with open(os.path.join(FOLDER_PATH, filename), "rb") as f:
                files.append({
                    "name": filename,
                    "data": f.read()
                })

            tipe_outlet = st.selectbox("Pilih Outlet (Hari Ini):", list(outlet_data.keys()))

            if tipe_outlet:
                st.write(f"Promo di {tipe_outlet}:")
                for i, promo in enumerate(outlet_data[tipe_outlet], start=1):
                    st.markdown(f"""
                    <div style='margin-bottom: 12px;'>
                        <span style='background-color: #000000; color: #FFFFFF; padding: 4px 8px; 
                                     border-radius: 5px; font-size: 90%; margin-right: 8px; font-weight: bold;'>
                            {i}.
                        </span>
                        <span style='background-color: #89AC46; color: #000000; padding: 4px 8px; 
                                     border-radius: 5px; font-size: 90%; margin-right: 5px; font-weight: bold;'>
                            {promo['nama_produk']}
                        </span>
                        <span style='background-color: #E50046; color: #000000; padding: 4px 8px; 
                                     border-radius: 5px; font-size: 90%; margin-right: 5px; font-weight: bold;'>
                            {promo['jenis_promo']}
                        </span>
                        <span style='background-color: #626F47; color: #000000; padding: 4px 8px; 
                                     border-radius: 5px; font-size: 90%; font-weight: bold;'>
                            Periode: {promo['periode_promo']}
                        </span>
                    </div>
                    """, unsafe_allow_html=True)
    return files


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
if 'files' not in st.session_state:
    st.session_state.files = load_files()

if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False

        # Daftar outlet yang termasuk Chain
        chain_outlet = [
            "Indomaret", "Indogrosir", "Alfamart", 
            "Alfamidi", "Lion Superindo", "Clandys", "Family Mart"
        ]

        # Tentukan tipe account berdasarkan outlet yang dipilih
        if tipe_outlet in chain_outlet:
            tipe_account_value = "Chain"
        else:
            tipe_account_value = "Lokal"

        # Tampilkan tipe account (disabled)
        tipe_account = st.selectbox("Tipe Account (HCO):", ["Chain", "Lokal"], index=0 if tipe_account_value == "Chain" else 1, disabled=True)
        kode_outlet = st.text_input("Kode Outlet:", key="kode_outlet")
        st.caption("Isikan - jika tidak tau")

        tanggal = st.date_input("Tanggal", value=datetime.today(),key="tanggal")
if 'confirm_delete_all' not in st.session_state:
    st.session_state.confirm_delete_all = False

st.session_state.files = load_files()

# Mulai Tabs
tab1, tab2, tab3 = st.tabs(["📦 POB", "Penggabungan Data - POB", "📝 RNL"])

with tab1:
    st.header("📊 Masukkan File POB")

    # Step 1: Pilih POB
    selected_pob = st.selectbox("Pilih POB", ('', 'POB - Dist', 'POB - SSO'))

    # Step 2: Jika sudah pilih POB, baru muncul pilihan MT/GT
    if selected_pob:
        selected_channel = st.selectbox("Pilih Channel", ('MT', 'GT'))

        # Step 3: Jika sudah pilih MT/GT, baru muncul upload
        if selected_channel:
            uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

            if uploaded_file is not None:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names

                all_data = {}  # Dictionary untuk menyimpan data per sheet
                all_results = []  # List untuk menggabungkan semua sheet

                for sheet_name in sheet_names:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

#--------------------------------------------------------KODE UNTUK POB-DIST, MT

                    if selected_pob == "POB - Dist" and selected_channel == "MT":
                        # Ambil informasi dari file
                        dist = df.iloc[1, 1]
                        area = df.iloc[2, 1]
                        cabang = df.iloc[3, 1]
                        bulan = df.iloc[4, 1]

                        # Mapping bulan
                        bulan_mapping = {
                            "January": "Januari", "February": "Februari", "March": "Maret",
                            "April": "April", "May": "Mei", "June": "Juni", "July": "Juli",
                            "August": "Agustus", "September": "September", "October": "Oktober",
                            "November": "November", "December": "Desember"
                        }

                        bulan_minus = {
                            "Januari": "Desember", "Februari": "Januari", "Maret": "Februari",
                            "April": "Maret", "Mei": "April", "Juni": "Mei", "Juli": "Juni",
                            "Agustus": "Juli", "September": "Agustus", "Oktober": "September",
                            "November": "Oktober", "Desember": "November"
                        }

                        bulan_plus = {
                            "Januari": "Februari", "Februari": "Maret", "Maret": "April",
                            "April": "Mei", "Mei": "Juni", "Juni": "Juli", "Juli": "Agustus",
                            "Agustus": "September", "September": "Oktober", "Oktober": "November",
                            "November": "Desember", "Desember": "Januari"
                        }

                        bulan_plus2 = {
                            "Januari": "Maret", "Februari": "April", "Maret": "Mei",
                            "April": "Juni", "Mei": "Juli", "Juni": "Agustus",
                            "Juli": "September", "Agustus": "Oktober", "September": "November",
                            "Oktober": "Desember", "November": "Januari", "Desember": "Februari"
                        }

                        tahun = datetime.now().year
                        bulan_str = bulan.strftime('%B') if isinstance(bulan, datetime) else str(bulan)
                        nama_bulan = bulan_mapping.get(bulan_str, bulan_str)
                        bulan_minus_1 = bulan_minus.get(nama_bulan, nama_bulan)
                        bulan_plus_fix = bulan_plus.get(nama_bulan, nama_bulan)
                        bulan_plus2_fix = bulan_plus2.get(nama_bulan, nama_bulan)
                        try:
                            if isinstance(bulan, str):  # Pastikan bulan berupa string sebelum dikonversi
                                bulan_dt = datetime.strptime(bulan, "%B")  # Ubah ke datetime (format nama bulan lengkap)
                            else:
                                bulan_dt = bulan  # Jika sudah datetime, langsung gunakan

                            bulan_formatted = bulan_dt.strftime('%b-%y')  # Format singkat (contoh: "Mar-24")
                        except ValueError:
                            bulan_formatted = bulan  # Jika gagal parsing, tetap gunakan nilai aslinya


                        # Data Cleaning sesuai POB dan Channel
                        item_products = df.iloc[9:, 1].dropna().astype(str).str.strip()
                        total_final = df.iloc[9:, 95].replace("-", 0).fillna(0)
                        forecast_1 = df.iloc[9:, 117].replace("-", 0).fillna(0)
                        forecast_2 = df.iloc[9:, 127].replace("-", 0).fillna(0)

                        # Konversi ke numerik agar aman
                        total_final = pd.to_numeric(total_final, errors='coerce').fillna(0)
                        forecast_1 = pd.to_numeric(forecast_1, errors='coerce').fillna(0)
                        forecast_2 = pd.to_numeric(forecast_2, errors='coerce').fillna(0)

                        # Hapus baris dengan teks tidak relevan
                        remove_keywords = ["total", "MOHON DIPERHATIKAN", "Jika", "adjust average", "Stock ED"]
                        mask_valid_products = ~item_products.astype(str).str.contains('|'.join(remove_keywords), case=False, na=False)

                        item_products = item_products[mask_valid_products].reset_index(drop=True)
                        total_final = total_final[mask_valid_products].reset_index(drop=True)
                        forecast_1 = forecast_1[mask_valid_products].reset_index(drop=True)
                        forecast_2 = forecast_2[mask_valid_products].reset_index(drop=True)

                        # Hapus baris terakhir (biasanya total)
                        item_products = item_products.iloc[:-1].reset_index(drop=True)
                        total_final = total_final.iloc[:-1].reset_index(drop=True)
                        forecast_1 = forecast_1.iloc[:-1].reset_index(drop=True)
                        forecast_2 = forecast_2.iloc[:-1].reset_index(drop=True)

                        # Buat dataframe dari sheet ini
                        result_df = pd.DataFrame({
                            'POB': [selected_pob] * len(item_products),
                            'Channel': [selected_channel] * len(item_products),
                            'Dist': [dist] * len(item_products),
                            'Area': [area] * len(item_products),
                            'Cabang': [cabang] * len(item_products),
                            #'Nama Sheet': [sheet_name] * len(item_products),  # Tambahkan nama sheet
                            'Bulan': [bulan_formatted] * len(item_products),
                            'Item Product': item_products,
                            'Total Final POB Adjust RM-AM / DISt': total_final,
                            f'Forecast {bulan_plus_fix}-{tahun}': forecast_1,
                            f'Forecast {bulan_plus2_fix}-{tahun}': forecast_2          
                        })
#-------------------------------------------------------KODE UNTUK POB-SSO, MT

                    elif selected_pob == "POB - SSO" and selected_channel == "MT":
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        now = datetime.now()
                        nama_bulan = now.strftime("%B")  # Nama bulan dalam format "January", "February", dll.
                        tahun = now.year  # Tahun dalam format 2025, 2026, dll.
                        df[0] = df[0].astype(str).str.strip().str.upper()  # Kolom A
                        df[1] = df[1].astype(str).str.strip().str.upper()  # Kolom B

                        total_index = df[(df[0] == "TOTAL") | (df[1] == "TOTAL")].index
                        if total_index.empty:
                            st.warning(f"Tidak ditemukan baris 'TOTAL' di sheet '{sheet_name}'. Lewati sheet ini!")

                        total_index = total_index[0]  # Ambil indeks pertama

                        # **Cabang**
                        try:
                            cabang = str(df.iloc[0, 1]).strip()  # Ambil cabang di baris pertama kolom ke-2
                            if pd.isna(cabang) or cabang == "":
                                cabang = "Unknown"
                        except IndexError:
                            cabang = "Unknown"

                        # **Produk**
                        products = df.iloc[2:total_index, 1].dropna().tolist()

                        # **Distributor**
                        distributors = []
                        col_start = 2  # Mulai dari kolom C
                        while col_start < df.shape[1]:
                            try:
                                dist_name = str(df.iloc[0, col_start]).strip()
                                if pd.notna(df.iloc[0, col_start + 1]):
                                    dist_name += " " + str(df.iloc[0, col_start + 1]).strip()
                                distributors.append(dist_name)
                            except IndexError:
                                break  # Jika kolom tidak ada, hentikan loop
                            col_start += 6  # Lompat 6 kolom

                        # **Buat Data**
                        data = []
                        col_start = 2  # Mulai dari kolom C
                        for distributor in distributors:
                            for product_idx, product in enumerate(products):
                                for week_idx, week in enumerate(["W1", "W2", "W3", "W4", "W5"]):
                                    try:
                                        nilai = df.iloc[2 + product_idx, col_start + week_idx]  # Ambil nilai berdasarkan Week
                                        nilai = 0 if pd.isna(nilai) else nilai  # Ganti NaN jadi 0
                                        data.append([timestamp, sheet_name, product, distributor, cabang, week, nilai])
                                    except IndexError:
                                        data.append([timestamp, sheet_name, product, distributor, cabang, week, 0])  # Jika out of bounds, isi 0
                            col_start += 6  # Lompat ke distributor berikutnya

                        # **Buat DataFrame**
                        result_df = pd.DataFrame(data, columns=["Timestamp", "Sheet", "Produk", "Distributor", "Cabang", "Week", "Nilai"])

                        # **Hapus angka "0 0 0 0" dari distributor secara aman**
                        result_df["Distributor"] = result_df["Distributor"].str.replace(r"(\s0+)+$", "", regex=True)

#-----------------------------------------KODE UNTUK POB-SSO, GT
                    elif selected_pob == "POB - SSO" and selected_channel == "GT":
                        1

#---------------------------------------RESULT DAN PENYIMPANAN

                   # Simpan data yang sudah dibersihkan
                    all_data[sheet_name] = result_df
                    all_results.append(result_df)

                # Gabungkan semua sheet jadi satu DataFrame
                if all_results:
                    final_result = pd.concat(all_results, ignore_index=True)

                    # Pilih sheet untuk ditampilkan
                    selected_sheet = st.selectbox("Pilih untuk Melihat Sheet!", sheet_names)

                    if selected_sheet in all_data:
                        st.write(f"### Data dari {selected_sheet}")
                        st.dataframe(all_data[selected_sheet])

        col1, col2 = st.columns(2)
                    if st.button("Proses dan Simpan!"):
                        safe_branch_name = cabang.replace("/", "_")

                        # Tentukan format nama file berdasarkan sheet yang dipilih
                        if "POB - SSO" in selected_sheet and "MT" in selected_sheet:
                            filename = f"POB - SSO - {nama_bulan} - {tahun}.xlsx"
                        elif "POB - DIST" in selected_sheet and "MT" in selected_sheet:
                            filename = f"PO {nama_bulan} {safe_branch_name} {tahun}.xlsx"
                        elif "POB - SSO" in selected_sheet and "GT" in selected_sheet:
                            filename = f"POB - SSO - GT - {nama_bulan} - {tahun}.xlsx"
                        elif "POB - DIST" in selected_sheet and "GT" in selected_sheet:
                            filename = f"PO {nama_bulan} {safe_branch_name} GT {tahun}.xlsx"
                        else:
                            filename = f"PO {nama_bulan} {safe_branch_name} {tahun}.xlsx"  # Default jika tidak ada kondisi yang cocok


                        def get_unique_filename(folder_path, filename):
                            base, ext = os.path.splitext(filename)
                            counter = 1
                            new_filename = filename

                            while os.path.exists(os.path.join(folder_path, new_filename)):
                                new_filename = f"{base} ({counter}){ext}"
                                counter += 1

                            return new_filename  # Harus ada return ini!

                        # Panggil fungsi untuk mendapatkan nama file unik
                        filename = get_unique_filename(FOLDER_PATH, filename)
                        file_path = os.path.join(FOLDER_PATH, filename)

                        # Simpan ke Excel
                        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                            final_result.to_excel(writer, sheet_name="POB Combined", index=False)

                        st.success(f"✅ File telah dicleaning dan berhasil disimpan di Overview!: `{filename}`")
                        st.rerun()  # Paksa Streamlit untuk refresh tanpa manual reload                       

    st.subheader("📂 Overview Saved Files")

    if st.session_state.files:
        selected_files = []
        select_all = st.checkbox("Select All")
        for file in st.session_state.files:
            checked = st.checkbox(file['name'], key=file['name'], value=select_all)
            if checked:
                selected_files.append(file['name'])

        col1, col2, _ = st.columns([1, 1, 5])
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

            if selected_files:
                if len(selected_files) == 1:
                    # Download langsung file XLSX tanpa membaca CSV
                    file_name = selected_files[0]
                    file_path = os.path.join(FOLDER_PATH, file_name)

#---------------------------------------------------------------DISPLAY PRODUK
        st.subheader(f"Detail Produk")
        produk_display = st.selectbox(
            f"Apakah produk {nama_produk} terdisplay di toko?", 
            [" ", "Iya", "Stock Kosong", "Tidak Jual"], 
            key="produk_display"
        )

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
                harga_produk = st.number_input(
                    f"Berapa harga produk {nama_produk} per pcs yang tertera di rak / server kasir?",
                    min_value=0, key="harga_produk"
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
                expired_date = st.date_input(
                    f"Tanggal Expired Date produk {nama_produk}", 
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
                        ["", "Ya", "Tidak"], 
                        key="promo_di_kasir"
                    )
                    info_struk = st.text_input("Informasi potongan harga yang tertera di struk:", key="info_struk")

            # --- SUBMIT ---
            if st.button("Submit"):
                errors = []

                # Validasi field wajib utama
                if not nama_surveyor or not kode_outlet or not kota:
                    errors.append("Nama Surveyor, Kode Outlet, dan Kota wajib diisi!")

                # --- Validasi Step Lainnya
                if produk_display in ["Iya", "Stock Kosong", "Tidak Jual"]:
                    # Field yang hanya wajib untuk Chain
                    if tipe_account_value == "Chain":
                        if promo_mailer.strip() == "":
                            errors.append("Promo Mailer wajib diisi untuk account Chain.")
                        if expired_date is None:
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
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()


                # Jika ada error
                if errors:
                    for err in errors:
                        st.error(err)
                        st.download_button(
                            label="📥 Download File",
                            data=file_bytes,
                            file_name=file_name,
                            mime="application/octet-stream",
                            key="download_single_btn"
                        )
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
                        "Timestamp Pengisian" : datetime.now(ZoneInfo('Asia/Jakarta')).strftime('%Y-%m-%d %H:%M:%S'),
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
                        "Sisa Stock", "Kode Stock", "Informasi Potongan Harga di Struk", "Expired Date", "Harga Produk"
                    # Jika lebih dari satu file, buat ZIP
                    if st.button("📥 Download Selected as ZIP"):
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zipf:
                            for file in st.session_state.files:
                                if file['name'] in selected_files:
                                    zipf.writestr(file['name'], file['data'])

                        st.download_button(
                            label="Download ZIP",
                            data=zip_buffer.getvalue(),
                            file_name="datasets.zip",
                            mime="application/zip",
                            key="download_zip_btn"
                        )

        with col2:
            if selected_files and not st.session_state.confirm_delete:
                if st.button("🗑️ Delete Selected"):
                    st.session_state.confirm_delete = True
                    st.rerun()

        if st.session_state.confirm_delete:
            st.warning("⚠️ Yakin ingin mendelete file yang dipilih?")
            col_ok, col_cancel = st.columns(2)
            with col_ok:
                if st.button("✅ Ya, Delete"):
                    for fname in selected_files:
                        file_path = os.path.join(FOLDER_PATH, fname)
                        if os.path.exists(file_path):
                            os.remove(file_path)
                    st.session_state.files = [
                        file for file in st.session_state.files if file['name'] not in selected_files
                    ]

                    df_new = pd.DataFrame([new_data])
                    df_existing = pd.concat([df_existing, df_new], ignore_index=True)
                    df_existing = df_existing.fillna("-")
                    df_existing = df_existing[order]
                    df_existing.to_excel(file_path, index=False, engine="openpyxl")
                    st.success("Data berhasil disimpan!")

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
                    st.session_state.confirm_delete = False
                    st.rerun()
            with col_cancel:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete = False
                    st.rerun()

        st.divider()
        if not st.session_state.confirm_delete_all:
            if st.button("🗑️ Delete All Files"):
                st.session_state.confirm_delete_all = True
                st.rerun()
            else:
                st.error("Password salah, coba lagi.")

        if st.session_state.confirm_delete_all:
            st.warning("⚠️ Yakin ingin mendelete SEMUA file?")
            col_ok_all, col_cancel_all = st.columns(2)
            with col_ok_all:
                if st.button("✅ Ya, Delete All"):
                    for file in os.listdir(FOLDER_PATH):
                        os.remove(os.path.join(FOLDER_PATH, file))
                    st.session_state.files = []
                    st.session_state.confirm_delete_all = False
                    st.rerun()
            with col_cancel_all:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete_all = False
                    st.rerun()
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
        st.info("Belum ada file yang disimpan.")

with tab2:
    st.header("🔗 Merge POB Files")

    if st.session_state.files:
        selected_files = st.multiselect("Pilih file untuk digabungkan:", [file["name"] if isinstance(file, dict) else file for file in st.session_state.files])

                with colB:
                    # Tombol awal untuk trigger konfirmasi
                    if "show_confirm" not in st.session_state:
                        st.session_state.show_confirm = False
        if st.button("🔄 Merge Files"):
            if selected_files:
                merged_data = []

                    if st.button("Hapus Semua Data"):
                        st.session_state.show_confirm = True
                for file in selected_files:
                    file_path = os.path.join(FOLDER_PATH, file)
                    df = pd.read_excel(file_path)
                    merged_data.append(df)

                    # Jika tombol sudah ditekan, baru munculkan konfirmasi
                    if st.session_state.show_confirm:
                        st.warning("⚠️ Apakah Anda yakin ingin menghapus semua data? Tindakan ini tidak dapat dibatalkan!")
                final_df = pd.concat(merged_data, ignore_index=True)

                        confirm = st.checkbox("Saya yakin ingin menghapus semua data")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='Merged POB', index=False)
                output.seek(0)

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
                st.download_button(
                    label="📥 Download Merged File",
                    data=output,
                    file_name="Merged_POB.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("✅ File berhasil digabung!")
            else:
                st.info("Belum ada data yang tersimpan.")
                st.warning("⚠️ Pilih minimal satu file untuk digabungkan.")
    else:
        st.info("Belum ada file yang tersedia untuk digabungkan.")
