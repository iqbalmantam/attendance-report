import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# ==========================================
# 1. KONFIGURASI GOOGLE SHEETS
# ==========================================
def get_gsheets_client():
    """Mengambil credentials dari Streamlit Secrets dan melakukan otorisasi."""
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    # Mengambil rahasia dari konfigurasi TOML di Streamlit Cloud
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

# URL Google Sheets Anda
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1lv0b68kBixRgYmJJk1J6Eppj7wXu5u6C8g81zzQiKg0/edit?gid=0#gid=0"

# ==========================================
# 2. ALGORITMA EKSTRAKSI CERDAS (SMART PARSER)
# ==========================================
def process_excel(file):
    """
    Membaca file Excel, mengabaikan sheet lembur, dan mencari baris tanggal
    untuk melakukan unpivot data secara otomatis.
    """
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    
    # Abaikan worksheet terkait lembur sesuai instruksi
    ignore_keywords = ['ot', 'lembur', 'summary ot']
    
    all_data = []
    
    for sheet in sheets:
        if any(keyword in sheet.lower() for keyword in ignore_keywords):
            continue
            
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # Heuristik: Cari baris yang berfungsi sebagai header tanggal (berisi angka 1-31)
            date_row_idx = None
            for i in range(min(15, len(df))): # Scan 15 baris pertama
                row_values = pd.to_numeric(df.iloc[i], errors='coerce')
                valid_days = row_values[(row_values >= 1) & (row_values <= 31)]
                
                # Jika ada minimal 15 kolom dengan angka hari, asumsikan ini baris tanggal
                if len(valid_days) >= 15: 
                    date_row_idx = i
                    break
                    
            if date_row_idx is not None:
                # Jadikan baris tersebut sebagai header sementara
                df.columns = df.iloc[date_row_idx].fillna('Deskripsi')
                
                # Ambil data di bawah header dan buang baris kosong
                df_data = df.iloc[date_row_idx+1:].dropna(how='all')
                
                # Pisahkan kolom Identitas (teks) dan kolom Tanggal (angka)
                id_vars = [col for col in df.columns if not isinstance(col, (int, float))]
                value_vars = [col for col in df.columns if isinstance(col, (int, float)) and 1 <= col <= 31]
                
                # UNPIVOT DATA (Melt)
                melted = df_data.melt(
                    id_vars=id_vars, 
                    value_vars=value_vars, 
                    var_name='Tanggal', 
                    value_name='Jumlah_MP'
                )
                
                # Tambahkan metadata
                melted['Sumber_Sheet'] = sheet
                melted['Bulan_Laporan'] = file.name.split('.')[0]
                
                # Cleaning data
                melted = melted.dropna(subset=['Jumlah_MP'])
                melted['Jumlah_MP'] = pd.to_numeric(melted['Jumlah_MP'], errors='coerce').fillna(0)
                melted = melted[melted['Jumlah_MP'] > 0] # Hanya ambil data yang ada isinya
                
                # Ambil kolom identitas utama (asumsi kolom pertama dari kiri yang bukan angka)
                if id_vars:
                    kolom_entitas = id_vars[0]
                    # Susun ulang kolom untuk database
                    final_df = melted[['Bulan_Laporan', 'Sumber_Sheet', kolom_entitas, 'Tanggal', 'Jumlah_MP']]
                    final_df.columns = ['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah']
                    all_data.append(final_df)
                    
        except Exception as e:
            st.warning(f"Gagal memproses sheet '{sheet}': {e}")
                
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()

# ==========================================
# 3. ANTARMUKA STREAMLIT (UI)
# ==========================================
st.set_page_config(page_title="Dashboard JDC Manpower", page_icon="📊", layout="wide")
st.title("📊 JDC Manpower Dashboard Report")

tab1, tab2 = st.tabs(["📈 Dashboard Utama", "⚙️ Upload & Sinkronisasi Data"])

# --- TAB 2: UPLOAD DATA ---
with tab2:
    st.subheader("Upload Laporan Harian (Excel)")
    st.markdown("Unggah file **JDC MP Daily Report** Anda. Sistem akan memproses data kehadiran secara otomatis.")
    
    uploaded_file = st.file_uploader("Pilih file .xlsx", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Memproses data..."):
            extracted_data = process_excel(uploaded_file)
            
            if not extracted_data.empty:
                st.success(f"Berhasil mengekstrak {len(extracted_data)} baris data!")
                st.dataframe(extracted_data.head(10))
                
                if st.button("💾 Simpan ke Google Sheets"):
                    with st.spinner("Menyinkronkan dengan Google Sheets..."):
                        try:
                            client = get_gsheets_client()
                            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
                            
                            # Konversi data ke format list untuk gspread
                            data_to_upload = extracted_data.astype(str).values.tolist()
                            sheet.append_rows(data_to_upload)
                            
                            st.success("Data berhasil ditambahkan ke Google Sheets! ✅")
                        except Exception as e:
                            st.error(f"Gagal menyimpan ke Google Sheets: {e}")
            else:
                st.error("Tidak ditemukan format tabel tanggal (1-31) yang valid di file ini.")

# --- TAB 1: DASHBOARD ---
with tab1:
    st.subheader("Ringkasan Data Tersimpan")
    
    # Tarik data dari Google Sheets
    df_db = pd.DataFrame()
    with st.spinner("Memuat data dari database..."):
        try:
            client = get_gsheets_client()
            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
            records = sheet.get_all_records()
            if records:
                df_db = pd.DataFrame(records)
        except Exception as e:
            st.warning("Gagal terhubung ke Google Sheets atau file masih kosong.")
        
    if not df_db.empty:
        # Konversi tipe data agar bisa dihitung
        df_db['Jumlah'] = pd.to_numeric(df_db['Jumlah'], errors='coerce').fillna(0)
        df_db['Tanggal'] = pd.to_numeric(df_db['Tanggal'], errors='coerce')
        
        # --- Filter Interaktif ---
        col1, col2 = st.columns(2)
        kategori_list = df_db['Kategori'].unique().tolist()
        pilihan_kategori = col1.multiselect("Filter Berdasarkan Kategori Sheet", kategori_list, default=kategori_list)
        
        df_filtered = df_db[df_db['Kategori'].isin(pilihan_kategori)]
        
        if df_filtered.empty:
            st.info("Pilih setidaknya satu kategori untuk menampilkan grafik.")
        else:
            # --- Indikator Utama (KPI) ---
            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manpower (Hari-Orang)", f"{int(df_filtered['Jumlah'].sum())}")
            m2.metric("Jumlah Kategori", len(df_filtered['Kategori'].unique()))
            m3.metric("Entitas/Posisi Terdaftar", len(df_filtered['Entitas_Posisi'].unique()))
            
            # --- Visualisasi ---
            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("### Tren Harian")
                # Agregasi data per tanggal dan kategori
                tren_harian = df_filtered.groupby(['Tanggal', 'Kategori'])['Jumlah'].sum().reset_index()
                # Urutkan berdasarkan tanggal agar garisnya rapi
                tren_harian = tren_harian.sort_values(by='Tanggal')
                
                fig_line = px.line(
                    tren_harian, 
                    x='Tanggal', 
                    y='Jumlah', 
                    color='Kategori', 
                    markers=True,
                    labels={'Jumlah': 'Total Manpower', 'Tanggal': 'Tanggal Laporan'}
                )
                # Atur sumbu X agar menampilkan angka bulat
                fig_line.update_xaxes(dtick=1)
                st.plotly_chart(fig_line, use_container_width=True)
                
            with chart_col2:
                st.markdown("### Komposisi Top 10 Entitas/Posisi")
                komposisi = df_filtered.groupby('Entitas_Posisi')['Jumlah'].sum().reset_index()
                komposisi = komposisi.sort_values(by='Jumlah', ascending=False).head(10)
                
                fig_bar = px.bar(
                    komposisi, 
                    x='Jumlah', 
                    y='Entitas_Posisi', 
                    orientation='h',
                    color='Jumlah',
                    color_continuous_scale='Blues',
                    labels={'Jumlah': 'Total', 'Entitas_Posisi': 'Posisi/Entitas'}
                )
                fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_bar, use_container_width=True)
                
            # Menampilkan Raw Data Database di bawah
            with st.expander("Lihat Data Mentah (Tersimpan di GSheet)"):
                st.dataframe(df_db)
    else:
        st.info("Belum ada data. Silakan ke tab 'Upload & Sinkronisasi Data' untuk mengunggah laporan pertama Anda.")