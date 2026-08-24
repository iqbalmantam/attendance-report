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
    Membaca file Excel, mengabaikan sheet lembur, membersihkan duplikasi kolom,
    dan melakukan unpivot data secara otomatis dari sheet operasional harian.
    """
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    
    # Abaikan worksheet terkait lembur
    ignore_keywords = ['ot', 'lembur', 'summary ot']
    
    all_data = []
    
    for sheet in sheets:
        if any(keyword in sheet.lower() for keyword in ignore_keywords):
            continue
            
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # --- PENANGANAN KHUSUS SHEET DW(ARU) ---
            if 'aru' in sheet.lower():
                header_row = 4
                date_row = 5
                data_start = 6
                
                if len(df) > data_start:
                    raw_cols = df.iloc[header_row].tolist()
                    date_vals = df.iloc[date_row].tolist()
                    
                    cols = []
                    for idx, val in enumerate(raw_cols):
                        if pd.notna(val):
                            cols.append(str(val).strip())
                        else:
                            cols.append(f"Col_{idx}")
                            
                    for idx, val in enumerate(date_vals):
                        if isinstance(val, pd.Timestamp) or hasattr(val, 'day'):
                            cols[idx] = val.day
                            
                    df.columns = cols
                    df_data = df.iloc[data_start:data_start+1].dropna(how='all') # Ambil baris headcount
                    
                    id_vars = [col for col in df.columns if not isinstance(col, int)]
                    value_vars = [col for col in df.columns if isinstance(col, int) and 1 <= col <= 31]
                    
                    if value_vars:
                        melted = df_data.melt(id_vars=id_vars, value_vars=value_vars, var_name='Tanggal', value_name='Jumlah_MP')
                        filename = getattr(file, 'name', 'Laporan Bulanan')
                        melted['Sumber_Sheet'] = sheet
                        melted['Bulan_Laporan'] = filename.split('.')[0]
                        melted['Jumlah_MP'] = pd.to_numeric(melted['Jumlah_MP'], errors='coerce').fillna(0)
                        melted = melted[melted['Jumlah_MP'] > 0]
                        
                        kolom_entitas = id_vars[0] if id_vars else 'Entity'
                        for col in id_vars:
                            if any(kw in str(col).lower() for kw in ['job', 'project', 'status', 'entity', 'vendor']):
                                kolom_entitas = col
                                break
                                
                        final_df = melted[['Bulan_Laporan', 'Sumber_Sheet', kolom_entitas, 'Tanggal', 'Jumlah_MP']].copy()
                        final_df.columns = ['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah']
                        final_df['Kategori'] = 'DW(ARU)'
                        all_data.append(final_df)
                continue

            # --- PARSER STANDAR (Temporary, DW(Dedicated), dll) ---
            date_row_idx = None
            for i in range(min(15, len(df))):
                row_values = df.iloc[i]
                numeric_days = 0
                for val in row_values:
                    if isinstance(val, (int, float)) and 1 <= val <= 31:
                        numeric_days += 1
                    elif isinstance(val, pd.Timestamp) or hasattr(val, 'day'):
                        numeric_days += 1
                if numeric_days >= 10:
                    date_row_idx = i
                    break
                    
            if date_row_idx is not None:
                raw_cols = df.iloc[date_row_idx].fillna('Deskripsi').tolist()
                seen = {}
                new_cols = []
                for c in raw_cols:
                    if isinstance(c, (int, float)) and 1 <= c <= 31:
                        new_cols.append(c)
                    elif isinstance(c, pd.Timestamp) or hasattr(c, 'day'):
                        new_cols.append(getattr(c, 'day', c))
                    else:
                        col_str = str(c).strip()
                        if col_str in seen:
                            seen[col_str] += 1
                            new_cols.append(f"{col_str}_{seen[col_str]}")
                        else:
                            seen[col_str] = 0
                            new_cols.append(col_str)
                            
                df.columns = new_cols
                df_data = df.iloc[date_row_idx+1:].dropna(how='all')
                
                id_vars = [col for col in df.columns if not isinstance(col, (int, float))]
                value_vars = [col for col in df.columns if isinstance(col, (int, float)) and 1 <= col <= 31]
                
                if not value_vars:
                    continue
                
                melted = df_data.melt(
                    id_vars=id_vars, 
                    value_vars=value_vars, 
                    var_name='Tanggal', 
                    value_name='Jumlah_MP'
                )
                
                filename = getattr(file, 'name', 'Laporan Bulanan')
                melted['Sumber_Sheet'] = sheet
                melted['Bulan_Laporan'] = filename.split('.')[0]
                
                melted = melted.dropna(subset=['Jumlah_MP'])
                melted['Jumlah_MP'] = pd.to_numeric(melted['Jumlah_MP'], errors='coerce').fillna(0)
                melted = melted[melted['Jumlah_MP'] > 0] 
                
                if id_vars:
                    kolom_entitas = id_vars[0]
                    for col in id_vars:
                        if isinstance(col, str) and any(kw in col.lower() for kw in ['company', 'entity', 'vendor', 'nama', 'name', 'job', 'project']):
                            kolom_entitas = col
                            break
                            
                    final_df = melted[['Bulan_Laporan', 'Sumber_Sheet', kolom_entitas, 'Tanggal', 'Jumlah_MP']]
                    final_df.columns = ['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah']
                    final_df['Kategori'] = sheet.strip()
                    all_data.append(final_df)
                    
        except Exception as e:
            st.warning(f"Catatan pada sheet '{sheet}': {e}")
                
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
    st.markdown("Unggah file **JDC MP Daily Report** Anda. Sistem akan memproses data kehadiran dan entitas secara otomatis.")
    
    uploaded_file = st.file_uploader("Pilih file .xlsx", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Memproses file Excel dengan algoritma cerdas..."):
            extracted_data = process_excel(uploaded_file)
            
            if not extracted_data.empty:
                st.success(f"Berhasil mengekstrak {len(extracted_data)} baris data operasional!")
                st.dataframe(extracted_data.head(10))
                
                if st.button("💾 Simpan ke Google Sheets"):
                    with st.spinner("Menyinkronkan dengan Google Sheets..."):
                        try:
                            client = get_gsheets_client()
                            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
                            
                            # Cek apakah sheet masih kosong, jika ya tambahkan header
                            existing_data = sheet.get_all_values()
                            if not existing_data:
                                sheet.append_row(['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah'])
                                
                            # PENTING: Bersihkan NaN menjadi string kosong agar compliant dengan JSON
                            extracted_data = extracted_data.fillna("")
                            data_to_upload = extracted_data.astype(str).values.tolist()
                            sheet.append_rows(data_to_upload)
                            
                            st.success("Data berhasil ditambahkan dan diamankan ke Google Sheets! ✅")
                        except Exception as e:
                            st.error(f"Gagal menyimpan ke Google Sheets: {e}")
            else:
                st.error("Tidak ditemukan format tabel tanggal (1-31) yang valid di file ini.")

# --- TAB 1: DASHBOARD ---
with tab1:
    st.subheader("Ringkasan Data Tersimpan")
    
    df_db = pd.DataFrame()
    with st.spinner("Memuat data dari database Google Sheets..."):
        try:
            client = get_gsheets_client()
            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
            records = sheet.get_all_records()
            if records:
                df_db = pd.DataFrame(records)
        except Exception as e:
            st.warning("Belum ada data tersimpan di Google Sheets atau koneksi gagal.")
        
    if not df_db.empty:
        # Konversi tipe data
        df_db['Jumlah'] = pd.to_numeric(df_db['Jumlah'], errors='coerce').fillna(0)
        df_db['Tanggal'] = pd.to_numeric(df_db['Tanggal'], errors='coerce')
        
        # --- Filter Interaktif ---
        col1, col2 = st.columns(2)
        kategori_list = df_db['Kategori'].unique().tolist()
        pilihan_kategori = col1.multiselect("Filter Berdasarkan Kategori / Sheet", kategori_list, default=kategori_list)
        
        file_list = df_db['File_Sumber'].unique().tolist() if 'File_Sumber' in df_db.columns else []
        pilihan_file = col2.multiselect("Filter Berdasarkan Bulan / File", file_list, default=file_list)
        
        df_filtered = df_db[df_db['Kategori'].isin(pilihan_kategori)]
        if pilihan_file and 'File_Sumber' in df_db.columns:
            df_filtered = df_filtered[df_filtered['File_Sumber'].isin(pilihan_file)]
        
        if df_filtered.empty:
            st.info("Pilih setidaknya satu kategori/file untuk menampilkan grafik.")
        else:
            # --- Indikator Utama (KPI) ---
            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manpower (Hari-Orang)", f"{int(df_filtered['Jumlah'].sum()):,}")
            m2.metric("Jumlah Kategori Aktif", len(df_filtered['Kategori'].unique()))
            m3.metric("Total Entitas / Posisi", len(df_filtered['Entitas_Posisi'].unique()))
            
            # --- Visualisasi ---
            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("### Tren Harian Manpower")
                tren_harian = df_filtered.groupby(['Tanggal', 'Kategori'])['Jumlah'].sum().reset_index()
                tren_harian = tren_harian.sort_values(by='Tanggal')
                
                fig_line = px.line(
                    tren_harian, 
                    x='Tanggal', 
                    y='Jumlah', 
                    color='Kategori', 
                    markers=True,
                    labels={'Jumlah': 'Total Manpower', 'Tanggal': 'Tanggal'}
                )
                fig_line.update_xaxes(dtick=1)
                st.plotly_chart(fig_line, use_container_width=True)
                
            with chart_col2:
                st.markdown("### Komposisi Top 10 Entitas / Posisi")
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
                st.dataframe(df_db, use_container_width=True)
    else:
        st.info("Belum ada data di Google Sheets. Silakan ke tab 'Upload & Sinkronisasi Data' untuk mengunggah laporan bulanan Anda.")