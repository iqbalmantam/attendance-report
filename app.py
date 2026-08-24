import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# ==========================================
# 0. KONFIGURASI HALAMAN & WATERMARK
# ==========================================
st.set_page_config(page_title="Dashboard JDC Manpower", page_icon="📊", layout="wide")

# Menyembunyikan menu footer dan header logo GitHub bawaan Streamlit
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

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
# 2. ALGORITMA EKSTRAKSI CERDAS (KHUSUS 5 SHEET PILIHAN)
# ==========================================
def process_excel(file):
    """
    Membaca file Excel dan HANYA memproses 5 worksheet pilihan:
    Attendance, OS, Fix, DW(Dedicated), DW(ARU).
    """
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    
    target_sheets = ['attendance', 'os', 'fix', 'dw(dedicated)', 'dw(aru)']
    
    all_data = []
    filename = getattr(file, 'name', 'Laporan Bulanan').split('.')[0]
    
    for sheet in sheets:
        clean_sheet_name = sheet.strip()
        matched = False
        for ts in target_sheets:
            if clean_sheet_name.lower().replace(" ", "") == ts.replace(" ", ""):
                matched = True
                break
        
        if not matched:
            continue
            
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # --- A. PENANGANAN KHUSUS SHEET DW(ARU) ---
            if 'aru' in clean_sheet_name.lower() and 'dw' in clean_sheet_name.lower():
                header_row = 4
                date_row = 5
                data_start = 6
                if len(df) > data_start:
                    raw_cols = df.iloc[header_row].tolist()
                    date_vals = df.iloc[date_row].tolist()
                    cols = [str(val).strip() if pd.notna(val) else f"Col_{idx}" for idx, val in enumerate(raw_cols)]
                    for idx, val in enumerate(date_vals):
                        if isinstance(val, pd.Timestamp) or hasattr(val, 'day'):
                            cols[idx] = val.day
                    df.columns = cols
                    df_data = df.iloc[data_start:data_start+1].dropna(how='all')
                    id_vars = [col for col in df.columns if not isinstance(col, int)]
                    value_vars = [col for col in df.columns if isinstance(col, int) and 1 <= col <= 31]
                    if value_vars:
                        melted = df_data.melt(id_vars=id_vars, value_vars=value_vars, var_name='Tanggal', value_name='Jumlah_MP')
                        melted['Sumber_Sheet'] = sheet
                        melted['Bulan_Laporan'] = filename
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
                
            # --- B. DW(Dedicated) atau Tabel Harian (1-31) ---
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
                
                if value_vars:
                    melted = df_data.melt(id_vars=id_vars, value_vars=value_vars, var_name='Tanggal', value_name='Jumlah_MP')
                    melted['Sumber_Sheet'] = sheet
                    melted['Bulan_Laporan'] = filename
                    melted['Jumlah_MP'] = pd.to_numeric(melted['Jumlah_MP'], errors='coerce').fillna(0)
                    melted = melted[melted['Jumlah_MP'] > 0]
                    
                    if id_vars:
                        kolom_entitas = id_vars[0]
                        for col in id_vars:
                            if isinstance(col, str) and any(kw in col.lower() for kw in ['company', 'entity', 'vendor', 'nama', 'name', 'job', 'project']):
                                kolom_entitas = col
                                break
                        final_df = melted[['Bulan_Laporan', 'Sumber_Sheet', kolom_entitas, 'Tanggal', 'Jumlah_MP']].copy()
                        final_df.columns = ['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah']
                        final_df['Kategori'] = sheet.strip()
                        all_data.append(final_df)
            else:
                # --- C. PARSING UNTUK SHEET FIX, OS, & ATTENDANCE ---
                if clean_sheet_name.lower() == 'fix':
                    for r in range(6, len(df)):
                        row_vals = df.iloc[r].tolist()
                        if len(row_vals) > 2 and pd.notna(row_vals[1]) and pd.notna(row_vals[2]):
                            comp = str(row_vals[1])
                            qty = pd.to_numeric(row_vals[2], errors='coerce')
                            if qty and qty > 0:
                                all_data.append(pd.DataFrame({
                                    'File_Sumber': [filename],
                                    'Kategori': ['Fix (Cost)'],
                                    'Entitas_Posisi': [comp],
                                    'Tanggal': [1],
                                    'Jumlah': [qty]
                                }))
                elif clean_sheet_name.lower() == 'os':
                    sub_df = df.iloc[5:].dropna(subset=[2]) if len(df) > 5 else pd.DataFrame()
                    if not sub_df.empty and 2 in sub_df.columns:
                        counts = sub_df[2].value_counts().reset_index()
                        counts.columns = ['Vendor', 'Count']
                        for _, row in counts.iterrows():
                            all_data.append(pd.DataFrame({
                                'File_Sumber': [filename],
                                'Kategori': ['OS (Outsourcing)'],
                                'Entitas_Posisi': [str(row['Vendor'])],
                                'Tanggal': [1],
                                'Jumlah': [row['Count']]
                            }))
                elif clean_sheet_name.lower() == 'attendance':
                    sub_df = df.iloc[17:].dropna(subset=[1]) if len(df) > 17 else pd.DataFrame()
                    if not sub_df.empty and 1 in sub_df.columns:
                        counts = sub_df[1].value_counts().reset_index()
                        counts.columns = ['Entity', 'Count']
                        for _, row in counts.iterrows():
                            all_data.append(pd.DataFrame({
                                'File_Sumber': [filename],
                                'Kategori': ['Attendance'],
                                'Entitas_Posisi': [str(row['Entity'])],
                                'Tanggal': [1],
                                'Jumlah': [row['Count']]
                            }))
                            
        except Exception as e:
            st.warning(f"Catatan pada sheet '{sheet}': {e}")
                
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()

# ==========================================
# 3. ANTARMUKA STREAMLIT (UI)
# ==========================================
st.title("📊 JDC Manpower Dashboard Report")

tab1, tab2 = st.tabs(["📈 Dashboard Utama", "⚙️ Upload & Sinkronisasi Data"])

# --- TAB 2: UPLOAD DATA ---
with tab2:
    st.subheader("Upload Laporan Harian (Excel)")
    st.markdown("Sistem dikonfigurasi khusus untuk memproses 5 worksheet: **Attendance, OS, Fix, DW(Dedicated), DW(ARU)**.")
    
    uploaded_file = st.file_uploader("Pilih file .xlsx", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Memproses file Excel terpilih..."):
            extracted_data = process_excel(uploaded_file)
            
            if not extracted_data.empty:
                st.success(f"Berhasil mengekstrak {len(extracted_data)} baris data dari 5 sheet utama!")
                st.dataframe(extracted_data.head(10))
                
                if st.button("💾 Simpan ke Google Sheets"):
                    with st.spinner("Menyinkronkan dengan Google Sheets..."):
                        try:
                            client = get_gsheets_client()
                            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
                            
                            existing_data = sheet.get_all_values()
                            if not existing_data:
                                sheet.append_row(['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah'])
                                
                            extracted_data = extracted_data.fillna("")
                            data_to_upload = extracted_data.astype(str).values.tolist()
                            sheet.append_rows(data_to_upload)
                            
                            st.success("Data berhasil ditambahkan dan diamankan ke Google Sheets! ✅")
                        except Exception as e:
                            st.error(f"Gagal menyimpan ke Google Sheets: {e}")
            else:
                st.error("Tidak ditemukan data valid dari sheet yang ditentukan.")

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
            st.warning("Gagal terhubung ke Google Sheets atau file masih kosong.")
        
    # --- SAFETY CHECK AGAR TIDAK ERROR KETIKA DATA KOSONG ---
    if not df_db.empty and 'Jumlah' in df_db.columns:
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
            m1.metric("Total Kumulatif", f"{int(df_filtered['Jumlah'].sum()):,}")
            m2.metric("Jumlah Kategori Aktif", len(df_filtered['Kategori'].unique()))
            m3.metric("Total Entitas / Posisi", len(df_filtered['Entitas_Posisi'].unique()))
            
            # --- Visualisasi ---
            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("### Tren / Distribusi Berdasarkan Kategori")
                tren_harian = df_filtered.groupby(['Tanggal', 'Kategori'])['Jumlah'].sum().reset_index()
                tren_harian = tren_harian.sort_values(by='Tanggal')
                
                fig_line = px.line(
                    tren_harian, 
                    x='Tanggal', 
                    y='Jumlah', 
                    color='Kategori', 
                    markers=True,
                    labels={'Jumlah': 'Total', 'Tanggal': 'Tanggal'}
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
        st.info("⚠️ Belum ada data di Google Sheets atau kolom belum terbentuk. Silakan pergi ke tab **'Upload & Sinkronisasi Data'**, lalu upload file Excel Anda dan klik tombol **'Simpan ke Google Sheets'**.")

# ==========================================
# 4. WATERMARK FOOTER
# ==========================================
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: grey; font-size: 14px;'>Developed by iqbalmantam</p>", 
    unsafe_allow_html=True
)
