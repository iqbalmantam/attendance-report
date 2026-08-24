import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# ==========================================
# 0. KONFIGURASI HALAMAN & WATERMARK
# ==========================================
st.set_page_config(page_title="Dashboard JDC Manpower", page_icon="📊", layout="wide")

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
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1lv0b68kBixRgYmJJk1J6Eppj7wXu5u6C8g81zzQiKg0/edit?gid=0#gid=0"

# ==========================================
# 2. ALGORITMA EKSTRAKSI PRESISI TINGGI
# ==========================================
def process_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    
    all_data = []
    filename = getattr(file, 'name', 'Laporan Bulanan').split('.')[0]
    
    for sheet in sheets:
        clean_name = sheet.strip().lower().replace(" ", "")
        
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # --- 1. ATTENDANCE (Total Days, Sakit/SCW, Cuti/AL) ---
            if 'attendance' in clean_name:
                row13 = df.iloc[13].tolist()
                total_days_col, al_col, scw_col = None, None, None
                
                for idx, val in enumerate(row13):
                    if pd.notna(val):
                        sval = str(val).strip().upper()
                        if 'TOTAL DAYS' in sval:
                            total_days_col = idx
                        elif sval == 'AL':
                            al_col = idx
                        elif sval in ['SCW', 'SAKIT']:
                            scw_col = idx
                            
                # Fallback jika header tidak terbaca persis
                if total_days_col is None:
                    total_days_col = 115 if df.shape[1] <= 117 else 118
                if al_col is None:
                    al_col = 104 if df.shape[1] <= 117 else 107
                if scw_col is None:
                    scw_col = 108 if df.shape[1] <= 117 else 111
                    
                sub_att = df.iloc[17:].dropna(subset=[1]) # Col 1 adalah Entitas/Entity
                for _, row in sub_att.iterrows():
                    entity = str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'Unknown'
                    
                    t_days = pd.to_numeric(row.iloc[total_days_col], errors='coerce') if total_days_col < len(row) else 0
                    al_val = pd.to_numeric(row.iloc[al_col], errors='coerce') if al_col < len(row) else 0
                    scw_val = pd.to_numeric(row.iloc[scw_col], errors='coerce') if scw_col < len(row) else 0
                    
                    if t_days and t_days > 0:
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['Attendance (Total Days)'],
                            'Entitas_Posisi': [entity],
                            'Tanggal': [1],
                            'Jumlah': [t_days]
                        }))
                    if al_val and al_val > 0:
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['Attendance (Cuti/AL)'],
                            'Entitas_Posisi': [entity],
                            'Tanggal': [1],
                            'Jumlah': [al_val]
                        }))
                    if scw_val and scw_val > 0:
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['Attendance (Sakit/SCW)'],
                            'Entitas_Posisi': [entity],
                            'Tanggal': [1],
                            'Jumlah': [scw_val]
                        }))
                continue

            # --- 2. OS (Total Cost dari Kolom AH / indeks 33) ---
            if clean_name == 'os':
                ah_idx = 33 # Kolom AH
                sub_os = df.iloc[7:].dropna(subset=[4]) # Col 4 adalah Nama Karyawan
                for _, row in sub_os.iterrows():
                    name = str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'Employee'
                    cost = pd.to_numeric(row.iloc[ah_idx], errors='coerce') if ah_idx < len(row) else 0
                    if cost and cost > 0:
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['OS (Total Cost)'],
                            'Entitas_Posisi': [name],
                            'Tanggal': [1],
                            'Jumlah': [cost]
                        }))
                continue

            # --- 3. FIX (Total Cost dari Cell E5 -> row 4, col 4) ---
            if clean_name == 'fix':
                if df.shape[0] > 4 and df.shape[1] > 4:
                    fix_val = pd.to_numeric(df.iloc[4, 4], errors='coerce')
                    if fix_val and fix_val > 0:
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['Fix (Total Cost E5)'],
                            'Entitas_Posisi': ['Fix Summary'],
                            'Tanggal': [1],
                            'Jumlah': [fix_val]
                        }))
                continue

            # --- 4. DW(Dedicated) & DW(ARU) (Total Cost Million IDR) ---
            if 'dw' in clean_name:
                for r in range(len(df)):
                    val = df.iloc[r, 3] if df.shape[1] > 3 else None
                    if pd.notna(val) and 'total cost' in str(val).lower():
                        total_cost = pd.to_numeric(df.iloc[r, 6], errors='coerce')
                        if total_cost and total_cost > 0:
                            all_data.append(pd.DataFrame({
                                'File_Sumber': [filename],
                                'Kategori': [f"{sheet.strip()} (Total Cost)"],
                                'Entitas_Posisi': ['Total Cost Summary'],
                                'Tanggal': [1],
                                'Jumlah': [total_cost]
                            }))
                continue
                            
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

with tab2:
    st.subheader("Upload Laporan Harian (Excel)")
    st.markdown("Sistem dikonfigurasi membaca metrik akurat dari **Attendance, OS (Col AH), Fix (Cell E5), DW(Dedicated), DW(ARU)** dengan proteksi anti-duplikasi.")
    
    uploaded_file = st.file_uploader("Pilih file .xlsx", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Memproses file Excel secara presisi..."):
            extracted_data = process_excel(uploaded_file)
            
            if not extracted_data.empty:
                filename_upload = extracted_data['File_Sumber'].iloc[0]
                st.success(f"Berhasil mengekstrak {len(extracted_data)} baris data untuk file: **{filename_upload}**")
                st.dataframe(extracted_data.head(10))
                
                if st.button("💾 Simpan ke Google Sheets"):
                    with st.spinner("Menyinkronkan dengan Google Sheets..."):
                        try:
                            client = get_gsheets_client()
                            sheet = client.open_by_url(SPREADSHEET_URL).sheet1
                            
                            existing_data = sheet.get_all_values()
                            
                            if not existing_data:
                                sheet.append_row(['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah'])
                                existing_data = sheet.get_all_values()
                                
                            # Proteksi Anti-Duplikasi Berdasarkan Nama File
                            if len(existing_data) > 1:
                                header = existing_data[0]
                                rows = existing_data[1:]
                                df_existing = pd.DataFrame(rows, columns=header)
                                
                                if 'File_Sumber' in df_existing.columns:
                                    filtered_rows = df_existing[df_existing['File_Sumber'] != filename_upload]
                                    sheet.clear()
                                    sheet.append_row(header)
                                    if not filtered_rows.empty:
                                        sheet.append_rows(filtered_rows.astype(str).values.tolist())
                                        
                            extracted_data = extracted_data.fillna("")
                            data_to_upload = extracted_data.astype(str).values.tolist()
                            sheet.append_rows(data_to_upload)
                            
                            st.success(f"Data untuk '{filename_upload}' berhasil disimpan (Pembaruan bersih tanpa duplikat)! ✅")
                        except Exception as e:
                            st.error(f"Gagal menyimpan ke Google Sheets: {e}")
            else:
                st.error("Tidak ditemukan data valid dari sheet yang ditentukan.")

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
        
    if not df_db.empty and 'Jumlah' in df_db.columns:
        df_db['Jumlah'] = pd.to_numeric(df_db['Jumlah'], errors='coerce').fillna(0)
        df_db['Tanggal'] = pd.to_numeric(df_db['Tanggal'], errors='coerce')
        
        col1, col2 = st.columns(2)
        kategori_list = df_db['Kategori'].unique().tolist()
        pilihan_kategori = col1.multiselect("Filter Berdasarkan Kategori", kategori_list, default=kategori_list)
        
        file_list = df_db['File_Sumber'].unique().tolist() if 'File_Sumber' in df_db.columns else []
        pilihan_file = col2.multiselect("Filter Berdasarkan Bulan / File", file_list, default=file_list)
        
        df_filtered = df_db[df_db['Kategori'].isin(pilihan_kategori)]
        if pilihan_file and 'File_Sumber' in df_db.columns:
            df_filtered = df_filtered[df_filtered['File_Sumber'].isin(pilihan_file)]
        
        if df_filtered.empty:
            st.info("Pilih setidaknya satu kategori/file untuk menampilkan grafik.")
        else:
            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Kumulatif Metrik", f"{int(df_filtered['Jumlah'].sum()):,}")
            m2.metric("Jumlah Kategori Aktif", len(df_filtered['Kategori'].unique()))
            m3.metric("Total Entitas / Posisi", len(df_filtered['Entitas_Posisi'].unique()))
            
            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("### Perbandingan Berdasarkan Kategori")
                kategori_summary = df_filtered.groupby('Kategori')['Jumlah'].sum().reset_index()
                
                fig_bar = px.bar(
                    kategori_summary, 
                    x='Kategori', 
                    y='Jumlah',
                    color='Kategori',
                    labels={'Jumlah': 'Total Nilai', 'Kategori': 'Kategori Laporan'}
                )
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with chart_col2:
                st.markdown("### Top Entitas / Karyawan")
                komposisi = df_filtered.groupby('Entitas_Posisi')['Jumlah'].sum().reset_index()
                komposisi = komposisi.sort_values(by='Jumlah', ascending=False).head(10)
                
                fig_bar2 = px.bar(
                    komposisi, 
                    x='Jumlah', 
                    y='Entitas_Posisi', 
                    orientation='h',
                    color='Jumlah',
                    color_continuous_scale='Blues',
                    labels={'Jumlah': 'Total', 'Entitas_Posisi': 'Entitas/Nama'}
                )
                fig_bar2.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_bar2, use_container_width=True)
                
            with st.expander("Lihat Data Mentah (Tersimpan di GSheet)"):
                st.dataframe(df_db, use_container_width=True)
    else:
        st.info("⚠️ Belum ada data di Google Sheets atau kolom belum terbentuk. Silakan ke tab **'Upload & Sinkronisasi Data'**, upload file Excel Anda, klik **'Simpan ke Google Sheets'**, lalu kembali ke tab ini.")

# ==========================================
# 4. WATERMARK FOOTER
# ==========================================
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: grey; font-size: 14px;'>Developed by iqbalmantam</p>", 
    unsafe_allow_html=True
)
