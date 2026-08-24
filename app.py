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
# 2. ALGORITMA EKSTRAKSI ROBUST (5 SHEET UTAMA)
# ==========================================
def process_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    target_sheets = ['attendance', 'os', 'fix', 'dw(dedicated)', 'dw(aru)']
    
    all_data = []
    filename = getattr(file, 'name', 'Laporan Bulanan').split('.')[0]
    
    for sheet in sheets:
        clean_name = sheet.strip().lower().replace(" ", "")
        matched = any(clean_name == ts.replace(" ", "") for ts in target_sheets)
        if not matched:
            continue
            
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            
            # --- 1. SHEET DW(ARU) ---
            if 'aru' in clean_name and 'dw' in clean_name:
                if len(df) > 6:
                    row_vals = df.iloc[6].tolist()
                    for idx, val in enumerate(row_vals):
                        num = pd.to_numeric(val, errors='coerce')
                        if num and num > 0 and idx >= 7:
                            all_data.append(pd.DataFrame({
                                'File_Sumber': [filename],
                                'Kategori': ['DW(ARU)'],
                                'Entitas_Posisi': [str(df.iloc[6, 3]) if pd.notna(df.iloc[6, 3]) else 'ARU'],
                                'Tanggal': [idx - 6],
                                'Jumlah': [num]
                            }))
                continue

            # --- 2. SHEET DW(Dedicated) ---
            if 'dw(dedicated)' in clean_name:
                date_row_idx = None
                for i in range(min(10, len(df))):
                    row_vals = pd.to_numeric(df.iloc[i], errors='coerce')
                    if len(row_vals[(row_vals >= 1) & (row_vals <= 31)]) >= 10:
                        date_row_idx = i
                        break
                if date_row_idx is not None:
                    for r in range(date_row_idx + 1, len(df)):
                        row = df.iloc[r]
                        project = row.iloc[3] if len(row) > 3 and pd.notna(row.iloc[3]) else 'Common'
                        job_pos = row.iloc[5] if len(row) > 5 and pd.notna(row.iloc[5]) else 'Worker'
                        entitas = f"{project} - {job_pos}"
                        
                        for c in range(7, len(row)):
                            day_val = df.iloc[date_row_idx, c]
                            if isinstance(day_val, (int, float)) and 1 <= day_val <= 31:
                                val_mp = pd.to_numeric(row.iloc[c], errors='coerce')
                                if val_mp and val_mp > 0:
                                    all_data.append(pd.DataFrame({
                                        'File_Sumber': [filename],
                                        'Kategori': ['DW(Dedicated)'],
                                        'Entitas_Posisi': [str(entitas)],
                                        'Tanggal': [int(day_val)],
                                        'Jumlah': [val_mp]
                                    }))
                continue

            # --- 3. SHEET FIX (Cost & Summary) ---
            if 'fix' in clean_name:
                for r in range(6, len(df)):
                    row_vals = df.iloc[r].tolist()
                    if len(row_vals) > 4 and pd.notna(row_vals[1]) and pd.notna(row_vals[4]):
                        comp = str(row_vals[1]).strip()
                        amount = pd.to_numeric(row_vals[4], errors='coerce')
                        if amount and amount > 0 and comp.lower() != 'nan':
                            all_data.append(pd.DataFrame({
                                'File_Sumber': [filename],
                                'Kategori': ['Fix (Cost Summary)'],
                                'Entitas_Posisi': [comp],
                                'Tanggal': [1],
                                'Jumlah': [amount]
                            }))
                continue

            # --- 4. SHEET OS (Outsourcing Roster) ---
            if 'os' in clean_name:
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
                continue

            # --- 5. SHEET Attendance ---
            if 'attendance' in clean_name:
                sub_df = df.iloc[17:].dropna(subset=[1]) if len(df) > 17 else pd.DataFrame()
                if not sub_df.empty and 1 in sub_df.columns:
                    counts = sub_df[1].value_counts().reset_index()
                    counts.columns = ['Entity', 'Count']
                    for _, row in counts.iterrows():
                        all_data.append(pd.DataFrame({
                            'File_Sumber': [filename],
                            'Kategori': ['Attendance Summary'],
                            'Entitas_Posisi': [str(row['Entity'])],
                            'Tanggal': [1],
                            'Jumlah': [row['Count']]
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
    st.markdown("Sistem dikonfigurasi khusus untuk memproses 5 worksheet: **Attendance, OS, Fix, DW(Dedicated), DW(ARU)** dengan proteksi anti-duplikasi data.")
    
    uploaded_file = st.file_uploader("Pilih file .xlsx", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Memproses file Excel secara mendalam..."):
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
                                # Jika sheet kosong, buat header dulu
                                sheet.append_row(['File_Sumber', 'Kategori', 'Entitas_Posisi', 'Tanggal', 'Jumlah'])
                                existing_data = sheet.get_all_values()
                                
                            # --- FITUR ANTI-DUPLIKASI ---
                            # Periksa apakah File_Sumber yang sama sudah ada di GSheet
                            if len(existing_data) > 1:
                                header = existing_data[0]
                                rows = existing_data[1:]
                                df_existing = pd.DataFrame(rows, columns=header)
                                
                                if 'File_Sumber' in df_existing.columns:
                                    # Filter keluar baris yang memiliki File_Sumber sama dengan yang di-upload
                                    filtered_rows = df_existing[df_existing['File_Sumber'] != filename_upload]
                                    
                                    # Tulis ulang sheet dengan data bersih + data baru
                                    sheet.clear()
                                    sheet.append_row(header)
                                    if not filtered_rows.empty:
                                        sheet.append_rows(filtered_rows.astype(str).values.tolist())
                                        
                            # Tambahkan data baru yang di-upload
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
        pilihan_kategori = col1.multiselect("Filter Berdasarkan Kategori / Sheet", kategori_list, default=kategori_list)
        
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
            m1.metric("Total Kumulatif", f"{int(df_filtered['Jumlah'].sum()):,}")
            m2.metric("Jumlah Kategori Aktif", len(df_filtered['Kategori'].unique()))
            m3.metric("Total Entitas / Posisi", len(df_filtered['Entitas_Posisi'].unique()))
            
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
