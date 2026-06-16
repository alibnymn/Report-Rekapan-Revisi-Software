import streamlit as st  # type: ignore[import]
import pandas as pd  # type: ignore[import]
import plotly.express as px  # type: ignore[import]
import datetime
from datetime import datetime, date, timedelta

def format_ke_str(val):
    """Konversi objek tanggal (date/datetime/Timestamp) ke format string string 'YYYY-MM-DD'.
    Jika datanya None atau kosong, akan dikembalikan sebagai string '-' atau None."""
    if pd.isna(val) or val is None:
        return "-"  # Atau bisa diganti None jika ingin sel Excel kosong
    
    # Jika tipenya sudah date atau datetime
    if isinstance(val, (date, datetime, pd.Timestamp)):
        return val.strftime('%Y-%m-%d') # Hasil: 2026-05-25
    
    # Jika tipenya string, kita bersihkan spasinya
    return str(val).strip()
# Konfigurasi Halaman
st.set_page_config(page_title="Software Revision Report App", layout="wide")

EXCEL_FILE = "Rekap Pembuatan Revisi Software.xlsx"
SHEET_NAME = "Rekap Pembuatan Revisi Software"

# Fungsi untuk membaca data dari Excel lengkap dengan mempertahankan indeks asli
@st.cache_data(ttl=2)
def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        
        # JAMINAN BERSIH: Ubah semua nama kolom menjadi string, hilangkan spasi di awal/akhir, dan jadikan UPPERCASE
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Hilangkan kolom hantu underscore jika ada
        df = df.drop(columns=['TGL_MULAI_QA', 'TGL_SELESAI_QA'], errors='ignore')
        
        # Reset index untuk pencarian baris asli Excel
        df = df.reset_index(names='ORIGINAL_INDEX')
        
        # Amankan format tanggal (nama kolom sudah kapital semua)
        date_cols = ['TANGGAL', 'TGL PENGERJAAN', 'DEADLINE', 'TGL SELESAI', 'TGL MULAI QA', 'TGL SELESAI QA', 'TGL VERIFIKASI']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
        return df
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return pd.DataFrame()

df = load_data()

# Definisikan list opsi dropdown
list_aplikasi = ["PRODUKSI", "QC", "KARYAWAN", "LOTUS EMPLOYEE","GUDANG", "RND", "INVENTORY", "HRCM", "PAYROLL", "MARKETING", "OPERATOR", "NOTIFICATION"]
list_developer = ["RANDY", "WILIAM", "ILMAN", "JONATHAN", "CHASTRO", "AGIS", "EDI"] 
list_jenis = ["REVISI MINOR", "REVISI MAJOR", "PEMBUATAN SOFTWARE"]
list_status = ["OK", "PENDING", "REJECT"]

# Filter otomatis di awal agar data developer lama tidak mengotori dashboard
if not df.empty:
    df_active = df[df['DEVELOPER'].isin(list_developer)].copy()
else:
    df_active = df.copy()

# Menu Navigasi
menu = st.sidebar.selectbox("Menu Utama", ["Dashboard Report & Action", "Input Data Baru", "Hapus Banyak Data"])

# ==================== MENU 1: DASHBOARD REPORT & ACTION ====================
if menu == "Dashboard Report & Action":
    st.title("📊 Software Revision Dashboard & Management")

    if not df_active.empty:
        # --- LOGIKA HITUNG KETERLAMBATAN DEVELOPER ---
        df_active['DEV_DELAY'] = df_active.apply(
            lambda r: r['TGL SELESAI'] > r['DEADLINE'] if pd.notna(r['TGL SELESAI']) and pd.notna(r['DEADLINE']) else False, 
            axis=1
        )
        
        # --- Bagian 1: Summary Cards ---
        total_active_tasks = len(df_active)
        dev_delayed_tasks = len(df_active[df_active['DEV_DELAY'] == True])
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Request (Aktif)", f"{total_active_tasks} Tugas")
        col2.metric("Status OK", f"{len(df_active[df_active['STATUS'] == 'OK'])} Selesai")
        col3.metric("Developer Delay", f"{dev_delayed_tasks} Tugas", delta=f"{dev_delayed_tasks} Overdue", delta_color="inverse")
        st.divider()

        # --- Bagian 2: Grafik Distribusi Aplikasi & Ketepatan Waktu Dev ---
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("Aplikasi Paling Sering Direvisi")
            app_counts = df_active['APLIKASI'].value_counts().reset_index()
            app_counts.columns = ['Aplikasi', 'Jumlah']
            fig_app = px.pie(app_counts, values='Jumlah', names='Aplikasi', hole=0.3)
            st.plotly_chart(fig_app, use_container_width=True, key="chart_pie_aplikasi_dashboard")
            
        with col_chart2:
            st.subheader("Ketepatan Waktu Penyelesaian Per-Developer")
            df_active['WAKTU_PENGERJAAN'] = df_active['DEV_DELAY'].map({True: 'Terlambat (Delay)', False: 'Tepat Waktu'})
            dev_time_counts = df_active.groupby(['DEVELOPER', 'WAKTU_PENGERJAAN']).size().reset_index(name='Jumlah Task')
            
            fig_dev = px.bar(dev_time_counts, 
                             x='Jumlah Task', 
                             y='DEVELOPER', 
                             color='WAKTU_PENGERJAAN',
                             orientation='h',
                             barmode='stack',
                             color_discrete_map={'Terlambat (Delay)': '#EF553B', 'Tepat Waktu': '#636EFA'},
                             text_auto=True)
            
            fig_dev.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_dev, use_container_width=True, key="chart_bar_developer_perf_dashboard")

        st.divider()

        # ---- REKAP QA / PEMERIKSA (ALI) ----
        st.subheader("📋 Analisis Performa Pemeriksaan QA")
        qa_ali_df = df_active[df_active['PEMERIKSA'].astype(str).str.lower() == 'ali'].copy()
        
        if not qa_ali_df.empty:
            qa_ali_df['STATUS_REKAP_QA'] = qa_ali_df['STATUS'].apply(lambda x: 'Selesai Diperiksa (OK)' if x == 'OK' else 'Dalam Proses / Pending')

            col_chart_qa, col_table_qa = st.columns([2, 1])
            
            with col_chart_qa:
                st.markdown("**Beban Kerja Pemeriksaan QA(Ali)**")
                qa_counts = qa_ali_df.groupby(['APLIKASI', 'STATUS_REKAP_QA']).size().reset_index(name='Jumlah Task')
                
                fig_qa = px.bar(qa_counts, 
                                x='Jumlah Task', 
                                y='APLIKASI', 
                                color='STATUS_REKAP_QA',
                                orientation='h', 
                                color_discrete_map={'Selesai Diperiksa (OK)': '#636EFA', 'Dalam Proses / Pending': '#FECB52'},
                                text_auto=True)
                
                fig_qa.update_layout(yaxis={'categoryorder':'total ascending'}, legend_title_text='Status QA')
                st.plotly_chart(fig_qa, use_container_width=True, key="chart_bar_qa_ali_dashboard")
                
            with col_table_qa:
                st.markdown("**Ringkasan Hasil Pemeriksaan**")
                qa_ali_df['STATUS_TABEL'] = qa_ali_df['STATUS'].replace({'TERLAMBAT': 'Dev Overdue / Delay'})
                
                qa_summary = qa_ali_df.pivot_table(
                    index='APLIKASI', 
                    columns='STATUS_TABEL', 
                    values='KETERANGAN', 
                    aggfunc='count', 
                    fill_value=0
                ).reset_index()
                
                qa_summary['Total Diperiksa'] = qa_summary.drop(columns=['APLIKASI'], errors='ignore').sum(axis=1)
                st.dataframe(qa_summary.sort_values(by='Total Diperiksa', ascending=False), use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data pemeriksaan dengan nama 'Ali' pada filter developer aktif saat ini.")

        st.divider()
# ==================== FITUR FILTER DATA COMPLETE (DI ATAS TABEL) ====================
        st.markdown("### 🔍 Filter Pencarian Data")
        
        # Membuat 4 kolom menjajar ke samping agar layout tetap rapi dan hemat ruang
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        with col_f1:
            # Filter Developer (Master List)
            list_dev_master = ["RANDY", "WILIAM", "ILMAN", "JONATHAN", "CHASTRO", "AGIS", "EDI"]
            pilihan_dev = st.multiselect("Filter Developer:", options=list_dev_master, placeholder="Semua Developer", key="filter_multiselect_dev")

        with col_f2:
            # Filter Aplikasi (Master List)
            list_app_master = ["PRODUKSI", "QC", "KARYAWAN", "LOTUS EMPLOYEE", "GUDANG", "RND", "INVENTORY", "HRCM", "PAYROLL", "MARKETING", "OPERATOR", "NOTIFICATION"]
            pilihan_app = st.multiselect("Filter Aplikasi:", options=list_app_master, placeholder="Semua Aplikasi", key="filter_multiselect_app")

        with col_f3:
            # Filter Jenis Revisi / Software (Master List)
            list_jenis_master = ["REVISI MINOR", "REVISI MAJOR", "PEMBUATAN SOFTWARE"]
            pilihan_jenis = st.multiselect("Filter Jenis:", options=list_jenis_master, placeholder="Semua Jenis", key="filter_multiselect_jenis")

        with col_f4:
            # Filter Keterangan (Pencarian Teks / Search Bar)
            cari_keterangan = st.text_input("Cari kata kunci Keterangan:", placeholder="Ketik kata kunci...", key="filter_text_keterangan")

        # --- LOGIKA EKSEKUSI FILTER ---
        df_hasil_filter = df_active.copy()

        # 1. Jalankan Filter Developer (Exact Match)
        if pilihan_dev:
            df_hasil_filter = df_hasil_filter[
                df_hasil_filter['DEVELOPER'].fillna('').astype(str).str.upper().str.strip().isin(pilihan_dev)
            ]
            
        # 2. Jalankan Filter Aplikasi (Exact Match)
        if pilihan_app:
            df_hasil_filter = df_hasil_filter[
                df_hasil_filter['APLIKASI'].fillna('').astype(str).str.upper().str.strip().isin(pilihan_app)
            ]
            
        # 3. Jalankan Filter Jenis Revisi (Exact Match)
        if pilihan_jenis:
            df_hasil_filter = df_hasil_filter[
                df_hasil_filter['JENIS'].fillna('').astype(str).str.upper().str.strip().isin(pilihan_jenis)
            ]
            
        # 4. Jalankan Filter Keterangan (Pencarian Kata Kunci)
        if cari_keterangan:
            df_hasil_filter = df_hasil_filter[
                df_hasil_filter['KETERANGAN'].fillna('').astype(str).str.lower().str.contains(cari_keterangan.lower(), na=False)
            ]
            
        st.info(f"💡 Menampilkan **{len(df_hasil_filter)}** tugas dari total data rekap saat ini.")
        st.divider()
        # ===========================================================================

        # ---- Data Rekap Revisi Software (Tabel Utama) ----
        st.subheader("Data Rekap Revisi Software")
        
        # Menggunakan hasil filter (df_hasil_filter) untuk dibuang kolom internalnya sebelum tampil
        clean_display_df = df_hasil_filter.drop(columns=['ORIGINAL_INDEX', 'DEV_DELAY', 'STATUS_REKAP_QA', 'STATUS_TABEL', 'WAKTU_PENGERJAAN'], errors='ignore')
        
        # Tampilkan tabel utama tanpa nomor index bawaan pandas
        st.dataframe(clean_display_df, use_container_width=True, hide_index=True)
        
        # ---- Tombol Download Data Yang Sudah Terfilter ----
        col_dl1, col_dl2, _ = st.columns([1, 1, 2])
        with col_dl1:
            csv_data = clean_display_df.to_csv(index=False).encode('utf-8')
            st.download_button(label="📥 Download Data (.CSV)", data=csv_data, file_name="Rekap_Revisi_Software_Filtered.csv", mime="text/csv", key="btn_download_csv")
        with col_dl2:
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                clean_display_df.to_excel(writer, index=False, sheet_name='Filtered Data')
            excel_data = buffer.getvalue()
            st.download_button(label="🟢 Download Data (.XLSX)", data=excel_data, file_name="Rekap_Revisi_Software_Filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_download_excel")

        # --- Bagian 3: Panel Aksi Edit / Delete ---
        st.divider()
        st.subheader("🛠️ Action Panel (Edit / Delete)")
        
        filtered_df = df_active.copy()
        filtered_df['SELECT_LABEL'] = filtered_df.apply(lambda row: f"Row {row['ORIGINAL_INDEX']+2} | [{row['APLIKASI']}] {str(row['KETERANGAN'])[:50]}...", axis=1)
        selected_task_label = st.selectbox("Pilih Data yang Akan Dikelola", options=["-- Pilih Data --"] + filtered_df['SELECT_LABEL'].tolist())
        
        if selected_task_label != "-- Pilih Data --":
            target_idx = filtered_df[filtered_df['SELECT_LABEL'] == selected_task_label]['ORIGINAL_INDEX'].values[0]
            
            # Baca Excel real-time murni untuk memuat isi form edit
            fresh_df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            fresh_df.columns = [str(c).strip() for c in fresh_df.columns]
            
            raw_row = fresh_df.loc[target_idx]
            row_data = {str(k).strip().upper(): v for k, v in raw_row.items()}
            
            action_mode = st.radio("Pilih Tindakan", ["Edit Data", "Hapus Data (Delete)"], horizontal=True)
            
            def parse_ke_date(val_kolom):
                if pd.isna(val_kolom) or str(val_kolom).strip() in ["", "-", "–", "—", "NaT", "nan"]:
                    return datetime.now().date()
                try:
                    if hasattr(val_kolom, 'date'):
                        return val_kolom.date()
                    val_str = str(val_kolom).strip().replace('–', '-').replace('—', '-')
                    mapping_bulan = {
                        "Jan": "Jan", "Feb": "Feb", "Mar": "Mar", "Apr": "Apr",
                        "Mei": "May", "Jun": "Jun", "Jul": "Jul", "Agu": "Aug",
                        "Sep": "Sep", "Okt": "Oct", "Nov": "Nov", "Des": "Dec"
                    }
                    parts = val_str.split('-')
                    if len(parts) == 3:
                        tgl = parts[0].strip()
                        bln = parts[1].strip()
                        thn = parts[2].strip()
                        if " " in thn:
                            thn = thn.split(" ")[0].strip()
                        if bln in mapping_bulan:
                            val_str = f"{int(tgl):02d}-{mapping_bulan[bln]}-{thn}"
                            fmt_tahun = "%y" if len(thn) == 2 else "%Y"
                            return datetime.strptime(val_str, f"%d-%b-{fmt_tahun}").date()
                    return pd.to_datetime(val_kolom).date()
                except:
                    return datetime.now().date()

            # --- PROSES EDIT DATA ---
            if action_mode == "Edit Data":
                with st.form("edit_form"):
                    st.warning(f"Anda sedang mengubah data pada baris Excel ke-{target_idx + 2}")
                    col_e1, col_e2 = st.columns(2)
                    
                    def clean_text_input(val):
                        if pd.isna(val) or str(val).strip().lower() in ["nan", "-", ""]:
                            return ""
                        return str(val).strip()
                    
                    with col_e1:
                        val_tanggal = st.date_input("Tanggal Request", parse_ke_date(row_data.get('TANGGAL')))
                        
                        val_app_curr = clean_text_input(row_data.get('APLIKASI')).upper()
                        idx_app = list_aplikasi.index(val_app_curr) if val_app_curr in list_aplikasi else 0
                        val_aplikasi = st.selectbox("Aplikasi", list_aplikasi, index=idx_app)
                        
                        val_dev_curr = clean_text_input(row_data.get('DEVELOPER')).upper()
                        idx_dev = list_developer.index(val_dev_curr) if val_dev_curr in list_developer else 0
                        val_developer = st.selectbox("Developer", list_developer, index=idx_dev)
                        
                        val_jen_curr = clean_text_input(row_data.get('JENIS'))
                        idx_jen = list_jenis.index(val_jen_curr) if val_jen_curr in list_jenis else 0
                        val_jenis = st.selectbox("Jenis Pembuatan", list_jenis, index=idx_jen)
                        
                        val_keterangan = st.text_area("Keterangan / Detail Revisi", value=clean_text_input(row_data.get('KETERANGAN')))
                        val_pembuat = st.text_input("Pembuat (User Request)", value=clean_text_input(row_data.get('PEMBUAT')))
                        
                    with col_e2:
                        val_tgl_pengerjaan = st.date_input("Tanggal Mulai Pengerjaan", parse_ke_date(row_data.get('TGL PENGERJAAN')))
                        val_deadline = st.date_input("Deadline Target", parse_ke_date(row_data.get('DEADLINE')))
                        val_tgl_selesai = st.date_input("Tanggal Selesai Pengerjaan", parse_ke_date(row_data.get('TGL SELESAI')))
                        
                        val_pemeriksa = st.selectbox("Pemeriksa (QA / Lead)", ["Ali"], index=0)
                        
                        old_start_qa = parse_ke_date(row_data.get('TGL MULAI QA')) 
                        old_end_qa = parse_ke_date(row_data.get('TGL SELESAI QA'))   
                        val_qa_range = st.date_input("Rentang Waktu Pemeriksaan QA (Mulai - Selesai)", value=(old_start_qa, old_end_qa))
                        
                        val_tgl_verifikasi = st.date_input("Tanggal Verifikasi", parse_ke_date(row_data.get('TGL VERIFIKASI')))
                        val_verifikator = st.text_input("Verifikator", value=clean_text_input(row_data.get('VERIFIKATOR')))
                        
                        val_stat_curr = clean_text_input(row_data.get('STATUS')).upper()
                        idx_stat = list_status.index(val_stat_curr) if val_stat_curr in list_status else 0
                        val_status = st.selectbox("Status", list_status, index=idx_stat)
                    
                    btn_update = st.form_submit_button("Simpan Perubahan (Update)")
                    
                    if btn_update:
                        if isinstance(val_qa_range, tuple) and len(val_qa_range) == 2:
                            edit_start_qa, edit_end_qa = val_qa_range
                        else:
                            edit_start_qa = val_qa_range[0] if isinstance(val_qa_range, tuple) else val_qa_range
                            edit_end_qa = edit_start_qa

                        full_df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                        full_df_raw.columns = [str(c).strip() for c in full_df_raw.columns]
                        
                        key_aplikasi = str(row_data.get('APLIKASI')).strip().upper()
                        key_keterangan = str(row_data.get('KETERANGAN')).strip().upper()
                        
                        match_condition = (
                            (full_df_raw['APLIKASI'].astype(str).str.strip().str.upper() == key_aplikasi) & 
                            (full_df_raw['KETERANGAN'].astype(str).str.strip().str.upper() == key_keterangan)
                        )
                        matched_indices = full_df_raw[match_condition].index
                        
                        if len(matched_indices) > 0:
                            real_excel_idx = matched_indices[0]
                        else:
                            real_excel_idx = target_idx 

                        kolom_tanggal_all = ['TANGGAL', 'TGL PENGERJAAN', 'DEADLINE', 'TGL SELESAI', 'TGL MULAI QA', 'TGL SELESAI QA', 'TGL VERIFIKASI']
                        for col_name in kolom_tanggal_all:
                            if col_name in full_df_raw.columns:
                                full_df_raw[col_name] = full_df_raw[col_name].astype(object)
                        
                        kolom_target = [
                            'TANGGAL', 'APLIKASI', 'KETERANGAN', 'DEVELOPER', 'JENIS', 
                            'TGL PENGERJAAN', 'DEADLINE', 'TGL SELESAI', 'PEMERIKSA', 
                            'TGL MULAI QA', 'TGL SELESAI QA', 'TGL VERIFIKASI', 'VERIFIKATOR', 'STATUS', 'PEMBUAT'
                        ]
                        
                        nilai_target = [
                            format_ke_str(val_tanggal), val_aplikasi, val_keterangan.strip(), val_developer, val_jenis,
                            format_ke_str(val_tgl_pengerjaan), format_ke_str(val_deadline), format_ke_str(val_tgl_selesai), val_pemeriksa,
                            format_ke_str(edit_start_qa), format_ke_str(edit_end_qa),
                            format_ke_str(val_tgl_verifikasi), val_verifikator, val_status, val_pembuat.strip()
                        ]
                        
                        full_df_raw.loc[real_excel_idx, kolom_target] = nilai_target
                        
                        try:
                            with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                                full_df_raw.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                            st.success("🎉 Data sukses ditimpa pas di baris aslinya!")
                            st.cache_data.clear() 
                            st.rerun()
                        except Exception as e:
                            st.error(f"Gagal mengupdate data ke excel: {e}")

            # --- PROSES DELETE BARIS SINGLE ---
            elif action_mode == "Hapus Data (Delete)":
                st.error(f"⚠️ Apakah Anda yakin ingin menghapus permanen baris ke-{target_idx + 2} ini?")
                st.caption(f"**Detail Data:** [{row_data.get('APLIKASI')}] {row_data.get('KETERANGAN')}")
                
                btn_delete = st.button("Ya, Hapus Permanen Data Ini")
                
                if btn_delete:
                    full_df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                    full_df_raw = full_df_raw.drop(index=target_idx)
                    
                    with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                        full_df_raw.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                        
                    st.success("🗑️ Data berhasil dihapus dari file Excel!")
                    st.cache_data.clear()
                    st.rerun()
    else:
        st.warning("Belum ada data aktif yang bisa ditampilkan.")

# ==================== MENU 2: INPUT DATA BARU ====================
elif menu == "Input Data Baru":
    st.title("📝 Penginputan Data Revisi")
    pilihan_input = st.radio("Pilih Metode Input:", ["✍️ Form Input Manual", "📥 Import Massal via Excel"], horizontal=True)
    
    if pilihan_input == "✍️ Form Input Manual":
        with st.form("input_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                tanggal = st.date_input("Tanggal Request", datetime.now().date())
                aplikasi = st.selectbox("Aplikasi", list_aplikasi)
                developer = st.selectbox("Developer", list_developer)
                jenis = st.selectbox("Jenis Pembuatan", list_jenis)
                keterangan = st.text_area("Keterangan / Detail Revisi")
                pembuat = st.text_input("Pembuat (User Request)")
            with col2:
                tgl_pengerjaan = st.date_input("Tanggal Mulai Pengerjaan", datetime.now().date())
                deadline = st.date_input("Deadline Target", datetime.now().date())
                tgl_selesai = st.date_input("Tanggal Selesai Pengerjaan", datetime.now().date())
                pemeriksa = st.selectbox("Pemeriksa (QA / Lead)", ["Ali"])

                # --- FIX KESALAHAN VARIABEL DI SINI ---
                # Menggunakan default waktu saat ini untuk input data baru murni, bukan mengambil dari 'selected_row'
                init_start = datetime.now().date()
                init_end = init_start + timedelta(days=1)

                qa_date_range = st.date_input(
                    "Rentang Waktu Pemeriksaan QA (Mulai - Selesai)",
                    value=(init_start, init_end)
                )
                
                tgl_verifikasi = st.date_input("Tanggal Verifikasi", datetime.now().date())
                verifikator = st.text_input("Verifikator", value="")
                status = st.selectbox("Status", list_status)
                
            submit_button = st.form_submit_button(label="Simpan Data ke Excel")
            
            if submit_button:
                if isinstance(qa_date_range, tuple) and len(qa_date_range) == 2:
                    tgl_mulai_qa, tgl_selesai_qa = qa_date_range
                else:
                    tgl_mulai_qa = qa_date_range[0] if isinstance(qa_date_range, tuple) else qa_date_range
                    tgl_selesai_qa = tgl_mulai_qa

                new_data = {
                    "TANGGAL": [format_ke_str(tanggal)], 
                    "APLIKASI": [aplikasi], 
                    "KETERANGAN": [keterangan.strip()],
                    "DEVELOPER": [developer], 
                    "JENIS": [jenis], 
                    "TGL PENGERJAAN": [format_ke_str(tgl_pengerjaan)],
                    "DEADLINE": [format_ke_str(deadline)], 
                    "TGL SELESAI": [format_ke_str(tgl_selesai)], 
                    "PEMERIKSA": [pemeriksa],
                    "TGL MULAI QA": [format_ke_str(tgl_mulai_qa)],     
                    "TGL SELESAI QA": [format_ke_str(tgl_selesai_qa)], 
                    "TGL VERIFIKASI": [format_ke_str(tgl_verifikasi)], 
                    "VERIFIKATOR": [verifikator], 
                    "STATUS": [status], 
                    "PEMBUAT": [pembuat.strip()]
                }
                new_df = pd.DataFrame(new_data)
                
                try:
                    existing_df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                    existing_df.columns = [str(c).strip() for c in existing_df.columns]
                    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                    
                    with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                        updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                    
                    st.success("✅ Data baru berhasil disimpan!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan data: {e}")

    elif pilihan_input == "📥 Import Massal via Excel":
        st.subheader("Upload File Excel untuk Import Data")
        
        kolom_template = ["TANGGAL", "APLIKASI", "KETERANGAN", "DEVELOPER", "JENIS", 
                          "TGL PENGERJAAN", "DEADLINE", "TGL SELESAI", "PEMERIKSA", 
                          "TGL MULAI QA", "TGL SELESAI QA",
                          "TGL VERIFIKASI", "VERIFIKATOR", "STATUS", "PEMBUAT"]
        
        uploaded_file = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])
        
        if uploaded_file is not None:
            try:
                import_df = pd.read_excel(uploaded_file)
                import_df.columns = [str(c).strip() for c in import_df.columns]
                
                missing_cols = [col for col in kolom_template if col not in import_df.columns]
                
                if missing_cols:
                    st.error(f"❌ Gagal Import! Kolom berikut tidak ditemukan di file Anda: {missing_cols}")
                else:
                    st.write("📊 **Pratinjau Data (Maksimal 5 Baris Pertama):**")
                    st.dataframe(import_df.head(5))
                    
                    proses_import = st.button("Konfirmasi & Gabungkan ke Master Excel")
                    
                    if proses_import:
                        existing_df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                        existing_df.columns = [str(c).strip() for c in existing_df.columns]
                        
                        import_df["TGL VERIFIKASI"] = import_df["TGL VERIFIKASI"].fillna("")
                        import_df["VERIFIKATOR"] = import_df["VERIFIKATOR"].fillna("")
                        
                        updated_df = pd.concat([existing_df, import_df], ignore_index=True)
                        
                        with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                            updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                            
                        st.success(f"🎉 Berhasil! {len(import_df)} baris data telah digabungkan.")
                        st.cache_data.clear()
                        st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file Excel: {e}")

# ==================== MENU 3: HAPUS BANYAK DATA ====================
elif menu == "Hapus Banyak Data":
    st.title("🗑️ Multiple Delete Data Master")
    
    try:
        df_master = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        
        if df_master.empty:
            st.info("Database kosong, tidak ada data yang bisa dihapus.")
        else:
            st.warning("💡 Cara Hapus: Centang kolom 'Pilih Hapus' pada baris yang ingin dibuang, lalu klik tombol merah di bawah.")
            
            df_with_checkbox = df_master.copy()
            df_with_checkbox.insert(0, "Pilih Hapus", False)
            
            edited_df = st.data_editor(
                df_with_checkbox,
                use_container_width=True,
                hide_index=False,
                disabled=[col for col in df_with_checkbox.columns if col != "Pilih Hapus"],
                key="bulk_delete_editor"
            )
            
            rows_to_delete = edited_df[edited_df["Pilih Hapus"] == True].index.tolist()
            
            if len(rows_to_delete) > 0:
                btn_bulk_delete = st.button(f"🔴 Hapus Permanen {len(rows_to_delete)} Baris Terpilih", type="primary")
                
                if btn_bulk_delete:
                    df_master_clean = df_master.drop(index=rows_to_delete)
                    
                    with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                        df_master_clean.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                        
                    st.success(f"🗑️ Sukses menghapus {len(rows_to_delete)} baris dari Excel!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.info("Silakan centang satu atau beberapa baris di atas untuk mulai menghapus.")
    except Exception as e:
        st.error(f"Gagal memuat menu hapus banyak data: {e}")