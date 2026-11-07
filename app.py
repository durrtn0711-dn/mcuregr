import streamlit as st
import requests
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="MCU Registrasi", page_icon="🏥", layout="centered")

# --- INITIALIZATION STATE ---
if "url" not in st.session_state:
    st.session_state.url = ""
if "pin" not in st.session_state:
    st.session_state.pin = ""
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "daftar_pt" not in st.session_state:
    st.session_state.daftar_pt = []
if "kegiatan_list" not in st.session_state:
    st.session_state.kegiatan_list = []
if "active_kegiatan" not in st.session_state:
    st.session_state.active_kegiatan = None

def run_apps_script(fn, *args):
    url = st.session_state.url
    pin = st.session_state.pin
    if not url or not pin:
        raise Exception("Alamat URL Server atau PIN belum diisi.")
    
    payload = {"pin": pin, "fn": fn, "args": list(args)}
    
    try:
        response = requests.post(url, json=payload, timeout=30, allow_redirects=True)
    except Exception as e:
        raise Exception(f"Gagal terhubung ke server: {e}")
    
    if response.status_code == 404:
        raise Exception("URL Apps Script 404. Cek kembali URL /exec kamu.")
    
    try:
        data = response.json()
    except Exception:
        raise Exception(f"Server mengembalikan respon non-JSON (Status {response.status_code}). Cek akses 'Anyone' di deployment Apps Script.")
    
    if not data.get("ok"):
        raise Exception(data.get("error", "Terjadi kesalahan pada backend Apps Script."))
    
    return data.get("data")

def load_home_data():
    home_data = run_apps_script("getHome")
    st.session_state.daftar_pt = [p["kode"] for p in home_data.get("pt", [])]
    st.session_state.kegiatan_list = home_data.get("keg", [])

def format_tgl(tgl_str):
    if not tgl_str or tgl_str == "-":
        return "-"
    parts = str(tgl_str).split("-")
    if len(parts) == 3:
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return str(tgl_str)

# ==========================================
# 1. HALAMAN LOGIN
# ==========================================
if not st.session_state.logged_in:
    st.title("🏥 MCU Registrasi")
    with st.form("login_form"):
        st.subheader("Masuk Aplikasi")
        input_url = st.text_input("Alamat Server (URL Web App Apps Script)", value=st.session_state.url)
        input_pin = st.text_input("PIN", type="password")
        btn_login = st.form_submit_button("Masuk")

        if btn_login:
            clean_url = input_url.strip()
            clean_pin = input_pin.strip()
            if not clean_url.startswith("https://script.google.com/"):
                st.error("Alamat server harus diawali https://script.google.com/ dan berakhiran /exec")
            elif not clean_pin:
                st.error("Isi PIN terlebih dahulu.")
            else:
                st.session_state.url = clean_url
                st.session_state.pin = clean_pin
                try:
                    with st.spinner("Memeriksa PIN dan memuat data..."):
                        load_home_data()
                        st.session_state.logged_in = True
                        st.success("Berhasil masuk!")
                        st.rerun()
                except Exception as e:
                    st.error(f"Gagal masuk: {e}")
    st.stop()

# ==========================================
# 2. HEADER
# ==========================================
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("🏥 MCU Registrasi")
with col_h2:
    if st.button("Keluar / Ganti PIN", type="secondary"):
        st.session_state.logged_in = False
        st.session_state.pin = ""
        st.session_state.active_kegiatan = None
        st.rerun()

# ==========================================
# 3. HOME (INPUT EXCEL & DAFTAR KEGIATAN)
# ==========================================
if st.session_state.active_kegiatan is None:
    st.subheader("Mulai Kegiatan atau Tambah Peserta")
    st.caption("Pilih perusahaan, tanggal kegiatan, lalu upload file mentah Excel.")

    if not st.session_state.daftar_pt:
        st.error("Daftar perusahaan kosong. Pastikan sheet PENGATURAN di Apps Script diisi.")
    else:
        pt_pilihan = st.selectbox("Pilih Perusahaan / PT", st.session_state.daftar_pt)
        tgl_kegiatan = st.date_input("Tanggal Kegiatan", datetime.now())
        uploaded_file = st.file_uploader("Upload File Excel Mentah (.xlsx, .xls)", type=["xlsx", "xls"])

        if uploaded_file and st.button("Periksa File Excel", use_container_width=True):
            with st.spinner("Membaca file Excel..."):
                try:
                    excel_data = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                    payload_sheets = []
                    
                    SKIP_REGEX = re.compile(r"^(PENDAFTARAN|HADIR|TIDAK HADIR|MCU KLINIK JEPARA|DATA MENTAH|MAILINGS)$", re.I)
                    
                    for sheet_name, df in excel_data.items():
                        if not SKIP_REGEX.match(sheet_name):
                            rows = df.fillna("").values.tolist()
                            rows_trimmed = [r[:40] for r in rows]
                            payload_sheets.append({
                                "name": sheet_name,
                                "rows": rows_trimmed
                            })
                    
                    tgl_str = tgl_kegiatan.strftime("%Y-%m-%d")
                    res = run_apps_script("previewUpload", pt_pilihan, tgl_str, payload_sheets)
                    
                    st.session_state.temp_upload = {
                        "kode": pt_pilihan,
                        "tgl": tgl_str,
                        "payload": payload_sheets,
                        "res": res,
                        "raw_bytes": uploaded_file.getvalue()
                    }
                except Exception as e:
                    st.error(f"Gagal membaca file: {e}")

    if "temp_upload" in st.session_state and st.session_state.temp_upload:
        up = st.session_state.temp_upload
        r = up["res"]
        
        st.write("---")
        st.subheader("📊 Hasil Pemeriksaan File")
        st.info(f"**{r['total']} Peserta di file** | {r['baru']} baru, {r['lewat']} dilewati karena sudah ada, {len(r['cek'])} perlu dicek")
        
        if r['cek']:
            st.warning("⚠️ **Perlu Dicek:**")
            for c in r['cek']:
                st.caption(f"• No {c['no']}. **{c['nama']}**: {c['masalah']}")
        
        c_simpan, c_batal = st.columns(2)
        with c_simpan:
            if st.button(f"Simpan {r['baru']} Peserta", disabled=(r['baru'] == 0), type="primary"):
                try:
                    with st.spinner("Menyimpan ke Google Sheets..."):
                        s = run_apps_script("simpanUpload", up["kode"], up["tgl"], up["payload"])
                        st.session_state.active_kegiatan = run_apps_script("getKegiatan", s["keg"])
                        del st.session_state.temp_upload
                        st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan data: {e}")
        with c_batal:
            if st.button("Batal Upload"):
                del st.session_state.temp_upload
                st.rerun()

    st.divider()
    st.subheader("Daftar Kegiatan MCU")
    
    if not st.session_state.kegiatan_list:
        st.caption("Belum ada kegiatan. Upload data mentah di atas untuk memulai.")
    else:
        for g in st.session_state.kegiatan_list:
            col_k1, col_k2 = st.columns([3, 1])
            with col_k1:
                st.write(f"**{g['kode']}** · {format_tgl(g['tgl'])}")
            with col_k2:
                status_label = "Draft" if g['status'] == "DRAFT" else "Ditutup"
                if st.button(f"Buka ({status_label})", key=f"keg_{g['id']}"):
                    with st.spinner("Memuat kegiatan..."):
                        st.session_state.active_kegiatan = run_apps_script("getKegiatan", g['id'])
                        st.rerun()

    st.divider()
    with st.expander("🗑️ Reset Semua Data"):
        if st.button("Ya, Hapus Semua Data", type="primary"):
            run_apps_script("resetSemua")
            load_home_data()
            st.rerun()

# ==========================================
# 4. DETAIL KEGIATAN & PRESENSI
# ==========================================
else:
    K = st.session_state.active_kegiatan
    rows = K.get("rows", [])

    col_b1, col_b2 = st.columns([1, 3])
    with col_b1:
        if st.button("⬅️ Kembali"):
            st.session_state.active_kegiatan = None
            if "pilih_beda" in st.session_state:
                del st.session_state.pilih_beda
            load_home_data()
            st.rerun()
    with col_b2:
        status_text = " · Ditutup" if K.get("status") == "DITUTUP" else ""
        st.subheader(f"{K['kode']} · {format_tgl(K['tgl'])}{status_text}")

    # Search Bar
    q = st.text_input("🔍 Cari nama, No RM, NIK, atau nomor TTD", "").strip().upper()

    sT = len(rows)
    sH = len([p for p in rows if p.get("hadir")])
    sK = len([p for p in rows if p.get("kumpul")])

    m1, m2, m3 = st.columns(3)
    m1.metric("Peserta", sT)
    m2.metric("Hadir", sH)
    m3.metric("Berkas", sK)

    # Filter Peserta
    filtered = []
    for p in rows:
        if not q:
            filtered.append(p)
        else:
            if q.isdigit() and int(q) == p.get("no"):
                filtered.append(p)
            elif (q in p.get("nama", "").upper() or 
                  q in p.get("rm", "").upper() or 
                  q in str(p.get("nik", ""))):
                filtered.append(p)

    st.divider()

    # List Cards Peserta
    for p in filtered:
        is_draft = K.get("status") == "DRAFT"
        hadir = p.get("hadir")
        kumpul = p.get("kumpul")
        
        badge = ""
        if K.get("status") == "DITUTUP" and not hadir:
            badge = "🔴 Tidak hadir"
        elif kumpul:
            badge = "🟢 Berkas ✓"
        elif hadir:
            badge = "🟡 Hadir"

        with st.expander(f"{p['no']}. {p['nama']} ({p['rm']}) {badge}"):
            st.write(f"**No RM:** {p.get('rm')} | **NIK:** {p.get('nik', '-')}")
            st.write(f"**Tgl Lahir:** {format_tgl(p.get('tgl'))} | **Usia:** {p.get('usia', '-')} thn | **JK:** {p.get('jk')}")
            if p.get("haid"):
                st.info(f"Keterangan Tambahan: **{p['haid']}**")
            if p.get("cek"):
                st.warning(f"Perlu dicek dengan KTP: {p['cek']}")

            if is_draft:
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    if not hadir:
                        haid_final_val = ""
                        jk_val = str(p.get("jk", "")).upper()
                        
                        if jk_val in ["P", "PEREMPUAN"]:
                            opt_h = st.radio("Status Khusus", ["Tidak ada", "HAID", "HAMIL"], key=f"rad_h_{p['no']}")
                            if opt_h == "HAID":
                                hari_h = st.text_input("Haid hari keberapa? (Opsional, cth: 2)", key=f"txt_h_{p['no']}").strip()
                                haid_final_val = f"HAID HARI KE-{hari_h}" if hari_h else "HAID"
                            elif opt_h == "HAMIL":
                                haid_final_val = "HAMIL"

                        if st.button("Set Hadir", key=f"btn_h_{p['no']}", type="primary"):
                            res_p = run_apps_script("setHadir", K["id"], p["no"], haid_final_val)
                            p.update(res_p)
                            st.rerun()

                with col_act2:
                    if hadir and not kumpul:
                        if st.button("Sudah mengumpulkan", key=f"btn_k_{p['no']}"):
                            res_p = run_apps_script("setKumpul", K["id"], p["no"])
                            p.update(res_p)
                            st.rerun()

    st.divider()

    # ==========================================
    # EXPORT EXCEL 4 SHEET PRESISI DENGAN BLOK WARNA (KUNING, IJO, MERAH)
    # ==========================================
    if rows:
        st.subheader("📥 Export File Excel Final (4 Sheet Berwarna)")
        
        kode_pt = str(K.get("kode", "")).upper()
        tgl_indo_str = format_tgl(K.get("tgl"))
        
        # DEFINISI WARNA STANDAR EXCEL SAMPLE
        FILL_GREEN = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # Hadir/Ijo
        FILL_YELLOW = PatternFill(start_color="FFF0A6", end_color="FFF0A6", fill_type="solid") # Hadir/Kuning
        FILL_RED = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")   # Tidak Hadir/Merah
        FILL_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")# Header Abu

        wb = openpyxl.Workbook()
        wb.remove(wb.active) # Hapus sheet default

        # HEADER BERDASARKAN PT
        if "SMJ" in kode_pt:
            judul_pt = "DAFTAR HADIR MEDICAL CHECK UP PT SAMI JEPARA FACTORY"
            headers = ['NO', 'NO RM', 'NAMA', 'JK', 'TANGGAL LAHIR (BLN-TGL-THN)', 'USIA', 'NO KTP', 'NO HP', 'JENIS MCU', 'TTD']
            cols_map = ['no', 'rm', 'nama', 'jk', 'tgl', 'usia', 'nik', 'hp', 'jenis']
        else:
            judul_pt = "DAFTAR HADIR MEDICAL CHECK UP PT PARKLAND WORLD INDONESIA"
            headers = ['NO', 'NO RM', 'NAMA', 'NO. URUT MCU', 'BAGIAN', 'JABATAN', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'JENIS KELAMIN', 'NO HP', 'NO. KTP', 'TTD', 'HAID/HAMIL']
            cols_map = ['no', 'rm', 'nama', 'urut', 'bagian', 'jabatan', 'tempat', 'tgl', 'usia', 'jk', 'hp', 'nik']

        # 1. SHEET MCU KLINIK JEPARA (DATA MENTAH)
        ws_raw = wb.create_sheet(title="MCU KLINIK JEPARA")
        raw_headers = [h for h in headers if h != 'TTD']
        ws_raw.append(raw_headers)
        for r_idx, p in enumerate(rows, start=1):
            row_data = [p.get(c, "") for c in cols_map]
            ws_raw.append(row_data)

        # FUNGSI MEMBUAT SHEET FORMAL (PENDAFTARAN, HADIR, TIDAK HADIR)
        def build_styled_sheet(sheet_title, data_subset, is_pendaftaran=False):
            ws = wb.create_sheet(title=sheet_title)
            
            # Header Judul Atas
            ws.cell(row=1, column=1, value=judul_pt).font = Font(bold=True, size=11)
            ws.cell(row=2, column=1, value=f"TANGGAL {tgl_indo_str}").font = Font(bold=True, size=11)
            
            # Row 3 Kosong
            # Row 4 Table Header
            for col_idx, text in enumerate(headers, start=1):
                cell = ws.cell(row=4, column=col_idx, value=text)
                cell.fill = FILL_HEADER
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Isi Data Pasien & Beri Warna Sesuai Aturan
            for r_idx, p in enumerate(data_subset, start=5):
                # Tentukan Warna Baris
                if is_pendaftaran:
                    if p.get("kumpul"):
                        row_fill = FILL_GREEN
                    elif p.get("hadir"):
                        row_fill = FILL_YELLOW
                    else:
                        row_fill = FILL_RED
                else:
                    row_fill = FILL_GREEN if sheet_title == "HADIR" else FILL_RED

                # Isi Kolom
                for c_idx, key in enumerate(cols_map, start=1):
                    val = p.get(key, "")
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.fill = row_fill

                # Kolom TTD
                cell_ttd = ws.cell(row=r_idx, column=len(cols_map) + 1, value=p.get("no", ""))
                cell_ttd.fill = row_fill
                cell_ttd.alignment = Alignment(horizontal="center")

                # Kolom HAID/HAMIL (Khusus PWJ)
                if "SMJ" not in kode_pt:
                    cell_haid = ws.cell(row=r_idx, column=len(cols_map) + 2, value=p.get("haid", ""))
                    cell_haid.fill = row_fill

        # 2. SHEET PENDAFTARAN (Semua Peserta + Hijau/Kuning/Merah)
        build_styled_sheet("PENDAFTARAN", rows, is_pendaftaran=True)

        # 3. SHEET HADIR (Hanya Peserta Hadir + Hijau Semua)
        peserta_hadir = [p for p in rows if p.get("hadir")]
        build_styled_sheet("HADIR", peserta_hadir)

        # 4. SHEET TIDAK HADIR (Hanya Peserta Tidak Hadir + Merah Semua)
        peserta_tidak = [p for p in rows if not p.get("hadir")]
        build_styled_sheet("TIDAK HADIR", peserta_tidak)

        # Export ke Streamlit Download
        buffer = io.BytesIO()
        wb.save(buffer)

        st.download_button(
            label=f"📄 Unduh Excel Final Presisi 4 Sheet ({kode_pt})",
            data=buffer.getvalue(),
            file_name=f"MCU_{kode_pt}_{K['tgl']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.divider()

    # --- CLOSE MCU & BUKA LAGI ---
    if K.get("status") == "DRAFT":
        if st.button("🔒 Close MCU", type="primary", use_container_width=True):
            try:
                with st.spinner("Memproses penutupan MCU..."):
                    res_close = run_apps_script("tutup", K["id"])
                    
                    if isinstance(res_close, dict) and res_close.get("perlu"):
                        st.session_state.pilih_beda = res_close
                        st.rerun()
                    else:
                        st.session_state.active_kegiatan = res_close
                        if "pilih_beda" in st.session_state:
                            del st.session_state.pilih_beda
                        st.success(f"MCU {K['kode']} berhasil ditutup!")
                        st.rerun()
            except Exception as e:
                st.error(f"Gagal Close MCU: {e}")

        # Konfirmasi perbedaan data awal vs olahan
        if "pilih_beda" in st.session_state and st.session_state.pilih_beda:
            pb = st.session_state.pilih_beda
            st.warning(f"⚠️ Ada {pb.get('jumlah', 0)} perbedaan data olahan dengan file pendaftaran awal.")
            
            col_aw, col_ol = st.columns(2)
            with col_aw:
                if st.button("Sesuai Data Awal"):
                    res_close = run_apps_script("tutup", K["id"], "AWAL")
                    st.session_state.active_kegiatan = res_close
                    del st.session_state.pilih_beda
                    st.rerun()
            with col_ol:
                if st.button("Sesuai Data Olahan"):
                    res_close = run_apps_script("tutup", K["id"], "OLAHAN")
                    st.session_state.active_kegiatan = res_close
                    del st.session_state.pilih_beda
                    st.rerun()

    else:
        if K.get("url"):
            st.markdown(f"[📂 Buka File Final Spreadsheet Drive]({K['url']})")
            st.markdown(f"[📥 Unduh File Excel Apps Script]({K['url'].replace('/edit', '/export?format=xlsx')})")
        
        st.info("Status MCU saat ini: **DITUTUP**.")
        if st.button("🔓 Buka lagi untuk diedit", use_container_width=True, type="primary"):
            try:
                with st.spinner("Membuka kembali MCU..."):
                    res_buka = run_apps_script("bukaLagi", K["id"])
                    st.session_state.active_kegiatan = res_buka
                    st.success("MCU berhasil dibuka kembali!")
                    st.rerun()
            except Exception as e:
                st.error(f"Gagal membuka MCU: {e}")

    # Reset MCU Spesifik
    with st.expander("⚠️ Reset Data MCU Ini"):
        if st.button("Hapus MCU Ini", type="secondary"):
            run_apps_script("resetKegiatan", K["id"])
            st.session_state.active_kegiatan = None
            load_home_data()
            st.rerun()
