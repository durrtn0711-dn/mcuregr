import streamlit as st
import requests
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Setup Halaman
st.set_page_config(page_title="MCU Registrasi", page_icon="🏥", layout="centered")

DEFAULT_URL = "https://script.google.com/macros/s/AKfycbxvKfpvDrlunxxrDcV2VvOKBFzDxDprL8mFvsucx0PgM6-lnSRDE5TRKSjOemSG85HwNw/exec"

# --- CSS RESET UNTUK TAMPILAN HP SUPER RINGKAS ---
st.markdown("""
<style>
.block-container {
    padding-top: 0.3rem !important;
    padding-bottom: 1rem !important;
    padding-left: 0.2rem !important;
    padding-right: 0.2rem !important;
}

/* Jarak vertikal Streamlit dipadatkan */
div[data-testid="stVerticalBlock"] > div {
    gap: 0.15rem !important;
}

/* Badge Statistik Sejajar */
.stat-box {
    background-color: #1e1e1e;
    border: 1px solid #333;
    border-radius: 6px;
    padding: 6px 10px;
    display: flex;
    justify-content: space-around;
    align-items: center;
    margin-bottom: 6px;
    font-size: 0.85rem;
}
.stat-num {
    font-weight: bold;
    font-size: 1.1rem;
}

/* Styling Details Native HTML (Bisa di-style 100%) */
details.card-item {
    border-radius: 6px;
    margin-bottom: 3px !important;
    padding: 6px 10px;
    font-size: 0.9rem;
}
details.card-item summary {
    font-weight: bold;
    cursor: pointer;
    outline: none;
    list-style: none;
}
details.card-item summary::-webkit-details-marker {
    display: none;
}

/* Warna Blok Murni */
details.bg-kuning {
    background-color: #F4F207 !important;
    color: #000000 !important;
    border: 1.5px solid #C4C200 !important;
}
details.bg-ijo {
    background-color: #92D050 !important;
    color: #000000 !important;
    border: 1.5px solid #6A9E32 !important;
}
details.bg-merah {
    background-color: #FF0000 !important;
    color: #FFFFFF !important;
    border: 1.5px solid #B30000 !important;
}
details.bg-putih {
    background-color: #262730 !important;
    color: #FFFFFF !important;
    border: 1px solid #444444 !important;
}

.card-body {
    background-color: #1a1a1a;
    color: #ffffff;
    padding: 8px;
    border-radius: 4px;
    margin-top: 6px;
    border: 1px solid #333;
}
</style>
""", unsafe_allow_html=True)

# --- INITIALIZATION STATE ---
if "url" not in st.session_state:
    st.session_state.url = DEFAULT_URL
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
    url = st.session_state.url or DEFAULT_URL
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
        raise Exception(f"Server mengembalikan respon non-JSON (Status {response.status_code}). Pastikan setelan deployment 'Anyone'.")
    
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

def analisa_nik_vs_tgl(nik, tgl_str, jk):
    if not nik or len(str(nik).strip()) < 12 or not tgl_str or tgl_str == "-":
        return None
    nik_str = str(nik).strip().replace(".0", "")
    if len(nik_str) < 12:
        return None
    
    try:
        dd_nik = int(nik_str[6:8])
        mm_nik = int(nik_str[8:10])
        yy_nik = int(nik_str[10:12])

        parts = str(tgl_str).split("-")
        if len(parts) != 3:
            return None
        
        yy_real = int(parts[0]) % 100
        mm_real = int(parts[1])
        dd_real = int(parts[2])

        jk_upper = str(jk).strip().upper()
        if jk_upper in ["P", "PEREMPUAN"]:
            dd_real_nik = dd_real + 40
        else:
            dd_real_nik = dd_real

        if (dd_nik != dd_real_nik) or (mm_nik != mm_real) or (yy_nik != yy_real):
            return f"⚠️ NIK ({nik_str[6:12]}) vs Tgl Lahir ({tgl_str}) TIDAK COCOK!"
    except Exception:
        return None
    return None

# ==========================================
# 1. HALAMAN LOGIN
# ==========================================
if not st.session_state.logged_in:
    st.title("🏥 MCU Registrasi")
    with st.form("login_form"):
        st.subheader("Masuk Aplikasi")
        input_url = st.text_input("Alamat Server (URL Web App)", value=DEFAULT_URL)
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
# 2. HEADER & LOGOUT
# ==========================================
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("### 🏥 MCU Registrasi")
with col_h2:
    if st.button("Keluar", type="secondary"):
        st.session_state.logged_in = False
        st.session_state.pin = ""
        st.session_state.active_kegiatan = None
        st.rerun()

# ==========================================
# 3. HOME (INPUT EXCEL & DAFTAR KEGIATAN)
# ==========================================
if st.session_state.active_kegiatan is None:
    st.subheader("Mulai Kegiatan")
    
    if not st.session_state.daftar_pt:
        st.error("Daftar perusahaan kosong. Cek sheet PENGATURAN di Apps Script.")
    else:
        pt_pilihan = st.selectbox("Pilih PT", st.session_state.daftar_pt)
        tgl_kegiatan = st.date_input("Tanggal", datetime.now())
        uploaded_file = st.file_uploader("Upload File Mentah Excel", type=["xlsx", "xls"])

        if uploaded_file and st.button("Periksa File Excel", use_container_width=True):
            with st.spinner("Membaca file..."):
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
                        "res": res
                    }
                except Exception as e:
                    st.error(f"Gagal membaca file: {e}")

    if "temp_upload" in st.session_state and st.session_state.temp_upload:
        up = st.session_state.temp_upload
        r = up["res"]
        
        st.write("---")
        st.info(f"**{r['total']} Peserta** | {r['baru']} baru, {r['lewat']} dilewati, {len(r['cek'])} perlu dicek")
        
        if r['cek']:
            for c in r['cek']:
                st.caption(f"• No {c['no']}. **{c['nama']}**: {c['masalah']}")
        
        c_simpan, c_batal = st.columns(2)
        with c_simpan:
            if st.button(f"Simpan {r['baru']} Peserta", disabled=(r['baru'] == 0), type="primary"):
                try:
                    with st.spinner("Menyimpan..."):
                        s = run_apps_script("simpanUpload", up["kode"], up["tgl"], up["payload"])
                        st.session_state.active_kegiatan = run_apps_script("getKegiatan", s["keg"])
                        del st.session_state.temp_upload
                        st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")
        with c_batal:
            if st.button("Batal"):
                del st.session_state.temp_upload
                st.rerun()

    st.divider()
    st.subheader("Daftar Kegiatan")
    
    if not st.session_state.kegiatan_list:
        st.caption("Belum ada kegiatan.")
    else:
        for g in st.session_state.kegiatan_list:
            col_k1, col_k2 = st.columns([3, 1])
            with col_k1:
                st.write(f"**{g['kode']}** · {format_tgl(g['tgl'])}")
            with col_k2:
                status_label = "Draft" if g['status'] == "DRAFT" else "Ditutup"
                if st.button(f"Buka ({status_label})", key=f"keg_{g['id']}"):
                    with st.spinner("Memuat..."):
                        st.session_state.active_kegiatan = run_apps_script("getKegiatan", g['id'])
                        st.rerun()

    st.divider()
    with st.expander("🗑️ Reset Semua Data"):
        if st.button("Ya, Hapus Semua Data", type="primary"):
            run_apps_script("resetSemua")
            load_home_data()
            st.rerun()

# ==========================================
# 4. DETAIL KEGIATAN (HTML DETAILS NATIVE - 100% BLOK WARNA & AUTO COLLAPSE)
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
        status_text = " (Ditutup)" if K.get("status") == "DITUTUP" else ""
        st.markdown(f"#### {K['kode']} · {format_tgl(K['tgl'])}{status_text}")

    q = st.text_input("🔍 Cari Nama, RM, NIK, No TTD", "", label_visibility="collapsed", placeholder="🔍 Cari Nama, RM, NIK, No TTD").strip().upper()

    sT = len(rows)
    sH = len([p for p in rows if p.get("hadir")])
    sK = len([p for p in rows if p.get("kumpul")])

    st.markdown(f"""
    <div class="stat-box">
        <div class="stat-item">Peserta: <span class="stat-num">{sT}</span></div>
        <div class="stat-item" style="color:#F4F207;">Hadir: <span class="stat-num">{sH}</span></div>
        <div class="stat-item" style="color:#92D050;">Berkas: <span class="stat-num">{sK}</span></div>
    </div>
    """, unsafe_allow_html=True)

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

    # RENDER DAFTAR KARTU HTML NATIVE DENGAN BLOK WARNA WARNI MULTI-BROWSER
    for p in filtered:
        is_draft = K.get("status") == "DRAFT"
        hadir = p.get("hadir")
        kumpul = p.get("kumpul")
        no_p = p['no']
        
        if K.get("status") == "DITUTUP" and not hadir:
            bg_cls = "bg-merah"
        elif kumpul:
            bg_cls = "bg-ijo"
        elif hadir:
            bg_cls = "bg-kuning"
        else:
            bg_cls = "bg-putih"

        # ELEMEN HTML DETAILS NATIVE (WARNA WARNI MANDIRI 100% BLOK + AUTO CLOSED PAS REFRESH)
        st.markdown(f"""
        <details class="card-item {bg_cls}">
            <summary>{no_p}. {p['nama']} ({p['rm']})</summary>
        </details>
        """, unsafe_allow_html=True)

        # AKSI PENDAFTARAN & EDIT HANYA BISA DIAKSES LEWAT TOMBOL DI SEBELAH / BAWAHNYA
        if is_draft:
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                if not hadir:
                    haid_final_val = ""
                    jk_val = str(p.get("jk", "")).upper()
                    
                    if jk_val in ["P", "PEREMPUAN"]:
                        opt_h = st.radio("Khusus", ["Tidak ada", "HAID", "HAMIL"], key=f"rad_h_{no_p}")
                        if opt_h == "HAID":
                            hari_h = st.selectbox("Hari Ke-", ["1", "2", "3", "4", "5", "6", "7", "Lainnya"], key=f"sel_h_{no_p}")
                            if hari_h == "Lainnya":
                                custom_h = st.text_input("Hari keberapa?", key=f"txt_h_{no_p}").strip()
                                haid_final_val = f"HAID+{custom_h}" if custom_h else "HAID"
                            else:
                                haid_final_val = f"HAID+{hari_h}"
                        elif opt_h == "HAMIL":
                            haid_final_val = "HAMIL"

                    if st.button("Set Hadir", key=f"btn_h_{no_p}", type="primary"):
                        res_p = run_apps_script("setHadir", K["id"], no_p, haid_final_val)
                        p.update(res_p)
                        st.rerun()

            with c2:
                if hadir and not kumpul:
                    if st.button("Berkas Terkumpul", key=f"btn_k_{no_p}"):
                        res_p = run_apps_script("setKumpul", K["id"], no_p)
                        p.update(res_p)
                        st.rerun()

            with c3:
                if st.button("✏️ Edit", key=f"btn_edit_{no_p}"):
                    st.session_state[f"edit_mode_{no_p}"] = True

        if st.session_state.get(f"edit_mode_{no_p}", False):
            with st.form(key=f"form_edit_{no_p}"):
                st.subheader(f"Edit Data No. {no_p}")
                e_nama = st.text_input("Nama Peserta", value=p.get("nama", ""))
                e_rm = st.text_input("No RM", value=p.get("rm", ""))
                e_nik = st.text_input("NIK", value=str(p.get("nik", "")))
                e_tgl = st.text_input("Tgl Lahir (YYYY-MM-DD)", value=str(p.get("tgl", "")))
                e_jk = st.selectbox("Jenis Kelamin", ["L", "P"], index=0 if str(p.get("jk","")).upper().startswith("L") else 1)
                
                c_save, c_canc = st.columns(2)
                with c_save:
                    if st.form_submit_button("💾 Simpan"):
                        data_update = {
                            "nama": e_nama,
                            "rm": e_rm,
                            "nik": e_nik,
                            "tgl": e_tgl,
                            "jk": e_jk
                        }
                        res_upd = run_apps_script("simpanPeserta", K["id"], no_p, data_update)
                        p.update(res_upd)
                        st.session_state[f"edit_mode_{no_p}"] = False
                        st.success("Tersimpan!")
                        st.rerun()
                with c_canc:
                    if st.form_submit_button("Batal"):
                        st.session_state[f"edit_mode_{no_p}"] = False
                        st.rerun()

    # TAMBAH MANUAL PESERTA
    if K.get("status") == "DRAFT":
        with st.expander("➕ Tambah Peserta Manual"):
            with st.form("form_tambah_manual"):
                m_nama = st.text_input("Nama Lengkap")
                m_rm = st.text_input("No RM")
                m_nik = st.text_input("NIK (16 Digit)")
                m_tgl = st.date_input("Tanggal Lahir", datetime(2000, 1, 1))
                m_jk = st.selectbox("Jenis Kelamin", ["L", "P"])
                
                if st.form_submit_button("Simpan Peserta Manual", type="primary"):
                    if m_nama and m_rm:
                        new_data = {
                            "nama": m_nama,
                            "rm": m_rm,
                            "nik": m_nik,
                            "tgl": m_tgl.strftime("%Y-%m-%d"),
                            "jk": m_jk,
                            "status": "BELUM"
                        }
                        res_new = run_apps_script("tambahPeserta", K["id"], new_data)
                        rows.append(res_new)
                        st.success(f"{m_nama} berhasil ditambahkan!")
                        st.rerun()
                    else:
                        st.error("Nama dan No RM wajib diisi!")

    st.divider()

    # ==========================================
    # EXPORT EXCEL 4 SHEET PERFEK
    # ==========================================
    if rows:
        st.subheader("📥 Export File Excel Final (4 Sheet)")
        
        kode_pt = str(K.get("kode", "")).upper()
        tgl_indo_str = format_tgl(K.get("tgl"))
        
        FILL_GREEN = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        FILL_YELLOW = PatternFill(start_color="F4F207", end_color="F4F207", fill_type="solid")
        FILL_RED = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        FILL_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        THIN_SIDE = Side(border_style="thin", color="000000")
        TABLE_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

        ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
        ALIGN_LEFT = Alignment(horizontal="left", vertical="center")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        if "SMJ" in kode_pt:
            judul_pt = "DAFTAR HADIR MEDICAL CHECK UP PT SAMI JEPARA FACTORY"
            headers = ['NO', 'NO RM', 'NAMA', 'JK', 'TANGGAL LAHIR (BLN-TGL-THN)', 'USIA', 'NO KTP', 'NO HP', 'JENIS MCU', 'TTD', 'HAID/HAMIL']
            cols_map = ['no', 'rm', 'nama', 'jk', 'tgl', 'usia', 'nik', 'hp', 'jenis']
        else:
            judul_pt = "DAFTAR HADIR MEDICAL CHECK UP PT PARKLAND WORLD INDONESIA"
            headers = ['NO', 'NO RM', 'NAMA', 'NO. URUT MCU', 'BAGIAN', 'JABATAN', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'JENIS KELAMIN', 'NO HP', 'NO. KTP', 'TTD', 'HAID/HAMIL']
            cols_map = ['no', 'rm', 'nama', 'urut', 'bagian', 'jabatan', 'tempat', 'tgl', 'usia', 'jk', 'hp', 'nik']

        ws_raw = wb.create_sheet(title="MCU KLINIK JEPARA")
        ws_raw.append(headers)
        for p in rows:
            r_val = [p.get(c, "") for c in cols_map] + [p.get("no", ""), p.get("haid", "")]
            ws_raw.append(r_val)

        def build_styled_sheet(sheet_title, data_subset, is_pendaftaran=False):
            ws = wb.create_sheet(title=sheet_title)
            
            ws.cell(row=1, column=1, value=judul_pt).font = Font(bold=True, size=11)
            ws.cell(row=2, column=1, value=f"TANGGAL {tgl_indo_str}").font = Font(bold=True, size=11)
            
            for col_idx, text in enumerate(headers, start=1):
                cell = ws.cell(row=4, column=col_idx, value=text)
                cell.fill = FILL_HEADER
                cell.font = Font(bold=True)
                cell.alignment = ALIGN_CENTER
                cell.border = TABLE_BORDER
            
            for r_idx, p in enumerate(data_subset, start=5):
                if is_pendaftaran:
                    if p.get("kumpul"):
                        row_fill = FILL_GREEN
                    elif p.get("hadir"):
                        row_fill = FILL_YELLOW
                    else:
                        row_fill = FILL_RED
                else:
                    row_fill = FILL_GREEN if sheet_title == "HADIR" else FILL_RED

                for c_idx, key in enumerate(cols_map, start=1):
                    val = p.get(key, "")
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.fill = row_fill
                    cell.border = TABLE_BORDER
                    
                    if key == "nama":
                        cell.alignment = ALIGN_LEFT
                    else:
                        cell.alignment = ALIGN_CENTER

                col_ttd_idx = len(cols_map) + 1
                cell_ttd = ws.cell(row=r_idx, column=col_ttd_idx, value=p.get("no", ""))
                cell_ttd.fill = row_fill
                cell_ttd.border = TABLE_BORDER
                cell_ttd.alignment = ALIGN_CENTER

                col_haid_idx = len(cols_map) + 2
                cell_haid = ws.cell(row=r_idx, column=col_haid_idx, value=p.get("haid", ""))
                cell_haid.fill = row_fill
                cell_haid.border = TABLE_BORDER
                cell_haid.alignment = ALIGN_CENTER

        build_styled_sheet("PENDAFTARAN", rows, is_pendaftaran=True)
        peserta_hadir = [p for p in rows if p.get("hadir")]
        build_styled_sheet("HADIR", peserta_hadir)
        peserta_tidak = [p for p in rows if not p.get("hadir")]
        build_styled_sheet("TIDAK HADIR", peserta_tidak)

        buffer = io.BytesIO()
        wb.save(buffer)

        st.download_button(
            label=f"📄 Unduh Excel Final ({kode_pt})",
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
