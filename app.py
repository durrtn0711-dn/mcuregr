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
# POP-UP DIALOG UNTUK PROSES PRESENSI / EDIT
# ==========================================
@st.dialog("📋 Detail & Presensi Peserta")
def popup_detail_peserta(p, K):
    is_draft = K.get("status") == "DRAFT"
    hadir = p.get("hadir")
    kumpul = p.get("kumpul")
    no_p = p['no']

    st.markdown(f"### {no_p}. {p['nama']}")
    st.write(f"**No RM:** {p.get('rm')} | **NIK:** {p.get('nik', '-')}")
    st.write(f"**Tgl Lahir:** {format_tgl(p.get('tgl'))} | **Usia:** {p.get('usia', '-')} thn | **JK:** {p.get('jk')}")
    
    err_nik = analisa_nik_vs_tgl(p.get('nik'), p.get('tgl'), p.get('jk'))
    if err_nik:
        st.error(err_nik)

    if p.get("haid"):
        st.warning(f"Keterangan: **{p['haid']}**")
    if p.get("cek"):
        st.error(f"Perlu dicek dengan KTP: {p['cek']}")

    st.divider()

    if is_draft:
        if not hadir:
            st.subheader("Opsi Kehadiran")
            haid_final_val = ""
            jk_val = str(p.get("jk", "")).upper()
            
            if jk_val in ["P", "PEREMPUAN"]:
                opt_h = st.radio("Status Khusus Perempuan", ["Tidak ada", "HAID", "HAMIL"], key=f"pop_rad_{no_p}")
                if opt_h == "HAID":
                    hari_h = st.selectbox("Hari Ke-", ["1", "2", "3", "4", "5", "6", "7", "Lainnya"], key=f"pop_sel_{no_p}")
                    if hari_h == "Lainnya":
                        custom_h = st.text_input("Hari keberapa?", key=f"pop_txt_{no_p}").strip()
                        haid_final_val = f"HAID+{custom_h}" if custom_h else "HAID"
                    else:
                        haid_final_val = f"HAID+{hari_h}"
                elif opt_h == "HAMIL":
                    haid_final_val = "HAMIL"

            if st.button("✅ Set Hadir", type="primary", use_container_width=True):
                res_p = run_apps_script("setHadir", K["id"], no_p, haid_final_val)
                p.update(res_p)
                st.rerun()

        elif hadir and not kumpul:
            if st.button("📁 Berkas Terkumpul", type="primary", use_container_width=True):
                res_p = run_apps_script("setKumpul", K["id"], no_p)
                p.update(res_p)
                st.rerun()

    # FORM EDIT
    with st.expander("✏️ Edit Data Peserta Ini"):
        with st.form(key=f"form_pop_edit_{no_p}"):
            e_nama = st.text_input("Nama Peserta", value=p.get("nama", ""))
            e_rm = st.text_input("No RM", value=p.get("rm", ""))
            e_nik = st.text_input("NIK", value=str(p.get("nik", "")))
            e_tgl = st.text_input("Tgl Lahir (YYYY-MM-DD)", value=str(p.get("tgl", "")))
            e_jk = st.selectbox("Jenis Kelamin", ["L", "P"], index=0 if str(p.get("jk","")).upper().startswith("L") else 1)
            
            if st.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
                data_update = {
                    "nama": e_nama,
                    "rm": e_rm,
                    "nik": e_nik,
                    "tgl": e_tgl,
                    "jk": e_jk
                }
                res_upd = run_apps_script("simpanPeserta", K["id"], no_p, data_update)
                p.update(res_upd)
                st.success("Tersimpan!")
                st.rerun()

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
# 4. DETAIL KEGIATAN & PRESENSI (DIRECT CSS TARGETING - 100% BLOK WARNA)
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

    # ENJEKSI CSS LANGSUNG MENARGETKAN KEY KARTU PESERTA (100% PASTI WARNA WARNI RAPAT)
    css_rules = """
    <style>
    .block-container { padding: 0.3rem 0.2rem !important; }
    div[data-testid="stVerticalBlock"] > div { gap: 0.1rem !important; }
    .stat-box {
        background-color: #1e1e1e; border: 1px solid #333; border-radius: 6px;
        padding: 6px 10px; display: flex; justify-content: space-around; align-items: center;
        margin-bottom: 6px; font-size: 0.85rem;
    }
    .stat-num { font-weight: bold; font-size: 1.1rem; }
    div.stButton > button {
        width: 100% !important; text-align: left !important; justify-content: flex-start !important;
        padding: 8px 10px !important; border-radius: 6px !important; font-weight: bold !important;
        font-size: 0.88rem !important; margin: 1px 0px !important; box-shadow: 0 1px 2px rgba(0,0,0,0.2);
    }
    """

    for p in filtered:
        no_p = p['no']
        hadir = p.get("hadir")
        kumpul = p.get("kumpul")
        
        if K.get("status") == "DITUTUP" and not hadir:
            bg_c, text_c, border_c = "#FF0000", "#FFFFFF", "#B30000" # MERAH
        elif kumpul:
            bg_c, text_c, border_c = "#92D050", "#000000", "#6A9E32" # HIJAU
        elif hadir:
            bg_c, text_c, border_c = "#F4F207", "#000000", "#C4C200" # KUNING
        else:
            bg_c, text_c, border_c = "#262730", "#FFFFFF", "#444444" # HITAM STANDAR

        # Kunci CSS Langsung per Tombol Key
        css_rules += f"""
        div.stButton > button[key="pcard_{no_p}"] {{
            background-color: {bg_c} !important;
            color: {text_c} !important;
            border: 1.5px solid {border_c} !important;
        }}
        div.stButton > button[key="pcard_{no_p}"]:hover {{
            background-color: {bg_c} !important;
            color: {text_c} !important;
            opacity: 0.9 !important;
        }}
        """

    css_rules += "</style>"
    st.markdown(css_rules, unsafe_allow_html=True)

    # RENDER DAFTAR KARTU NAMA
    for p in filtered:
        no_p = p['no']
        label_btn = f"{no_p}. {p['nama']} ({p['rm']})"
        
        # KLIK NAMA PESERTA -> MUNCUL POP-UP DIALOG DI TENGAH LAYAR
        if st.button(label_btn, key=f"pcard_{no_p}"):
            popup_detail_peserta(p, K)

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

        ws_raw = wb.create
