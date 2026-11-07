import streamlit as st
import requests
import pandas as pd
import json
import re
from datetime import datetime

# Setup Halaman
st.set_page_config(page_title="MCU Registrasi", page_icon="🏥", layout="centered")

# --- INITIALIZATION STATE (Menggantikan localStorage JS) ---
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

# --- FUNGSI RUN (Sama persis fungsi run() di JS untuk panggil Apps Script) ---
def run_apps_script(fn, *args):
    url = st.session_state.url
    pin = st.session_state.pin
    if not url or not pin:
        raise Exception("URL Server atau PIN belum diset.")
    
    payload = {
        "pin": pin,
        "fn": fn,
        "args": list(args)
    }
    
    response = requests.post(url, json=payload, timeout=15, allow_redirects=True)
    
    try:
        data = response.json()
    except Exception:
        # Kalau bukan JSON, tampilkan teks HTML/error asli dari Google biar ketahuan penyebabnya
        raise Exception(f"Server tidak mengembalikan JSON. Status code: {response.status_code}. Respon server: {response.text[:200]}")
    
    if not data.get("ok"):
        raise Exception(data.get("error", "Terjadi kesalahan pada server Apps Script."))
    return data.get("data")
# --- HELPER FUNGSIONAL ---
def load_home_data():
    home_data = run_apps_script("getHome")
    # Ambil daftar PT dari sheet PENGATURAN
    st.session_state.daftar_pt = [p["kode"] for p in home_data.get("pt", [])]
    # Ambil daftar kegiatan yang ada
    st.session_state.kegiatan_list = home_data.get("keg", [])

def format_tgl(tgl_str):
    if not tgl_str:
        return "-"
    parts = str(tgl_str).split("-")
    if len(parts) == 3:
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return str(tgl_str)

# --- 1. HALAMAN LOGIN (Memeriksa URL & PIN ke Server) ---
if not st.session_state.logged_in:
    st.title("🏥 MCU Registrasi")
    with st.form("login_form"):
        st.subheader("Masuk")
        st.caption("Isi sekali, tersimpan di sesi ini.")
        
        input_url = st.text_input("Alamat server (URL web app Apps Script)")
        input_pin = st.text_input("PIN", type="password")
        btn_login = st.form_submit_button("Masuk")

        if btn_login:
            if not input_url.startswith("https://script.google.com/"):
                st.error("Alamat server harus diawali https://script.google.com/")
            elif not input_pin:
                st.error("Isi PIN.")
            else:
                st.session_state.url = input_url.strip()
                st.session_state.pin = input_pin.strip()
                try:
                    with st.spinner("Memeriksa PIN ke server..."):
                        load_home_data()
                        st.session_state.logged_in = True
                        st.rerun()
                except Exception as e:
                    st.error(f"Gagal masuk: {e}")
    st.stop()

# --- 2. HEADER & NAVIGASI (Keluar / Ganti PIN) ---
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("🏥 MCU Registrasi")
with col_h2:
    if st.button("Keluar / Ganti PIN", type="secondary"):
        st.session_state.logged_in = False
        st.session_state.pin = ""
        st.session_state.active_kegiatan = None
        st.rerun()

# --- 3. HALAMAN UTAMA (HOME) ---
if st.session_state.active_kegiatan is None:
    st.subheader("Mulai kegiatan atau tambah peserta")
    st.caption("Pilih perusahaan, tanggal, lalu file Excel mentah.")

    # Pilih PT (Otomatis sesuai aturan Sheet PENGATURAN)
    if not st.session_state.daftar_pt:
        st.error("Daftar perusahaan kosong. Cek sheet PENGATURAN atau jalankan siapkan() lagi di Apps Script.")
    else:
        pt_pilihan = st.selectbox("Pilih Perusahaan / PT", st.session_state.daftar_pt)
        tgl_kegiatan = st.date_input("Tanggal Kegiatan", datetime.now())
        uploaded_file = st.file_uploader("Upload File Excel Mentah (.xlsx, .xls)", type=["xlsx", "xls"])

        if uploaded_file and st.button("Periksa file", use_container_width=True):
            with st.spinner("Membaca file..."):
                try:
                    excel_data = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                    payload_sheets = []
                    
                    SKIP_REGEX = re.compile(r"^(PENDAFTARAN|HADIR|TIDAK HADIR|MCU KLINIK JEPARA|DATA MENTAH|MAILINGS)$", re.I)
                    
                    for sheet_name, df in excel_data.items():
                        if not SKIP_REGEX.match(sheet_name):
                            rows = df.fillna("").values.tolist()
                            # Ambil maksimal 40 kolom (sama persis JS: r.slice(0,40))
                            rows_trimmed = [r[:40] for r in rows]
                            payload_sheets.append({
                                "name": sheet_name,
                                "rows": rows_trimmed
                            })
                    
                    # Panggil previewUpload ke Apps Script
                    tgl_str = tgl_kegiatan.strftime("%Y-%m-%d")
                    res = run_apps_script("previewUpload", pt_pilihan, tgl_str, payload_sheets)
                    
                    st.session_state.temp_upload = {
                        "kode": pt_pilihan,
                        "tgl": tgl_str,
                        "payload": payload_sheets,
                        "res": res
                    }
                except Exception as e:
                    st.error(f"Error: {e}")

    # Display Preview Upload
    if "temp_upload" in st.session_state and st.session_state.temp_upload:
        up = st.session_state.temp_upload
        r = up["res"]
        st.info(f"**{r['total']} peserta di file** | {r['baru']} baru, {r['lewat']} dilewati karena sudah ada, {len(r['cek'])} perlu dicek")
        
        if r['cek']:
            for c in r['cek']:
                st.warning(f"{c['no']}. {c['nama']}: {c['masalah']}")
        
        c_simpan, c_batal = st.columns(2)
        with c_simpan:
            if st.button(f"Simpan {r['baru']} peserta", disabled=(r['baru'] == 0)):
                try:
                    with st.spinner("Menyimpan ke Google Sheets..."):
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
    st.subheader("Kegiatan")
    
    if not st.session_state.kegiatan_list:
        st.caption("Belum ada kegiatan. Upload data mentah untuk memulai.")
    else:
        for g in st.session_state.kegiatan_list:
            col_k1, col_k2 = st.columns([3, 1])
            with col_k1:
                st.write(f"**{g['kode']}** {format_tgl(g['tgl'])}")
            with col_k2:
                status_label = "Draft" if g['status'] == "DRAFT" else "Ditutup"
                if st.button(f"Buka ({status_label})", key=f"keg_{g['id']}"):
                    with st.spinner("Memuat kegiatan..."):
                        st.session_state.active_kegiatan = run_apps_script("getKegiatan", g['id'])
                        st.rerun()

    st.divider()
    if st.button("Reset semua data", type="primary", use_container_width=True):
        if st.checkbox("Saya yakin ingin menghapus SEMUA kegiatan & peserta"):
            run_apps_script("resetSemua")
            load_home_data()
            st.rerun()

# --- 4. HALAMAN DETAIL KEGIATAN & PRESENSI ---
else:
    K = st.session_state.active_kegiatan
    rows = K.get("rows", [])

    col_b1, col_b2 = st.columns([1, 3])
    with col_b1:
        if st.button("⬅️ Kembali"):
            st.session_state.active_kegiatan = None
            load_home_data()
            st.rerun()
    with col_b2:
        status_text = " · Ditutup" if K.get("status") == "DITUTUP" else ""
        st.subheader(f"{K['kode']} · {format_tgl(K['tgl'])}{status_text}")

    # Search Bar
    q = st.text_input("🔍 Cari nama, No RM, NIK, atau nomor TTD", "").strip().upper()

    # Metrics Total / Hadir / Berkas
    sT = len(rows)
    sH = len([p for p in rows if p.get("hadir")])
    sK = len([p for p in rows if p.get("kumpul")])

    m1, m2, m3 = st.columns(3)
    m1.metric("Peserta", sT)
    m2.metric("Hadir", sH)
    m3.metric("Berkas", sK)

    # Filter Peserta (Logic Cocok JS)
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

    # Render List Peserta
    for p in filtered:
        is_draft = K.get("status") == "DRAFT"
        hadir = p.get("hadir")
        kumpul = p.get("kumpul")
        
        # Penentuan Warna Card (Sesuai CSS HTML)
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
            if p.get("cek"):
                st.warning(f"Perlu dicek dengan KTP: {p['cek']}")

            if is_draft:
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    if not hadir:
                        haid_val = ""
                        if p.get("jk") == "P":
                            haid_val = st.radio("Haid/Hamil", ["Tidak ada", "HAID", "HAMIL"], key=f"haid_{p['no']}")
                            haid_val = "" if haid_val == "Tidak ada" else haid_val
                        
                        if st.button("Hadir", key=f"btn_h_{p['no']}"):
                            res_p = run_apps_script("setHadir", K["id"], p["no"], haid_val)
                            p.update(res_p)
                            st.rerun()
                with col_act2:
                    if hadir and not kumpul:
                        if st.button("Sudah mengumpulkan", key=f"btn_k_{p['no']}"):
                            res_p = run_apps_script("setKumpul", K["id"], p["no"])
                            p.update(res_p)
                            st.rerun()
# Penanganan jika Apps Script mengembalikan perbedaan data (pilihBeda)
if "pilih_beda" in st.session_state and st.session_state.pilih_beda:
    pb = st.session_state.pilih_beda
    st.warning(f"⚠️ Ada {pb['jumlah']} perbedaan data olahan dengan file pendaftaran awal.")
    
    col_aw, col_ol = st.columns(2)
    with col_aw:
        if st.button("Pakai Data Awal"):
            res_close = run_apps_script("tutup", K["id"], "AWAL")
            st.session_state.active_kegiatan = res_close
            del st.session_state.pilih_beda
            st.rerun()
    with col_ol:
        if st.button("Pakai Data Olahan"):
            res_close = run_apps_script("tutup", K["id"], "OLAHAN")
            st.session_state.active_kegiatan = res_close
            del st.session_state.pilih_beda
            st.rerun()
    st.divider()

    # Tombol Aksi Bawah (Close MCU / Re-open / Reset)
    if K.get("status") == "DRAFT":
        if st.button("➕ Tambah peserta manual", use_container_width=True):
            st.warning("Gunakan menu tambah manual pada Apps Script.")
if st.button("🔒 Close MCU", type="primary", use_container_width=True):
    try:
        with st.spinner("Menutup MCU & memproses file final..."):
            res_close = run_apps_script("tutup", K["id"])
            
            # Jika ada perbedaan data (perlu konfirmasi AWAL / OLAHAN)
            if isinstance(res_close, dict) and res_close.get("perlu"):
                st.session_state.pilih_beda = res_close
                st.rerun()
            else:
                st.session_state.active_kegiatan = res_close
                st.success(f"MCU {K['kode']} berhasil ditutup!")
                st.rerun()
    except Exception as e:
        st.error(f"Gagal Close MCU dari Apps Script: {e}")
    else:
        if K.get("url"):
            st.markdown(f"[📂 Buka File Final]({K['url']})", unsafe_allow_group=True)
        if st.button("🔓 Buka lagi untuk diedit", use_container_width=True):
            st.session_state.active_kegiatan = run_apps_script("bukaLagi", K["id"])
            st.rerun()

    if st.button("Reset data MCU ini", type="secondary", use_container_width=True):
        if st.checkbox("Hapus seluruh peserta kegiatan ini?"):
            run_apps_script("resetKegiatan", K["id"])
            st.session_state.active_kegiatan = None
            load_home_data()
            st.rerun()
