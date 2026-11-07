import streamlit as st
import pandas as pd
import json
import re
from datetime import datetime

# Setup Halaman
st.set_page_config(page_title="MCU Registrasi", page_icon="🏥", layout="centered")

# --- INITIALIZATION / STATE (Menggantikan localStorage & State JS) ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "kegiatan_list" not in st.session_state:
    st.session_state.kegiatan_list = []
if "active_keg_id" not in st.session_state:
    st.session_state.active_keg_id = None
if "peserta_db" not in st.session_state:
    # Key: keg_id, Value: List of Dict Peserta
    st.session_state.peserta_db = {}

# Daftar Perusahaan (Menggantikan sheet PENGATURAN)
DAFTAR_PT = ["PT A", "PT B", "PT C", "MCU KLINIK JEPARA"]

# --- HELPER FUNCTIONS ---
def get_usia(tgl_lahir):
    if not tgl_lahir:
        return "-"
    try:
        born = datetime.strptime(str(tgl_lahir), "%Y-%m-%d")
        today = datetime.today()
        return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
    except Exception:
        return "-"

# --- 1. HALAMAN LOGIN ---
if not st.session_state.logged_in:
    st.title("🏥 MCU Registrasi")
    with st.form("login_form"):
        st.subheader("Masuk Server / System")
        server_url = st.text_input("Alamat Server (URL App)", value="https://script.google.com/macros/s/.../exec")
        pin = st.text_input("PIN", type="password")
        submit_login = st.form_submit_button("Masuk")

        if submit_login:
            if pin == "1234":  # Ganti dengan logika PIN Anda
                st.session_state.logged_in = True
                st.success("Berhasil masuk!")
                st.rerun()
            else:
                st.error("PIN salah, coba lagi.")
    st.stop()

# --- 2. NAVIGASI ATAS & LOGOUT ---
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.title("🏥 MCU Registrasi")
with col_head2:
    if st.button("Keluar"):
        st.session_state.logged_in = False
        st.session_state.active_keg_id = None
        st.rerun()

# --- 3. HALAMAN UTAMA (HOME / DASHBOARD KEGIATAN) ---
if st.session_state.active_keg_id is None:
    st.subheader("Mulai Kegiatan atau Tambah Peserta")
    
    with st.expander("➕ Upload File Excel Mentah Baru", expanded=True):
        pt_pilihan = st.selectbox("Pilih Perusahaan", DAFTAR_PT)
        tgl_kegiatan = st.date_input("Tanggal Kegiatan", datetime.now())
        uploaded_file = st.file_uploader("Unggah File Excel Mentah (.xlsx, .xls)", type=["xlsx", "xls"])
        
        if uploaded_file and st.button("Periksa & Import File"):
            try:
                # Membaca Sheet Excel (Menggantikan library SheetJS / XLSX JS)
                excel_data = pd.read_excel(uploaded_file, sheet_name=None)
                total_peserta = 0
                new_rows = []
                
                # Header normalisasi
                for sheet_name, df in excel_data.items():
                    if re.match(r"^(PENDAFTARAN|HADIR|TIDAK HADIR|DATA MENTAH)$", sheet_name, re.I):
                        continue
                    for idx, row in df.iterrows():
                        total_peserta += 1
                        new_rows.append({
                            "no": len(new_rows) + 1,
                            "nama": str(row.get("NAMA", row.get("Nama", f"Peserta {total_peserta}"))),
                            "rm": str(row.get("NO RM", row.get("RM", f"RM-{total_peserta:04d}"))),
                            "nik": str(row.get("NIK", "")),
                            "tgl": str(row.get("TGL LAHIR", "")),
                            "jk": str(row.get("JK", "L")),
                            "hadir": False,
                            "kumpul": False,
                            "haid": "",
                            "cek": ""
                        })
                
                # Simpan kegiatan baru
                keg_id = f"{pt_pilihan}_{tgl_kegiatan.strftime('%Y%m%d')}"
                st.session_state.kegiatan_list.append({
                    "id": keg_id,
                    "kode": pt_pilihan,
                    "tgl": str(tgl_kegiatan),
                    "status": "DRAFT"
                })
                st.session_state.peserta_db[keg_id] = new_rows
                st.success(f"Berhasil mengunggah {len(new_rows)} peserta!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal membaca file Excel: {e}")

    st.divider()
    st.subheader("📋 Daftar Kegiatan MCU")
    if not st.session_state.kegiatan_list:
        st.info("Belum ada kegiatan. Upload data mentah untuk memulai.")
    else:
        for keg in st.session_state.kegiatan_list:
            col_a, col_b, col_c = st.columns([3, 1, 1])
            with col_a:
                st.write(f"**{keg['kode']}** · {keg['tgl']}")
            with col_b:
                st.caption(f"Status: {keg['status']}")
            with col_c:
                if st.button("Buka", key=f"open_{keg['id']}"):
                    st.session_state.active_keg_id = keg['id']
                    st.rerun()

    if st.button("Reset Semua Data", type="primary"):
        st.session_state.kegiatan_list = []
        st.session_state.peserta_db = {}
        st.rerun()

# --- 4. HALAMAN DETAIL KEGIATAN & PRESENSI ---
else:
    keg_id = st.session_state.active_keg_id
    keg_info = next((k for k in st.session_state.kegiatan_list if k["id"] == keg_id), None)
    peserta_list = st.session_state.peserta_db.get(keg_id, [])

    if st.button("⬅️ Kembali ke Daftar Kegiatan"):
        st.session_state.active_keg_id = None
        st.rerun()

    st.subheader(f"{keg_info['kode']} · {keg_info['tgl']} ({keg_info['status']})")

    # Metrics Stat (Sesuai elemen stat di HTML)
    total_p = len(peserta_list)
    hadir_p = len([p for p in peserta_list if p["hadir"]])
    kumpul_p = len([p for p in peserta_list if p["kumpul"]])

    m1, m2, m3 = st.columns(3)
    m1.metric("Peserta", total_p)
    m2.metric("Hadir", hadir_p)
    m3.metric("Berkas Terkumpul", kumpul_p)

    # Search Filter
    query = st.text_input("🔍 Cari Nama, No RM, NIK, atau No. TTD", "")

    # Filter Peserta
    filtered_peserta = peserta_list
    if query:
        q = query.strip().upper()
        filtered_peserta = [
            p for p in peserta_list
            if q in str(p["no"]) or q in p["nama"].upper() or q in p["rm"].upper() or q in str(p["nik"])
        ]

    st.divider()

    # Tampilan Daftar Peserta (Cards)
    for p in filtered_peserta:
        status_str = "Berkas ✓" if p["kumpul"] else ("Hadir" if p["hadir"] else "Belum Hadir")
        bg_color = "🟢" if p["kumpul"] else ("🟡" if p["hadir"] else "⚪")
        
        with st.expander(f"{bg_color} {p['no']}. {p['nama']} ({p['rm']}) — {status_str}"):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.write(f"**No RM:** {p['rm']}")
                st.write(f"**NIK:** {p['nik']}")
                st.write(f"**Tgl Lahir:** {p['tgl']} (Usia: {get_usia(p['tgl'])})")
            with col_d2:
                st.write(f"**Jenis Kelamin:** {'Laki-laki' if p['jk'] == 'L' else 'Perempuan'}")
                if p["jk"] == "P":
                    st.write(f"**Haid/Hamil:** {p['haid'] if p['haid'] else 'Tidak ada'}")

            # Tombol Aksi Presensi per Peserta
            if keg_info["status"] == "DRAFT":
                c1, c2, c3 = st.columns(3)
                with c1:
                    if not p["hadir"]:
                        if p["jk"] == "P":
                            haid_opt = st.radio("Status Haid/Hamil", ["Tidak ada", "HAID", "HAMIL"], key=f"h_{p['no']}")
                        if st.button("Set Hadir", key=f"btn_h_{p['no']}"):
                            p["hadir"] = True
                            if p["jk"] == "P":
                                p["haid"] = "" if haid_opt == "Tidak ada" else haid_opt
                            st.rerun()
                with c2:
                    if p["hadir"] and not p["kumpul"]:
                        if st.button("Berkas Terkumpul", key=f"btn_k_{p['no']}"):
                            p["kumpul"] = True
                            st.rerun()

    st.divider()

    # Tombol Aksi Kegiatan (Close MCU / Re-open)
    if keg_info["status"] == "DRAFT":
        if st.button("➕ Tambah Peserta Manual", use_container_width=True):
            new_no = len(peserta_list) + 1
            peserta_list.append({
                "no": new_no,
                "nama": f"Peserta Manual {new_no}",
                "rm": f"RM-M-{new_no:03d}",
                "nik": "",
                "tgl": "",
                "jk": "L",
                "hadir": True,
                "kumpul": False,
                "haid": "",
                "cek": ""
            })
            st.rerun()

        if st.button("🔒 Close MCU", type="primary", use_container_width=True):
            keg_info["status"] = "DITUTUP"
            st.success("Kegiatan MCU telah ditutup!")
            st.rerun()
    else:
        if st.button("🔓 Buka Lagi MCU untuk Diedit", use_container_width=True):
            keg_info["status"] = "DRAFT"
            st.rerun()