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

DEFAULT_URL = "https://script.google.com/macros/s/AKfycbxvKfpvDrlunxxrDxV2VvOKBFzDxDprL8mFvsucx0PgM6-lnSRDE5TRKSjOemSG85HwNw/exec"

# --- CSS PRESISI SUPER RAPAT KHUSUS LAYAR HP ---
st.markdown("""
<style>
/* Reset Padding Layar HP */
.block-container {
    padding-top: 0.3rem !important;
    padding-bottom: 1rem !important;
    padding-left: 0.2rem !important;
    padding-right: 0.2rem !important;
}

/* Hilangkan Jarak Kosong Antar Komponen */
div[data-testid="stVerticalBlock"] > div {
    gap: 0.15rem !important;
    margin-bottom: 0px !important;
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
    color: #ffffff;
}
.stat-num {
    font-weight: bold;
    font-size: 1.1rem;
}

/* STYLING POPOVER (NAMA PESERTA ADALAH TOMBOL UTAMA) */
div[data-testid="stPopover"] {
    width: 100% !important;
    margin-bottom: 2px !important;
}
div[data-testid="stPopover"] > button {
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 10px 12px !important;
    border-radius: 6px !important;
    font-weight: bold !important;
    font-size: 0.9rem !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.2);
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
        yy_nik = int(nik
