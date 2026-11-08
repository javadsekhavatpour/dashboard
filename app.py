#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TSE Screener - Streamlit Dashboard
==================================
Run: streamlit run app.py
"""

import streamlit as st
import os
import sys
import subprocess
from datetime import datetime

# ============================================================
# Page config
# ============================================================
st.set_page_config(
    page_title="TSE Screener",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .main > div { direction: rtl; max-width: 1200px; margin: 0 auto; }
    h1, h2, h3 { text-align: right; }
    .stTextInput > div > div > input { text-align: center; }
    .stButton > button { font-family: Tahoma; font-size: 16px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Helper functions
# ============================================================
def gregorian_to_jalali(gy, gm, gd):
    g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334]
    gy2 = gy - 1600; gm2 = gm - 1; gd2 = gd - 1
    g_day_no = 365*gy2 + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 + g_d_m[gm2] + gd2
    g_day_no -= 79
    if g_day_no < 0: jy = 0; g_day_no = g_day_no + 120356
    else: jy = 979 + 33*(g_day_no // 12053); g_day_no = g_day_no % 12053
    jy += 4*(g_day_no // 1461); g_day_no = g_day_no % 1461
    if g_day_no > 365: jy += (g_day_no-1)//365; g_day_no = (g_day_no-1) % 365
    if g_day_no < 186: jm = 1 + g_day_no//31; jd = 1 + (g_day_no % 31)
    else: jm = 7 + (g_day_no-186)//30; jd = 1 + ((g_day_no-186) % 30)
    return jy, jm, jd

J_MONTHS = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور",
            "مهر","آبان","آذر","دی","بهمن","اسفند"]

def now_jalali():
    now = datetime.now()
    jy, jm, jd = gregorian_to_jalali(now.year, now.month, now.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

def now_jalali_fa():
    now = datetime.now()
    jy, jm, jd = gregorian_to_jalali(now.year, now.month, now.day)
    return f"{J_MONTHS[jm-1]} {jd}، {jy}"

@st.cache_data(show_spinner=False, ttl=300)
def find_ins_code(symbol_name):
    import requests
    from urllib.parse import quote
    try:
        encoded = quote(symbol_name)
        url = f"https://cdn.tsetmc.com/api/Instrument/GetInstrumentSearch/{encoded}"
        r = requests.get(url, headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
            "Referer": "https://old.tsetmc.com/"
        }, timeout=15)
        r.raise_for_status()
        data = r.json()
        candidates = data.get("instrumentSearch", [])
        if not candidates: return None, None
        for item in candidates:
            if item.get("lVal18AFC", "").strip() == symbol_name:
                return str(item["insCode"]), item.get("lVal30","").strip()
        item = candidates[0]
        return str(item["insCode"]), item.get("lVal30","").strip()
    except: return None, None

def get_last_update_date(symbol):
    # Check in input/ subfolder first, then current dir
    input_file = os.path.join("input", f"{symbol}_input.xlsx")
    if not os.path.exists(input_file):
        input_file = f"{symbol}_input.xlsx"
    if not os.path.exists(input_file): return None
    try:
        from openpyxl import load_workbook
        wb = load_workbook(input_file, data_only=True)
        for sn in wb.sheetnames:
            if "قیمت" in sn and "سابقه" in sn:
                ws = wb[sn]
                d = ws.cell(row=3, column=1).value
                if d: return str(d)
        return None
    except: return None

def run_operation(symbol, mode, days=60, build_dash=True):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_dir, "update_input.py")
    if not os.path.exists(script_path): return False, "update_input.py not found"
    cmd = [sys.executable, script_path, "--symbol", symbol, f"--{mode}", "--days", str(days)]
    if build_dash: cmd.append("--dashboard")
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300, cwd=script_dir)
        return result.returncode == 0, result.stdout[-2000:] + "\n" + result.stderr[-500:]
    except subprocess.TimeoutExpired: return False, "timeout"
    except Exception as e: return False, str(e)

def show_stock_dashboard(symbol):
    # Check in dashboard/ subfolder first, then current dir
    dashboard_file = os.path.join("dashboard", f"{symbol}_dashboard.xlsx")
    if not os.path.exists(dashboard_file):
        dashboard_file = f"{symbol}_dashboard.xlsx"
    if not os.path.exists(dashboard_file):
        st.warning(f"File {symbol}_dashboard.xlsx not found")
        return
    try:
        from openpyxl import load_workbook
        import pandas as pd
        wb = load_workbook(dashboard_file, data_only=True)
        if "داشبورد" not in wb.sheetnames:
            st.warning("Dashboard sheet not found")
            return
        ws = wb["داشبورد"]
        st.markdown(f"### {ws.cell(row=1, column=2).value}")
        data = []
        for r in range(3, ws.max_row + 1):
            d = ws.cell(row=r, column=2).value
            if not d or d == "میانگین": continue
            data.append({
                "تاریخ": d,
                "جریان پول": str(ws.cell(row=r, column=3).value or ""),
                "حجم": ws.cell(row=r, column=4).value,
                "قدرت": ws.cell(row=r, column=5).value,
                "قیمت": ws.cell(row=r, column=6).value,
                "تداوم": ws.cell(row=r, column=7).value,
                "کل": ws.cell(row=r, column=8).value,
                "شدت": ws.cell(row=r, column=9).value,
                "R_c": str(ws.cell(row=r, column=10).value or ""),
                "سیگنال": ws.cell(row=r, column=11).value,
            })
        if not data:
            st.warning("No data")
            return
        df = pd.DataFrame(data)
        latest = data[0]
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("امتیاز کل", latest["کل"])
        with col2: st.metric("سیگنال", latest["سیگنال"])
        with col3: st.metric("شدت", latest["شدت"])
        with col4: st.metric("R_c", latest["R_c"])
        st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error: {e}")

# ============================================================
# State
# ============================================================
if 'stock_list' not in st.session_state:
    st.session_state.stock_list = [""] * 15
if 'last_updates' not in st.session_state:
    st.session_state.last_updates = {}

# ============================================================
# Header
# ============================================================
st.markdown("""
<div style='background: linear-gradient(135deg, #1B2A4A 0%, #2C3E50 100%);
            padding: 25px; border-radius: 12px; margin-bottom: 20px; text-align: center;'>
    <h1 style='color: white; margin: 0; font-size: 24px;'>📊 سامانه غربالگری سهام بورس تهران</h1>
    <p style='color: #E8ECF4; margin: 8px 0 0 0; font-size: 14px;'>
        ایجاد و به‌روزرسانی فایل input + ساخت داشبورد تحلیلی
    </p>
</div>
""", unsafe_allow_html=True)

# Market status
now = datetime.now()
wd = now.weekday()
if wd == 4: market_status = "🔴 بازار بسته (جمعه)"
elif wd == 3: market_status = "🔴 بازار بسته (پنجشنبه)"
elif 8 <= now.hour < 13: market_status = f"🟢 بازار باز — {now.strftime('%H:%M:%S')}"
else: market_status = "🟡 بازار بسته"

st.markdown(f"""
<div style='background: #F7F7F5; padding: 12px; border-radius: 8px; margin-bottom: 15px; text-align: center;'>
    <strong>{market_status}</strong> | امروز: {now_jalali_fa()}
</div>
""", unsafe_allow_html=True)

# ============================================================
# Stock table
# ============================================================
st.markdown("## 📋 جدول نمادها")

# Header row
cols_hdr = st.columns([0.5, 2, 1, 1, 2, 1.5])
with cols_hdr[0]: st.markdown("**#**")
with cols_hdr[1]: st.markdown("**نماد**")
with cols_hdr[2]: st.markdown("**ایجاد ⬇️**")
with cols_hdr[3]: st.markdown("**آپدیت 🌐**")
with cols_hdr[4]: st.markdown("**آخرین بروزرسانی**")
with cols_hdr[5]: st.markdown("**وضعیت**")
st.markdown("---")

# Data rows
for i in range(15):
    col_num, col_sym, col_create, col_update, col_date, col_status = st.columns([0.5, 2, 1, 1, 2, 1.5])
    
    with col_num:
        st.markdown(f"<div style='text-align:center; padding-top:10px;'><b>{i+1}</b></div>", unsafe_allow_html=True)
    
    with col_sym:
        symbol = st.text_input(f"نماد {i+1}", value=st.session_state.stock_list[i],
                              key=f"sym_{i}", placeholder="—", label_visibility="collapsed")
        st.session_state.stock_list[i] = symbol.strip()
    
    with col_create:
        create_btn = st.button("⬇️", key=f"create_{i}")
    
    with col_update:
        update_btn = st.button("🌐", key=f"update_{i}")
    
    with col_date:
        last_date = None
        if symbol.strip():
            last_date = get_last_update_date(symbol.strip())
        if last_date:
            st.markdown(f"<div style='text-align:center; padding-top:8px; color:#1B7D46; font-size:13px; font-weight:bold;'>✓ {last_date}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align:center; padding-top:8px; color:#8C8A84; font-size:13px;'>—</div>", unsafe_allow_html=True)
    
    with col_status:
        if symbol.strip():
            ins_code, _ = find_ins_code(symbol.strip())
            if ins_code:
                st.markdown(f"<div style='text-align:center; padding-top:8px; color:#1B7D46; font-size:12px;'>✓ یافت شد</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='text-align:center; padding-top:8px; color:#C0392B; font-size:12px;'>❌ یافت نشد</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align:center; padding-top:8px; color:#8C8A84; font-size:12px;'>—</div>", unsafe_allow_html=True)
    
    # Process buttons
    if create_btn and symbol.strip():
        with st.spinner(f"Creating {symbol}..."):
            success, output = run_operation(symbol.strip(), "create")
        if success:
            st.success(f"✅ {symbol}_input.xlsx created")
            st.session_state.last_updates[symbol.strip()] = now_jalali()
            st.rerun()
        else:
            st.error(f"❌ Error: {symbol}")
            with st.expander("Details"): st.text(output)
    
    if update_btn and symbol.strip():
        with st.spinner(f"Updating {symbol}..."):
            success, output = run_operation(symbol.strip(), "update")
        if success:
            st.success(f"✅ {symbol} updated")
            st.session_state.last_updates[symbol.strip()] = now_jalali()
            st.rerun()
        else:
            st.error(f"❌ Error: {symbol}")
            with st.expander("Details"): st.text(output)

# ============================================================
# Batch operations
# ============================================================
st.markdown("---")
st.markdown("## 🚀 اجرای گروهی")

col_left, col_right, col_center = st.columns([2, 2, 1])

with col_left:
    group_mode = st.radio("نوع عملیات:", options=["create", "update"],
        format_func=lambda x: "🆕 ایجاد فایل (۶۰ روز)" if x == "create" else "🔄 به‌روزرسانی",
        horizontal=True)

with col_right:
    group_days = st.slider("تعداد روزها", min_value=10, max_value=120, value=60, step=10)
    group_dash = st.checkbox("ساخت داشبورد", value=True)

with col_center:
    st.markdown("<div style='padding-top:25px;'></div>", unsafe_allow_html=True)
    run_all = st.button("▶️ اجرا", type="primary", use_container_width=True)

valid_symbols = [s for s in st.session_state.stock_list if s.strip()]

if run_all:
    if not valid_symbols:
        st.warning("هیچ نمادی وارد نشده است!")
    else:
        st.markdown(f"**نمادها:** {', '.join(valid_symbols)}")
        progress = st.progress(0)
        status_text = st.empty()
        results = []
        
        for i, sym in enumerate(valid_symbols):
            progress.progress(i / len(valid_symbols))
            status_text.text(f"Processing {sym} ({i+1}/{len(valid_symbols)})...")
            success, output = run_operation(sym.strip(), group_mode, group_days, group_dash)
            results.append((sym, success))
            if success: st.session_state.last_updates[sym.strip()] = now_jalali()
        
        progress.progress(1.0)
        status_text.empty()
        
        st.markdown("### 📊 نتایج")
        success_count = sum(1 for _, s in results if s)
        fail_count = len(results) - success_count
        
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("کل", len(results))
        with col2: st.metric("موفق", success_count)
        with col3: st.metric("ناموفق", fail_count)
        
        for sym, success in results:
            if success: st.markdown(f"✅ **{sym}** — موفق")
            else: st.markdown(f"❌ **{sym}** — ناموفق")
        
        if group_dash and success_count > 0:
            st.markdown("---")
            st.markdown("## 📊 داشبوردها")
            successful = [s for s, ok in results if ok]
            if len(successful) == 1:
                show_stock_dashboard(successful[0])
            elif len(successful) > 1:
                tabs = st.tabs(successful)
                for i, sym in enumerate(successful):
                    with tabs[i]: show_stock_dashboard(sym)

# ============================================================
# Footer
# ============================================================
st.markdown(f"""
<div style='text-align: center; color: #8C8A84; font-size: 12px; margin-top: 30px;'>
    TSE Screener v2.1 | {now.strftime('%Y-%m-%d %H:%M')}
</div>
""", unsafe_allow_html=True)
