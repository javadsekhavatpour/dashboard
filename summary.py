#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد خلاصه — آخرین روز همه نمادها
=====================================

این اسکریپت تمام فایل‌های *_dashboard.xlsx در پوشه dashboard/ را می‌خواند
و فقط آخرین سطر (جدیدترین روز) هر نماد را نمایش می‌دهد.

همچنین برای هر نماد دو دکمه دارد:
  🔄 بروزرسانی: دریافت داده‌های جدید از TSETMC و بازسازی داشبورد
  📊 تاریخچه: نمایش داشبورد کامل تاریخی نماد

روش اجرا:
  streamlit run summary.py
"""

import streamlit as st
import os
import sys
import glob
import subprocess
from datetime import datetime

st.set_page_config(page_title="خلاصه نمادها", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main > div { direction: rtl; max-width: 1400px; margin: 0 auto; }
    h1, h2, h3 { text-align: right; }
    .stDataFrame { direction: rtl; }
    .stMetric { text-align: center; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# خواندن آخرین سطر از فایل داشبورد
# ============================================================
def read_latest_from_dashboard(dashboard_file):
    """خواندن آخرین سطر (ردیف ۳) از شیت داشبورد فایل *_dashboard.xlsx"""
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(dashboard_file, data_only=True)
    except:
        return None
    
    if "داشبورد" not in wb.sheetnames:
        return None
    
    ws = wb["داشبورد"]
    
    data = {
        "date": ws.cell(row=3, column=2).value or "—",
        "score_mf": ws.cell(row=3, column=3).value,
        "score_vol": ws.cell(row=3, column=4).value,
        "score_bp": ws.cell(row=3, column=5).value,
        "score_price": ws.cell(row=3, column=6).value,
        "score_pers": ws.cell(row=3, column=7).value,
        "total_score": ws.cell(row=3, column=8).value,
        "intensity": ws.cell(row=3, column=9).value or "—",
        "rc": ws.cell(row=3, column=10).value,
        "signal": ws.cell(row=3, column=11).value or "—",
    }
    
    # پاکسازی مقادیر
    for key in ["score_mf", "score_vol", "score_bp", "score_price", "score_pers", "total_score"]:
        val = data[key]
        if val is not None:
            if isinstance(val, str):
                val = val.replace("●", "").replace("●", "").strip()
                try:
                    val = int(float(val))
                except:
                    val = 0
            data[key] = val
        else:
            data[key] = 0
    
    rc_val = data["rc"]
    if rc_val is not None:
        if isinstance(rc_val, str):
            rc_val = rc_val.replace("●", "").replace("●", "").replace("+", "").strip()
            try:
                rc_val = float(rc_val)
            except:
                rc_val = 0.0
        data["rc"] = rc_val
    else:
        data["rc"] = 0.0
    
    return data

# ============================================================
# بروزرسانی نماد
# ============================================================
def update_symbol(symbol):
    """بروزرسانی فایل input و بازسازی داشبورد برای یک نماد"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # ۱. آپدیت فایل input
    update_script = os.path.join(script_dir, "update_input.py")
    input_file = os.path.join("input", f"{symbol}_input.xlsx")
    
    if os.path.exists(input_file):
        # فایل موجود است، آپدیت کن
        cmd = [sys.executable, update_script, "--symbol", symbol, "--update"]
    else:
        # فایل موجود نیست، ایجاد کن
        cmd = [sys.executable, update_script, "--symbol", symbol, "--create"]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, cwd=script_dir)
        if result.returncode != 0:
            return False, result.stderr[-500:] if result.stderr else "Unknown error"
    except Exception as e:
        return False, str(e)
    
    # ۲. بازسازی داشبورد
    dash_script = os.path.join(script_dir, "make_dashboard.py")
    dash_dir = os.path.join(script_dir, "dashboard")
    os.makedirs(dash_dir, exist_ok=True)
    dash_file = os.path.join(dash_dir, f"{symbol}_dashboard.xlsx")
    
    try:
        result = subprocess.run(
            [sys.executable, dash_script, input_file, "--name", symbol, "--output", dash_file],
            capture_output=True, text=True, timeout=120, cwd=script_dir
        )
        return result.returncode == 0, result.stdout[-500:] if result.returncode == 0 else result.stderr[-500:]
    except Exception as e:
        return False, str(e)

# ============================================================
# نمایش داشبورد تاریخی نماد
# ============================================================
def show_symbol_history(symbol):
    """نمایش داشبورد کامل تاریخی یک نماد"""
    from openpyxl import load_workbook
    import pandas as pd
    
    dashboard_file = os.path.join("dashboard", f"{symbol}_dashboard.xlsx")
    if not os.path.exists(dashboard_file):
        dashboard_file = f"{symbol}_dashboard.xlsx"
    
    if not os.path.exists(dashboard_file):
        st.warning(f"فایل {symbol}_dashboard.xlsx یافت نشد")
        return
    
    try:
        wb = load_workbook(dashboard_file, data_only=True)
    except:
        st.error("خطا در خواندن فایل")
        return
    
    if "داشبورد" not in wb.sheetnames:
        st.warning("شیت داشبورد یافت نشد")
        return
    
    ws = wb["داشبورد"]
    st.markdown(f"### {ws.cell(row=1, column=2).value}")
    
    data = []
    for r in range(3, ws.max_row + 1):
        d = ws.cell(row=r, column=2).value
        if not d or d == "میانگین": continue
        
        def clean(val):
            if val is None: return "—"
            if isinstance(val, str):
                return val.replace("●", "").replace("●", "").strip()
            return val
        
        data.append({
            "تاریخ": d,
            "جریان پول": clean(ws.cell(row=r, column=3).value),
            "حجم": clean(ws.cell(row=r, column=4).value),
            "قدرت خریدار": clean(ws.cell(row=r, column=5).value),
            "واکنش قیمت": clean(ws.cell(row=r, column=6).value),
            "تداوم": clean(ws.cell(row=r, column=7).value),
            "کل": clean(ws.cell(row=r, column=8).value),
            "شدت": clean(ws.cell(row=r, column=9).value),
            "R_c": clean(ws.cell(row=r, column=10).value),
            "سیگنال": clean(ws.cell(row=r, column=11).value),
        })
    
    if not data:
        st.warning("داده‌ای وجود ندارد")
        return
    
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True, hide_index=True)

# ============================================================
# رابط کاربری
# ============================================================

st.markdown("""
<div style='background: linear-gradient(135deg, #1B2A4A 0%, #2C3E50 100%);
            padding: 25px; border-radius: 12px; margin-bottom: 20px; text-align: center;'>
    <h1 style='color: white; margin: 0; font-size: 24px;'>📊 خلاصه نمادها — آخرین روز</h1>
    <p style='color: #E8ECF4; margin: 8px 0 0 0; font-size: 14px;'>
        نمایش جدیدترین سطر از فایل‌های داشبورد
    </p>
</div>
""", unsafe_allow_html=True)

# پیدا کردن همه فایل‌های dashboard
dashboard_dir = "dashboard"
if not os.path.exists(dashboard_dir):
    dashboard_files = glob.glob("*_dashboard.xlsx")
else:
    dashboard_files = glob.glob(os.path.join(dashboard_dir, "*_dashboard.xlsx"))

# همچنین فایل‌های input را هم پیدا کن (برای نمادهایی که هنوز داشبورد ندارند)
input_dir = "input"
if os.path.exists(input_dir):
    input_files = glob.glob(os.path.join(input_dir, "*_input.xlsx"))
else:
    input_files = glob.glob("*_input.xlsx")

# لیست همه نمادها (از داشبورد و input)
all_symbols = set()
for f in dashboard_files:
    sym = os.path.basename(f).replace("_dashboard.xlsx", "")
    all_symbols.add(sym)
for f in input_files:
    sym = os.path.basename(f).replace("_input.xlsx", "")
    all_symbols.add(sym)

all_symbols = sorted(all_symbols)

if not all_symbols:
    st.warning("هیچ نمادی یافت نشد. ابتدا فایل input ایجاد کنید.")
    st.stop()

st.info(f"تعداد نمادها: {len(all_symbols)}")

# ============================================================
# جدول اصلی با دکمه‌های عملیات
# ============================================================
st.markdown("## 📋 جدول خلاصه نمادها")

import pandas as pd

# هدر جدول
cols_hdr = st.columns([1, 1.5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1.5])
with cols_hdr[0]: st.markdown("**نماد**")
with cols_hdr[1]: st.markdown("**تاریخ**")
with cols_hdr[2]: st.markdown("**جریان پول**\n(۰-۳۰)")
with cols_hdr[3]: st.markdown("**حجم**\n(۰-۱۵)")
with cols_hdr[4]: st.markdown("**قدرت خریدار**\n(۰-۲۰)")
with cols_hdr[5]: st.markdown("**واکنش قیمت**\n(۰-۲۵)")
with cols_hdr[6]: st.markdown("**تداوم**\n(۰-۱۰)")
with cols_hdr[7]: st.markdown("**کل**\n(۰-۱۰۰)")
with cols_hdr[8]: st.markdown("**شدت**")
with cols_hdr[9]: st.markdown("**R_c**")
with cols_hdr[10]: st.markdown("**سیگنال**")
with cols_hdr[11]: st.markdown("**🔄 آپدیت**")
with cols_hdr[12]: st.markdown("**📊 تاریخچه**")
st.markdown("---")

# ردیف‌های جدول
for symbol in all_symbols:
    # خواندن داده‌های داشبورد
    dash_file = os.path.join(dashboard_dir, f"{symbol}_dashboard.xlsx")
    if not os.path.exists(dash_file):
        dash_file = f"{symbol}_dashboard.xlsx"
    
    data = None
    if os.path.exists(dash_file):
        data = read_latest_from_dashboard(dash_file)
    
    cols = st.columns([1, 1.5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1.5])
    
    with cols[0]:
        st.markdown(f"**{symbol}**")
    
    with cols[1]:
        if data:
            st.markdown(data["date"])
        else:
            st.markdown("—")
    
    with cols[2]:
        if data:
            st.markdown(f"● {data['score_mf']}")
        else:
            st.markdown("—")
    
    with cols[3]:
        if data:
            st.markdown(f"● {data['score_vol']}")
        else:
            st.markdown("—")
    
    with cols[4]:
        if data:
            st.markdown(f"● {data['score_bp']}")
        else:
            st.markdown("—")
    
    with cols[5]:
        if data:
            st.markdown(f"● {data['score_price']}")
        else:
            st.markdown("—")
    
    with cols[6]:
        if data:
            st.markdown(f"● {data['score_pers']}")
        else:
            st.markdown("—")
    
    with cols[7]:
        if data:
            st.markdown(f"● {data['total_score']}")
        else:
            st.markdown("—")
    
    with cols[8]:
        if data:
            st.markdown(data["intensity"])
        else:
            st.markdown("—")
    
    with cols[9]:
        if data:
            rc = data["rc"]
            st.markdown(f"{'+' if rc >= 0 else ''}{rc:.2f}")
        else:
            st.markdown("—")
    
    with cols[10]:
        if data:
            signal = data["signal"]
            if signal in ["خرید قوی", "خرید"]:
                st.markdown(f"🟢 {signal}")
            elif signal in ["فروش", "هشدار"]:
                st.markdown(f"🔴 {signal}")
            else:
                st.markdown(f"⚪ {signal}")
        else:
            st.markdown("—")
    
    with cols[11]:
        if st.button("🔄", key=f"update_{symbol}", help=f"بروزرسانی {symbol}"):
            with st.spinner(f"در حال بروزرسانی {symbol}..."):
                success, msg = update_symbol(symbol)
            if success:
                st.success(f"✅ {symbol} بروزرسانی شد")
                st.rerun()
            else:
                st.error(f"❌ خطا: {msg}")
    
    with cols[12]:
        if st.button("📊", key=f"history_{symbol}", help=f"داشبورد {symbol}"):
            st.session_state["show_history"] = symbol
    
    st.markdown("")

# ============================================================
# نمایش داشبورد تاریخی
# ============================================================
if "show_history" in st.session_state and st.session_state["show_history"]:
    symbol = st.session_state["show_history"]
    st.markdown("---")
    st.markdown(f"## 📊 داشبورد تاریخی {symbol}")
    show_symbol_history(symbol)
    if st.button("✖ بستن", key="close_history"):
        st.session_state["show_history"] = None
        st.rerun()

# ============================================================
# کارت‌های خلاصه
# ============================================================
st.markdown("---")
st.markdown("## 📈 خلاصه کلی")

# خواندن همه داده‌ها برای محاسبه خلاصه
all_data = []
for symbol in all_symbols:
    dash_file = os.path.join(dashboard_dir, f"{symbol}_dashboard.xlsx")
    if not os.path.exists(dash_file):
        dash_file = f"{symbol}_dashboard.xlsx"
    if os.path.exists(dash_file):
        data = read_latest_from_dashboard(dash_file)
        if data:
            data["symbol"] = symbol
            all_data.append(data)

total_symbols = len(all_data)
if total_symbols > 0:
    total_score = sum(d["total_score"] for d in all_data)
    avg_score = total_score / total_symbols
    buy_signals = sum(1 for d in all_data if d["signal"] in ["خرید قوی", "خرید"])
    sell_signals = sum(1 for d in all_data if d["signal"] in ["فروش", "هشدار"])
else:
    avg_score = 0
    buy_signals = 0
    sell_signals = 0

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("تعداد نمادها", total_symbols)

with col2:
    st.metric("میانگین امتیاز", f"{avg_score:.1f}")

with col3:
    st.metric("سیگنال خرید", f"🟢 {buy_signals}")

with col4:
    st.metric("سیگنال فروش/هشدار", f"🔴 {sell_signals}")

# ============================================================
# فوتر
# ============================================================
now = datetime.now()
st.markdown(f"""
<div style='text-align: center; color: #8C8A84; font-size: 12px; margin-top: 30px;'>
    خلاصه نمادها | {now.strftime('%Y-%m-%d %H:%M')}
</div>
""", unsafe_allow_html=True)
