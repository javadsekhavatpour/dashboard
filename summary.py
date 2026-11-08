#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد خلاصه — آخرین روز همه نمادها
=====================================

این اسکریپت تمام فایل‌های *_dashboard.xlsx در پوشه dashboard/ را می‌خواند
و فقط آخرین سطر (جدیدترین روز) هر نماد را نمایش می‌دهد.

روش اجرا:
  streamlit run summary.py
"""

import streamlit as st
import os
import glob
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
    
    # ردیف ۳ = جدیدترین روز (ردیف ۱ عنوان، ردیف ۲ هدر)
    # ستون‌ها (مطابق make_dashboard.py):
    # B=تاریخ, C=امتیاز جریان پول, D=امتیاز حجم, E=امتیاز قدرت خریدار,
    # F=امتیاز واکنش قیمت, G=امتیاز تداوم, H=امتیاز کل, I=شدت,
    # J=٪ قیمت پایانی (R_c), K=سیگنال
    
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
    
    # پاکسازی مقادیر (حذف ● و تبدیل به عدد)
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
    
    # پاکسازی R_c
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

if not dashboard_files:
    st.warning("هیچ فایل داشبورد یافت نشد. ابتدا با build_all_dashboards.bat داشبورد بسازید.")
    st.stop()

st.info(f"تعداد نمادها: {len(dashboard_files)}")

# خواندن داده‌های همه نمادها
all_data = []
for f in sorted(dashboard_files):
    basename = os.path.basename(f)
    symbol = basename.replace("_dashboard.xlsx", "")
    
    data = read_latest_from_dashboard(f)
    if data:
        data["symbol"] = symbol
        all_data.append(data)

if not all_data:
    st.error("داده‌ای قابل خواندن نبود")
    st.stop()

# ============================================================
# جدول اصلی — دقیقاً مثل داشبورد
# ============================================================
st.markdown("## 📋 جدول خلاصه نمادها")

import pandas as pd

rows = []
for d in all_data:
    rc = d.get("rc", 0)
    rc_str = f"{'+' if rc >= 0 else ''}{rc:.2f}"
    
    # رنگ سیگنال
    signal = d.get("signal", "خنثی")
    if signal in ["خرید قوی", "خرید"]:
        signal_str = f"🟢 {signal}"
    elif signal in ["فروش", "هشدار"]:
        signal_str = f"🔴 {signal}"
    else:
        signal_str = f"⚪ {signal}"
    
    rows.append({
        "نماد": d["symbol"],
        "تاریخ": d["date"],
        "امتیاز جریان پول\n(۰-۳۰)": f"● {d['score_mf']}",
        "امتیاز حجم\n(۰-۱۵)": f"● {d['score_vol']}",
        "امتیاز قدرت خریدار\n(۰-۲۰)": f"● {d['score_bp']}",
        "امتیاز واکنش قیمت\n(۰-۲۵)": f"● {d['score_price']}",
        "امتیاز تداوم\n(۰-۱۰)": f"● {d['score_pers']}",
        "امتیاز کل\n(۰-۱۰۰)": f"● {d['total_score']}",
        "شدت": d["intensity"],
        "٪ قیمت پایانی\n(R_c)": f"● {rc_str}",
        "سیگنال": signal_str,
    })

df = pd.DataFrame(rows)

# نمایش جدول
st.dataframe(df, use_container_width=True, hide_index=True)

# ============================================================
# کارت‌های خلاصه
# ============================================================
st.markdown("---")
st.markdown("## 📈 خلاصه کلی")

total_symbols = len(all_data)
total_score = sum(d["total_score"] for d in all_data)
avg_score = total_score / total_symbols if total_symbols else 0
buy_signals = sum(1 for d in all_data if d["signal"] in ["خرید قوی", "خرید"])
sell_signals = sum(1 for d in all_data if d["signal"] in ["فروش", "هشدار"])

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
