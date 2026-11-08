#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد خلاصه — آخرین روز همه نمادها
=====================================

این اسکریپت تمام فایل‌های *_input.xlsx در پوشه input/ را می‌خواند
و آخرین روز هر نماد را در یک جدول نمایش می‌دهد.

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
    .main > div { direction: rtl; max-width: 1200px; margin: 0 auto; }
    h1, h2, h3 { text-align: right; }
    .stDataFrame { direction: rtl; }
    .stMetric { text-align: center; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# تبدیل تاریخ
# ============================================================
def gregorian_to_jalali(gy, gm, gd):
    g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334]
    gy2 = gy - 1600; gm2 = gm - 1; gd2 = gd - 1
    g_day_no = 365*gy2 + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 + g_d_m[gm2] + gd2
    g_day_no -= 79
    if g_day_no < 0:
        jy = 0; g_day_no = g_day_no + 120356
    else:
        jy = 979 + 33*(g_day_no // 12053); g_day_no = g_day_no % 12053
    jy += 4*(g_day_no // 1461); g_day_no = g_day_no % 1461
    if g_day_no > 365:
        jy += (g_day_no-1)//365; g_day_no = (g_day_no-1) % 365
    if g_day_no < 186:
        jm = 1 + g_day_no//31; jd = 1 + (g_day_no % 31)
    else:
        jm = 7 + (g_day_no-186)//30; jd = 1 + ((g_day_no-186) % 30)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

# ============================================================
# خواندن آخرین روز از فایل input
# ============================================================
def read_latest_data(input_file):
    """خواندن آخرین روز از فایل input.xlsx"""
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(input_file, data_only=True)
    except:
        return None
    
    # پیدا کردن شیت‌ها
    ws_price = None
    ws_hqa = None
    for sn in wb.sheetnames:
        if "قیمت" in sn and "سابقه" in sn:
            ws_price = wb[sn]
        elif "حقیقی" in sn and "حقوقی" in sn:
            ws_hqa = wb[sn]
    
    if ws_price is None or ws_hqa is None:
        return None
    
    # آخرین روز قیمت (ردیف ۳ = جدیدترین)
    price_data = {}
    price_data["date"] = ws_price.cell(row=3, column=1).value or "—"
    price_data["close"] = ws_price.cell(row=3, column=7).value or 0
    price_data["prev_close"] = ws_price.cell(row=3, column=5).value or 0
    price_data["last"] = ws_price.cell(row=3, column=10).value or 0
    price_data["high"] = ws_price.cell(row=3, column=14).value or 0
    price_data["low"] = ws_price.cell(row=3, column=13).value or 0
    price_data["volume"] = ws_price.cell(row=3, column=3).value or 0
    price_data["value"] = ws_price.cell(row=3, column=4).value or 0
    price_data["count"] = ws_price.cell(row=3, column=2).value or 0
    price_data["change"] = ws_price.cell(row=3, column=8).value or 0
    pct_decimal = ws_price.cell(row=3, column=9).value
    if pct_decimal is not None and isinstance(pct_decimal, (int, float)):
        price_data["pct"] = pct_decimal * 100 if abs(pct_decimal) < 1 else pct_decimal
    else:
        if price_data["prev_close"] > 0:
            price_data["pct"] = (price_data["change"] / price_data["prev_close"]) * 100
        else:
            price_data["pct"] = 0
    
    # آخرین روز حقیقی/حقوقی (ردیف ۵ = جدیدترین)
    hqa_data = {}
    hqa_data["buy_real_val"] = ws_hqa.cell(row=5, column=4).value or 0
    hqa_data["sell_real_val"] = ws_hqa.cell(row=5, column=12).value or 0
    hqa_data["buy_real_cnt"] = ws_hqa.cell(row=5, column=2).value or 0
    hqa_data["sell_real_cnt"] = ws_hqa.cell(row=5, column=10).value or 0
    hqa_data["buy_legal_val"] = ws_hqa.cell(row=5, column=8).value or 0
    hqa_data["money_flow"] = ws_hqa.cell(row=5, column=20).value or 0
    hqa_data["power_ratio"] = ws_hqa.cell(row=5, column=19).value or 0
    hqa_data["ownership_change"] = ws_hqa.cell(row=5, column=18).value or 0
    
    # محاسبه‌ی خلاصه
    net_flow = hqa_data["buy_real_val"] - hqa_data["sell_real_val"]
    hqa_data["net_flow"] = net_flow
    
    return {**price_data, **hqa_data}

# ============================================================
# رابط کاربری
# ============================================================

st.markdown("""
<div style='background: linear-gradient(135deg, #1B2A4A 0%, #2C3E50 100%);
            padding: 25px; border-radius: 12px; margin-bottom: 20px; text-align: center;'>
    <h1 style='color: white; margin: 0; font-size: 24px;'>📊 خلاصه نمادها — آخرین روز</h1>
    <p style='color: #E8ECF4; margin: 8px 0 0 0; font-size: 14px;'>
        نمایش جدیدترین داده‌های همه نمادها از پوشه input/
    </p>
</div>
""", unsafe_allow_html=True)

# پیدا کردن همه فایل‌های input
input_dir = "input"
if not os.path.exists(input_dir):
    # اگر پوشه input نبود، فایل‌های پوشه فعلی را بگرد
    input_files = glob.glob("*_input.xlsx")
else:
    input_files = glob.glob(os.path.join(input_dir, "*_input.xlsx"))

if not input_files:
    st.warning("هیچ فایل input یافت نشد. ابتدا فایل ایجاد کنید.")
    st.stop()

st.info(f"تعداد نمادها: {len(input_files)}")

# خواندن داده‌های همه نمادها
all_data = []
for f in sorted(input_files):
    # استخراج نام نماد از نام فایل
    basename = os.path.basename(f)
    symbol = basename.replace("_input.xlsx", "")
    
    data = read_latest_data(f)
    if data:
        data["symbol"] = symbol
        all_data.append(data)

if not all_data:
    st.error("داده‌ای قابل خواندن نبود")
    st.stop()

# ============================================================
# جدول خلاصه
# ============================================================
st.markdown("## 📋 جدول خلاصه نمادها")

import pandas as pd

# ساخت DataFrame
rows = []
for d in all_data:
    pct = d.get("pct", 0)
    pct_str = f"{pct:+.2f}%" if isinstance(pct, (int, float)) else "—"
    
    net_flow = d.get("net_flow", 0)
    money_flow = d.get("money_flow", 0)
    
    # رنگ سیگنال
    if money_flow > 0:
        flow_str = f"🟢 +{money_flow:.2f}"
    elif money_flow < 0:
        flow_str = f"🔴 {money_flow:.2f}"
    else:
        flow_str = "⚪ 0.00"
    
    rows.append({
        "نماد": d["symbol"],
        "تاریخ": d["date"],
        "قیمت پایانی": f"{int(d['close']):,}",
        "درصد تغییر": pct_str,
        "آخرین": f"{int(d['last']):,}",
        "حجم": f"{int(d['volume']):,}",
        "خرید حقیقی (م.ریال)": f"{d['buy_real_val']:.3f}",
        "فروش حقیقی (م.ریال)": f"{d['sell_real_val']:.3f}",
        "ورود/خروج پول (م.تومان)": flow_str,
        "قدرت خریدار": f"{d['power_ratio']:.2f}",
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
total_volume = sum(d["volume"] for d in all_data)
total_buy = sum(d["buy_real_val"] for d in all_data)
total_sell = sum(d["sell_real_val"] for d in all_data)
total_net = (total_buy - total_sell) / 10  # میلیارد تومان

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("تعداد نمادها", total_symbols)

with col2:
    st.metric("مجموع حجم", f"{total_volume:,}")

with col3:
    st.metric("مجموع خرید حقیقی", f"{total_buy:.1f} م.ریال")

with col4:
    color = "🟢" if total_net >= 0 else "🔴"
    st.metric(f"{color} خالص پول حقیقی", f"{total_net:+.1f} م.تومان")

# ============================================================
# جدول جزئیات قیمت
# ============================================================
st.markdown("---")
st.markdown("## 💰 جزئیات قیمت")

price_rows = []
for d in all_data:
    pct = d.get("pct", 0)
    pct_str = f"{pct:+.2f}%" if isinstance(pct, (int, float)) else "—"
    
    price_rows.append({
        "نماد": d["symbol"],
        "تاریخ": d["date"],
        "دیروز": f"{int(d['prev_close']):,}",
        "اولین": f"{int(d['first']):,}",
        "پایانی": f"{int(d['close']):,}",
        "آخرین": f"{int(d['last']):,}",
        "کمترین": f"{int(d['low']):,}",
        "بیشترین": f"{int(d['high']):,}",
        "تغییر": f"{int(d['change']):+,}",
        "درصد": pct_str,
        "تعداد معامله": f"{int(d['count']):,}",
    })

df_price = pd.DataFrame(price_rows)
st.dataframe(df_price, use_container_width=True, hide_index=True)

# ============================================================
# جدول جزئیات حقیقی/حقوقی
# ============================================================
st.markdown("---")
st.markdown("## 👥 جزئیات حقیقی/حقوقی")

hqa_rows = []
for d in all_data:
    money_flow = d.get("money_flow", 0)
    if money_flow > 0:
        flow_str = f"🟢 +{money_flow:.2f}"
    elif money_flow < 0:
        flow_str = f"🔴 {money_flow:.2f}"
    else:
        flow_str = "⚪ 0.00"
    
    ownership = d.get("ownership_change", 0)
    ownership_str = f"{int(ownership):+,}" if ownership else "—"
    
    hqa_rows.append({
        "نماد": d["symbol"],
        "خرید حقیقی (تعداد)": f"{int(d['buy_real_cnt']):,}",
        "خرید حقیقی (ارزش م.ریال)": f"{d['buy_real_val']:.3f}",
        "فروش حقیقی (تعداد)": f"{int(d['sell_real_cnt']):,}",
        "فروش حقیقی (ارزش م.ریال)": f"{d['sell_real_val']:.3f}",
        "خرید حقوقی (ارزش م.ریال)": f"{d['buy_legal_val']:.3f}",
        "ورود/خروج (م.تومان)": flow_str,
        "قدرت خریدار": f"{d['power_ratio']:.2f}",
        "تغییر مالکیت (سهم)": ownership_str,
    })

df_hqa = pd.DataFrame(hqa_rows)
st.dataframe(df_hqa, use_container_width=True, hide_index=True)

# ============================================================
# فوتر
# ============================================================
now = datetime.now()
st.markdown(f"""
<div style='text-align: center; color: #8C8A84; font-size: 12px; margin-top: 30px;'>
    خلاصه نمادها | {now.strftime('%Y-%m-%d %H:%M')}
</div>
""", unsafe_allow_html=True)
