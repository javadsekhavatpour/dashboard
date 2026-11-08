#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد خلاصه — آخرین روز همه نمادها
=====================================

این اسکریپت تمام فایل‌های *_dashboard.xlsx در پوشه dashboard/ را می‌خواند
و فقط آخرین سطر (جدیدترین روز) هر نماد را نمایش می‌دهد.

همچنین برای هر نماد دو دکمه دارد:
  🔄 بروزرسانی: دریافت داده‌های زنده امروز از TSETMC و بازسازی داشبورد
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
# وارد کردن توابع از update_input.py
# ============================================================
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from update_input import (
        find_ins_code, fetch_live_price, fetch_live_hqa,
        gregorian_to_jalali, deven_to_jdate_full, deven_to_jdate_fa
    )
except ImportError:
    st.error("فایل update_input.py یافت نشد!")
    st.stop()

# ============================================================
# توابع محاسباتی داشبورد
# ============================================================
def safe_div(num, den):
    if den is None or den == 0: return 0
    return num / den

def calc_money_flow_5d(mf_list, idx):
    if idx + 4 >= len(mf_list): return None
    b5 = sum(mf_list[k][1] for k in range(idx, idx+5))
    s5 = sum(mf_list[k][2] for k in range(idx, idx+5))
    total = b5 + s5
    if total == 0: return 0
    return ((b5 - s5) / total) * 100

def calc_money_flow_score(ratio_5d):
    if ratio_5d is None: return None
    if ratio_5d <= 0: return 0
    elif ratio_5d < 2: return 5
    elif ratio_5d < 5: return 10
    elif ratio_5d < 10: return 20
    else: return 30

def calc_volume_avg(vol_list, idx, days):
    if idx + days >= len(vol_list): return None
    return sum(vol_list[k][1] for k in range(idx+1, idx+1+days)) / days

def calc_volume_score(ratio_30d):
    if ratio_30d is None: return None
    if ratio_30d < 0.75: return 0
    elif ratio_30d < 1: return 3
    elif ratio_30d < 1.25: return 5
    elif ratio_30d < 1.5: return 8
    elif ratio_30d < 2: return 11
    elif ratio_30d < 3: return 13
    else: return 15

def calc_buyer_power_score(power_ratio, count_ratio, real_buy_share):
    if power_ratio >= 2: s1 = 10
    elif power_ratio >= 1.5: s1 = 8
    elif power_ratio >= 1.2: s1 = 6
    elif power_ratio >= 1: s1 = 4
    elif power_ratio >= 0.8: s1 = 2
    else: s1 = 0
    if count_ratio < 0.7: s2 = 5
    elif count_ratio < 0.9: s2 = 4
    elif count_ratio < 1.1: s2 = 3
    elif count_ratio < 1.3: s2 = 2
    elif count_ratio < 1.5: s2 = 1
    else: s2 = 0
    if real_buy_share >= 60: s3 = 5
    elif real_buy_share >= 55: s3 = 4
    elif real_buy_share >= 50: s3 = 3
    elif real_buy_share >= 45: s3 = 2
    elif real_buy_share >= 40: s3 = 1
    else: s3 = 0
    return s1 + s2 + s3

def calc_price_scores(close, prev_close, last, high, low, rc_pct=None):
    if rc_pct is not None: rc = rc_pct
    else: rc = 0
    if rc > 2: s1 = 8
    elif rc > 0: s1 = 6
    elif rc == 0: s1 = 4
    elif rc > -2: s1 = 2
    else: s1 = 0
    rl = safe_div(last - close, close) * 100
    if rl > 1: s2 = 5
    elif rl > 0: s2 = 4
    elif rl == 0: s2 = 3
    elif rl > -1: s2 = 1
    else: s2 = 0
    if high == low: position = 100
    else: position = safe_div(close - low, high - low) * 100
    if position >= 80: s3 = 7
    elif position >= 60: s3 = 5
    elif position >= 40: s3 = 3
    elif position >= 20: s3 = 1
    else: s3 = 0
    if high == low: recovery = 100
    else: recovery = safe_div(last - low, high - low) * 100
    if recovery >= 80: s4 = 5
    elif recovery >= 60: s4 = 4
    elif recovery >= 40: s4 = 3
    elif recovery >= 20: s4 = 1
    else: s4 = 0
    return rc, s1+s2+s3+s4

def calc_persistence(mf_list, idx):
    if idx + 4 >= len(mf_list): return None, None, None
    statuses = []
    for k in range(idx, idx+5):
        bv = mf_list[k][1]; sv = mf_list[k][2]
        ratio = safe_div(bv - sv, bv + sv) * 100
        if ratio > 0.5: status = 1
        elif ratio < -0.5: status = -1
        else: status = 0
        statuses.append(status)
    positive_days = sum(1 for s in statuses if s == 1)
    if positive_days == 5: score = 10
    elif positive_days == 4: score = 8
    elif positive_days == 3: score = 6
    elif positive_days == 2: score = 3
    elif positive_days == 1: score = 1
    else: score = 0
    return positive_days, score, None

# ============================================================
# خواندن همه داده‌های تاریخی از فایل input
# ============================================================
def load_all_data(input_file):
    """خواندن همه داده‌ها از فایل input برای محاسبات"""
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(input_file, data_only=True)
    except:
        return None
    
    ws_price = None; ws_hqa = None
    for sn in wb.sheetnames:
        if "قیمت" in sn and "سابقه" in sn: ws_price = wb[sn]
        elif "حقیقی" in sn and "حقوقی" in sn: ws_hqa = wb[sn]
    
    if ws_price is None or ws_hqa is None:
        return None
    
    MONTH_ORDER = {"فروردین":1,"اردیبهشت":2,"خرداد":3,"تیر":4,"مرداد":5,"شهریور":6,
                   "مهر":7,"آبان":8,"آذر":9,"دی":10,"بهمن":11,"اسفند":12}
    J_MONTHS_MAP = {"01":"فروردین","02":"اردیبهشت","03":"خرداد","04":"تیر","05":"مرداد","06":"شهریور",
                    "07":"مهر","08":"آبان","09":"آذر","10":"دی","11":"بهمن","12":"اسفند"}
    
    def j_to_persian(j_date):
        parts = str(j_date).split("/")
        if len(parts) != 3: return str(j_date)
        y, m, d = parts
        return f"{J_MONTHS_MAP.get(m,'?')} {int(d)}"
    
    def normalize_date(d):
        parts = str(d).split(" ")
        return f"{parts[0]} {int(parts[1])}" if len(parts)==2 else str(d)
    
    def sort_dates(dates):
        return sorted(dates, key=lambda x: (MONTH_ORDER.get(x.split(" ")[0],0), int(x.split(" ")[1])))
    
    prices = {}
    for row in range(3, 200):
        dj = ws_price.cell(row=row, column=1).value
        if not dj or "جمع" in str(dj): break
        df = j_to_persian(dj)
        c = ws_price.cell(row=row, column=7).value
        pc = ws_price.cell(row=row, column=5).value
        pct_decimal = ws_price.cell(row=row, column=9).value
        if pct_decimal is not None and isinstance(pct_decimal, (int, float)):
            rc_pct = pct_decimal * 100 if abs(pct_decimal) < 1 else pct_decimal
        else:
            rc_pct = None
        if c and pc:
            prices[df] = {
                "close": c, "prev_close": pc,
                "last": ws_price.cell(row=row, column=10).value,
                "high": ws_price.cell(row=row, column=14).value,
                "low": ws_price.cell(row=row, column=13).value,
                "rc_pct": rc_pct,
            }
    
    hqa = {}
    for row in range(5, 200):
        d = ws_hqa.cell(row=row, column=1).value
        if not d or "جمع" in str(d) or "تاریخ" in str(d): break
        df = normalize_date(d)
        hqa[df] = {
            "buy_real_value": ws_hqa.cell(row=row, column=4).value or 0,
            "sell_real_value": ws_hqa.cell(row=row, column=12).value or 0,
            "buy_real_count": ws_hqa.cell(row=row, column=2).value or 0,
            "sell_real_count": ws_hqa.cell(row=row, column=10).value or 0,
            "buy_legal_value": ws_hqa.cell(row=row, column=8).value or 0,
            "volume": (ws_hqa.cell(row=row, column=3).value or 0) + (ws_hqa.cell(row=row, column=7).value or 0),
        }
    
    common = set(prices.keys()) & set(hqa.keys())
    common_sorted = sort_dates(common)
    if not common_sorted:
        return None
    
    mf = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"]) for d in reversed(common_sorted)]
    vol = [(d, hqa[d]["volume"]) for d in common_sorted]
    bp = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"], hqa[d]["buy_real_count"],
           hqa[d]["sell_real_count"], hqa[d]["buy_legal_value"]) for d in common_sorted]
    pri = [{"date":d, **prices[d]} for d in common_sorted if d in prices]
    
    return {"money_flow": mf, "volume": vol, "buyer_power": bp,
            "prices": pri, "persistence": list(reversed(mf))}

# ============================================================
# محاسبه امتیازات برای داده زنده
# ============================================================
def calculate_scores_from_live(live_price, live_hqa, ds):
    """محاسبه امتیازات با داده زنده امروز + داده تاریخی"""
    
    # اضافه کردن داده زنده به داده تاریخی
    if live_price:
        live_date_fa = deven_to_jdate_full(live_price["dEven"])
        live_pct = None
        if live_price.get("priceYesterday", 0) > 0:
            change = live_price["pClosing"] - live_price["priceYesterday"]
            live_pct = (change / live_price["priceYesterday"]) * 100
        
        live_p = {
            "date": live_date_fa,
            "close": live_price["pClosing"],
            "prev_close": live_price["priceYesterday"],
            "last": live_price["pDrCotVal"],
            "high": live_price["priceMax"],
            "low": live_price["priceMin"],
            "rc_pct": live_pct,
        }
    else:
        return None
    
    # اضافه کردن به لیست قیمت‌ها (در ابتدا = جدیدترین)
    pri_rev = list(reversed(ds["prices"]))
    pri_rev.insert(0, live_p)
    
    # اضافه کردن به جریان پول
    mf = ds["money_flow"].copy()
    if live_hqa:
        buy_real = live_hqa["buy_I_Value"] / 1e9  # تبدیل به میلیارد ریال
        sell_real = live_hqa["sell_I_Value"] / 1e9
        mf.insert(0, (live_date_fa, buy_real, sell_real))
    
    # اضافه کردن به حجم
    vol_rev = list(reversed(ds["volume"]))
    if live_hqa:
        live_vol = live_hqa["buy_I_Volume"] + live_hqa["buy_N_Volume"]
        vol_rev.insert(0, (live_date_fa, live_vol))
    
    # اضافه کردن به قدرت خریدار
    bp_rev = list(reversed(ds["buyer_power"]))
    if live_hqa:
        buy_real_val = live_hqa["buy_I_Value"] / 1e9
        sell_real_val = live_hqa["sell_I_Value"] / 1e9
        buy_real_cnt = int(live_hqa["buy_I_Count"])
        sell_real_cnt = int(live_hqa["sell_I_Count"])
        buy_legal_val = live_hqa["buy_N_Value"] / 1e9
        bp_rev.insert(0, (live_date_fa, buy_real_val, sell_real_val, buy_real_cnt, sell_real_cnt, buy_legal_val))
    
    # محاسبه امتیازات (idx=0 = امروز)
    i = 0
    
    # جریان پول
    ratio_5d_mf = calc_money_flow_5d(mf, i)
    score_mf = calc_money_flow_score(ratio_5d_mf) or 0
    
    # حجم
    avg_30 = calc_volume_avg(vol_rev, i, 30)
    if avg_30 is not None:
        v_today = vol_rev[i][1]
        ratio_30 = safe_div(v_today, avg_30)
        score_vol = calc_volume_score(ratio_30) or 0
    else:
        score_vol = 0
    
    # قدرت خریدار
    if bp_rev:
        d, bv, sv, bc, sc, lv = bp_rev[i]
    else:
        d, bv, sv, bc, sc, lv = (None, 0, 0, 0, 0, 0)
    
    avg_buy = safe_div(bv, bc)
    avg_sell = safe_div(sv, sc)
    power_ratio = safe_div(avg_buy, avg_sell)
    count_ratio = safe_div(bc, sc)
    total_value = bv + lv
    real_buy_share = safe_div(bv, total_value) * 100
    score_bp = calc_buyer_power_score(power_ratio, count_ratio, real_buy_share) or 0
    
    # واکنش قیمت
    p = pri_rev[i] if pri_rev else {"close":0, "prev_close":0, "last":0, "high":0, "low":0, "rc_pct":None}
    rc, score_price = calc_price_scores(
        p["close"], p["prev_close"], p["last"], p["high"], p["low"],
        rc_pct=p.get("rc_pct")
    )
    
    # تداوم
    pos_days, score_pers, state_pers = calc_persistence(mf, i)
    score_pers = score_pers or 0
    
    # امتیاز کل
    total_score = score_mf + score_vol + score_bp + score_price + score_pers
    
    # شدت
    if total_score >= 70: intensity = "بسیار قوی"
    elif total_score >= 50: intensity = "قوی"
    elif total_score >= 30: intensity = "متوسط"
    else: intensity = "ضعیف"
    
    # سیگنال
    net_mf = bv - sv if bv and sv else 0
    if total_score >= 50 and net_mf > 0: signal = "خرید قوی"
    elif total_score >= 30 and net_mf > 0: signal = "خرید"
    elif total_score <= 30 and net_mf < 0: signal = "فروش"
    elif total_score <= 50 and net_mf < 0: signal = "هشدار"
    else: signal = "خنثی"
    
    return {
        "date": live_date_fa,
        "close": live_price["pClosing"],
        "score_mf": score_mf,
        "score_vol": score_vol,
        "score_bp": score_bp,
        "score_price": score_price,
        "score_pers": score_pers,
        "total_score": total_score,
        "rc": rc,
        "intensity": intensity,
        "signal": signal,
    }

# ============================================================
# خواندن آخرین سطر از فایل داشبورد
# ============================================================
def read_latest_from_dashboard(dashboard_file):
    """خواندن آخرین سطر (ردیف ۳) از شیت داشبورد"""
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
# بروزرسانی زنده نماد
# ============================================================
def live_update_symbol(symbol):
    """دریافت داده زنده امروز و محاسبه امتیازات"""
    
    # ۱. پیدا کردن کد نماد
    ins_code, full_name, sym_name = find_ins_code(symbol)
    if not ins_code:
        return None, f"نماد {symbol} یافت نشد"
    
    # ۲. خواندن داده‌های تاریخی از فایل input
    input_file = os.path.join("input", f"{symbol}_input.xlsx")
    if not os.path.exists(input_file):
        input_file = f"{symbol}_input.xlsx"
    
    if not os.path.exists(input_file):
        return None, f"فایل input برای {symbol} یافت نشد"
    
    ds = load_all_data(input_file)
    if ds is None:
        return None, "خطا در خواندن فایل input"
    
    # ۳. دریافت داده زنده امروز
    live_price = fetch_live_price(ins_code)
    if not live_price:
        return None, "دریافت قیمت زنده ناموفق بود"
    
    live_hqa = fetch_live_hqa(ins_code)
    if not live_hqa:
        return None, "دریافت حقیقی/حقوقی زنده ناموفق بود"
    
    # ۴. محاسبه امتیازات
    result = calculate_scores_from_live(live_price, live_hqa, ds)
    if result is None:
        return None, "خطا در محاسبه امتیازات"
    
    return result, None

# ============================================================
# رابط کاربری
# ============================================================

st.markdown("""
<div style='background: linear-gradient(135deg, #1B2A4A 0%, #2C3E50 100%);
            padding: 25px; border-radius: 12px; margin-bottom: 20px; text-align: center;'>
    <h1 style='color: white; margin: 0; font-size: 24px;'>📊 خلاصه نمادها — آخرین روز</h1>
    <p style='color: #E8ECF4; margin: 8px 0 0 0; font-size: 14px;'>
        نمایش جدیدترین سطر از فایل‌های داشبورد + بروزرسانی زنده
    </p>
</div>
""", unsafe_allow_html=True)

# پیدا کردن همه فایل‌های dashboard
dashboard_dir = "dashboard"
if not os.path.exists(dashboard_dir):
    dashboard_files = glob.glob("*_dashboard.xlsx")
else:
    dashboard_files = glob.glob(os.path.join(dashboard_dir, "*_dashboard.xlsx"))

# همچنین فایل‌های input
input_dir = "input"
if os.path.exists(input_dir):
    input_files = glob.glob(os.path.join(input_dir, "*_input.xlsx"))
else:
    input_files = glob.glob("*_input.xlsx")

# لیست همه نمادها
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
    # خواندن داده‌های داشبورد (سطر اول = داده تاریخی)
    dash_file = os.path.join(dashboard_dir, f"{symbol}_dashboard.xlsx")
    if not os.path.exists(dash_file):
        dash_file = f"{symbol}_dashboard.xlsx"
    
    hist_data = None
    if os.path.exists(dash_file):
        hist_data = read_latest_from_dashboard(dash_file)
    
    # بررسی وجود داده زنده در session_state
    live_key = f"live_{symbol}"
    live_data = st.session_state.get(live_key)
    
    # ===== ردیف اول: داده تاریخی =====
    data = hist_data
    cols = st.columns([1, 1.5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1.5])
    
    with cols[0]:
        st.markdown(f"**{symbol}**")
    with cols[1]:
        st.markdown(data["date"] if data else "—")
    with cols[2]:
        st.markdown(f"● {data['score_mf']}" if data else "—")
    with cols[3]:
        st.markdown(f"● {data['score_vol']}" if data else "—")
    with cols[4]:
        st.markdown(f"● {data['score_bp']}" if data else "—")
    with cols[5]:
        st.markdown(f"● {data['score_price']}" if data else "—")
    with cols[6]:
        st.markdown(f"● {data['score_pers']}" if data else "—")
    with cols[7]:
        st.markdown(f"● {data['total_score']}" if data else "—")
    with cols[8]:
        st.markdown(data["intensity"] if data else "—")
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
        if st.button("🔄", key=f"update_{symbol}", help=f"بروزرسانی زنده {symbol}"):
            with st.spinner(f"در حال دریافت داده زنده {symbol}..."):
                result, error = live_update_symbol(symbol)
            if result:
                st.session_state[live_key] = result
                st.success(f"✅ {symbol} بروزرسانی شد — {result['date']}")
                st.rerun()
            else:
                st.error(f"❌ {error}")
    with cols[12]:
        if st.button("📊", key=f"history_{symbol}", help=f"داشبورد {symbol}"):
            st.session_state["show_history"] = symbol
    
    # ===== ردیف دوم: داده زنده (اگر وجود داشته باشد) =====
    if live_data:
        cols2 = st.columns([1, 1.5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1.5])
        
        with cols2[0]:
            st.markdown(f"🟢 **{symbol}**")
        with cols2[1]:
            st.markdown(f"🟢 {live_data['date']}")
        with cols2[2]:
            st.markdown(f"🟢 ● {live_data['score_mf']}")
        with cols2[3]:
            st.markdown(f"🟢 ● {live_data['score_vol']}")
        with cols2[4]:
            st.markdown(f"🟢 ● {live_data['score_bp']}")
        with cols2[5]:
            st.markdown(f"🟢 ● {live_data['score_price']}")
        with cols2[6]:
            st.markdown(f"🟢 ● {live_data['score_pers']}")
        with cols2[7]:
            st.markdown(f"🟢 ● {live_data['total_score']}")
        with cols2[8]:
            st.markdown(f"🟢 {live_data['intensity']}")
        with cols2[9]:
            rc = live_data["rc"]
            st.markdown(f"🟢 {'+' if rc >= 0 else ''}{rc:.2f}")
        with cols2[10]:
            signal = live_data["signal"]
            if signal in ["خرید قوی", "خرید"]:
                st.markdown(f"🟢 {signal}")
            elif signal in ["فروش", "هشدار"]:
                st.markdown(f"🔴 {signal}")
            else:
                st.markdown(f"⚪ {signal}")
        with cols2[11]:
            if st.button("🗑️", key=f"clear_{symbol}", help=f"حذف داده زنده {symbol}"):
                del st.session_state[live_key]
                st.rerun()
        with cols2[12]:
            st.markdown("")
    
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

all_data = []
for symbol in all_symbols:
    live_key = f"live_{symbol}"
    if live_key in st.session_state:
        all_data.append(st.session_state[live_key])
    else:
        dash_file = os.path.join(dashboard_dir, f"{symbol}_dashboard.xlsx")
        if not os.path.exists(dash_file):
            dash_file = f"{symbol}_dashboard.xlsx"
        if os.path.exists(dash_file):
            data = read_latest_from_dashboard(dash_file)
            if data:
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
