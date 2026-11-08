#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد خلاصه — آخرین روز همه نمادها
=====================================

این اسکریپت تمام فایل‌های *_input.xlsx در پوشه input/ را می‌خواند
و آخرین روز هر نماد را با همان ستون‌های داشبورد تحلیلی نمایش می‌دهد.

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
# توابع محاسباتی (همان توابع make_dashboard.py)
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
    if rc_pct is not None:
        rc = rc_pct
    else:
        rc = 0
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
# خواندن داده‌ها از فایل input و محاسبه امتیازات
# ============================================================
def read_latest_data(input_file):
    """خواندن آخرین روز از فایل input.xlsx و محاسبه امتیازات"""
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
    
    # خواندن همه داده‌های تاریخی برای محاسبات ۵ و ۳۰ روزه
    # شیت قیمت
    prices = {}
    MONTH_ORDER = {"فروردین":1,"اردیبهشت":2,"خرداد":3,"تیر":4,"مرداد":5,"شهریور":6,
                   "مهر":7,"آبان":8,"آذر":9,"دی":10,"بهمن":11,"اسفند":12}
    J_MONTHS = {"01":"فروردین","02":"اردیبهشت","03":"خرداد","04":"تیر","05":"مرداد","06":"شهریور",
                "07":"مهر","08":"آبان","09":"آذر","10":"دی","11":"بهمن","12":"اسفند"}
    
    def j_to_persian(j_date):
        parts = str(j_date).split("/")
        if len(parts) != 3: return str(j_date)
        y, m, d = parts
        return f"{J_MONTHS.get(m,'?')} {int(d)}"
    
    def normalize_date(d):
        parts = str(d).split(" ")
        return f"{parts[0]} {int(parts[1])}" if len(parts)==2 else str(d)
    
    def sort_dates(dates):
        return sorted(dates, key=lambda x: (MONTH_ORDER.get(x.split(" ")[0],0), int(x.split(" ")[1])))
    
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
    
    # شیت حقیقی/حقوقی
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
    
    # ساخت داده‌ساختار
    mf = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"]) for d in reversed(common_sorted)]
    vol = [(d, hqa[d]["volume"]) for d in common_sorted]
    bp = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"], hqa[d]["buy_real_count"],
           hqa[d]["sell_real_count"], hqa[d]["buy_legal_value"]) for d in common_sorted]
    pri = [{"date":d, **prices[d]} for d in common_sorted if d in prices]
    
    # محاسبه امتیازات برای آخرین روز (idx=0)
    i = 0
    
    # جریان پول
    ratio_5d_mf = calc_money_flow_5d(mf, i)
    score_mf = calc_money_flow_score(ratio_5d_mf) or 0
    
    # حجم
    vol_rev = list(reversed(vol))
    avg_30 = calc_volume_avg(vol_rev, i, 30)
    if avg_30 is not None:
        v_today = vol_rev[i][1]
        ratio_30 = safe_div(v_today, avg_30)
        score_vol = calc_volume_score(ratio_30) or 0
    else:
        score_vol = 0
    
    # قدرت خریدار
    d, bv, sv, bc, sc, lv = bp[i] if i < len(bp) else (None, 0, 0, 0, 0, 0)
    # bp is reversed (newest first), so bp[0] = today
    # Actually bp is oldest->newest, need to reverse
    bp_rev = list(reversed(bp))
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
    pri_rev = list(reversed(pri))
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
    
    # داده‌های خام آخرین روز
    latest_price = pri[-1] if pri else {"date":"—", "close":0, "prev_close":0, "last":0, "high":0, "low":0}
    latest_hqa_date = mf[0][0] if mf else "—"
    latest_buy_real = hqa.get(latest_hqa_date, {}).get("buy_real_value", 0)
    latest_sell_real = hqa.get(latest_hqa_date, {}).get("sell_real_value", 0)
    latest_volume = hqa.get(latest_hqa_date, {}).get("volume", 0)
    latest_money_flow = (latest_buy_real - latest_sell_real) / 10 if latest_buy_real and latest_sell_real else 0
    
    return {
        "date": latest_price.get("date", "—"),
        "close": latest_price.get("close", 0),
        "prev_close": latest_price.get("prev_close", 0),
        "last": latest_price.get("last", 0),
        "high": latest_price.get("high", 0),
        "low": latest_price.get("low", 0),
        "volume": latest_volume,
        "buy_real_val": latest_buy_real,
        "sell_real_val": latest_sell_real,
        "buy_real_cnt": bc or 0,
        "sell_real_cnt": sc or 0,
        "power_ratio": power_ratio,
        "money_flow": latest_money_flow,
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
# رابط کاربری
# ============================================================

st.markdown("""
<div style='background: linear-gradient(135deg, #1B2A4A 0%, #2C3E50 100%);
            padding: 25px; border-radius: 12px; margin-bottom: 20px; text-align: center;'>
    <h1 style='color: white; margin: 0; font-size: 24px;'>📊 خلاصه نمادها — آخرین روز</h1>
    <p style='color: #E8ECF4; margin: 8px 0 0 0; font-size: 14px;'>
        نمایش جدیدترین داده‌های همه نمادها با محاسبات داشبورد
    </p>
</div>
""", unsafe_allow_html=True)

# پیدا کردن همه فایل‌های input
input_dir = "input"
if not os.path.exists(input_dir):
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
