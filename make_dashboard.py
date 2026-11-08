#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ساخت داشبورد از فایل input — نسخه مستقل (بدون نیاز به LibreOffice)
====================================================================

این اسکریپت یک فایل input.xlsx (شامل دو شیت «سابقه قیمت و معاملات» و «حقیقی و حقوقی»)
را می‌خواند و فایل داشبورد سهام را می‌سازد.

تمام محاسبات در Python انجام می‌شود — نیازی به LibreOffice یا Excel نیست.

روش استفاده (ساده):
  python make_dashboard.py
  python make_dashboard.py سرود_input.xlsx
  python make_dashboard.py سرود_input.xlsx --name سرود
  python make_dashboard.py سرود_input.xlsx --output داشبورد.xlsx
"""

import sys, os, argparse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# استایل‌ها و رنگ‌ها
# ============================================================
FONT_NAME = "Vazirmatn"
PRIMARY = "1B2A4A"
PRIMARY_LIGHT = "E8ECF4"
NEUTRAL_900 = "37352F"
NEUTRAL_600 = "8C8A84"
NEUTRAL_200 = "E9E9E8"
NEUTRAL_100 = "F7F7F5"
NEUTRAL_0 = "FFFFFF"
COLOR_RED = "C0392B"
COLOR_ORANGE = "D4820A"
COLOR_YELLOW = "F1C40F"
COLOR_YELLOW_BG = "F9E79F"
COLOR_RED_BG = "F5B7B1"
COLOR_BLUE = "1B2A4A"
COLOR_GREEN = "1B7D46"
COLOR_DARK_GREEN = "0E5A2E"

H_FILL = PatternFill("solid", fgColor=PRIMARY)
H_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
H_ALIGN = Alignment(horizontal="center", vertical="center", readingOrder=2, wrap_text=True)
H_BORDER = Border(bottom=Side(style="thin", color=NEUTRAL_200))
B_FONT = Font(name=FONT_NAME, size=9, color=NEUTRAL_900)
B_ALIGN = Alignment(horizontal="center", vertical="center", readingOrder=2)
B_ALIGN_NUM = Alignment(horizontal="left", vertical="center")
INPUT_FILL = PatternFill("solid", fgColor="F0F4FA")
INPUT_FONT = Font(name=FONT_NAME, size=9, color="0000FF")

MONTH_ORDER = {"فروردین":1,"اردیبهشت":2,"خرداد":3,"تیر":4,"مرداد":5,"شهریور":6,
               "مهر":7,"آبان":8,"آذر":9,"دی":10,"بهمن":11,"اسفند":12}

# ============================================================
# توابع کمکی
# ============================================================
def j_to_persian(j_date):
    """تبدیل '1405/08/13' به 'آبان 13'"""
    parts = str(j_date).split("/")
    if len(parts) != 3: return str(j_date)
    y, m, d = parts
    mm = {"01":"فروردین","02":"اردیبهشت","03":"خرداد","04":"تیر","05":"مرداد","06":"شهریور",
          "07":"مهر","08":"آبان","09":"آذر","10":"دی","11":"بهمن","12":"اسفند"}
    return f"{mm.get(m,'?')} {int(d)}"

def normalize_date(d):
    """'آبان 13' → 'آبان 13' (نرمال‌سازی فاصله)"""
    parts = str(d).split(" ")
    return f"{parts[0]} {int(parts[1])}" if len(parts)==2 else str(d)

def sort_dates(dates):
    """مرتب‌سازی تاریخ‌های فارسی به ترتیب زمانی"""
    return sorted(dates, key=lambda x: (MONTH_ORDER.get(x.split(" ")[0],0),
                                         int(x.split(" ")[1])))

def safe_div(num, den):
    """تقسیم امن - اگر مخرج صفر باشد، 0 برمی‌گرداند"""
    if den is None or den == 0: return 0
    return num / den

def get_score_color(s):
    if s is None or not isinstance(s,(int,float)): return NEUTRAL_600
    if s<=0: return COLOR_RED
    elif s<2: return COLOR_ORANGE
    elif s<10: return COLOR_BLUE
    elif s<20: return COLOR_GREEN
    else: return COLOR_DARK_GREEN

def get_rc_color(r):
    if r is None or not isinstance(r,(int,float)): return NEUTRAL_600
    if r>=3: return COLOR_DARK_GREEN
    elif r>=2: return COLOR_GREEN
    elif r>0: return COLOR_BLUE
    elif r>-2: return COLOR_ORANGE
    else: return COLOR_RED

def set_widths(ws, widths):
    ws.column_dimensions["A"].width = 3
    for c, w in widths.items(): ws.column_dimensions[c].width = w

def find_price_sheet(wb):
    """پیدا کردن شیت قیمت در workbook"""
    candidates = ["سابقه قیمت و معاملات", "سابقه قیمت"]
    for sn in wb.sheetnames:
        if sn in candidates: return wb[sn]
    for sn in wb.sheetnames:
        if "قیمت" in sn and "سابقه" in sn: return wb[sn]
    return None

def find_hqa_sheet(wb):
    """پیدا کردن شیت حقیقی/حقوقی در workbook"""
    candidates = ["حقیقی و حقوقی", "حقیقی/حقوقی"]
    for sn in wb.sheetnames:
        if sn in candidates: return wb[sn]
    for sn in wb.sheetnames:
        if "حقیقی" in sn and "حقوقی" in sn: return wb[sn]
    return None

# ============================================================
# توابع محاسباتی (جایگزین فرمول‌های Excel)
# ============================================================
def calc_money_flow_5d(mf_list, idx):
    """محاسبه نسبت ۵ روزه جریان پول برای روز idx.
    mf_list: لیست (date, buy_value, sell_value) به ترتیب newest first
    idx: اندیس روز جاری (0 = آخرین روز)
    """
    if idx + 4 >= len(mf_list): return None
    b5 = sum(mf_list[k][1] for k in range(idx, idx+5))
    s5 = sum(mf_list[k][2] for k in range(idx, idx+5))
    total = b5 + s5
    if total == 0: return 0
    return ((b5 - s5) / total) * 100

def calc_money_flow_score(ratio_5d):
    """محاسبه امتیاز جریان پول (۰-۳۰) از نسبت ۵ روزه"""
    if ratio_5d is None: return None
    if ratio_5d <= 0: return 0
    elif ratio_5d < 2: return 5
    elif ratio_5d < 5: return 10
    elif ratio_5d < 10: return 20
    else: return 30

def calc_volume_avg(vol_list, idx, days):
    """محاسبه میانگین حجم N روز گذشته"""
    if idx + days >= len(vol_list): return None
    return sum(vol_list[k][1] for k in range(idx+1, idx+1+days)) / days

def calc_volume_score(ratio_30d):
    """محاسبه امتیاز حجم (۰-۱۵)"""
    if ratio_30d is None: return None
    if ratio_30d < 0.75: return 0
    elif ratio_30d < 1: return 3
    elif ratio_30d < 1.25: return 5
    elif ratio_30d < 1.5: return 8
    elif ratio_30d < 2: return 11
    elif ratio_30d < 3: return 13
    else: return 15

def calc_volume_intensity(score):
    """محاسبه شدت حجم از امتیاز"""
    if score is None: return "—"
    mapping = {0:"بسیار کم", 3:"کم", 5:"عادی", 8:"بالا",
               11:"بسیار بالا", 13:"غیرعادی", 15:"بسیار غیرعادی"}
    return mapping.get(score, "—")

def calc_buyer_power_score(power_ratio, count_ratio, real_buy_share):
    """محاسبه امتیاز قدرت خریدار (۰-۲۰)"""
    # مؤلفه ۱: قدرت (۱۰ امتیاز)
    if power_ratio >= 2: s1 = 10
    elif power_ratio >= 1.5: s1 = 8
    elif power_ratio >= 1.2: s1 = 6
    elif power_ratio >= 1: s1 = 4
    elif power_ratio >= 0.8: s1 = 2
    else: s1 = 0
    # مؤلفه ۲: تعداد (۵ امتیاز)
    if count_ratio < 0.7: s2 = 5
    elif count_ratio < 0.9: s2 = 4
    elif count_ratio < 1.1: s2 = 3
    elif count_ratio < 1.3: s2 = 2
    elif count_ratio < 1.5: s2 = 1
    else: s2 = 0
    # مؤلفه ۳: سهم (۵ امتیاز)
    if real_buy_share >= 60: s3 = 5
    elif real_buy_share >= 55: s3 = 4
    elif real_buy_share >= 50: s3 = 3
    elif real_buy_share >= 45: s3 = 2
    elif real_buy_share >= 40: s3 = 1
    else: s3 = 0
    return s1 + s2 + s3

def calc_price_scores(close, prev_close, last, high, low, rc_pct=None):
    """محاسبه امتیاز واکنش قیمت (۰-۲۵)
    اگر rc_pct داده شود (از فایل input)، مستقیم استفاده می‌کنیم
    در غیر این صورت، محاسبه نمی‌کنیم و امتیاز صفر می‌گیرد.
    """
    # مؤلفه ۱: بازده پایانی (۸ امتیاز)
    # فقط از درصد تغییر قیمت پایانی فایل input استفاده می‌کنیم
    # اگر نبود، rc = None و امتیاز صفر
    if rc_pct is not None:
        rc = rc_pct
    else:
        rc = 0  # اگر داده‌ای در فایل نبود، صفر در نظر بگیر
    if rc > 2: s1 = 8
    elif rc > 0: s1 = 6
    elif rc == 0: s1 = 4
    elif rc > -2: s1 = 2
    else: s1 = 0
    # مؤلفه ۲: آخرین (۵ امتیاز)
    rl = safe_div(last - close, close) * 100
    if rl > 1: s2 = 5
    elif rl > 0: s2 = 4
    elif rl == 0: s2 = 3
    elif rl > -1: s2 = 1
    else: s2 = 0
    # مؤلفه ۳: موقعیت (۷ امتیاز)
    if high == low:
        position = 100
    else:
        position = safe_div(close - low, high - low) * 100
    if position >= 80: s3 = 7
    elif position >= 60: s3 = 5
    elif position >= 40: s3 = 3
    elif position >= 20: s3 = 1
    else: s3 = 0
    # مؤلفه ۴: بازیابی (۵ امتیاز)
    if high == low:
        recovery = 100
    else:
        recovery = safe_div(last - low, high - low) * 100
    if recovery >= 80: s4 = 5
    elif recovery >= 60: s4 = 4
    elif recovery >= 40: s4 = 3
    elif recovery >= 20: s4 = 1
    else: s4 = 0
    return rc, s1, rl, s2, position, s3, recovery, s4, s1+s2+s3+s4

def calc_persistence(mf_list, idx):
    """محاسبه امتیاز تداوم جریان پول (۰-۱۰)"""
    if idx + 4 >= len(mf_list): return None, None, None
    # محاسبه status برای هر روز
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
    if score >= 8: state = "تداوم قوی"
    elif score >= 6: state = "تداوم متوسط"
    elif score >= 3: state = "تداوم ضعیف"
    else: state = "بدون تداوم"
    return positive_days, score, state

# ============================================================
# خواندن داده‌ها از فایل input
# ============================================================
def load_data(input_file):
    """خواندن داده‌ها از فایل input و ساخت داده‌ساختار (ds)"""
    print(f"📥 خواندن فایل: {input_file}")
    wb = load_workbook(input_file, data_only=True)

    ws_price = find_price_sheet(wb)
    ws_hqa   = find_hqa_sheet(wb)

    if ws_price is None:
        raise ValueError(f"شیت قیمت یافت نشد! شیت‌های موجود: {wb.sheetnames}")
    if ws_hqa is None:
        raise ValueError(f"شیت حقیقی/حقوقی یافت نشد! شیت‌های موجود: {wb.sheetnames}")

    print(f"   ✓ شیت قیمت: '{ws_price.title}'")
    print(f"   ✓ شیت حقیقی/حقوقی: '{ws_hqa.title}'")

    # استخراج قیمت
    prices = {}
    for row in range(3, 200):
        dj = ws_price.cell(row=row, column=1).value
        if not dj or "جمع" in str(dj): break
        df = j_to_persian(dj)
        c  = ws_price.cell(row=row, column=7).value   # قیمت پایانی
        pc = ws_price.cell(row=row, column=5).value   # قیمت دیروز
        # درصد تغییر قیمت پایانی (ستون ۹) - مستقیم از فایل input
        # این مقدار به صورت اعشاری است (0.0196 = 1.96%)
        pct_decimal = ws_price.cell(row=row, column=9).value
        if pct_decimal is not None and isinstance(pct_decimal, (int, float)):
            # تبدیل به درصد نمایشی (1.96)
            rc_pct = pct_decimal * 100 if abs(pct_decimal) < 1 else pct_decimal
        else:
            rc_pct = None  # اگر نبود، در build محاسبه می‌کنیم
        if c and pc:
            prices[df] = {
                "close": c, "prev_close": pc,
                "last":  ws_price.cell(row=row, column=10).value,
                "high":  ws_price.cell(row=row, column=14).value,
                "low":   ws_price.cell(row=row, column=13).value,
                "rc_pct": rc_pct,  # درصد تغییر قیمت پایانی از فایل input
            }

    # استخراج حقیقی/حقوقی
    hqa = {}
    for row in range(5, 200):
        d = ws_hqa.cell(row=row, column=1).value
        if not d or "جمع" in str(d) or "تاریخ" in str(d): break
        df = normalize_date(d)
        hqa[df] = {
            "buy_real_value":   ws_hqa.cell(row=row, column=4).value or 0,
            "sell_real_value":  ws_hqa.cell(row=row, column=12).value or 0,
            "buy_real_count":   ws_hqa.cell(row=row, column=2).value or 0,
            "sell_real_count":  ws_hqa.cell(row=row, column=10).value or 0,
            "buy_legal_value":  ws_hqa.cell(row=row, column=8).value or 0,
            "volume": (ws_hqa.cell(row=row, column=3).value or 0) + (ws_hqa.cell(row=row, column=7).value or 0),
        }

    common = set(prices.keys()) & set(hqa.keys())
    common_sorted = sort_dates(common)
    print(f"   ✓ قیمت: {len(prices)} روز، حقیقی/حقوقی: {len(hqa)} روز، مشترک: {len(common_sorted)} روز")

    mf  = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"]) for d in reversed(common_sorted)]
    vol = [(d, hqa[d]["volume"]) for d in common_sorted]
    bp  = [(d, hqa[d]["buy_real_value"], hqa[d]["sell_real_value"], hqa[d]["buy_real_count"],
           hqa[d]["sell_real_count"], hqa[d]["buy_legal_value"]) for d in common_sorted]
    pri = [{"date":d, **prices[d]} for d in common_sorted if d in prices]

    return {"money_flow":mf, "volume":vol, "buyer_power":bp,
            "prices":pri, "persistence":list(reversed(mf))}

# ============================================================
# ساخت شیت‌های معیار — با محاسبات Python
# ============================================================
def build_jaryan_pul(wb, ds, company):
    """معیار ۱: جریان پول حقیقی — همه محاسبات در Python"""
    mf = ds["money_flow"]; n = len(mf)
    if "جریان پول حقیقی" in wb.sheetnames: del wb["جریان پول حقیقی"]
    ws = wb.create_sheet("جریان پول حقیقی", 1)
    ws.sheet_view.showGridLines = False; ws.sheet_view.rightToLeft = True
    set_widths(ws, {"B":12,"C":17,"D":17,"E":17,"F":22,"G":14,"H":14,"I":10,"J":14})
    ws.row_dimensions[2].height=32; ws.row_dimensions[4].height=36
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=10)
    ws.cell(row=2, column=2, value=f"معیار ۱: جریان پول حقیقی — {company} ۱۴۰۵").font = Font(name=FONT_NAME, size=18, bold=True, color=PRIMARY)
    ws.cell(row=2, column=2).alignment = Alignment(horizontal="right", vertical="center", readingOrder=2)
    for i,h in enumerate(["تاریخ","ارزش خرید حقیقی\n(م. ریال)","ارزش فروش حقیقی\n(م. ریال)","خالص جریان پول\n(م. ریال)","گردش معاملات\n(م. ریال)","نسبت جریان پول\n(٪)","نسبت ۵ روزه\n(٪)","امتیاز\n(۰-۳۰)","وضعیت"]):
        c = ws.cell(row=4, column=2+i, value=h); c.fill=H_FILL; c.font=H_FONT; c.alignment=H_ALIGN; c.border=H_BORDER

    fdr=5
    for i,(d,bv,sv) in enumerate(mf):
        r=fdr+i; ws.row_dimensions[r].height=22
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,11): ws.cell(row=r,column=c).fill=fill; ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=d).alignment=B_ALIGN
        ws.cell(row=r,column=3,value=bv).alignment=B_ALIGN_NUM; ws.cell(row=r,column=3).number_format="#,##0.000"
        ws.cell(row=r,column=4,value=sv).alignment=B_ALIGN_NUM; ws.cell(row=r,column=4).number_format="#,##0.000"
        net = bv - sv
        c=ws.cell(row=r,column=5,value=net);c.alignment=B_ALIGN_NUM;c.number_format="#,##0.000;[Red]-#,##0.000"
        c.font=Font(name=FONT_NAME,size=10,bold=True,color=COLOR_GREEN if net>=0 else COLOR_RED)
        turnover = bv + sv
        ws.cell(row=r,column=6,value=turnover).alignment=B_ALIGN_NUM; ws.cell(row=r,column=6).number_format="#,##0.000"
        daily_ratio = safe_div(net, turnover) * 100
        ws.cell(row=r,column=7,value=daily_ratio).alignment=B_ALIGN_NUM; ws.cell(row=r,column=7).number_format="0.00"
        ratio_5d = calc_money_flow_5d(mf, i)
        score = calc_money_flow_score(ratio_5d)
        if ratio_5d is not None:
            ws.cell(row=r,column=8,value=ratio_5d).alignment=B_ALIGN_NUM; ws.cell(row=r,column=8).number_format="0.00"
            ws.cell(row=r,column=9,value=score).alignment=B_ALIGN_NUM; ws.cell(row=r,column=9).number_format="0"
            if ratio_5d>=10: st,sc="بسیار قوی",COLOR_DARK_GREEN
            elif ratio_5d>=5: st,sc="قوی",COLOR_GREEN
            elif ratio_5d>=2: st,sc="متوسط",PRIMARY
            elif ratio_5d>0: st,sc="کم",COLOR_ORANGE
            else: st,sc="ضعیف",COLOR_RED
        else:
            for c in [8,9,10]: ws.cell(row=r,column=c,value="—").alignment=B_ALIGN; ws.cell(row=r,column=c).font=Font(name=FONT_NAME,size=9,color=NEUTRAL_600)
            continue
        ws.cell(row=r,column=10,value=st).alignment=B_ALIGN; ws.cell(row=r,column=10).font=Font(name=FONT_NAME,size=10,bold=True,color=sc)
    ws.freeze_panes="C5"


def build_hajm(wb, ds, company):
    """معیار ۲: تحلیل حجم — همه محاسبات در Python"""
    vol=list(reversed(ds["volume"]));n=len(vol)
    if "تحلیل حجم" in wb.sheetnames: del wb["تحلیل حجم"]
    ws=wb.create_sheet("تحلیل حجم",2)
    ws.sheet_view.showGridLines=False;ws.sheet_view.rightToLeft=True
    set_widths(ws,{"B":6,"C":12,"D":16,"E":18,"F":18,"G":14,"H":14,"I":10,"J":14})
    ws.row_dimensions[1].height=24;ws.row_dimensions[2].height=40
    ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=10)
    ws.cell(row=1,column=2,value=f"معیار ۲: تحلیل حجم — {company} ({vol[-1][0]} تا {vol[0][0]})").font=Font(name=FONT_NAME,size=12,bold=True,color=PRIMARY)
    ws.cell(row=1,column=2).alignment=Alignment(horizontal="right",vertical="center",readingOrder=2)
    for i,h in enumerate(["ردیف","تاریخ","حجم امروز\n(سهم)","میانگین ۵ روز","میانگین ۳۰ روز","نسبت ۵ روزه","نسبت ۳۰ روزه","امتیاز\n(۰-۱۵)","شدت"]):
        c=ws.cell(row=2,column=2+i,value=h);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
    fdr=3
    for i,(d,v) in enumerate(vol):
        r=fdr+i;ws.row_dimensions[r].height=20
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,11):ws.cell(row=r,column=c).fill=fill;ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=i+1).alignment=B_ALIGN
        ws.cell(row=r,column=3,value=d).alignment=B_ALIGN
        ws.cell(row=r,column=4,value=v).alignment=B_ALIGN_NUM;ws.cell(row=r,column=4).number_format="#,##0"
        # میانگین ۵ روز
        avg_5 = calc_volume_avg(vol, i, 5)
        if avg_5 is not None:
            ws.cell(row=r,column=5,value=avg_5).alignment=B_ALIGN_NUM;ws.cell(row=r,column=5).number_format="#,##0"
            ratio_5 = safe_div(v, avg_5)
            ws.cell(row=r,column=7,value=ratio_5).alignment=B_ALIGN_NUM;ws.cell(row=r,column=7).number_format="0.00"
        else:
            ws.cell(row=r,column=5,value="—").alignment=B_ALIGN;ws.cell(row=r,column=5).font=Font(name=FONT_NAME,size=9,color=NEUTRAL_600)
            ws.cell(row=r,column=7,value="—").alignment=B_ALIGN;ws.cell(row=r,column=7).font=Font(name=FONT_NAME,size=9,color=NEUTRAL_600)
        # میانگین ۳۰ روز
        avg_30 = calc_volume_avg(vol, i, 30)
        if avg_30 is not None:
            ws.cell(row=r,column=6,value=avg_30).alignment=B_ALIGN_NUM;ws.cell(row=r,column=6).number_format="#,##0"
            ratio_30 = safe_div(v, avg_30)
            ws.cell(row=r,column=8,value=ratio_30).alignment=B_ALIGN_NUM;ws.cell(row=r,column=8).number_format="0.00"
            score = calc_volume_score(ratio_30)
            ws.cell(row=r,column=9,value=score).alignment=B_ALIGN
            ws.cell(row=r,column=10,value=calc_volume_intensity(score)).alignment=B_ALIGN
        else:
            for c in [6,8,9,10]:ws.cell(row=r,column=c,value="—").alignment=B_ALIGN;ws.cell(row=r,column=c).font=Font(name=FONT_NAME,size=9,color=NEUTRAL_600)
    ws.freeze_panes="D3"


def build_qudrat(wb, ds, company):
    """معیار ۳: قدرت خریدار — همه محاسبات در Python"""
    bp=list(reversed(ds["buyer_power"]));n=len(bp)
    if "قدرت خریدار" in wb.sheetnames: del wb["قدرت خریدار"]
    ws=wb.create_sheet("قدرت خریدار",3)
    ws.sheet_view.showGridLines=False;ws.sheet_view.rightToLeft=True
    set_widths(ws,{"B":6,"C":11,"D":16,"E":16,"F":14,"G":14,"H":18,"I":14,"J":14,"K":13,"L":9,"M":13,"N":9,"O":14,"P":9,"Q":12})
    ws.row_dimensions[1].height=24;ws.row_dimensions[2].height=24;ws.row_dimensions[3].height=36
    ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=17)
    ws.cell(row=1,column=2,value=f"معیار ۳: قدرت خریدار — {company}").font=Font(name=FONT_NAME,size=12,bold=True,color=PRIMARY)
    ws.cell(row=1,column=2).alignment=Alignment(horizontal="right",vertical="center",readingOrder=2)
    for s,e,t in [(2,2,"ردیف"),(3,3,"تاریخ"),(4,8,"ورودی‌های اصلی"),(9,12,"مؤلفه ۱ (۱۰)"),(13,14,"مؤلفه ۲ (۵)"),(15,16,"مؤلفه ۳ (۵)"),(17,17,"امتیاز کل\n(۰-۲۰)")]:
        if s==e:ws.merge_cells(start_row=2,start_column=s,end_row=3,end_column=s)
        else:ws.merge_cells(start_row=2,start_column=s,end_row=2,end_column=e)
        c=ws.cell(row=2,column=s,value=t);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
        for cc in range(s,e+1):ws.cell(row=2,column=cc).fill=H_FILL;ws.cell(row=2,column=cc).border=H_BORDER
    for col,h in [(4,"BuyRealValue"),(5,"SellRealValue"),(6,"BuyRealCount"),(7,"SellRealCount"),(8,"TotalTradeValue"),(9,"AverageBuy"),(10,"AverageSell"),(11,"PowerRatio"),(12,"امتیاز"),(13,"CountRatio"),(14,"امتیاز"),(15,"RealBuyShare\n(٪)"),(16,"امتیاز")]:
        c=ws.cell(row=3,column=col,value=h);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
    fdr=4
    for i,(d,bv,sv,bc,sc,lv) in enumerate(bp):
        r=fdr+i;ws.row_dimensions[r].height=20
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,18):ws.cell(row=r,column=c).fill=fill;ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=i+1).alignment=B_ALIGN
        ws.cell(row=r,column=3,value=d).alignment=B_ALIGN
        ws.cell(row=r,column=4,value=bv).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=5,value=sv).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=6,value=bc).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=7,value=sc).alignment=B_ALIGN_NUM
        total_value = bv + lv
        ws.cell(row=r,column=8,value=total_value).alignment=B_ALIGN_NUM
        avg_buy = safe_div(bv, bc)
        avg_sell = safe_div(sv, sc)
        power_ratio = safe_div(avg_buy, avg_sell)
        ws.cell(row=r,column=9,value=avg_buy).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=10,value=avg_sell).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=11,value=power_ratio).alignment=B_ALIGN_NUM
        count_ratio = safe_div(bc, sc)
        ws.cell(row=r,column=13,value=count_ratio).alignment=B_ALIGN_NUM
        real_buy_share = safe_div(bv, total_value) * 100
        ws.cell(row=r,column=15,value=real_buy_share).alignment=B_ALIGN_NUM
        total_score = calc_buyer_power_score(power_ratio, count_ratio, real_buy_share)
        # امتیاز جزئی برای نمایش
        if power_ratio >= 2: s_pow = 10
        elif power_ratio >= 1.5: s_pow = 8
        elif power_ratio >= 1.2: s_pow = 6
        elif power_ratio >= 1: s_pow = 4
        elif power_ratio >= 0.8: s_pow = 2
        else: s_pow = 0
        if count_ratio < 0.7: s_cnt = 5
        elif count_ratio < 0.9: s_cnt = 4
        elif count_ratio < 1.1: s_cnt = 3
        elif count_ratio < 1.3: s_cnt = 2
        elif count_ratio < 1.5: s_cnt = 1
        else: s_cnt = 0
        if real_buy_share >= 60: s_shr = 5
        elif real_buy_share >= 55: s_shr = 4
        elif real_buy_share >= 50: s_shr = 3
        elif real_buy_share >= 45: s_shr = 2
        elif real_buy_share >= 40: s_shr = 1
        else: s_shr = 0
        ws.cell(row=r,column=12,value=s_pow).alignment=B_ALIGN
        ws.cell(row=r,column=14,value=s_cnt).alignment=B_ALIGN
        ws.cell(row=r,column=16,value=s_shr).alignment=B_ALIGN
        c=ws.cell(row=r,column=17,value=total_score);c.alignment=B_ALIGN;c.font=Font(name=FONT_NAME,size=11,bold=True,color=PRIMARY)
    ws.freeze_panes="D4"


def build_vakonesh(wb, ds, company):
    """معیار ۴: واکنش قیمت — همه محاسبات در Python"""
    pri=list(reversed(ds["prices"]));n=len(pri)
    if "واکنش قیمت" in wb.sheetnames: del wb["واکنش قیمت"]
    ws=wb.create_sheet("واکنش قیمت",4)
    ws.sheet_view.showGridLines=False;ws.sheet_view.rightToLeft=True
    set_widths(ws,{"B":6,"C":11,"D":13,"E":16,"F":13,"G":13,"H":13,"I":13,"J":9,"K":13,"L":9,"M":13,"N":9,"O":13,"P":9,"Q":12})
    ws.row_dimensions[1].height=24;ws.row_dimensions[2].height=24;ws.row_dimensions[3].height=36
    ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=17)
    ws.cell(row=1,column=2,value=f"معیار ۴: واکنش قیمت — {company}").font=Font(name=FONT_NAME,size=12,bold=True,color=PRIMARY)
    ws.cell(row=1,column=2).alignment=Alignment(horizontal="right",vertical="center",readingOrder=2)
    for s,e,t in [(2,2,"ردیف"),(3,3,"تاریخ"),(4,8,"ورودی‌های قیمت"),(9,10,"مؤلفه ۱ (۸)"),(11,12,"مؤلفه ۲ (۵)"),(13,14,"مؤلفه ۳ (۷)"),(15,16,"مؤلفه ۴ (۵)"),(17,17,"امتیاز کل\n(۰-۲۵)")]:
        if s==e:ws.merge_cells(start_row=2,start_column=s,end_row=3,end_column=s)
        else:ws.merge_cells(start_row=2,start_column=s,end_row=2,end_column=e)
        c=ws.cell(row=2,column=s,value=t);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
        for cc in range(s,e+1):ws.cell(row=2,column=cc).fill=H_FILL;ws.cell(row=2,column=cc).border=H_BORDER
    for col,h in [(4,"Close\n(پایانی)"),(5,"PrevClose\n(دیروز)"),(6,"Last\n(آخرین)"),(7,"High\n(بیشترین)"),(8,"Low\n(کمترین)"),(9,"R_c (٪)"),(10,"امتیاز"),(11,"R_l (٪)"),(12,"امتیاز"),(13,"Position (٪)"),(14,"امتیاز"),(15,"Recovery (٪)"),(16,"امتیاز")]:
        c=ws.cell(row=3,column=col,value=h);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
    fdr=4
    for i,p in enumerate(pri):
        r=fdr+i;ws.row_dimensions[r].height=20
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,18):ws.cell(row=r,column=c).fill=fill;ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=i+1).alignment=B_ALIGN
        ws.cell(row=r,column=3,value=p["date"]).alignment=B_ALIGN
        for col,key in [(4,"close"),(5,"prev_close"),(6,"last"),(7,"high"),(8,"low")]:
            c=ws.cell(row=r,column=col,value=p[key]);c.alignment=B_ALIGN_NUM;c.number_format="#,##0";c.fill=INPUT_FILL;c.font=INPUT_FONT
        # محاسبه همه مؤلفه‌ها — استفاده از rc_pct مستقیم از فایل input
        rc, s1, rl, s2, position, s3, recovery, s4, total = calc_price_scores(
            p["close"], p["prev_close"], p["last"], p["high"], p["low"],
            rc_pct=p.get("rc_pct")
        )
        ws.cell(row=r,column=9,value=rc).alignment=B_ALIGN_NUM;ws.cell(row=r,column=9).number_format="0.00"
        ws.cell(row=r,column=10,value=s1).alignment=B_ALIGN
        ws.cell(row=r,column=11,value=rl).alignment=B_ALIGN_NUM;ws.cell(row=r,column=11).number_format="0.00"
        ws.cell(row=r,column=12,value=s2).alignment=B_ALIGN
        ws.cell(row=r,column=13,value=position).alignment=B_ALIGN_NUM;ws.cell(row=r,column=13).number_format="0.00"
        ws.cell(row=r,column=14,value=s3).alignment=B_ALIGN
        ws.cell(row=r,column=15,value=recovery).alignment=B_ALIGN_NUM;ws.cell(row=r,column=15).number_format="0.00"
        ws.cell(row=r,column=16,value=s4).alignment=B_ALIGN
        c=ws.cell(row=r,column=17,value=total);c.alignment=B_ALIGN;c.font=Font(name=FONT_NAME,size=11,bold=True,color=PRIMARY)
    ws.freeze_panes="D4"


def build_tadavom(wb, ds, company):
    """معیار ۵: تداوم جریان — همه محاسبات در Python"""
    mf=ds["money_flow"];n=len(mf)
    if "تداوم جریان پول حقیقی" in wb.sheetnames: del wb["تداوم جریان پول حقیقی"]
    ws=wb.create_sheet("تداوم جریان پول حقیقی",5)
    ws.sheet_view.showGridLines=False;ws.sheet_view.rightToLeft=True
    set_widths(ws,{"B":6,"C":11,"D":14,"E":14,"F":14,"G":14,"H":16,"I":11,"J":14,"K":14,"L":14})
    ws.row_dimensions[1].height=24;ws.row_dimensions[2].height=40
    ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=12)
    ws.cell(row=1,column=2,value=f"معیار ۵: تداوم جریان — {company}").font=Font(name=FONT_NAME,size=12,bold=True,color=PRIMARY)
    ws.cell(row=1,column=2).alignment=Alignment(horizontal="right",vertical="center",readingOrder=2)
    for i,h in enumerate(["ردیف","تاریخ","BuyRealValue","SellRealValue","NetFlow","RealTurnover","DailyFlowRatio\n(٪)","Status","PositiveDays","Score\n(۰-۱۰)","وضعیت"]):
        c=ws.cell(row=2,column=2+i,value=h);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER
    fdr=3
    for i,(d,bv,sv) in enumerate(mf):
        r=fdr+i;ws.row_dimensions[r].height=20
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,13):ws.cell(row=r,column=c).fill=fill;ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=i+1).alignment=B_ALIGN
        ws.cell(row=r,column=3,value=d).alignment=B_ALIGN
        ws.cell(row=r,column=4,value=bv).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=5,value=sv).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=6,value=bv-sv).alignment=B_ALIGN_NUM
        ws.cell(row=r,column=7,value=bv+sv).alignment=B_ALIGN_NUM
        ratio = safe_div(bv - sv, bv + sv) * 100
        ws.cell(row=r,column=8,value=ratio).alignment=B_ALIGN_NUM;ws.cell(row=r,column=8).number_format="0.00"
        if ratio > 0.5: status = 1
        elif ratio < -0.5: status = -1
        else: status = 0
        ws.cell(row=r,column=9,value=status).alignment=B_ALIGN
        pos_days, score, state = calc_persistence(mf, i)
        if pos_days is not None:
            ws.cell(row=r,column=10,value=pos_days).alignment=B_ALIGN
            ws.cell(row=r,column=11,value=score).alignment=B_ALIGN
            ws.cell(row=r,column=12,value=state).alignment=B_ALIGN
        else:
            for c in [10,11,12]:ws.cell(row=r,column=c,value="—").alignment=B_ALIGN;ws.cell(row=r,column=c).font=Font(name=FONT_NAME,size=9,color=NEUTRAL_600)
    ws.freeze_panes="D3"


def build_dashboard_sheet(wb, ds, company):
    """ساخت شیت اصلی داشبورد — همه محاسبات در Python"""
    mf = ds["money_flow"]
    vol = list(reversed(ds["volume"]))
    pri = list(reversed(ds["prices"]))
    bp = list(reversed(ds["buyer_power"]))
    n = len(mf)
    dates = [mf[i][0] for i in range(n)]  # newest first

    if "داشبورد" in wb.sheetnames: del wb["داشبورد"]
    ws=wb.create_sheet("داشبورد",0)
    ws.sheet_view.showGridLines=False; ws.sheet_view.rightToLeft=True; ws.column_dimensions["A"].width=3
    for c,w in {"B":11,"C":16,"D":12,"E":14,"F":13,"G":10,"H":10,"I":11,"J":14,"K":12}.items():
        ws.column_dimensions[c].width=w
    ws.row_dimensions[1].height=24; ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=11)
    ws.cell(row=1,column=2,value=f"داشبورد {company} — {dates[-1]} تا {dates[0]} ({n} روز) | امتیاز: ۰-۱۰۰").font=Font(name=FONT_NAME,size=12,bold=True,color=PRIMARY)
    ws.cell(row=1,column=2).alignment=Alignment(horizontal="right",vertical="center",readingOrder=2)
    ws.row_dimensions[2].height=40
    for i,h in enumerate(["تاریخ","امتیاز\nجریان پول\n(۰-۳۰)","امتیاز\nحجم\n(۰-۱۵)","امتیاز\nقدرت خریدار\n(۰-۲۰)","امتیاز\nواکنش قیمت\n(۰-۲۵)","امتیاز\nتداوم\n(۰-۱۰)","امتیاز کل\n(۰-۱۰۰)","شدت","٪ قیمت\nپایانی (R_c)","سیگنال"]):
        c=ws.cell(row=2,column=2+i,value=h);c.fill=H_FILL;c.font=H_FONT;c.alignment=H_ALIGN;c.border=H_BORDER

    fdr=3
    all_data = []
    for i in range(n):
        # محاسبه همه امتیازها برای روز i
        # جریان پول
        ratio_5d_mf = calc_money_flow_5d(mf, i)
        score_mf = calc_money_flow_score(ratio_5d_mf) or 0
        # حجم
        # vol = reversed یعنی newest اول، پس vol[i] = امروز
        avg_30 = calc_volume_avg(vol, i, 30)  # میانگین ۳۰ روز گذشته (از فردا به بعد)
        if avg_30 is not None:
            v_today = vol[i][1]
            ratio_30 = safe_div(v_today, avg_30)
            score_vol = calc_volume_score(ratio_30) or 0
        else:
            score_vol = 0
        # قدرت خریدار
        d, bv, sv, bc, sc, lv = bp[i]  # bp reversed یعنی newest اول
        avg_buy = safe_div(bv, bc)
        avg_sell = safe_div(sv, sc)
        power_ratio = safe_div(avg_buy, avg_sell)
        count_ratio = safe_div(bc, sc)
        total_value = bv + lv
        real_buy_share = safe_div(bv, total_value) * 100
        score_bp = calc_buyer_power_score(power_ratio, count_ratio, real_buy_share) or 0
        # واکنش قیمت
        # pri = list(reversed(ds["prices"])) یعنی newest اول است
        # پس برای i=0 (امروز) باید p = pri[0] باشد، نه pri[n-1-0]
        p = pri[i]
        rc, s1, rl, s2, position, s3, recovery, s4, score_price = calc_price_scores(
            p["close"], p["prev_close"], p["last"], p["high"], p["low"],
            rc_pct=p.get("rc_pct")
        )
        # تداوم
        pos_days, score_pers, state_pers = calc_persistence(mf, i)
        score_pers = score_pers or 0
        # امتیاز کل
        total_score = score_mf + score_vol + score_bp + score_price + score_pers
        all_data.append({
            "date": dates[i],
            "score_mf": score_mf,
            "score_vol": score_vol,
            "score_bp": score_bp,
            "score_price": score_price,
            "score_pers": score_pers,
            "total": total_score,
            "rc": rc,
        })

    # نوشتن داده‌ها
    for i, data in enumerate(all_data):
        r = fdr + i
        ws.row_dimensions[r].height=20
        fill=PatternFill("solid",fgColor=NEUTRAL_0 if i%2==0 else NEUTRAL_100)
        for c in range(2,12): ws.cell(row=r,column=c).fill=fill; ws.cell(row=r,column=c).font=B_FONT
        ws.cell(row=r,column=2,value=data["date"]).alignment=B_ALIGN
        # جریان پول
        sc_mf = get_score_color(data["score_mf"])
        si_mf = int(data["score_mf"]) if data["score_mf"] == int(data["score_mf"]) else data["score_mf"]
        c=ws.cell(row=r,column=3,value=f"● {si_mf}");c.alignment=B_ALIGN;c.font=Font(name=FONT_NAME,size=11,bold=True,color=sc_mf)
        # حجم
        ws.cell(row=r,column=4,value=data["score_vol"]).alignment=B_ALIGN
        # قدرت خریدار
        ws.cell(row=r,column=5,value=data["score_bp"]).alignment=B_ALIGN
        # واکنش قیمت
        ws.cell(row=r,column=6,value=data["score_price"]).alignment=B_ALIGN
        # تداوم
        if data["score_pers"] > 0:
            ws.cell(row=r,column=7,value=data["score_pers"]).alignment=B_ALIGN
        else:
            ws.cell(row=r,column=7,value=0).alignment=B_ALIGN
        # امتیاز کل
        ws.cell(row=r,column=8,value=data["total"]).alignment=B_ALIGN
        ws.cell(row=r,column=8).font=Font(name=FONT_NAME,size=10,bold=True,color=PRIMARY)
        # شدت
        tv = data["total"]
        if tv>=70: it,ic="بسیار قوی",COLOR_GREEN
        elif tv>=50: it,ic="قوی",COLOR_GREEN
        elif tv>=30: it,ic="متوسط",COLOR_YELLOW
        else: it,ic="ضعیف",COLOR_RED
        ws.cell(row=r,column=9,value=it).alignment=B_ALIGN; ws.cell(row=r,column=9).font=Font(name=FONT_NAME,size=9,bold=True,color=ic)
        # ٪ قیمت پایانی
        rv = data["rc"]
        rclr=get_rc_color(rv); rs="+" if rv>=0 else ""
        c=ws.cell(row=r,column=10,value=f"● {rs}{rv:.2f}");c.alignment=B_ALIGN;c.font=Font(name=FONT_NAME,size=10,bold=True,color=rclr)
        # سیگنال
        # خالص جریان پول روز جاری
        bv_today, sv_today = mf[i][1], mf[i][2]
        net_mf = bv_today - sv_today
        if tv>=50 and net_mf>0: sg,sgc="خرید قوی",COLOR_GREEN
        elif tv>=30 and net_mf>0: sg,sgc="خرید",COLOR_GREEN
        elif tv<=30 and net_mf<0: sg,sgc="فروش",COLOR_RED
        elif tv<=50 and net_mf<0: sg,sgc="هشدار",COLOR_YELLOW
        else: sg,sgc="خنثی",NEUTRAL_900
        ri = (it == "ضعیف"); rss = (sg == "فروش")
        if ri and rss: c.fill=PatternFill("solid",fgColor=COLOR_RED_BG)
        elif ri or rss: c.fill=PatternFill("solid",fgColor=COLOR_YELLOW_BG)
        ws.cell(row=r,column=11,value=sg).alignment=B_ALIGN; ws.cell(row=r,column=11).font=Font(name=FONT_NAME,size=9,bold=True,color=sgc)

    # ردیف میانگین
    ar = fdr + n
    ws.row_dimensions[ar].height=24
    af=PatternFill("solid",fgColor=PRIMARY_LIGHT); afo=Font(name=FONT_NAME,size=10,bold=True,color=PRIMARY)
    ab=Border(top=Side(style="medium",color=NEUTRAL_200))
    for c in range(2,12): ws.cell(row=ar,column=c).fill=af;ws.cell(row=ar,column=c).font=afo;ws.cell(row=ar,column=c).border=ab
    ws.cell(row=ar,column=2,value="میانگین").alignment=B_ALIGN
    avg_mf = sum(d["score_mf"] for d in all_data) / n
    ws.cell(row=ar,column=3,value=f"● {avg_mf:.1f}").alignment=B_ALIGN
    ws.cell(row=ar,column=3).font=Font(name=FONT_NAME,size=11,bold=True,color=get_score_color(avg_mf))
    avg_vol = sum(d["score_vol"] for d in all_data) / n
    ws.cell(row=ar,column=4,value=f"{avg_vol:.1f}").alignment=B_ALIGN
    avg_bp = sum(d["score_bp"] for d in all_data) / n
    ws.cell(row=ar,column=5,value=f"{avg_bp:.1f}").alignment=B_ALIGN
    avg_price = sum(d["score_price"] for d in all_data) / n
    ws.cell(row=ar,column=6,value=f"{avg_price:.1f}").alignment=B_ALIGN
    avg_pers = sum(d["score_pers"] for d in all_data) / n
    ws.cell(row=ar,column=7,value=f"{avg_pers:.1f}").alignment=B_ALIGN
    avg_total = sum(d["total"] for d in all_data) / n
    ws.cell(row=ar,column=8,value=f"{avg_total:.1f}").alignment=B_ALIGN
    ws.cell(row=ar,column=9,value="میانگین").alignment=B_ALIGN
    avg_rc = sum(d["rc"] for d in all_data) / n
    rs = "+" if avg_rc >= 0 else ""
    ws.cell(row=ar,column=10,value=f"● {rs}{avg_rc:.2f}").alignment=B_ALIGN
    ws.cell(row=ar,column=10).font=Font(name=FONT_NAME,size=10,bold=True,color=get_rc_color(avg_rc))
    ws.cell(row=ar,column=11,value="—").alignment=B_ALIGN
    ws.freeze_panes="C3"


# ============================================================
# تابع اصلی
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="ساخت داشبورد از فایل input (مستقل از LibreOffice)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
مثال‌ها:
  python make_dashboard.py
  python make_dashboard.py سرود_input.xlsx
  python make_dashboard.py سرود_input.xlsx --name سرود
  python make_dashboard.py سرود_input.xlsx --output داشبورد.xlsx
        """)
    parser.add_argument("input_file", nargs="?", help="فایل ورودی input.xlsx")
    parser.add_argument("--name", default=None, help="نام شرکت برای عنوان داشبورد")
    parser.add_argument("--output", default=None, help="مسیر فایل خروجی داشبورد")
    args = parser.parse_args()

    print("=" * 60)
    print("ساخت داشبورد از فایل input (نسخه مستقل — بدون LibreOffice)")
    print("=" * 60)

    # پیدا کردن فایل ورودی
    input_file = args.input_file
    if not input_file:
        for fn in os.listdir("."):
            if fn.endswith("_input.xlsx") and not fn.startswith("~"):
                input_file = fn
                break
        if not input_file:
            print("\n❌ فایل ورودی مشخص نشد!")
            print("\nروش استفاده:")
            print("  python make_dashboard.py سرود_input.xlsx")
            return 1
        print(f"\n📁 فایل ورودی به‌صورت خودکار پیدا شد: {input_file}")

    if not os.path.exists(input_file):
        print(f"\n❌ فایل یافت نشد: {input_file}")
        return 1

    company = args.name or os.path.splitext(os.path.basename(input_file))[0].replace("_input", "")

    if args.output:
        output_file = args.output
    else:
        base = os.path.splitext(os.path.basename(input_file))[0].replace("_input", "")
        output_file = f"{base}_dashboard.xlsx"

    print(f"\n📦 فایل ورودی: {input_file}")
    print(f"   نام شرکت:   {company}")
    print(f"   خروجی:      {output_file}")

    # مرحله ۱: خواندن داده‌ها
    print(f"\n━━━ مرحله ۱: خواندن داده‌ها ━━━")
    try:
        ds = load_data(input_file)
    except Exception as e:
        print(f"\n❌ خطا در خواندن فایل: {e}")
        return 1

    # مرحله ۲: ساخت شیت‌های داشبورد
    print(f"\n━━━ مرحله ۲: ساخت شیت‌های داشبورد (محاسبات در Python) ━━━")
    if os.path.exists(output_file):
        wb = load_workbook(output_file)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames: del wb["Sheet"]

    print(f"   ساخت شیت ۱: جریان پول حقیقی...")
    build_jaryan_pul(wb, ds, company)
    print(f"   ساخت شیت ۲: تحلیل حجم...")
    build_hajm(wb, ds, company)
    print(f"   ساخت شیت ۳: قدرت خریدار...")
    build_qudrat(wb, ds, company)
    print(f"   ساخت شیت ۴: واکنش قیمت...")
    build_vakonesh(wb, ds, company)
    print(f"   ساخت شیت ۵: تداوم جریان پول...")
    build_tadavom(wb, ds, company)
    print(f"   ساخت شیت اصلی داشبورد...")
    build_dashboard_sheet(wb, ds, company)

    wb.save(output_file)
    print(f"\n{'='*60}")
    print(f"✓ داشبورد ساخته شد: {output_file}")
    print(f"{'='*60}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
