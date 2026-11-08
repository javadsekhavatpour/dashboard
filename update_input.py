#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
فاز ۱: ایجاد/به‌روزرسانی فایل input با داده ۶۰ روزه از TSETMC
================================================================

روش استفاده:
  python update_input.py --symbol سرود --create
  python update_input.py --symbol سرود --update
  python update_input.py --symbol سرود,دلر,کپارس --create --dashboard
  python update_input.py --ins-code 3358793329461860 --create
"""

import sys, os, argparse, json, re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# تنظیمات
# ============================================================
TSETMC_API_BASE = "https://cdn.tsetmc.com/api"
TSETMC_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Referer": "https://old.tsetmc.com/",
}
DEFAULT_DAYS = 60

# ============================================================
# استایل‌ها
# ============================================================
FONT_NAME = "Vazirmatn"
PRIMARY = "1B2A4A"
H_FILL = PatternFill("solid", fgColor=PRIMARY)
H_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
H_ALIGN = Alignment(horizontal="center", vertical="center", readingOrder=2, wrap_text=True)
H_BORDER = Border(bottom=Side(style="thin", color="E9E9E8"))
B_FONT = Font(name=FONT_NAME, size=9, color="37352F")
B_ALIGN = Alignment(horizontal="center", vertical="center", readingOrder=2)
B_ALIGN_NUM = Alignment(horizontal="left", vertical="center")
TITLE_FONT = Font(name=FONT_NAME, size=12, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1B2A4A")
GROUP_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
GROUP_FILL = PatternFill("solid", fgColor="2C3E50")
TOT_FONT = Font(name=FONT_NAME, size=10, bold=True, color="1B2A4A")
TOT_FILL = PatternFill("solid", fgColor="E8ECF4")
EVEN_FILL = PatternFill("solid", fgColor="F7F7F5")
ODD_FILL = PatternFill("solid", fgColor="FFFFFF")

# ============================================================
# تبدیل تاریخ میلادی → شمسی
# ============================================================
J_MONTHS = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور",
            "مهر","آبان","آذر","دی","بهمن","اسفند"]

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
    return jy, jm, jd

def deven_to_jdate_full(deven):
    deven = int(deven)
    gy = deven // 10000; gm = (deven // 100) % 100; gd = deven % 100
    jy, jm, jd = gregorian_to_jalali(gy, gm, gd)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

def deven_to_jdate_fa(deven):
    deven = int(deven)
    gy = deven // 10000; gm = (deven // 100) % 100; gd = deven % 100
    jy, jm, jd = gregorian_to_jalali(gy, gm, gd)
    return f"{J_MONTHS[jm-1]} {jd}"

# ============================================================
# ارتباط با TSETMC API
# ============================================================
def api_get(path, timeout=30):
    import requests
    url = f"{TSETMC_API_BASE}{path}"
    r = requests.get(url, headers=TSETMC_HEADERS, timeout=timeout)
    r.raise_for_status()
    return r.json()

def find_ins_code(symbol_name):
    from urllib.parse import quote
    encoded = quote(symbol_name)
    data = api_get(f"/Instrument/GetInstrumentSearch/{encoded}")
    candidates = data.get("instrumentSearch", [])
    if not candidates:
        return None, None, None
    for item in candidates:
        if item.get("lVal18AFC", "").strip() == symbol_name:
            return str(item["insCode"]), item.get("lVal30","").strip(), item.get("lVal18AFC","").strip()
    item = candidates[0]
    return str(item["insCode"]), item.get("lVal30","").strip(), item.get("lVal18AFC","").strip()

def fetch_price_history(ins_code, days=60):
    data = api_get(f"/ClosingPrice/GetClosingPriceDailyList/{ins_code}/{days}")
    items = data.get("closingPriceDaily", [])
    return sorted(items, key=lambda x: x["dEven"])

def fetch_hqa_history(ins_code):
    data = api_get(f"/ClientType/GetClientTypeHistory/{ins_code}")
    items = data.get("clientType", [])
    return sorted(items, key=lambda x: x["recDate"])

# ============================================================
# دریافت داده‌های زنده (روز جاری - لایو)
# ============================================================
def fetch_live_price(ins_code):
    """دریافت قیمت زنده روز جاری از TSETMC.
    این داده‌ها در طول روز به‌روز می‌شوند.
    """
    try:
        data = api_get(f"/ClosingPrice/GetClosingPriceInfo/{ins_code}")
        info = data.get("closingPriceInfo", {})
        if not info:
            print(f"  WARNING: No live price data")
            return None
        
        # ساخت دیکشنری با همان ساختار historical data
        live = {
            "dEven": info.get("dEven", 0),
            "priceYesterday": float(info.get("priceYesterday", 0) or 0),
            "priceFirst": float(info.get("priceFirst", 0) or 0),
            "pClosing": float(info.get("pClosing", 0) or 0),
            "pDrCotVal": float(info.get("pDrCotVal", 0) or 0),
            "priceMin": float(info.get("priceMin", 0) or 0),
            "priceMax": float(info.get("priceMax", 0) or 0),
            "zTotTran": float(info.get("zTotTran", 0) or 0),
            "qTotTran5J": float(info.get("qTotTran5J", 0) or 0),
            "qTotCap": float(info.get("qTotCap", 0) or 0),
        }
        
        # محاسبه priceChange از قیمت پایانی
        if live["priceYesterday"] > 0:
            live["priceChange"] = live["pClosing"] - live["priceYesterday"]
        else:
            live["priceChange"] = 0
        
        print(f"  ✓ Live price: close={int(live['pClosing'])}, last={int(live['pDrCotVal'])}")
        return live
    except Exception as e:
        print(f"  WARNING: Cannot fetch live price: {e}")
        return None

def fetch_live_hqa(ins_code):
    """دریافت داده‌های حقیقی/حقوقی زنده روز جاری."""
    try:
        data = api_get(f"/ClientType/GetClientType/{ins_code}/0/0")
        ct = data.get("clientType", {})
        if not ct:
            print(f"  WARNING: No live HQA data")
            return None
        
        # ساخت دیکشنری با همان ساختار historical data
        live = {
            "recDate": int(datetime.now().strftime("%Y%m%d")),  # تاریخ امروز
            "buy_I_Volume": float(ct.get("buy_I_Volume", 0) or 0),
            "buy_I_Value": float(ct.get("buy_I_Value", 0) or 0),
            "buy_I_Count": float(ct.get("buy_I_Count", 0) or 0),
            "buy_N_Volume": float(ct.get("buy_N_Volume", 0) or 0),
            "buy_N_Value": float(ct.get("buy_N_Value", 0) or 0),
            "buy_N_Count": float(ct.get("buy_N_Count", 0) or 0),
            "sell_I_Volume": float(ct.get("sell_I_Volume", 0) or 0),
            "sell_I_Value": float(ct.get("sell_I_Value", 0) or 0),
            "sell_I_Count": float(ct.get("sell_I_Count", 0) or 0),
            "sell_N_Volume": float(ct.get("sell_N_Volume", 0) or 0),
            "sell_N_Value": float(ct.get("sell_N_Value", 0) or 0),
            "sell_N_Count": float(ct.get("sell_N_Count", 0) or 0),
        }
        
        buy_real = live["buy_I_Value"] / 1e9
        sell_real = live["sell_I_Value"] / 1e9
        print(f"  ✓ Live HQA: buy={buy_real:.3f}, sell={sell_real:.3f} B IRR")
        return live
    except Exception as e:
        print(f"  WARNING: Cannot fetch live HQA: {e}")
        return None

# ============================================================
# ساخت فایل input
# ============================================================
def build_input_file(symbol, ins_code, full_name, prices, hqa_data, output_file=None):
    if output_file is None:
        output_file = f"{symbol}_input.xlsx"
    wb = Workbook()
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    
    # ===== شیت ۱: سابقه قیمت و معاملات =====
    ws1 = wb.create_sheet("سابقه قیمت و معاملات")
    ws1.sheet_view.rightToLeft = True
    ws1.sheet_view.showGridLines = False
    widths_p = {"A":86,"B":20,"C":22,"D":23,"E":21,"F":21,"G":22,"H":28,
                "I":26,"J":21,"K":27,"L":26,"M":22,"N":23}
    for col, w in widths_p.items(): ws1.column_dimensions[col].width = w
    
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    c = ws1.cell(row=1, column=1, value=f"سابقه جامع قیمت و معاملات {full_name} — بازار بورس")
    c.font = TITLE_FONT; c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    ws1.row_dimensions[1].height = 28
    
    headers_p = ["تاریخ","تعداد معاملات","حجم معاملات","ارزش معاملات (ریال)","قیمت دیروز (ریال)",
                 "اولین قیمت (ریال)","قیمت پایانی (ریال)","تغییر قیمت پایانی (ریال)","درصد تغییر قیمت پایانی",
                 "آخرین قیمت (ریال)","تغییر آخرین قیمت (ریال)","درصد تغییر آخرین قیمت",
                 "کمترین قیمت (ریال)","بیشترین قیمت (ریال)"]
    for i, h in enumerate(headers_p, 1):
        c = ws1.cell(row=2, column=i, value=h)
        c.font = H_FONT; c.fill = H_FILL; c.alignment = H_ALIGN; c.border = H_BORDER
    ws1.row_dimensions[2].height = 36
    
    prices_newest_first = list(reversed(prices))
    for i, p in enumerate(prices_newest_first):
        r = 3 + i
        deven = p["dEven"]
        prev_close = int(p["priceYesterday"])
        first = int(p["priceFirst"])
        close = int(p["pClosing"])
        last = int(p["pDrCotVal"])
        low = int(p["priceMin"])
        high = int(p["priceMax"])
        # محاسبه تغییر قیمت پایانی (نه آخرین قیمت)
        change = close - prev_close
        count = int(p["zTotTran"])
        volume = int(p["qTotTran5J"])
        value = int(p["qTotCap"])
        date_fa = deven_to_jdate_full(deven)
        close_pct = change / prev_close if prev_close else 0
        last_change = last - prev_close
        last_pct = last_change / prev_close if prev_close else 0
        fill = ODD_FILL if i % 2 == 0 else EVEN_FILL
        ws1.row_dimensions[r].height = 18
        vals = [
            (1,  date_fa, "General", B_ALIGN),
            (2,  count, "#,##0", B_ALIGN_NUM),
            (3,  volume, "#,##0", B_ALIGN_NUM),
            (4,  value, "#,##0", B_ALIGN_NUM),
            (5,  prev_close, "#,##0", B_ALIGN_NUM),
            (6,  first, "#,##0", B_ALIGN_NUM),
            (7,  close, "#,##0", B_ALIGN_NUM),
            (8,  change, "#,##0;[Red]-#,##0", B_ALIGN_NUM),
            (9,  close_pct, "+0.00%;-0.00%;0.00%", B_ALIGN_NUM),
            (10, last, "#,##0", B_ALIGN_NUM),
            (11, last_change, "#,##0;[Red]-#,##0", B_ALIGN_NUM),
            (12, last_pct, "+0.00%;-0.00%;0.00%", B_ALIGN_NUM),
            (13, low, "#,##0", B_ALIGN_NUM),
            (14, high, "#,##0", B_ALIGN_NUM),
        ]
        for col, val, fmt, align in vals:
            cell = ws1.cell(row=r, column=col, value=val)
            cell.font = B_FONT; cell.fill = fill
            cell.alignment = align; cell.number_format = fmt
    ws1.freeze_panes = "B3"
    
    # ===== شیت ۲: حقیقی و حقوقی =====
    ws2 = wb.create_sheet("حقیقی و حقوقی")
    ws2.sheet_view.rightToLeft = True
    ws2.sheet_view.showGridLines = False
    widths_h = {"A":86,"B":11,"C":15,"D":15,"E":15,"F":11,"G":15,"H":15,"I":15,
                "J":11,"K":15,"L":15,"M":15,"N":11,"O":15,"P":15,"Q":15,"R":18,"S":15,"T":18}
    for col, w in widths_h.items(): ws2.column_dimensions[col].width = w
    
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    c = ws2.cell(row=1, column=1, value=f"اطلاعات حقیقی و حقوقی {full_name} — سال ۱۴۰۵")
    c.font = TITLE_FONT; c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    ws2.row_dimensions[1].height = 26
    ws2.row_dimensions[2].height = 6
    
    groups = [(1,1,"تاریخ",False),(2,5,"خرید حقیقی",True),(6,9,"خرید حقوقی",True),
              (10,13,"فروش حقیقی",True),(14,17,"فروش حقوقی",True),(18,20,"خلاصه و تحلیل",True)]
    for s, e, t, _ in groups:
        if s == e:
            ws2.merge_cells(start_row=3, start_column=s, end_row=4, end_column=s)
        else:
            ws2.merge_cells(start_row=3, start_column=s, end_row=3, end_column=e)
        c = ws2.cell(row=3, column=s, value=t)
        c.font = GROUP_FONT; c.fill = GROUP_FILL
        c.alignment = H_ALIGN; c.border = H_BORDER
        for cc in range(s, e+1):
            ws2.cell(row=3, column=cc).fill = GROUP_FILL
            ws2.cell(row=3, column=cc).border = H_BORDER
    ws2.row_dimensions[3].height = 22
    
    sub_headers = [None,"تعداد","حجم (سهم)","ارزش (میلیارد ریال)","قیمت میانگین (ریال)",
                   "تعداد","حجم (سهم)","ارزش (میلیارد ریال)","قیمت میانگین (ریال)",
                   "تعداد","حجم (سهم)","ارزش (میلیارد ریال)","قیمت میانگین (ریال)",
                   "تعداد","حجم (سهم)","ارزش (میلیارد ریال)","قیمت میانگین (ریال)",
                   "تغییر مالکیت حقوقی به حقیقی (سهم)","قدرت خریدار حقیقی به فروشنده","ورود/خروج پول حقیقی (میلیارد تومان)"]
    for i, h in enumerate(sub_headers, 1):
        if h is None: continue
        c = ws2.cell(row=4, column=i, value=h)
        c.font = H_FONT; c.fill = H_FILL; c.alignment = H_ALIGN; c.border = H_BORDER
    ws2.row_dimensions[4].height = 36
    
    hqa_newest_first = list(reversed(hqa_data))
    for i, h in enumerate(hqa_newest_first):
        r = 5 + i
        rec_date = h["recDate"]
        buy_I_vol = float(h.get("buy_I_Volume", 0) or 0)
        buy_I_val = float(h.get("buy_I_Value", 0) or 0)
        buy_I_cnt = float(h.get("buy_I_Count", 0) or 0)
        buy_N_vol = float(h.get("buy_N_Volume", 0) or 0)
        buy_N_val = float(h.get("buy_N_Value", 0) or 0)
        buy_N_cnt = float(h.get("buy_N_Count", 0) or 0)
        sell_I_vol = float(h.get("sell_I_Volume", 0) or 0)
        sell_I_val = float(h.get("sell_I_Value", 0) or 0)
        sell_I_cnt = float(h.get("sell_I_Count", 0) or 0)
        sell_N_vol = float(h.get("sell_N_Volume", 0) or 0)
        sell_N_val = float(h.get("sell_N_Value", 0) or 0)
        sell_N_cnt = float(h.get("sell_N_Count", 0) or 0)
        date_fa = deven_to_jdate_fa(rec_date)
        buy_I_avg = buy_I_val / buy_I_vol if buy_I_vol else 0
        buy_N_avg = buy_N_val / buy_N_vol if buy_N_vol else 0
        sell_I_avg = sell_I_val / sell_I_vol if sell_I_vol else 0
        sell_N_avg = sell_N_val / sell_N_vol if sell_N_vol else 0
        buy_I_val_b = buy_I_val / 1e9
        buy_N_val_b = buy_N_val / 1e9
        sell_I_val_b = sell_I_val / 1e9
        sell_N_val_b = sell_N_val / 1e9
        ownership_change = buy_I_vol - sell_I_vol
        power_ratio = buy_I_val / sell_I_val if sell_I_val else 0
        money_flow = (buy_I_val - sell_I_val) / 1e10
        fill = ODD_FILL if i % 2 == 0 else EVEN_FILL
        ws2.row_dimensions[r].height = 18
        vals = [
            (1,  date_fa, "General", B_ALIGN),
            (2,  int(buy_I_cnt), "#,##0", B_ALIGN_NUM),
            (3,  int(buy_I_vol), "#,##0", B_ALIGN_NUM),
            (4,  buy_I_val_b, "#,##0.000", B_ALIGN_NUM),
            (5,  buy_I_avg, "#,##0", B_ALIGN_NUM),
            (6,  int(buy_N_cnt), "#,##0", B_ALIGN_NUM),
            (7,  int(buy_N_vol), "#,##0", B_ALIGN_NUM),
            (8,  buy_N_val_b, "#,##0.000", B_ALIGN_NUM),
            (9,  buy_N_avg, "#,##0", B_ALIGN_NUM),
            (10, int(sell_I_cnt), "#,##0", B_ALIGN_NUM),
            (11, int(sell_I_vol), "#,##0", B_ALIGN_NUM),
            (12, sell_I_val_b, "#,##0.000", B_ALIGN_NUM),
            (13, sell_I_avg, "#,##0", B_ALIGN_NUM),
            (14, int(sell_N_cnt), "#,##0", B_ALIGN_NUM),
            (15, int(sell_N_vol), "#,##0", B_ALIGN_NUM),
            (16, sell_N_val_b, "#,##0.000", B_ALIGN_NUM),
            (17, sell_N_avg, "#,##0", B_ALIGN_NUM),
            (18, int(ownership_change), "#,##0;[Red]-#,##0", B_ALIGN_NUM),
            (19, power_ratio, "0.00", B_ALIGN_NUM),
            (20, money_flow, "#,##0.00;[Red]-#,##0.00", B_ALIGN_NUM),
        ]
        for col, val, fmt, align in vals:
            cell = ws2.cell(row=r, column=col, value=val)
            cell.font = B_FONT; cell.fill = fill
            cell.alignment = align; cell.number_format = fmt
    ws2.freeze_panes = "B5"
    
    wb.save(output_file)
    return output_file

# ============================================================
# به‌روزرسانی فایل موجود
# ============================================================
def update_input_file(input_file, ins_code, full_name):
    print(f"  Updating: {input_file}")
    wb = load_workbook(input_file)
    ws_price = None
    ws_hqa = None
    for sn in wb.sheetnames:
        if "قیمت" in sn and "سابقه" in sn:
            ws_price = wb[sn]
        elif "حقیقی" in sn and "حقوقی" in sn:
            ws_hqa = wb[sn]
    
    if ws_price is None or ws_hqa is None:
        print(f"  ERROR: Standard sheets not found!")
        return False
    
    # Find last date (newest date is in row 3, oldest is at the bottom)
    last_date_full = ws_price.cell(row=3, column=1).value  # newest date
    last_date_fa = ws_hqa.cell(row=5, column=1).value  # newest date
    
    print(f"  Last price date: {last_date_full}")
    print(f"  Last HQA date: {last_date_fa}")
    
    # Fetch historical data from TSETMC
    print(f"  Fetching data from TSETMC...")
    prices = fetch_price_history(ins_code, DEFAULT_DAYS)
    hqa_data = fetch_hqa_history(ins_code)
    
    # Filter only NEW dates (newer than last_date)
    # prices and hqa_data are sorted oldest to newest
    new_prices = []
    for p in prices:
        deven = p["dEven"]
        date_fa = deven_to_jdate_full(deven)
        if date_fa == last_date_full:
            idx = prices.index(p)
            new_prices = prices[idx+1:]  # all items after last_date
            break
    
    new_hqa = []
    for h in hqa_data:
        rec_date = h["recDate"]
        date_fa = deven_to_jdate_fa(rec_date)
        if date_fa == last_date_fa:
            idx = hqa_data.index(h)
            new_hqa = hqa_data[idx+1:]
            break
    
    new_prices = list(reversed(new_prices))
    new_hqa = list(reversed(new_hqa))
    
    print(f"  New price days: {len(new_prices)}")
    print(f"  New HQA days: {len(new_hqa)}")
    
    if not new_prices and not new_hqa:
        print(f"  No new data to add")
        return True
    
    # Shift existing data and insert new
    # Price sheet
    if new_prices:
        last_data_row = 3
        for r in range(3, 200):
            d = ws_price.cell(row=r, column=1).value
            if not d or "جمع" in str(d): break
            last_data_row = r
        shift_count = len(new_prices)
        for r in range(last_data_row, 2, -1):
            for c in range(1, 15):
                src = ws_price.cell(row=r, column=c)
                dst = ws_price.cell(row=r + shift_count, column=c)
                dst.value = src.value
                if src.has_style:
                    dst.font = src.font.copy()
                    dst.fill = src.fill.copy()
                    dst.alignment = src.alignment.copy()
                    dst.number_format = src.number_format
        for i, p in enumerate(new_prices):
            r = 3 + i
            deven = p["dEven"]
            prev_close = int(p["priceYesterday"])
            first = int(p["priceFirst"])
            close = int(p["pClosing"])
            last = int(p["pDrCotVal"])
            low = int(p["priceMin"])
            high = int(p["priceMax"])
            # محاسبه تغییر قیمت پایانی (نه آخرین قیمت)
            change = close - prev_close
            count = int(p["zTotTran"])
            volume = int(p["qTotTran5J"])
            value = int(p["qTotCap"])
            date_fa = deven_to_jdate_full(deven)
            close_pct = change / prev_close if prev_close else 0
            last_change = last - prev_close
            last_pct = last_change / prev_close if prev_close else 0
            fill = ODD_FILL if i % 2 == 0 else EVEN_FILL
            ws_price.row_dimensions[r].height = 18
            vals = [
                (1,  date_fa, "General", B_ALIGN),
                (2,  count, "#,##0", B_ALIGN_NUM),
                (3,  volume, "#,##0", B_ALIGN_NUM),
                (4,  value, "#,##0", B_ALIGN_NUM),
                (5,  prev_close, "#,##0", B_ALIGN_NUM),
                (6,  first, "#,##0", B_ALIGN_NUM),
                (7,  close, "#,##0", B_ALIGN_NUM),
                (8,  change, "#,##0;[Red]-#,##0", B_ALIGN_NUM),
                (9,  close_pct, "+0.00%;-0.00%;0.00%", B_ALIGN_NUM),
                (10, last, "#,##0", B_ALIGN_NUM),
                (11, last_change, "#,##0;[Red]-#,##0", B_ALIGN_NUM),
                (12, last_pct, "+0.00%;-0.00%;0.00%", B_ALIGN_NUM),
                (13, low, "#,##0", B_ALIGN_NUM),
                (14, high, "#,##0", B_ALIGN_NUM),
            ]
            for col, val, fmt, align in vals:
                cell = ws_price.cell(row=r, column=col, value=val)
                cell.font = B_FONT; cell.fill = fill
                cell.alignment = align; cell.number_format = fmt
    
    # HQA sheet
    if new_hqa:
        last_data_row = 5
        for r in range(5, 200):
            d = ws_hqa.cell(row=r, column=1).value
            if not d or "جمع" in str(d): break
            last_data_row = r
        shift_count = len(new_hqa)
        for r in range(last_data_row, 4, -1):
            for c in range(1, 21):
                src = ws_hqa.cell(row=r, column=c)
                dst = ws_hqa.cell(row=r + shift_count, column=c)
                dst.value = src.value
                if src.has_style:
                    dst.font = src.font.copy()
                    dst.fill = src.fill.copy()
                    dst.alignment = src.alignment.copy()
                    dst.number_format = src.number_format
        for i, h in enumerate(new_hqa):
            r = 5 + i
            rec_date = h["recDate"]
            buy_I_vol = float(h.get("buy_I_Volume", 0) or 0)
            buy_I_val = float(h.get("buy_I_Value", 0) or 0)
            buy_I_cnt = float(h.get("buy_I_Count", 0) or 0)
            buy_N_vol = float(h.get("buy_N_Volume", 0) or 0)
            buy_N_val = float(h.get("buy_N_Value", 0) or 0)
            buy_N_cnt = float(h.get("buy_N_Count", 0) or 0)
            sell_I_vol = float(h.get("sell_I_Volume", 0) or 0)
            sell_I_val = float(h.get("sell_I_Value", 0) or 0)
            sell_I_cnt = float(h.get("sell_I_Count", 0) or 0)
            sell_N_vol = float(h.get("sell_N_Volume", 0) or 0)
            sell_N_val = float(h.get("sell_N_Value", 0) or 0)
            sell_N_cnt = float(h.get("sell_N_Count", 0) or 0)
            date_fa = deven_to_jdate_fa(rec_date)
            buy_I_avg = buy_I_val / buy_I_vol if buy_I_vol else 0
            buy_N_avg = buy_N_val / buy_N_vol if buy_N_vol else 0
            sell_I_avg = sell_I_val / sell_I_vol if sell_I_vol else 0
            sell_N_avg = sell_N_val / sell_N_vol if sell_N_vol else 0
            buy_I_val_b = buy_I_val / 1e9
            buy_N_val_b = buy_N_val / 1e9
            sell_I_val_b = sell_I_val / 1e9
            sell_N_val_b = sell_N_val / 1e9
            ownership_change = buy_I_vol - sell_I_vol
            power_ratio = buy_I_val / sell_I_val if sell_I_val else 0
            money_flow = (buy_I_val - sell_I_val) / 1e10
            fill = ODD_FILL if i % 2 == 0 else EVEN_FILL
            ws_hqa.row_dimensions[r].height = 18
            vals = [
                (1,  date_fa, "General", B_ALIGN),
                (2,  int(buy_I_cnt), "#,##0", B_ALIGN_NUM),
                (3,  int(buy_I_vol), "#,##0", B_ALIGN_NUM),
                (4,  buy_I_val_b, "#,##0.000", B_ALIGN_NUM),
                (5,  buy_I_avg, "#,##0", B_ALIGN_NUM),
                (6,  int(buy_N_cnt), "#,##0", B_ALIGN_NUM),
                (7,  int(buy_N_vol), "#,##0", B_ALIGN_NUM),
                (8,  buy_N_val_b, "#,##0.000", B_ALIGN_NUM),
                (9,  buy_N_avg, "#,##0", B_ALIGN_NUM),
                (10, int(sell_I_cnt), "#,##0", B_ALIGN_NUM),
                (11, int(sell_I_vol), "#,##0", B_ALIGN_NUM),
                (12, sell_I_val_b, "#,##0.000", B_ALIGN_NUM),
                (13, sell_I_avg, "#,##0", B_ALIGN_NUM),
                (14, int(sell_N_cnt), "#,##0", B_ALIGN_NUM),
                (15, int(sell_N_vol), "#,##0", B_ALIGN_NUM),
                (16, sell_N_val_b, "#,##0.000", B_ALIGN_NUM),
                (17, sell_N_avg, "#,##0", B_ALIGN_NUM),
                (18, int(ownership_change), "#,##0;[Red]-#,##0", B_ALIGN_NUM),
                (19, power_ratio, "0.00", B_ALIGN_NUM),
                (20, money_flow, "#,##0.00;[Red]-#,##0.00", B_ALIGN_NUM),
            ]
            for col, val, fmt, align in vals:
                cell = ws_hqa.cell(row=r, column=col, value=val)
                cell.font = B_FONT; cell.fill = fill
                cell.alignment = align; cell.number_format = fmt
    
    try:
        wb.save(input_file)
        print(f"  Updated: {input_file}")
        return True
    except PermissionError:
        print(f"  ERROR: File is open in Excel")
        return False

# ============================================================
# اجرای داشبورد
# ============================================================
def run_dashboard(input_file, name):
    import subprocess
    script_dir = os.path.dirname(os.path.abspath(__file__))
    make_script = os.path.join(script_dir, "make_dashboard.py")
    if not os.path.exists(make_script):
        print(f"  WARNING: make_dashboard.py not found")
        return False
    
    # ساخت پوشه dashboard
    dash_dir = os.path.join(script_dir, "dashboard")
    os.makedirs(dash_dir, exist_ok=True)
    
    # خروجی در پوشه dashboard
    dash_file = os.path.join(dash_dir, f"{name}_dashboard.xlsx")
    
    try:
        result = subprocess.run(
            [sys.executable, make_script, input_file, "--name", name, "--output", dash_file],
            capture_output=True, text=True, timeout=120
        )
        return result.returncode == 0
    except Exception as e:
        print(f"  ERROR: {e}")
        return False

# ============================================================
# تابع اصلی
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Create/Update input file from TSETMC")
    parser.add_argument("--symbol", default=None, help="Stock symbol (e.g. sarood)")
    parser.add_argument("--ins-code", default=None, help="TSETMC instrument code")
    parser.add_argument("--create", action="store_true", help="Create new file with 60 days")
    parser.add_argument("--update", action="store_true", help="Update existing file")
    parser.add_argument("--dashboard", action="store_true", help="Build dashboard after")
    parser.add_argument("--days", type=int, default=DEFAULT_DAYS, help=f"Number of days (default: {DEFAULT_DAYS})")
    args = parser.parse_args()
    
    print("=" * 60)
    print("Phase 1: Create/Update input file")
    print("=" * 60)
    
    if not args.create and not args.update:
        print("\nERROR: Must specify --create or --update")
        return 1
    
    symbols = []
    if args.symbol:
        symbols = [s.strip() for s in args.symbol.split(",")]
    elif args.ins_code:
        symbols = [args.ins_code]
    else:
        print("\nERROR: Must specify --symbol or --ins-code")
        return 1
    
    for sym in symbols:
        print(f"\n{'='*60}")
        
        if args.ins_code and sym == args.ins_code:
            ins_code = sym
            full_name = sym
            symbol_name = sym
        else:
            print(f"Searching symbol '{sym}'...")
            ins_code, full_name, symbol_name = find_ins_code(sym)
            if not ins_code:
                print(f"  NOT FOUND")
                continue
            print(f"  Found: {full_name} (code: {ins_code})")
        
        input_file = os.path.join("input", f"{symbol_name}_input.xlsx")
        
        # Create input folder if not exists
        os.makedirs("input", exist_ok=True)
        
        if args.create:
            print(f"\nFetching {args.days} days data...")
            print(f"  Prices...")
            prices = fetch_price_history(ins_code, args.days)
            print(f"  Got {len(prices)} days")
            print(f"  HQA...")
            hqa_data = fetch_hqa_history(ins_code)
            print(f"  Got {len(hqa_data)} days")
            
            if len(prices) > args.days:
                prices = prices[-args.days:]
            if len(hqa_data) > args.days:
                hqa_data = hqa_data[-args.days:]
            
            print(f"\nBuilding input file...")
            output = build_input_file(symbol_name, ins_code, full_name, prices, hqa_data, input_file)
            print(f"  Created: {output}")
            
            if prices:
                latest = prices[-1]
                date_fa = deven_to_jdate_full(latest["dEven"])
                print(f"\n  Last day: {date_fa}")
                print(f"  Close: {int(latest['pClosing']):,}")
                print(f"  Volume: {int(latest['qTotTran5J']):,}")
        
        elif args.update:
            if not os.path.exists(input_file):
                print(f"  File {input_file} not found!")
                print(f"  Use --create first")
                continue
            success = update_input_file(input_file, ins_code, full_name)
            if not success:
                continue
        
        if args.dashboard:
            print(f"\nBuilding dashboard...")
            if run_dashboard(input_file, symbol_name):
                print(f"  Dashboard created: {symbol_name}_dashboard.xlsx")
            else:
                print(f"  Dashboard failed")
    
    print(f"\n{'='*60}")
    print("Done")
    print(f"{'='*60}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
