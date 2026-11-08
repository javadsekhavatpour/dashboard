# سامانه غربالگری سهام بورس تهران (TSE Screener)

## فایل‌ها

| فایل | شرح |
|---|---|
| `app.py` | داشبورد اصلی Streamlit (جدول نمادها + ایجاد/آپدیت فایل input) |
| `summary.py` | داشبورد خلاصه (نمایش آخرین روز همه نمادها) |
| `update_input.py` | دریافت داده از TSETMC و ساخت/آپدیت فایل input |
| `make_dashboard.py` | ساخت داشبورد تحلیلی از فایل input |
| `run.bat` | اجرای داشبورد اصلی |
| `run_summary.bat` | اجرای داشبورد خلاصه |
| `install.bat` | نصب پیش‌نیازها (streamlit, requests, openpyxl) |
| `build_all_dashboards.bat` | ساخت داشبورد همه فایل‌های input |

## نصب

```bash
pip install streamlit requests openpyxl
```

## اجرا

```bash
# داشبورد اصلی (ایجاد/آپدیت فایل input)
streamlit run app.py

# داشبورد خلاصه (آخرین روز همه نمادها)
streamlit run summary.py
```

## ساختار پوشه‌ها

```
پوشه اصلی/
├── app.py
├── summary.py
├── update_input.py
├── make_dashboard.py
├── run.bat
├── run_summary.bat
├── install.bat
├── build_all_dashboards.bat
├── input/                    ← فایل‌های input
│   ├── سرود_input.xlsx
│   └── دلر_input.xlsx
└── dashboard/               ← فایل‌های داشبورد
    ├── سرود_dashboard.xlsx
    └── دلر_dashboard.xlsx
```
