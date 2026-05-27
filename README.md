# Time Income Tracker

یک ابزار شخصی آفلاین، سبک و سریع برای مدیریت زمان کاری، ثبت فعالیت پروژه، محاسبه درآمد ساعتی و خروجی Excel.

## ویژگی‌ها
- تایمر دقیق Start / Pause / Resume / Stop & Save / Cancel
- مدیریت پروژه با نرخ ساعتی (تومان)
- ذخیره `hourly_rate_snapshot` برای پایداری گزارش‌ها
- System Tray + Hotkey سراسری (پیش‌فرض Shift+Q)
- تشخیص idle ویندوز با Windows API (`ctypes` + `GetLastInputInfo`)
- auto-pause هنگام idle
- نوتیفیکیشن دوره‌ای هنگام فعال بودن تایمر
- خروجی Excel invoice-ready با جمع کل زمان و مبلغ
- ذخیره‌سازی کامل داده در `%APPDATA%\TimeIncomeTracker`

## مسیر ذخیره داده‌ها
- Database: `%APPDATA%\TimeIncomeTracker\time_income.db`
- Logs: `%APPDATA%\TimeIncomeTracker\logs\app.log`
- تنظیمات: جدول `settings` در SQLite
- خروجی اکسل: مسیر تنظیم‌شده در Settings (پیش‌فرض `%APPDATA%\TimeIncomeTracker`)

## نیازمندی توسعه‌دهنده
- Python 3.12+
- Windows 10/11 (برای تست کامل tray/hotkey/idle)

## اجرای سورس (برای توسعه‌دهنده)
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## ساخت فایل exe
```bat
build.bat
```
یا دستی:
```bash
pyinstaller --noconfirm --onefile --windowed --name TimeIncomeTracker main.py
```
خروجی:
- `dist\TimeIncomeTracker.exe`

## نکته مهم برای کاربر نهایی
**کاربر نهایی نیازی به نصب Python، pip یا ابزار توسعه ندارد و فقط فایل `TimeIncomeTracker.exe` را اجرا می‌کند.**

## Troubleshooting
- اگر Hotkey ثبت نشد: برنامه ادامه می‌دهد و خطا در status/log نمایش داده می‌شود.
- اگر Notification محدود بود: برنامه با fallback tray message اجرا می‌شود.
- اگر Excel خروجی نداد: مسیر خروجی و سطح دسترسی پوشه را بررسی کنید.
- اگر دیتابیس نبود: برنامه خودکار جدول‌ها را ایجاد می‌کند.
