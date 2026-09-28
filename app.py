# -*- coding: utf-8 -*-
APP_VERSION = "3.7.0"
APP_NAME = "Plate Archive Pro"

import os
import sys
import math
import random
import sqlite3
import shutil
import ctypes
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import xlwings as xw
except ImportError:
    xw = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

EXCEL_LOCK = threading.RLock()

base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
icon_path = os.path.join(base_path, "app_icon.ico")
if sys.platform.startswith("win"):
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("platearchive.app.standalone")
    except Exception:
        pass

DHIKR_AND_WISDOM_LIST = [
    "🌿 سُبْحَانَ اللَّهِ وَبِحَمْدِهِ ، سُبْحَانَ اللَّهِ الْعَظِيمِ",
    "🤍 أَسْتَغْفِرُ اللَّهَ الْعَظِيمَ الَّذِي لا إِلَهَ إِلا هُوَ الْحَيُّ الْقَيُّومُ وَأَتُوبُ إِلَيْهِ",
    "✨ لا حَوْلَ وَلا قُوَّةَ إِلا بِاللَّهِ الْعَلِيِّ الْعَظِيمِ كَنْزٌ مِنْ كُنُوزِ الْجَنَّةِ",
    "🌸 اللَّهُمَّ صَلِّ وَسَلِّمْ وَبَارِكْ عَلَى نَبِيِّنَا مُحَمَّدٍ",
    "🌿 لا إِلَهَ إِلا أَنْتَ سُبْحَانَكَ إِنِّي كُنْتُ مِنَ الظَّالِمِينَ",
    "🤍 سُبْحَانَ اللَّهِ ، وَالْحَمْدُ لِلَّهِ ، وَلا إِلَهَ إِلا اللَّهُ ، وَاللَّهُ أَكْبَرُ",
    "💡 حكمة: مَنْ تَأَنَّى أَصَابَ أَوْ كَادَ، وَمَنْ عَجِلَ أَخْطَأَ أَوْ كَادَ.",
    "💡 حكمة: الإِخْلاصُ فِي الْعَمَلِ سَبَبٌ لِلْقَبُولِ وَالتَّوْفِيقِ.",
    "💡 حكمة: الْوَقْتُ هُوَ رَأْسُ مَالِكَ الْحَقِيقِيِّ، فَلا تُنْفِقْهُ فِي الْفَرَاغِ."
]

def compute_prayer_times_local(lat=35.93, lon=36.63):
    now = datetime.now()
    day_of_year = now.timetuple().tm_yday
    b = 2 * math.pi * (day_of_year - 81) / 365
    eot = 9.87 * math.sin(2 * b) - 7.53 * math.cos(b) - 1.5 * math.sin(b)
    decl = 23.45 * math.sin(math.radians((360 / 365) * (day_of_year - 81)))
    timezone = 3.0
    solar_noon = 12 + (timezone * 15 - lon) / 15 - (eot / 60)
    phi = math.radians(lat)
    delta = math.radians(decl)

    def hour_angle(alpha):
        val = (math.sin(math.radians(alpha)) - math.sin(phi) * math.sin(delta)) / (math.cos(phi) * math.cos(delta))
        return math.degrees(math.acos(max(-1.0, min(1.0, val)))) / 15

    ha_fajr = hour_angle(-18.0)
    ha_maghrib = hour_angle(-0.833)
    asr_alt = math.degrees(math.atan(1 / (1 + math.tan(abs(phi - delta)))))
    ha_asr = hour_angle(asr_alt)

    def to_dt(dec_hour):
        h = int(dec_hour)
        m = int((dec_hour - h) * 60)
        s = int((((dec_hour - h) * 60) - m) * 60)
        return now.replace(hour=h % 24, minute=m, second=s, microsecond=0)

    return [
        ("الفجر", to_dt(solar_noon - ha_fajr)),
        ("الظهر", to_dt(solar_noon)),
        ("العصر", to_dt(solar_noon + ha_asr)),
        ("المغرب", to_dt(solar_noon + ha_maghrib)),
        ("العشاء", to_dt(solar_noon + hour_angle(-17.0)))
    ]

def normalize_plate(v):
    if v is None: return ""
    s = str(v).strip()
    if s.endswith(".0"): s = s[:-2]
    if s.startswith("'"): s = s[1:]
    return s

def normalize_search_text(v):
    digits = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = normalize_plate(v).strip().casefold().translate(digits)
    return (s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
             .replace("ة", "ه").replace("ى", "ي"))

def normalize_box(v):
    if v is None: return ""
    s = str(v).strip()
    if s.endswith(".0"): s = s[:-2]
    s = s.upper().replace("صندوق", "").strip()
    if s.startswith("B"): s = s[1:].strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return str(int(digits)) if digits else ""

def excel_date_text(v):
    if not v: return ""
    if isinstance(v, datetime): return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if s.startswith("'"): s = s[1:]
    s = s[:10]
    if "-" in s:
        p = s.split("-")
        if len(p) == 3 and len(p[0]) == 4: return f"{p[2]}/{p[1]}/{p[0]}"
    return s

class PlateManagerMilitaryApp:
    def __init__(self, root):
        self.root = root
        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            pass
        self.root.title(f"{APP_NAME} | الإصدار {APP_VERSION}")
        self.root.geometry("1340x860")
        self.root.configure(bg="#F0F4F8")

        self.excel_path = None
        self.busy = False
        self.edit_mode = None
        self.found_record = None
        self.undo_history = []
        self.max_undo = 10
        self.data_cache = []
        self.audit_path = None
        self.backup_dir = None
        self._box_debounce_id = None

        # تثبيت خيارات إضافية
        self.pinned_options = tk.BooleanVar(value=False)
        self.last_options = ""

        # إعدادات الأعمدة
        self.start_row = 5
        self.col_plate = "A"
        self.col_box = "B"
        self.col_date = "C"
        self.col_notes = "D"
        self.col_status = "E"
        self.col_options = "F"

        # خيارات الرخصة (مرفقة / غير مرفقة)
        self.license_options = [
            "مرفقة برخصة اثناء الارشفة",
            "غير مرفقة برخصة اثناء الارشفة"
        ]
        self.license_index = 0

        # خيارات الحالة (جديدة / قديمة)
        self.status_options = ["جديدة", "قديمة"]
        self.status_index = 0

        self.prayer_visible = True

        self.setup_ui()
        self.setup_keyboard_nav()
        self.update_clock()
        self.update_dhikr_cycle()
        self.update_prayer_countdown()

    def setup_keyboard_nav(self):
        self.root.bind("<Control-z>", lambda e: self.undo_last())
        self.root.bind("<Up>", lambda e: self.toggle_license(-1))
        self.root.bind("<Down>", lambda e: self.toggle_license(1))
        self.root.bind("<Right>", lambda e: self.set_status(0))
        self.root.bind("<Left>", lambda e: self.set_status(1))

    def toggle_license(self, step=1):
        self.license_index = (self.license_index + step) % len(self.license_options)
        val = self.license_options[self.license_index]
        self.btn_license_display.config(
            text=f"✔ {val}",
            bg="#0284C7" if self.license_index == 0 else "#64748B"
        )

    def set_status(self, idx):
        self.status_index = idx
        val = self.status_options[self.status_index]
        self.btn_status_display.config(
            text=f"الحالة: {val}",
            bg="#059669" if self.status_index == 0 else "#D97706"
        )

    def toggle_status_click(self):
        self.set_status(1 - self.status_index)

    def update_clock(self):
        if not self.root.winfo_exists(): return
        now = datetime.now()
        days = {"Monday":"الإثنين","Tuesday":"الثلاثاء","Wednesday":"الأربعاء","Thursday":"الخميس","Friday":"الجمعة","Saturday":"السبت","Sunday":"الأحد"}
        self.lbl_clock.config(text=f"🕒 {days.get(now.strftime('%A'), '')} : {now.strftime('%d/%m/%Y  |  %I:%M:%S %p')}")
        self.root.after(1000, self.update_clock)

    def update_dhikr_cycle(self):
        if self.prayer_visible and hasattr(self, "lbl_dhikr"):
            self.lbl_dhikr.config(text=random.choice(DHIKR_AND_WISDOM_LIST))
        self.root.after(120000, self.update_dhikr_cycle)

    def update_prayer_countdown(self):
        if self.prayer_visible and hasattr(self, "lbl_prayer"):
            now = datetime.now()
            prayers = compute_prayer_times_local()
            next_name, next_dt = None, None
            for name, dt in prayers:
                if dt > now:
                    next_name, next_dt = name, dt
                    break
            if not next_name:
                next_name = "الفجر"
                next_dt = prayers[0][1] + timedelta(days=1)

            diff = next_dt - now
            rem_sec = int(diff.total_seconds())
            h, m, s = rem_sec // 3600, (rem_sec % 3600) // 60, rem_sec % 60
            self.lbl_prayer.config(text=f"🕌 الصلاة القادمة: {next_name} | {h:02d}:{m:02d}:{s:02d}")
        self.root.after(1000, self.update_prayer_countdown)

    def setup_ui(self):
        self.top_bar = tk.Frame(self.root, bg="#0A1D37", height=85)
        self.top_bar.pack(fill="x", side="top")

        title_frame = tk.Frame(self.top_bar, bg="#0A1D37")
        title_frame.pack(side="right", padx=20, pady=6)
        tk.Label(title_frame, text=APP_NAME, bg="#0A1D37", fg="#FFFFFF", font=("Segoe UI", 15, "bold")).pack(anchor="e")
        self.lbl_clock = tk.Label(title_frame, text="", bg="#0A1D37", fg="#38BDF8", font=("Segoe UI", 10, "bold"))
        self.lbl_clock.pack(anchor="e", pady=(2, 0))

        self.prayer_bar = tk.Frame(self.top_bar, bg="#0A1D37")
        self.prayer_bar.place(relx=0.5, rely=0.5, anchor="center")

        self.frame_dhikr = tk.Frame(self.prayer_bar, bg="#0F2942", bd=1, relief="solid", padx=14, pady=3)
        self.frame_dhikr.pack(fill="x", pady=(0, 3))
        self.lbl_dhikr = tk.Label(self.frame_dhikr, text=random.choice(DHIKR_AND_WISDOM_LIST), bg="#0F2942", fg="#F1F5F9", font=("Segoe UI", 9, "bold"), wraplength=560, justify="center")
        self.lbl_dhikr.pack()

        self.frame_prayer = tk.Frame(self.prayer_bar, bg="#0F2942", bd=1, relief="solid", padx=14, pady=3)
        self.frame_prayer.pack(fill="x", pady=(3, 0))
        self.lbl_prayer = tk.Label(self.frame_prayer, text="🕌 جاري الحساب...", bg="#0F2942", fg="#38BDF8", font=("Segoe UI", 9, "bold"), justify="center")
        self.lbl_prayer.pack()

        tools_box = tk.Frame(self.top_bar, bg="#0A1D37")
        tools_box.pack(side="left", padx=15, pady=10)

        btn_menu = tk.Button(tools_box, text=" ☰ الخيارات والتقارير ", bg="#1E293B", fg="#38BDF8", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", padx=12, pady=5, command=self.open_large_modern_menu)
        btn_menu.pack(side="left", padx=4)

        btn_browse = tk.Button(tools_box, text="📂 ربط إكسل", bg="#0284C7", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=12, pady=5, command=self.select_file)
        btn_browse.pack(side="left", padx=4)

        self.lbl_file = tk.Label(tools_box, text="لم يتم تحديد ملف", bg="#0A1D37", fg="#94A3B8", font=("Segoe UI", 9))
        self.lbl_file.pack(side="left", padx=6)

        main_container = tk.Frame(self.root, bg="#F0F4F8")
        main_container.pack(fill="both", expand=True, padx=20, pady=8)
        self.setup_stats_row(main_container)

        body_frame = tk.Frame(main_container, bg="#F0F4F8")
        body_frame.pack(fill="both", expand=True, pady=4)
        body_frame.grid_columnconfigure(0, weight=1, uniform="main_cards")
        body_frame.grid_columnconfigure(1, weight=1, uniform="main_cards")
        body_frame.grid_rowconfigure(0, weight=1)

        left_card = tk.Frame(body_frame, bg="#FFFFFF", bd=1, relief="solid")
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.setup_left_card(left_card)

        right_card = tk.Frame(body_frame, bg="#FFFFFF", bd=1, relief="solid")
        right_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.setup_right_card(right_card)

        self.toast_frame = tk.Frame(self.root, bg="#F0F4F8", height=35)
        self.toast_frame.pack(fill="x", side="bottom", padx=20, pady=(0, 4))
        self.lbl_toast = tk.Label(self.toast_frame, text="", font=("Segoe UI", 10, "bold"), pady=4)

    def setup_stats_row(self, parent):
        stats_frame = tk.Frame(parent, bg="#F0F4F8")
        stats_frame.pack(fill="x", pady=(0, 6))

        def make_stat_badge(title, accent):
            f = tk.Frame(stats_frame, bg="#FFFFFF", bd=1, relief="solid", padx=6, pady=5)
            f.pack(side="right", fill="both", expand=True, padx=2)
            tk.Label(f, text=title, bg="#FFFFFF", fg="#64748B", font=("Segoe UI", 8, "bold")).pack(anchor="e")
            v = tk.Label(f, text="0", bg="#FFFFFF", fg=accent, font=("Segoe UI", 13, "bold"))
            v.pack(anchor="e")
            return v

        self.card_total = make_stat_badge("إجمالي اللوحات", "#0A1D37")
        self.card_new = make_stat_badge("لوحات جديدة", "#059669")
        self.card_old = make_stat_badge("لوحات قديمة", "#D97706")
        self.card_box = make_stat_badge("عداد الحزمة المكتوبة", "#1E293B")
        self.card_box.config(text="-")

    def setup_right_card(self, parent):
        tk.Label(parent, text="➕ تسجيل وإضافة لوحة جديدة", bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 13, "bold")).pack(anchor="e", padx=20, pady=(8, 4))
        body = tk.Frame(parent, bg="#FFFFFF")
        body.pack(fill="both", expand=True, padx=20, pady=2)

        # 1. رقم اللوحة
        tk.Label(body, text="* رقم اللوحة:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_plate = tk.Entry(body, font=("Segoe UI", 15, "bold"), bg="#F8FAFC", fg="#0F172A", relief="solid", bd=1, justify="center")
        self.ent_plate.pack(fill="x", ipady=4, pady=(0, 6))
        self.ent_plate.bind("<Return>", lambda e: self.save_plate_direct())

        # 2. رقم الحزمة (مثبت تلقائياً حتى تغييره)
        tk.Label(body, text="* رقم الحزمة (B) [مثبت تلقائياً حتى تغييره]:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_box = tk.Entry(body, font=("Segoe UI", 12), bg="#F8FAFC", relief="solid", bd=1, justify="center")
        self.ent_box.pack(fill="x", ipady=3, pady=(0, 6))
        self.ent_box.bind("<KeyRelease>", self.on_box_type_delayed)
        self.ent_box.bind("<Return>", lambda e: self.save_plate_direct())

        # 3. العمود D: حالة الرخصة (تبديل بالأسهم أو بالنقر)
        tk.Label(body, text="* العمود D: حالة الرخصة (تبديل بالأسهم ↑ و ↓ أو بالنقر):", bg="#FFFFFF", fg="#0284C7", font=("Segoe UI", 9, "bold")).pack(anchor="e", pady=(2, 2))
        self.btn_license_display = tk.Button(
            body,
            text=f"✔ {self.license_options[0]}",
            bg="#0284C7",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="solid",
            bd=1,
            pady=7,
            cursor="hand2",
            command=lambda: self.toggle_license(1)
        )
        self.btn_license_display.pack(fill="x", pady=(0, 8))

        # 4. العمود E: الحالة (تبديل بالأسهم أو بالنقر)
        tk.Label(body, text="* العمود E: الحالة (تبديل بالأسهم → و ← أو بالنقر):", bg="#FFFFFF", fg="#059669", font=("Segoe UI", 9, "bold")).pack(anchor="e", pady=(2, 2))
        self.btn_status_display = tk.Button(
            body,
            text="الحالة: جديدة",
            bg="#059669",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="solid",
            bd=1,
            pady=7,
            cursor="hand2",
            command=self.toggle_status_click
        )
        self.btn_status_display.pack(fill="x", pady=(0, 8))

        # 5. العمود F: خيارات إضافية كتابة حرة مع تثبيت
        tk.Label(body, text="* العمود F: خيارات إضافية (كتابة يدوية):", bg="#FFFFFF", fg="#64748B", font=("Segoe UI", 9, "bold")).pack(anchor="e", pady=(2, 2))
        opt_box = tk.Frame(body, bg="#FFFFFF")
        opt_box.pack(fill="x", pady=(0, 8))
        self.ent_options = tk.Entry(opt_box, font=("Segoe UI", 11), bg="#F8FAFC", relief="solid", bd=1)
        self.ent_options.pack(side="right", fill="x", expand=True, ipady=4)
        tk.Checkbutton(opt_box, text="📌 تثبيت", variable=self.pinned_options, bg="#FFFFFF", font=("Segoe UI", 9, "bold"), command=self.toggle_pin_options).pack(side="left", padx=4)
        tk.Button(opt_box, text="مسح", bg="#FEE2E2", fg="#B91C1C", relief="flat", command=lambda: self.ent_options.delete(0, tk.END)).pack(side="left")

        # أزرار الحفظ والتراجع
        bf = tk.Frame(body, bg="#FFFFFF")
        bf.pack(fill="x", pady=(4, 0))
        self.btn_save = tk.Button(bf, text="⚡ حفظ مباشر في الإكسل (Enter)", bg="#0284C7", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", pady=7, command=self.save_plate_direct)
        self.btn_save.pack(side="right", fill="x", expand=True, padx=2)
        self.btn_undo = tk.Button(bf, text="↩ تراجع", bg="#64748B", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", pady=7, command=self.undo_last)
        self.btn_undo.pack(side="left", fill="x", expand=True, padx=2)

        self.lbl_mode = tk.Label(body, text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857", font=("Segoe UI", 9, "bold"), pady=4)
        self.lbl_mode.pack(fill="x", pady=4)

    def setup_left_card(self, parent):
        tk.Label(parent, text="🔍 البحث السريع + آخر الإدخالات", bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 13, "bold")).pack(anchor="e", padx=20, pady=(8, 4))
        body = tk.Frame(parent, bg="#FFFFFF")
        body.pack(fill="both", expand=True, padx=20, pady=2)

        tk.Label(body, text="اكتب رقم اللوحة للبحث:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_search = tk.Entry(body, font=("Segoe UI", 13), bg="#F8FAFC", relief="solid", bd=1, justify="center")
        self.ent_search.pack(fill="x", ipady=4, pady=(0, 5))
        self.ent_search.bind("<Return>", lambda e: self.search_plate())

        tk.Button(body, text="فحص السجل 🔍", bg="#0A1D37", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", pady=4, command=self.search_plate).pack(fill="x", pady=(0, 4))
        self.lbl_search_status = tk.Label(body, text="لم يتم إجراء بحث", bg="#F1F5F9", fg="#475569", font=("Segoe UI", 9), pady=6, wraplength=540)
        self.lbl_search_status.pack(fill="x", pady=(0, 4))

        action_row = tk.Frame(body, bg="#FFFFFF")
        action_row.pack(fill="x", pady=2)
        self.btn_edit_found = tk.Button(action_row, text="✏ تعديل السجل", bg="#D97706", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", state="disabled", command=self.start_edit_found)
        self.btn_edit_found.pack(side="right", fill="x", expand=True, padx=2)
        self.btn_go_excel = tk.Button(action_row, text="📊 الانتقال إلى خلية اللوحة في Excel", bg="#334155", fg="#FFFFFF", font=("Segoe UI", 9, "bold"), relief="flat", state="disabled", command=self.go_to_found_excel)
        self.btn_go_excel.pack(side="left", fill="x", expand=True, padx=2)

        lf = tk.LabelFrame(body, text="آخر 10 إدخالات", bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 9, "bold"), padx=5, pady=5)
        lf.pack(fill="both", expand=True, pady=6)
        
        cols = ("date", "row", "opt", "status", "lic", "box", "plate")
        self.recent_tree = ttk.Treeview(lf, columns=cols, show="headings", height=8)
        for c, t, w in (
            ("date","التاريخ",90),
            ("row","صف",45),
            ("opt","خيارات F",80),
            ("status","الحالة",65),
            ("lic","الرخصة D",140),
            ("box","الحزمة",60),
            ("plate","اللوحة",95)
        ):
            self.recent_tree.heading(c, text=t); self.recent_tree.column(c, width=w, anchor="center")

        sb = ttk.Scrollbar(lf, orient="vertical", command=self.recent_tree.yview)
        self.recent_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.recent_tree.pack(fill="both", expand=True)

    def toggle_pin_options(self):
        if self.pinned_options.get():
            self.last_options = self.ent_options.get().strip()
            self.show_toast("📌 تم تثبيت خيارات العمود F")
        else:
            self.show_toast("📌 تم إلغاء التثبيت")

    def select_file(self):
        path = filedialog.askopenfilename(title="اختر ملف الإكسل", filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")])
        if path:
            self.excel_path = os.path.normpath(path)
            self._init_storage_paths()
            self.lbl_file.config(text=os.path.basename(path), fg="#38BDF8")
            self.update_stats_async()
            self.on_box_type()
            self.show_toast("✔ تم ربط ملف Excel بنجاح")

    def get_sheet(self):
        if not self.excel_path: raise RuntimeError("يرجى ربط ملف الإكسل أولاً!")
        if xw is None: raise RuntimeError("مكتبة xlwings غير مثبتة")
        wb = xw.Book(self.excel_path)
        return wb, wb.sheets[0]

    def with_com(self, func, *args, **kwargs):
        with EXCEL_LOCK:
            if pythoncom: pythoncom.CoInitialize()
            try:
                return func(*args, **kwargs)
            finally:
                if pythoncom: pythoncom.CoUninitialize()

    def get_last_row(self, ws, col):
        return ws.range(f"{col}{ws.cells.rows.count}").end("up").row

    def get_target_row_smart(self, ws, col):
        last = self.get_last_row(ws, col)
        if last < self.start_row: return self.start_row
        vals = ws.range(f"{col}{self.start_row}:{col}{last}").value
        if not isinstance(vals, list): vals = [vals]
        for i, v in enumerate(vals):
            p = normalize_plate(v)
            if not p or p == "0": return self.start_row + i
        return last + 1

    def _init_storage_paths(self):
        if not self.excel_path: return
        base = os.path.dirname(self.excel_path)
        self.backup_dir = os.path.join(base, "PlateArchive_Backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        self.audit_path = os.path.join(self.backup_dir, "audit_log.jsonl")

    def _safety_backup(self, reason="before-change"):
        if not self.excel_path or not os.path.exists(self.excel_path): return
        self._init_storage_paths()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = os.path.splitext(self.excel_path)[1]
        dst = os.path.join(self.backup_dir, f"backup_{stamp}_{reason}{ext}")
        try: shutil.copy2(self.excel_path, dst)
        except Exception: pass

    def get_cached_records(self):
        if not self.excel_path: return []
        def work():
            wb, ws = self.get_sheet()
            last_r = self.get_last_row(ws, self.col_plate)
            if last_r < self.start_row: return []
            count = last_r - self.start_row + 1
            def tl(v): return v if isinstance(v, list) else [v] * count
            p_vals = tl(ws.range(f"{self.col_plate}{self.start_row}:{self.col_plate}{last_r}").value)
            b_vals = tl(ws.range(f"{self.col_box}{self.start_row}:{self.col_box}{last_r}").value)
            d_vals = tl(ws.range(f"{self.col_date}{self.start_row}:{self.col_date}{last_r}").value)
            n_vals = tl(ws.range(f"{self.col_notes}{self.start_row}:{self.col_notes}{last_r}").value)
            s_vals = tl(ws.range(f"{self.col_status}{self.start_row}:{self.col_status}{last_r}").value)
            o_vals = tl(ws.range(f"{self.col_options}{self.start_row}:{self.col_options}{last_r}").value)
            records = []
            for i in range(count):
                plate = normalize_plate(p_vals[i])
                if plate and plate != "0":
                    records.append({
                        "plate": plate,
                        "box": normalize_box(b_vals[i]),
                        "date": excel_date_text(d_vals[i]),
                        "license": normalize_plate(n_vals[i]),
                        "status": normalize_plate(s_vals[i]) or "جديدة",
                        "options": normalize_plate(o_vals[i]) or "",
                        "row": self.start_row + i
                    })
            return records
        return self.with_com(work)

    def save_plate_direct(self):
        if self.busy: return
        if self.edit_mode:
            self.save_edit(); return
        if not self.excel_path:
            self.show_toast("⚠ اختر ملف Excel أولاً!", True); return
        plate = normalize_plate(self.ent_plate.get())
        box = normalize_box(self.ent_box.get())
        license_val = self.license_options[self.license_index]
        status = self.status_options[self.status_index]
        options_val = self.ent_options.get().strip()
        if self.pinned_options.get() and not options_val:
            options_val = self.last_options

        if not plate or not box:
            self.show_toast("رقم اللوحة والحزمة مطلوبان!", True); return

        now = datetime.now().strftime("%d/%m/%Y")
        self.busy = True
        self.btn_save.config(state="disabled")
        self.lbl_mode.config(text="جارٍ الحفظ في Excel…", bg="#FEF3C7", fg="#92400E")

        def worker():
            def work():
                self._safety_backup("add")
                wb, ws = self.get_sheet()
                last = self.get_last_row(ws, self.col_plate)
                if last >= self.start_row:
                    vals = ws.range(f"{self.col_plate}{self.start_row}:{self.col_plate}{last}").value
                    vals = vals if isinstance(vals, list) else [vals]
                    for v in vals:
                        if normalize_plate(v) == plate:
                            raise RuntimeError(f"اللوحة [{plate}] مسجلة مسبقاً!")
                nxt = self.get_target_row_smart(ws, self.col_plate)
                old = {
                    c: ws.range(f"{c}{nxt}").value 
                    for c in (self.col_plate, self.col_box, self.col_date, self.col_notes, self.col_status, self.col_options)
                }
                ws.range(f"{self.col_plate}{nxt}").value = f"'{plate}"
                ws.range(f"{self.col_box}{nxt}").value = int(box)
                ws.range(f"{self.col_date}{nxt}").value = f"'{now}"
                ws.range(f"{self.col_notes}{nxt}").value = license_val
                ws.range(f"{self.col_status}{nxt}").value = status
                ws.range(f"{self.col_options}{nxt}").value = options_val
                wb.save()
                return {
                    "plate": plate, "box": box, "license": license_val, 
                    "status": status, "options": options_val, "date": now, 
                    "row": nxt, "old": old
                }

            try:
                res = self.with_com(work)
                self.root.after(0, lambda: self._finish_add(res))
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda: self._finish_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_add(self, res):
        self.busy = False
        self.btn_save.config(state="normal")
        self.lbl_mode.config(text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857")
        self.undo_history.append(res)
        self.refresh_recent()
        self.ent_plate.delete(0, tk.END)
        
        # خيارات العمود F
        if not self.pinned_options.get():
            self.ent_options.delete(0, tk.END)
        else:
            self.last_options = res["options"]
            self.ent_options.delete(0, tk.END)
            self.ent_options.insert(0, self.last_options)

        self.ent_plate.focus()
        self.update_stats_async()
        self.on_box_type()  # تحديث العداد تلقائياً
        self.show_toast(f"✔ تم حفظ {res['plate']} في الحزمة B{res['box']}")

    def _finish_error(self, msg):
        self.busy = False
        self.btn_save.config(state="normal")
        self.lbl_mode.config(text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857")
        self.show_toast(msg, True)

    def refresh_recent(self):
        for x in self.recent_tree.get_children(): self.recent_tree.delete(x)
        for r in reversed(self.undo_history[-10:]):
            self.recent_tree.insert("", "end", values=(
                r["date"], r["row"], r.get("options",""),
                r["status"], r.get("license",""), r["box"], r["plate"]
            ))

    def search_plate(self):
        query = normalize_search_text(self.ent_search.get().strip())
        if not self.excel_path or not query: return
        self.lbl_search_status.config(text="جارٍ البحث في السجلات...", fg="#475569", bg="#F1F5F9")

        def worker():
            recs = self.get_cached_records()
            found = None
            for r in recs:
                if query == normalize_search_text(r["plate"]):
                    found = r
                    break
            self.root.after(0, lambda: self._finish_search(found))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_search(self, found):
        if found:
            self.found_record = found
            self.lbl_search_status.config(
                text=f"✔ مسجلة: B{found['box']} | {found['status']} | {found['license']} | {found.get('options','-')} | صف {found['row']}", 
                fg="#047857", bg="#ECFDF5"
            )
            self.btn_edit_found.config(state="normal")
            self.btn_go_excel.config(state="normal")
        else:
            self.found_record = None
            self.lbl_search_status.config(text="❌ غير مسجلة في السجلات", fg="#B91C1C", bg="#FEE2E2")
            self.btn_edit_found.config(state="disabled")
            self.btn_go_excel.config(state="disabled")

    def go_to_found_excel(self):
        if not self.found_record or not self.excel_path: return
        row = self.found_record.get("row")
        def work():
            wb, ws = self.get_sheet()
            ws.range(f"{self.col_plate}{row}").select()
            wb.app.visible = True
        try:
            self.with_com(work)
            self.show_toast("✔ تم تحديد خلية اللوحة في Excel")
        except Exception as e:
            self.show_toast(str(e), True)

    def start_edit_found(self):
        if not self.found_record: return
        r = self.found_record
        self.ent_plate.delete(0, tk.END); self.ent_plate.insert(0, r["plate"])
        self.ent_box.delete(0, tk.END); self.ent_box.insert(0, r["box"])
        self.ent_options.delete(0, tk.END); self.ent_options.insert(0, r.get("options", ""))
        
        lic = r.get("license", "")
        if lic in self.license_options:
            self.license_index = self.license_options.index(lic)
        else:
            self.license_index = 0
        self.toggle_license(0)

        st = r.get("status", "جديدة")
        if st in self.status_options:
            self.set_status(self.status_options.index(st))

        self.edit_mode = r.copy()
        self.lbl_mode.config(text=f"✏ وضع التعديل — الصف {r['row']}", bg="#FFF7ED", fg="#9A3412")
        self.btn_save.config(text="تحديث وحفظ (Enter)", bg="#D97706")

    def save_edit(self):
        r = self.edit_mode
        new_p = normalize_plate(self.ent_plate.get())
        new_b = normalize_box(self.ent_box.get())
        license_val = self.license_options[self.license_index]
        status = self.status_options[self.status_index]
        opt_val = self.ent_options.get().strip()

        self.busy = True
        self.btn_save.config(state="disabled")
        self.lbl_mode.config(text="جارٍ تحديث السجل…", bg="#FEF3C7", fg="#92400E")

        def worker():
            def work():
                wb, ws = self.get_sheet()
                row = r["row"]
                ws.range(f"{self.col_plate}{row}").value = f"'{new_p}"
                ws.range(f"{self.col_box}{row}").value = int(new_b)
                ws.range(f"{self.col_notes}{row}").value = license_val
                ws.range(f"{self.col_status}{row}").value = status
                ws.range(f"{self.col_options}{row}").value = opt_val
                wb.save()

            try:
                self.with_com(work)
                self.root.after(0, self._finish_edit)
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda: self._finish_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_edit(self):
        self.busy = False
        self.edit_mode = None
        self.btn_save.config(text="⚡ حفظ مباشر في الإكسل (Enter)", bg="#0284C7", state="normal")
        self.lbl_mode.config(text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857")
        self.ent_plate.delete(0, tk.END)
        self.update_stats_async()
        self.on_box_type()  # تحديث العداد تلقائياً
        self.show_toast("✔ تم تعديل السجل بنجاح")

    def undo_last(self):
        if self.busy or not self.undo_history: return
        rec = self.undo_history[-1]
        self.busy = True

        def worker():
            def work():
                wb, ws = self.get_sheet()
                for c, v in rec["old"].items(): ws.range(f"{c}{rec['row']}").value = v
                wb.save()

            try:
                self.with_com(work)
                self.root.after(0, lambda: self._finish_undo(rec))
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda: self._finish_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_undo(self, rec):
        self.busy = False
        self.undo_history.pop()
        self.refresh_recent()
        self.update_stats_async()
        self.on_box_type()  # تحديث العداد تلقائياً
        self.show_toast(f"↩ تم التراجع عن {rec['plate']}")

    def update_stats_async(self):
        if not self.excel_path: return
        def worker():
            recs = self.get_cached_records()
            total = len(recs)
            new = sum(1 for r in recs if r.get("status") == "جديدة")
            old = sum(1 for r in recs if r.get("status") == "قديمة")
            self.root.after(0, lambda: (
                self.card_total.config(text=str(total)),
                self.card_new.config(text=str(new)),
                self.card_old.config(text=str(old))
            ))
        threading.Thread(target=worker, daemon=True).start()

    def on_box_type_delayed(self, event=None):
        if self._box_debounce_id:
            self.root.after_cancel(self._box_debounce_id)
        self._box_debounce_id = self.root.after(350, self.on_box_type)

    def on_box_type(self):
        b = normalize_box(self.ent_box.get())
        if not b or not self.excel_path:
            self.card_box.config(text="-"); return
        def worker():
            recs = self.get_cached_records()
            c = sum(1 for r in recs if r.get("box") == b)
            self.root.after(0, lambda: self.card_box.config(text=f"{c} لوحة"))
        threading.Thread(target=worker, daemon=True).start()

    def show_toast(self, msg, is_error=False):
        self.lbl_toast.config(text=msg, bg="#DC2626" if is_error else "#15803D", fg="white")
        self.toast_frame.pack(fill="x", side="bottom", padx=20, pady=(0, 4))
        self.lbl_toast.pack(fill="x")
        self.root.after(3500, lambda: self.lbl_toast.pack_forget())

    # ==================== قائمة الخيارات والتقارير ====================
    def open_large_modern_menu(self):
        win = tk.Toplevel(self.root)
        win.transient(self.root); win.lift(); win.focus_force()
        win.title("الخيارات والتقارير المتقدمة")
        win.geometry("480x580")
        win.configure(bg="#0A1D37")
        win.resizable(False, False)

        tk.Label(win, text="الخيارات والعمليات", bg="#0A1D37", fg="#38BDF8", font=("Segoe UI", 14, "bold")).pack(pady=(16, 12))

        def make_btn(txt, cmd, bg_col, sub_txt):
            f = tk.Frame(win, bg="#112B4E", cursor="hand2")
            f.pack(fill="x", padx=20, pady=5)
            f.bind("<Button-1>", lambda e: [win.destroy(), cmd()])
            b = tk.Button(f, text=txt, bg=bg_col, fg="white", font=("Segoe UI", 10, "bold"), relief="flat", width=18, command=lambda: [win.destroy(), cmd()])
            b.pack(side="right", padx=6, pady=6)
            lbl = tk.Label(f, text=sub_txt, bg="#112B4E", fg="#94A3B8", font=("Segoe UI", 8))
            lbl.pack(side="left", padx=10)
            lbl.bind("<Button-1>", lambda e: [win.destroy(), cmd()])

        make_btn("⚙ تخصيص الأعمدة", self.open_column_config_popup, "#D97706", "تعديل أحرف الأعمدة وصف البداية")
        make_btn("📄 تصدير تقرير PDF", self.open_pdf_export_popup, "#DC2626", "طباعة كشف الصندوق مباشرة")
        make_btn("📝 تصدير Excel مفلتر", self.open_notes_export_popup, "#EA580C", "تصدير حسب الرخصة أو الصندوق")
        make_btn("📋 الفهرس الشامل", self.open_records_window, "#2563EB", "استعراض كافة اللوحات المسجلة")
        make_btn("📦 إدارة الصناديق", self.open_box_manager, "#0EA5E9", "كشف الصناديق وعدد اللوحات")
        make_btn("💾 النسخ والاستعادة", self.open_backup_manager, "#059669", "نسخ احتياطي واستعادة يدوية")
        make_btn("🗄 تصدير قاعدة SQLite", self.export_sqlite, "#0F2942", "تصدير نسخة بصيغة DB")

    def open_column_config_popup(self):
        win = tk.Toplevel(self.root)
        win.transient(self.root); win.lift(); win.focus_force()
        win.title("تخصيص الأعمدة وصف البداية")
        win.geometry("450x450")
        win.configure(bg="#F8FAFC")
        win.resizable(False, False)

        tk.Label(win, text="⚙ تخصيص إعدادات ملف Excel", bg="#0A1D37", fg="white", font=("Segoe UI", 13, "bold"), pady=10).pack(fill="x")
        
        frame = tk.Frame(win, bg="#F8FAFC")
        frame.pack(fill="both", expand=True, padx=30, pady=15)

        fields = [
            ("صف بداية البيانات:", "start_row", str(self.start_row)),
            ("عمود اللوحة:", "col_plate", self.col_plate),
            ("عمود الحزمة:", "col_box", self.col_box),
            ("عمود التاريخ:", "col_date", self.col_date),
            ("عمود الرخصة D:", "col_notes", self.col_notes),
            ("عمود الحالة E:", "col_status", self.col_status),
            ("عمود الخيارات F:", "col_options", self.col_options),
        ]

        entries = {}
        for i, (label, key, val) in enumerate(fields):
            tk.Label(frame, text=label, bg="#F8FAFC", font=("Segoe UI", 10, "bold")).grid(row=i, column=1, sticky="e", pady=4)
            e = tk.Entry(frame, justify="center", font=("Segoe UI", 10, "bold"), width=8)
            e.insert(0, val)
            e.grid(row=i, column=0, pady=4, padx=(0, 20))
            entries[key] = e

        def save_cfg():
            try:
                self.start_row = int(entries["start_row"].get().strip())
                self.col_plate = entries["col_plate"].get().strip().upper()
                self.col_box = entries["col_box"].get().strip().upper()
                self.col_date = entries["col_date"].get().strip().upper()
                self.col_notes = entries["col_notes"].get().strip().upper()
                self.col_status = entries["col_status"].get().strip().upper()
                self.col_options = entries["col_options"].get().strip().upper()
                self.update_stats_async()
                self.on_box_type()
                win.destroy()
                self.show_toast("✔ تم حفظ تخصيص الأعمدة بنجاح")
            except Exception as ex:
                messagebox.showerror("خطأ", f"يرجى التأكد من صحة المدخلات: {ex}", parent=win)

        tk.Button(win, text="💾 حفظ الإعدادات", bg="#059669", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", pady=8, command=save_cfg).pack(fill="x", padx=30, pady=(0, 15))

    def open_pdf_export_popup(self):
        if not self.excel_path: self.show_toast("اختر ملف إكسل أولاً!", True); return
        win = tk.Toplevel(self.root); win.transient(self.root); win.lift(); win.focus_force()
        win.title("تصدير تقرير PDF حسب الصندوق"); win.geometry("380x230"); win.configure(bg="white"); win.resizable(False, False)
        tk.Label(win, text="تصدير كشف رسمي للوحات الصندوق", bg="white", fg="#0A1D37", font=("Segoe UI", 12, "bold")).pack(pady=(20, 5))
        tk.Label(win, text="اكتب رقم الصندوق المطلوب (مثال: 10):", bg="white", fg="#64748B", font=("Segoe UI", 9)).pack()
        e = tk.Entry(win, justify="center", font=("Segoe UI", 13), relief="solid", bd=1); e.pack(pady=8, ipady=3); e.focus()

        def do():
            box = normalize_box(e.get())
            if not box: messagebox.showwarning("تنبيه", "يرجى إدخال رقم الصندوق!", parent=win); return
            path = filedialog.asksaveasfilename(title=f"حفظ تقرير صندوق B{box}", initialfile=f"تقرير_صندوق_B{box}.pdf", defaultextension=".pdf", filetypes=[("PDF File", "*.pdf")])
            if not path: return
            win.destroy()
            self._export_pdf_instant(path, box)

        tk.Button(win, text="📄 إنشاء وتصدير PDF", bg="#DC2626", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=6, command=do).pack(fill="x", padx=35, pady=10)

    def _export_pdf_instant(self, path, box):
        b_target = normalize_box(box)
        self.show_toast("جارٍ إنشاء ملف PDF...")
        def worker():
            def work():
                wb, _ = self.get_sheet()
                app = wb.app
                records = [r for r in self.get_cached_records() if normalize_box(r.get("box")) == b_target]
                if not records: raise RuntimeError(f"لا توجد لوحات مسجلة للصندوق B{b_target}!")
                app.screen_updating = False; app.display_alerts = False
                sheet_name = "_PDF_Temp_Print_"
                for sh in wb.sheets:
                    if sh.name == sheet_name: sh.delete()
                ws = wb.sheets.add(sheet_name, after=wb.sheets[len(wb.sheets)-1])
                ws.api.DisplayRightToLeft = True

                # سطر عنوان واحد مدمج يجمع اسم التقرير والصندوق والتاريخ لتوفير المساحة بالكامل
                ws.range("A1:F1").merge()
                ws.range("A1").value = f"أرشيف اللوحات — كشف لوحات الصندوق (B{b_target}) — تاريخ الطباعة: {datetime.now().strftime('%d/%m/%Y')}"
                ws.range("A1").font.bold = True
                ws.range("A1").font.size = 11
                ws.range("A1").api.HorizontalAlignment = -4108
                ws.range("A1:F1").row_height = 24

                # ترويسة الجدول تبدأ مباشرة من السطر 2
                ws.range("A2:F2").value = ["ت", "رقم اللوحة", "الحزمة", "التاريخ", "الرخصة", "الحالة"]
                ws.range("A2:F2").color = "#0A1D37"
                ws.range("A2:F2").api.Font.Color = 0xFFFFFF
                ws.range("A2:F2").api.Font.Bold = True
                ws.range("A2:F2").api.HorizontalAlignment = -4108
                ws.range("A2:F2").row_height = 20

                rows = [[i, f"'{r['plate']}", f"B{r['box']}", f"'{r['date']}", r.get('license',''), r.get('status','')] for i, r in enumerate(records, 1)]
                end_r = 2 + len(rows)
                ws.range(f"A3:F{end_r}").value = rows
                ws.range(f"A2:F{end_r}").api.Borders.LineStyle = 1
                ws.range(f"A2:F{end_r}").api.HorizontalAlignment = -4108
                ws.range(f"A3:F{end_r}").row_height = 18

                for col, w in zip("ABCDEF", [5, 14, 10, 12, 28, 10]):
                    ws.range(f"{col}:{col}").column_width = w

                # إعداد الصفحة: هوامش ضيقة وضغط تلقائي ليتسع بصفحة واحدة (1 Page Wide x 1 Page Tall)
                ps = ws.api.PageSetup
                ps.Orientation = 1  # طولي Portrait
                ps.PaperSize = 9    # A4
                ps.TopMargin = app.api.InchesToPoints(0.3)
                ps.BottomMargin = app.api.InchesToPoints(0.3)
                ps.LeftMargin = app.api.InchesToPoints(0.3)
                ps.RightMargin = app.api.InchesToPoints(0.3)
                ps.CenterHorizontally = True
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = 1

                ws.api.ExportAsFixedFormat(0, os.path.abspath(path))
                ws.delete(); app.screen_updating = True; app.display_alerts = True

            try:
                self.with_com(work)
                self.root.after(0, lambda: self.show_toast(f"✔ تم إنشاء تقرير PDF للصندوق B{b_target}"))
                try: os.startfile(path)
                except Exception: pass
            except Exception as ex:
                msg = str(ex)
                self.root.after(0, lambda: self.show_toast(msg, True))

        threading.Thread(target=worker, daemon=True).start()

    def open_notes_export_popup(self):
        if not self.excel_path: self.show_toast("اختر ملف إكسل أولاً!", True); return
        win = tk.Toplevel(self.root); win.transient(self.root); win.lift(); win.focus_force()
        win.title("تصدير Excel مفلتر"); win.geometry("480x300"); win.configure(bg="white"); win.resizable(False, False)
        tk.Label(win, text="📊 تصدير سجلات مخصصة إلى ملف Excel", bg="white", fg="#0A1D37", font=("Segoe UI", 12, "bold")).pack(pady=(15, 8))

        form = tk.Frame(win, bg="white"); form.pack(fill="x", padx=30, pady=5)
        tk.Label(form, text="حالة الرخصة:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="e", pady=5)
        e_lic = ttk.Combobox(form, values=["الكل"] + self.license_options, state="readonly", justify="center")
        e_lic.set("الكل")
        e_lic.grid(row=0, column=0, pady=5, padx=5, sticky="ew")

        tk.Label(form, text="الصندوق (اختياري):", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=1, sticky="e", pady=5)
        e_box = tk.Entry(form, justify="center", font=("Segoe UI", 10), relief="solid", bd=1)
        e_box.grid(row=1, column=0, pady=5, padx=5, sticky="ew")
        form.grid_columnconfigure(0, weight=1)

        def export():
            lq = e_lic.get()
            bq = normalize_box(e_box.get())
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="تصدير_مخصص.xlsx")
            if not path: return
            win.destroy()
            self.show_toast("جارٍ تجهيز ملف Excel المفلتر...")

            def worker():
                records = self.get_cached_records()
                filtered = [
                    r for r in records 
                    if (lq == "الكل" or r.get("license") == lq) and (not bq or normalize_box(r.get("box")) == bq)
                ]
                if not filtered:
                    self.root.after(0, lambda: self.show_toast("لا توجد سجلات تطابق الشروط!", True))
                    return

                def work():
                    app = xw.App(visible=False, add_book=False)
                    wb = app.books.add(); ws = wb.sheets[0]; ws.api.DisplayRightToLeft = True
                    ws.range("A1:G1").value = ["رقم اللوحة", "رقم الحزمة", "تاريخ الإضافة", "حالة الرخصة", "الحالة", "خيارات F", "صف Excel"]
                    rows = [
                        [f"'{r['plate']}", f"B{r['box']}", f"'{r['date']}", r.get('license',''), r.get('status',''), r.get('options',''), r['row']] 
                        for r in filtered
                    ]
                    ws.range(f"A2:G{len(rows)+1}").value = rows
                    ws.range("A1:G1").font.bold = True; ws.range("A1:G1").color = "#0A1D37"; ws.range("A1:G1").api.Font.Color = 0xFFFFFF
                    wb.save(path); wb.close(); app.quit()

                try:
                    self.with_com(work)
                    self.root.after(0, lambda: self.show_toast(f"✔ تم تصدير {len(filtered)} سجل إلى Excel"))
                    try: os.startfile(path)
                    except Exception: pass
                except Exception as e:
                    msg = str(e)
                    self.root.after(0, lambda: self.show_toast(msg, True))

            threading.Thread(target=worker, daemon=True).start()

        tk.Button(win, text="📊 تصدير الآن إلى Excel", bg="#EA580C", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=7, command=export).pack(fill="x", padx=30, pady=15)

    def open_records_window(self):
        if not self.excel_path: self.show_toast("اختر ملف إكسل أولاً!", True); return
        win = tk.Toplevel(self.root); win.title("الفهرس الشامل للوحات"); win.geometry("960x550"); win.configure(bg="#F0F4F8")
        lbl_head = tk.Label(win, text="جارٍ تحميل السجلات...", bg="#0A1D37", fg="white", font=("Segoe UI", 12, "bold"), pady=8)
        lbl_head.pack(fill="x")
        frame = tk.Frame(win, bg="white"); frame.pack(fill="both", expand=True, padx=12, pady=10)
        cols = ("row", "plate", "box", "date", "lic", "status", "opt")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c, t, w in (
            ("row","صف",45),
            ("plate","اللوحة",110),
            ("box","الحزمة",65),
            ("date","التاريخ",90),
            ("lic","حالة الرخصة D",170),
            ("status","الحالة E",75),
            ("opt","خيارات F",90)
        ):
            tree.heading(c, text=t); tree.column(c, width=w, anchor="center")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set); sb.pack(side="left", fill="y"); tree.pack(fill="both", expand=True)

        def worker():
            records = self.get_cached_records()
            def update():
                lbl_head.config(text=f"الفهرس الشامل | إجمالي اللوحات: {len(records)}")
                for r in records: 
                    tree.insert("", "end", values=(r["row"], r["plate"], f"B{r['box']}", r["date"], r["license"], r["status"], r["options"]))
            self.root.after(0, update)

        threading.Thread(target=worker, daemon=True).start()

    def open_box_manager(self):
        if not self.excel_path: self.show_toast("اختر ملف إكسل أولاً!", True); return
        win = tk.Toplevel(self.root); win.title("إدارة الصناديق"); win.geometry("650x450"); win.configure(bg="#F0F4F8")
        lbl_head = tk.Label(win, text="جارٍ إحصاء الصناديق...", bg="#0A1D37", fg="white", font=("Segoe UI", 12, "bold"), pady=8)
        lbl_head.pack(fill="x")
        frame = tk.Frame(win, bg="white"); frame.pack(fill="both", expand=True, padx=12, pady=10)
        tree = ttk.Treeview(frame, columns=("box", "count", "new", "old"), show="headings")
        for c, t, w in (("box","الحزمة/الصندوق",140),("count","إجمالي اللوحات",120),("new","جديدة",100),("old","قديمة",100)):
            tree.heading(c, text=t); tree.column(c, width=w, anchor="center")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set); sb.pack(side="left", fill="y"); tree.pack(fill="both", expand=True)

        def worker():
            recs = self.get_cached_records()
            boxes = {}
            for r in recs: boxes.setdefault(r.get("box",""), []).append(r)
            def update():
                lbl_head.config(text=f"كشف الصناديق | إجمالي الحزم: {len(boxes)}")
                for b, items in sorted(boxes.items(), key=lambda x: str(x[0])):
                    tree.insert("", "end", values=(f"B{b}" if b else "-", len(items), sum(1 for x in items if x.get("status")=="جديدة"), sum(1 for x in items if x.get("status")=="قديمة")))
            self.root.after(0, update)

        threading.Thread(target=worker, daemon=True).start()

    def open_backup_manager(self):
        if not self.excel_path: self.show_toast("اختر ملف إكسل أولاً!", True); return
        self._init_storage_paths()
        win = tk.Toplevel(self.root); win.title("النسخ الاحتياطي والاستعادة"); win.geometry("680x420"); win.configure(bg="#F0F4F8")
        tk.Label(win, text="النسخ الاحتياطي والاستعادة اليدوية", bg="#0A1D37", fg="white", font=("Segoe UI", 12, "bold"), pady=8)
        tree = ttk.Treeview(win, columns=("name", "date"), show="headings")
        for c, t, w in (("name","اسم ملف النسخة",420),("date","التاريخ",180)):
            tree.heading(c, text=t); tree.column(c, width=w, anchor="center")
        tree.pack(fill="both", expand=True, padx=15, pady=10)
        def refresh():
            tree.delete(*tree.get_children())
            if self.backup_dir and os.path.isdir(self.backup_dir):
                files = [os.path.join(self.backup_dir, f) for f in os.listdir(self.backup_dir) if f.endswith((".xlsx", ".xlsm"))]
                for f in sorted(files, key=os.path.getmtime, reverse=True):
                    tree.insert("", "end", iid=f, values=(os.path.basename(f), datetime.fromtimestamp(os.path.getmtime(f)).strftime("%d/%m/%Y %H:%M:%S")))
        def do_backup():
            self._safety_backup("manual"); refresh(); self.show_toast("✔ تم أخذ نسخة احتياطية الآن")
        def do_restore():
            sel = tree.selection()
            if not sel: return
            if messagebox.askyesno("تأكيد", "هل تريد استعادة النسخة المحددة واستبدال الملف الحالي؟", parent=win):
                shutil.copy2(sel[0], self.excel_path)
                self.update_stats_async(); self.on_box_type(); self.show_toast("✔ تمت الاستعادة بنجاح")
        bar = tk.Frame(win, bg="#F0F4F8"); bar.pack(fill="x", padx=15, pady=8)
        tk.Button(bar, text="💾 إنشاء نسخة الآن", command=do_backup, bg="#059669", fg="white", relief="flat", pady=6).pack(side="right", padx=5)
        tk.Button(bar, text="♻ استعادة المحددة", command=do_restore, bg="#DC2626", fg="white", relief="flat", pady=6).pack(side="right", padx=5)
        refresh()

    def export_sqlite(self):
        if not self.excel_path: return
        path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite DB", "*.db")])
        if not path: return
        self.show_toast("جارٍ تصدير قاعدة بيانات SQLite...")
        def worker():
            recs = self.get_cached_records()
            with sqlite3.connect(path) as conn:
                c = conn.cursor()
                c.execute("DROP TABLE IF EXISTS plates")
                c.execute("CREATE TABLE plates (plate TEXT UNIQUE, box TEXT, date TEXT, license TEXT, status TEXT, options TEXT)")
                for r in recs:
                    c.execute("INSERT OR REPLACE INTO plates VALUES (?,?,?,?,?,?)", (r["plate"], f"B{r['box']}", r["date"], r["license"], r["status"], r["options"]))
                conn.commit()
            self.root.after(0, lambda: self.show_toast("✔ تم تصدير SQLite بنجاح"))
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = PlateManagerMilitaryApp(root)
    root.mainloop()
