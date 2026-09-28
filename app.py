# -*- coding: utf-8 -*-
APP_VERSION = "3.1.0"
APP_NAME = "Plate Archive Pro"

import os
import sys
import math
import random
import sqlite3
import shutil
import ctypes
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog

try:
    import xlwings as xw
except ImportError:
    xw = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

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
    "✨ لا حَوْلَ وَلا قُوَّةَ إِلا بِاللَّهِ الْعَلِيِّ الْعَظِيمِ",
    "🌸 اللَّهُمَّ صَلِّ وَسَلِّمْ وَبَارِكْ عَلَى نَبِيِّنَا مُحَمَّدٍ",
    "💡 حكمة: مَنْ تَأَنَّى أَصَابَ أَوْ كَادَ، وَمَنْ عَجِلَ أَخْطَأَ أَوْ كَادَ.",
    "💡 حكمة: الإِخْلاصُ فِي الْعَمَلِ سَبَبٌ لِلْقَبُولِ وَالتَّوْفِيقِ."
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
        self.root.geometry("1340x840")
        self.root.configure(bg="#F0F4F8")

        self.excel_path = None
        self.busy = False
        self.edit_mode = None
        self.editing_record_id = None
        self.undo_history = []
        self.max_undo = 10
        self.pinned_notes = tk.BooleanVar(value=False)
        self.last_note = ""
        self.data_cache = []
        self.audit_path = None
        self.backup_dir = None

        self.start_row = 5
        self.col_plate = "A"
        self.col_box = "B"
        self.col_date = "C"
        self.col_notes = "D"
        self.col_status = "E"
        self.status_options = ["جديدة", "قديمة"]
        self.status_var = tk.StringVar(value="جديدة")

        self.prayer_visible = True

        self.setup_ui()
        self.setup_keyboard_nav()
        self.update_clock()
        self.update_dhikr_cycle()
        self.update_prayer_countdown()

    def setup_keyboard_nav(self):
        self.root.bind("<Control-z>", lambda e: self.undo_last())

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

        btn_menu = tk.Button(tools_box, text=" ☰ الخيارات المتقدمة ", bg="#1E293B", fg="#38BDF8", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", padx=12, pady=5, command=self.open_large_modern_menu)
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

        tk.Label(body, text="* رقم اللوحة:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_plate = tk.Entry(body, font=("Segoe UI", 16, "bold"), bg="#F8FAFC", fg="#0F172A", relief="solid", bd=1, justify="center")
        self.ent_plate.pack(fill="x", ipady=5, pady=(0, 8))
        self.ent_plate.bind("<Return>", lambda e: self.save_plate_direct())

        tk.Label(body, text="* رقم الحزمة (مثال: 15 أو B15):", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_box = tk.Entry(body, font=("Segoe UI", 13), bg="#F8FAFC", relief="solid", bd=1, justify="center")
        self.ent_box.pack(fill="x", ipady=4, pady=(0, 8))
        self.ent_box.bind("<KeyRelease>", self.on_box_type)
        self.ent_box.bind("<Return>", lambda e: self.save_plate_direct())

        tk.Label(body, text="تاريخ الإضافة:", bg="#FFFFFF", fg="#64748B", font=("Segoe UI", 9, "bold")).pack(anchor="e")
        self.lbl_date_preview = tk.Label(body, text=datetime.now().strftime("%d/%m/%Y"), bg="#F1F5F9", fg="#334155", font=("Segoe UI", 11, "bold"), pady=5)
        self.lbl_date_preview.pack(fill="x", pady=(0, 8))

        tk.Label(body, text="الملاحظات:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        nf = tk.Frame(body, bg="#FFFFFF"); nf.pack(fill="x", pady=(0, 8))
        self.ent_notes = tk.Entry(nf, font=("Segoe UI", 11), bg="#F8FAFC", relief="solid", bd=1)
        self.ent_notes.pack(side="right", fill="x", expand=True, ipady=4)
        tk.Checkbutton(nf, text="📌 تثبيت", variable=self.pinned_notes, bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        tk.Button(nf, text="مسح", bg="#FEE2E2", fg="#B91C1C", relief="flat", command=lambda: self.ent_notes.delete(0, tk.END)).pack(side="left")

        tk.Label(body, text="* حالة اللوحة:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        status_row = tk.Frame(body, bg="#FFFFFF"); status_row.pack(fill="x", pady=(0, 8))
        for value in self.status_options:
            tk.Radiobutton(status_row, text=value, value=value, variable=self.status_var, bg="#FFFFFF", fg="#0F172A", font=("Segoe UI", 11, "bold"), selectcolor="#E0F2FE", indicatoron=False, padx=20, pady=5).pack(side="right", fill="x", expand=True, padx=2)

        bf = tk.Frame(body, bg="#FFFFFF"); bf.pack(fill="x", pady=(5, 0))
        self.btn_save = tk.Button(bf, text="⚡ حفظ مباشر في الإكسل (Enter)", bg="#0284C7", fg="#FFFFFF", font=("Segoe UI", 12, "bold"), relief="flat", cursor="hand2", pady=7, command=self.save_plate_direct)
        self.btn_save.pack(side="right", fill="x", expand=True, padx=2)
        self.btn_undo = tk.Button(bf, text="↩ تراجع واستبدال الفراغ", bg="#64748B", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", pady=7, command=self.undo_last)
        self.btn_undo.pack(side="left", fill="x", expand=True, padx=2)

        self.lbl_mode = tk.Label(body, text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857", font=("Segoe UI", 9, "bold"), pady=5)
        self.lbl_mode.pack(fill="x", pady=5)

    def setup_left_card(self, parent):
        tk.Label(parent, text="🔍 البحث السريع + آخر الإدخالات", bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 13, "bold")).pack(anchor="e", padx=20, pady=(8, 4))
        body = tk.Frame(parent, bg="#FFFFFF"); body.pack(fill="both", expand=True, padx=20, pady=2)

        tk.Label(body, text="اكتب رقم اللوحة للبحث:", bg="#FFFFFF", fg="#1E293B", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(2, 2))
        self.ent_search = tk.Entry(body, font=("Segoe UI", 13), bg="#F8FAFC", relief="solid", bd=1, justify="center")
        self.ent_search.pack(fill="x", ipady=5, pady=(0, 6))
        self.ent_search.bind("<Return>", lambda e: self.search_plate())

        tk.Button(body, text="فحص السجل 🔍", bg="#0A1D37", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", pady=5, command=self.search_plate).pack(fill="x", pady=(0, 4))
        self.lbl_search_status = tk.Label(body, text="لم يتم إجراء بحث", bg="#F1F5F9", fg="#475569", font=("Segoe UI", 10), pady=6, wraplength=540)
        self.lbl_search_status.pack(fill="x", pady=(0, 4))

        self.btn_edit_found = tk.Button(body, text="✏ تعديل السجل الموجود مباشرة", bg="#D97706", fg="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat", state="disabled", command=self.start_edit_found)
        self.btn_edit_found.pack(fill="x", pady=2)

        lf = tk.LabelFrame(body, text="آخر 10 إدخالات", bg="#FFFFFF", fg="#0A1D37", font=("Segoe UI", 9, "bold"), padx=5, pady=5)
        lf.pack(fill="both", expand=True, pady=6)
        
        cols = ("date", "row", "status", "box", "plate")
        self.recent_tree = ttk.Treeview(lf, columns=cols, show="headings", height=8)
        for c, t, w in (("date","التاريخ",125),("row","صف",65),("status","الحالة",90),("box","الحزمة",80),("plate","اللوحة",130)):
            self.recent_tree.heading(c, text=t); self.recent_tree.column(c, width=w, anchor="center")

        sb = ttk.Scrollbar(lf, orient="vertical", command=self.recent_tree.yview)
        self.recent_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.recent_tree.pack(fill="both", expand=True)

    def select_file(self):
        path = filedialog.askopenfilename(title="اختر ملف الإكسل", filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")])
        if path:
            self.excel_path = os.path.normpath(path)
            self._init_storage_paths()
            self.lbl_file.config(text=os.path.basename(path), fg="#38BDF8")
            self.update_stats_async()
            self.show_toast("✔ تم ربط ملف Excel بنجاح")

    def get_sheet(self):
        if not self.excel_path: raise RuntimeError("يرجى ربط ملف الإكسل أولاً!")
        if xw is None: raise RuntimeError("مكتبة xlwings غير مثبتة")
        wb = xw.Book(self.excel_path)
        return wb, wb.sheets[0]

    def with_com(self, func, *args, **kwargs):
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
            records = []
            for i in range(count):
                plate = normalize_plate(p_vals[i])
                if plate and plate != "0":
                    records.append({"plate": plate, "box": normalize_box(b_vals[i]), "date": excel_date_text(d_vals[i]), "notes": normalize_plate(n_vals[i]), "status": normalize_plate(s_vals[i]) or "جديدة", "row": self.start_row + i})
            return records
        return self.with_com(work)

    def save_plate_direct(self):
        if self.edit_mode:
            self.save_edit(); return
        if not self.excel_path:
            self.show_toast("⚠ اختر ملف Excel أولاً!", True); return
        plate = normalize_plate(self.ent_plate.get())
        box = normalize_box(self.ent_box.get())
        status = self.status_var.get()
        if not plate or not box:
            self.show_toast("رقم اللوحة والحزمة مطلوبان!", True); return

        notes = self.ent_notes.get().strip()
        now = datetime.now().strftime("%d/%m/%Y")
        
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
            old = {c: ws.range(f"{c}{nxt}").value for c in (self.col_plate, self.col_box, self.col_date, self.col_notes, self.col_status)}
            ws.range(f"{self.col_plate}{nxt}").value = f"'{plate}"
            ws.range(f"{self.col_box}{nxt}").value = int(box)
            ws.range(f"{self.col_date}{nxt}").value = f"'{now}"
            ws.range(f"{self.col_notes}{nxt}").value = notes
            ws.range(f"{self.col_status}{nxt}").value = status
            wb.save()
            return {"plate": plate, "box": box, "notes": notes, "status": status, "date": now, "row": nxt, "old": old}

        try:
            res = self.with_com(work)
            self.undo_history.append(res)
            self.refresh_recent()
            self.ent_plate.delete(0, tk.END)
            self.ent_box.delete(0, tk.END)
            if not self.pinned_notes.get(): self.ent_notes.delete(0, tk.END)
            self.ent_plate.focus()
            self.update_stats_async()
            self.show_toast(f"✔ تم حفظ {plate} في الحزمة B{box}")
        except Exception as e:
            self.show_toast(str(e), True)

    def refresh_recent(self):
        for x in self.recent_tree.get_children(): self.recent_tree.delete(x)
        for r in reversed(self.undo_history[-10:]):
            self.recent_tree.insert("", "end", values=(r["date"], r["row"], r["status"], r["box"], r["plate"]))

    def search_plate(self):
        query = normalize_search_text(self.ent_search.get().strip())
        if not self.excel_path or not query: return
        recs = self.get_cached_records()
        for r in recs:
            if query == normalize_search_text(r["plate"]):
                self.found_record = r
                self.lbl_search_status.config(text=f"✔ اللوحة مسجلة: B{r['box']} | {r['status']} | صف {r['row']}", fg="#047857", bg="#ECFDF5")
                self.btn_edit_found.config(state="normal")
                return
        self.found_record = None
        self.lbl_search_status.config(text="❌ غير مسجلة في السجلات", fg="#B91C1C", bg="#FEE2E2")
        self.btn_edit_found.config(state="disabled")

    def start_edit_found(self):
        if not self.found_record: return
        r = self.found_record
        self.ent_plate.delete(0, tk.END); self.ent_plate.insert(0, r["plate"])
        self.ent_box.delete(0, tk.END); self.ent_box.insert(0, r["box"])
        self.ent_notes.delete(0, tk.END); self.ent_notes.insert(0, r["notes"])
        self.status_var.set(r["status"])
        self.edit_mode = r.copy()
        self.lbl_mode.config(text=f"✏ وضع التعديل — الصف {r['row']}", bg="#FFF7ED", fg="#9A3412")
        self.btn_save.config(text="تحديث وحفظ (Enter)", bg="#D97706")

    def save_edit(self):
        r = self.edit_mode
        new_p = normalize_plate(self.ent_plate.get())
        new_b = normalize_box(self.ent_box.get())
        notes = self.ent_notes.get().strip()
        status = self.status_var.get()

        def work():
            wb, ws = self.get_sheet()
            row = r["row"]
            ws.range(f"{self.col_plate}{row}").value = f"'{new_p}"
            ws.range(f"{self.col_box}{row}").value = int(new_b)
            ws.range(f"{self.col_notes}{row}").value = notes
            ws.range(f"{self.col_status}{row}").value = status
            wb.save()

        try:
            self.with_com(work)
            self.edit_mode = None
            self.btn_save.config(text="⚡ حفظ مباشر في الإكسل (Enter)", bg="#0284C7")
            self.lbl_mode.config(text="وضع: إضافة جديدة", bg="#ECFDF5", fg="#047857")
            self.ent_plate.delete(0, tk.END); self.ent_box.delete(0, tk.END)
            self.update_stats_async()
            self.show_toast("✔ تم تعديل السجل بنجاح")
        except Exception as e:
            self.show_toast(str(e), True)

    def undo_last(self):
        if not self.undo_history: return
        rec = self.undo_history[-1]
        def work():
            wb, ws = self.get_sheet()
            for c, v in rec["old"].items(): ws.range(f"{c}{rec['row']}").value = v
            wb.save()
        try:
            self.with_com(work)
            self.undo_history.pop()
            self.refresh_recent()
            self.update_stats_async()
            self.show_toast(f"↩ تم التراجع عن {rec['plate']}")
        except Exception as e:
            self.show_toast(str(e), True)

    def update_stats_async(self):
        recs = self.get_cached_records()
        total = len(recs)
        new = sum(1 for r in recs if r.get("status") == "جديدة")
        old = sum(1 for r in recs if r.get("status") == "قديمة")
        self.card_total.config(text=str(total))
        self.card_new.config(text=str(new))
        self.card_old.config(text=str(old))

    def on_box_type(self, event=None):
        b = normalize_box(self.ent_box.get())
        if not b: self.card_box.config(text="-"); return
        recs = self.get_cached_records()
        c = sum(1 for r in recs if r.get("box") == b)
        self.card_box.config(text=f"{c} لوحة")

    def show_toast(self, msg, is_error=False):
        self.lbl_toast.config(text=msg, bg="#DC2626" if is_error else "#15803D", fg="white")
        self.toast_frame.pack(fill="x", side="bottom", padx=20, pady=(0, 4))
        self.lbl_toast.pack(fill="x")
        self.root.after(3500, lambda: self.lbl_toast.pack_forget())

    def open_large_modern_menu(self):
        win = tk.Toplevel(self.root)
        win.title("الخيارات")
        win.geometry("400x300")
        win.configure(bg="#0A1D37")
        tk.Label(win, text="الخيارات المتاحة", bg="#0A1D37", fg="#38BDF8", font=("Segoe UI", 12, "bold")).pack(pady=15)
        tk.Button(win, text="📂 استبدال ملف الإكسل", bg="#0284C7", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=6, command=lambda: [win.destroy(), self.select_file()]).pack(fill="x", padx=40, pady=5)
        tk.Button(win, text="🗄 تصدير قاعدة بيانات SQLite", bg="#0F2942", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=6, command=lambda: [win.destroy(), self.export_sqlite()]).pack(fill="x", padx=40, pady=5)

    def export_sqlite(self):
        if not self.excel_path: return
        path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite DB", "*.db")])
        if not path: return
        recs = self.get_cached_records()
        with sqlite3.connect(path) as conn:
            c = conn.cursor()
            c.execute("DROP TABLE IF EXISTS plates")
            c.execute("CREATE TABLE plates (plate TEXT UNIQUE, box TEXT, date TEXT, notes TEXT, status TEXT)")
            for r in recs:
                c.execute("INSERT OR REPLACE INTO plates VALUES (?,?,?,?,?)", (r["plate"], f"B{r['box']}", r["date"], r["notes"], r["status"]))
            conn.commit()
        self.show_toast("✔ تم تصدير SQLite بنجاح")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlateManagerMilitaryApp(root)
    root.mainloop()
