#!/usr/bin/env python3
"""Builds the simple data entry workbook: one sheet per month, horizontal, one row per indicator (two for ratio
indicators: البسط and المقام) and one column per day. All visible text is Arabic. The platform's Data Management page
imports it (dmLocateSimpleSheet in index.html): rows are matched by the hidden code column, or by the Arabic indicator name.

Usage: python3 tools/build_simple_template.py [--year 2026] [--out templates/APH_KPI_Simple_2026.xlsx]
"""
import argparse, datetime as dt, json, os, subprocess

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHEET_KPIS, SHEET_HELP = "المؤشرات", "التعليمات"
MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
DATE_HEADER, DAY_HEADER = "التاريخ", "اليوم"
NUM_SUFFIX, DEN_SUFFIX = "البسط", "المقام"
WEEKDAYS = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
FREQ_AR = {"daily": "يومي", "weekly": "أسبوعي", "monthly": "شهري", "quarterly": "ربع سنوي", "annual": "سنوي"}
ENTRY_DAY = {"daily": "كل يوم", "weekly": "يوم الأحد (بداية الأسبوع)", "monthly": "أول يوم في الشهر",
             "quarterly": "أول يوم في الربع", "annual": "أول يوم في السنة"}

NAVY = "0B1F33"
F_TITLE = Font(name="Arial", size=14, bold=True, color=NAVY)
F_HEAD = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_DEP = Font(name="Arial", size=9, bold=True, color="334155")
F_BODY = Font(name="Arial", size=10)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_DEP = PatternFill("solid", fgColor="E2E8F0")
FILL_DATE = PatternFill("solid", fgColor="F1F5F9")
FILL_OFF = PatternFill("solid", fgColor="D9D9D9")   # not this indicator's entry day
FILL_FRI = PatternFill("solid", fgColor="F8FAFC")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def platform(year):
    out = subprocess.run(["node", os.path.join(ROOT, "tools", "extract_kpis.js"), str(year)], capture_output=True, text=True, check=True).stdout
    return json.loads(out)


def entry_day(freq, d, year):
    """True on the row where this indicator's reading is entered (the first day of its period in the year)."""
    if freq == "daily": return True
    if freq == "weekly": return d.weekday() == 6 or d == dt.date(year, 1, 1)
    if freq == "monthly": return d.day == 1
    if freq == "quarterly": return d.day == 1 and d.month in (1, 4, 7, 10)
    return d.month == 1 and d.day == 1


def build(year, out):
    P = platform(year)
    derived = set(P["derived"])
    kpis = [k for k in P["kpis"] if k["code"] not in derived and k.get("active") is not False]
    for k in kpis:   # Arabic-only text in the workbook
        ar = P["arabicLabels"].get(k["code"], {})
        k["name"], k["definition"] = ar.get("label") or k["name"], ar.get("definition") or k.get("definition") or ""
    deps = [d["nameAr"] for d in P["departments"]]
    kpis.sort(key=lambda k: deps.index(k["department"]) if k["department"] in deps else 99)
    ratio = lambda k: P["calcTypes"][k["calcType"]]["mult"] > 0

    # columns: (kpi, part) part = "" value | "NUM" | "DEN"
    cols = []
    for k in kpis:
        if ratio(k): cols += [(k, "NUM"), (k, "DEN")]
        else: cols.append((k, ""))

    wb = Workbook()
    wb.remove(wb.active)
    # One sheet per month, horizontal: one row per indicator (two for ratio indicators), one column per day.
    # Column A (hidden) holds the indicator code, read by the platform so a renamed row still imports.
    # Row 1: dates. Row 2: weekday. Rows 3+: indicators. Column B: department. Column C: indicator (Arabic).
    dv_err = ("اكتب رقماً فقط (صفر أو أكثر) بدون رموز أو نص.", "قيمة غير صحيحة")
    for month in range(1, 13):
        ws = wb.create_sheet(MONTHS[month - 1])
        ws.sheet_view.rightToLeft = True
        days = []
        d = dt.date(year, month, 1)
        while d.month == month: days.append(d); d += dt.timedelta(days=1)
        ws.cell(row=1, column=1, value="CODE")
        for col, text in ((2, "الإدارة"), (3, DATE_HEADER)):
            ws.cell(row=1, column=col, value=text)
        ws.cell(row=2, column=3, value=DAY_HEADER)
        for j, day in enumerate(days, 4):
            a = ws.cell(row=1, column=j, value=day)
            a.number_format = "dd/mm"
            ws.cell(row=2, column=j, value=WEEKDAYS[day.weekday()])
        for j in range(1, len(days) + 4):
            for r in (1, 2):
                c = ws.cell(row=r, column=j)
                c.font, c.fill, c.alignment, c.border = (F_HEAD, FILL_HEAD, CENTER, BORDER) if r == 1 else (F_DEP, FILL_DEP, CENTER, BORDER)
        for i, (k, part) in enumerate(cols, 3):
            ws.cell(row=i, column=1, value=k["code"] + (":" + part if part else ""))
            b = ws.cell(row=i, column=2, value=k["department"])
            label = k["name"] + (f" — {NUM_SUFFIX}" if part == "NUM" else f" — {DEN_SUFFIX}" if part == "DEN" else "")
            c = ws.cell(row=i, column=3, value=label)
            for x in (b, c): x.font, x.fill, x.border, x.alignment = F_BODY, FILL_DATE, BORDER, RIGHT
            for j, day in enumerate(days, 4):
                v = ws.cell(row=i, column=j)
                v.border = BORDER
                v.protection = Protection(locked=False)
                if not entry_day(k["frequency"], day, year): v.fill = FILL_OFF
        ws.column_dimensions["A"].hidden = True
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 44
        for j in range(4, len(days) + 4): ws.column_dimensions[get_column_letter(j)].width = 8
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
        dv.error, dv.errorTitle = dv_err
        dv.add(f"D3:{get_column_letter(len(days) + 3)}{len(cols) + 2}")
        ws.add_data_validation(dv)
        ws.freeze_panes = "D3"
        ws.protection.sheet = True      # protects dates and names from accidental edits; no password
        ws.protection.formatColumns = ws.protection.formatRows = False
    # ---- indicators sheet
    wk = wb.create_sheet(SHEET_KPIS)
    wk.sheet_view.rightToLeft = True
    heads = ["المؤشر", "الإدارة", "الدورية", "يوم الإدخال", "طريقة الإدخال", "الوحدة", f"المستهدف {year}", "التعريف"]
    for j, h in enumerate(heads, 1):
        c = wk.cell(row=1, column=j, value=h)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    for i, k in enumerate(kpis, 2):
        row = [k["name"], k["department"], FREQ_AR.get(k["frequency"], ""), ENTRY_DAY.get(k["frequency"], ""),
               "البسط والمقام (تحسب المنصة النسبة)" if ratio(k) else "القيمة", k.get("unit") or "",
               k.get("templateTarget"), k.get("definition") or ""]
        for j, v in enumerate(row, 1):
            c = wk.cell(row=i, column=j, value=v)
            c.font, c.border, c.alignment = F_BODY, BORDER, RIGHT
    for j, w in enumerate([40, 26, 11, 22, 26, 10, 12, 70], 1): wk.column_dimensions[get_column_letter(j)].width = w
    wk.freeze_panes = "B2"

    # ---- instructions sheet
    wh = wb.create_sheet(SHEET_HELP)
    wh.sheet_view.rightToLeft = True
    lines = [
        (f"قالب إدخال بيانات المؤشرات — {year}", F_TITLE),
        ("", None),
        (f"1. لكل شهر ورقة باسمه (يناير إلى ديسمبر)، فيها كل أيام الشهر من 01/01/{year} إلى 31/12/{year}: المؤشرات في الصفوف والأيام في الأعمدة.", F_BODY),
        ("2. اكتب الرقم في صف المؤشر تحت تاريخ اليوم. الخلايا البيضاء هي أيام الإدخال، والرمادية ليست يوم إدخال هذا المؤشر.", F_BODY),
        ("3. المؤشر اليومي يُدخل كل يوم، والأسبوعي يوم الأحد، والشهري أول يوم في الشهر، والربعي أول يوم في الربع.", F_BODY),
        ("4. مؤشرات النسب لها صفّان: البسط والمقام. أدخلهما، وتحسب المنصة النسبة بنفسها.", F_BODY),
        ("5. أرقام فقط بدون رموز (اكتب 92.5 وليس 92.5%). اترك الخلية فارغة إذا لم تتوفر القراءة، ولا تكتب صفراً بدلاً منها.", F_BODY),
        ("6. لا تغيّر عناوين الأعمدة ولا ترتيب التواريخ.", F_BODY),
        ("7. للرفع: من المنصة افتح «إدارة البيانات» ثم «رفع ملف» واختر هذا الملف. تظهر شاشة فحص بالسجلات الجديدة والمحدّثة والأخطاء قبل الحفظ.", F_BODY),
        ("8. يمكن رفع الملف نفسه أكثر من مرة: القراءة الموجودة تُحدَّث ولا تتكرر.", F_BODY),
        ("", None),
        (f"تفاصيل كل مؤشر (الدورية والوحدة والمستهدف) في ورقة «{SHEET_KPIS}».", F_BODY),
    ]
    for i, (t, f) in enumerate(lines, 1):
        c = wh.cell(row=i, column=1, value=t)
        if f: c.font = f
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    wh.column_dimensions["A"].width = 120

    wb.move_sheet(SHEET_HELP, offset=-13)
    wb.active = 1
    wb.properties.title = f"قالب إدخال بيانات المؤشرات {year}"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)
    print(f"{out}: 12 monthly sheets x {len(cols)} rows ({len(kpis)} indicators), {os.path.getsize(out) / 1e3:.0f} KB")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--out")
    a = ap.parse_args()
    build(a.year, a.out or os.path.join(ROOT, "templates", f"APH_KPI_Simple_{a.year}.xlsx"))
