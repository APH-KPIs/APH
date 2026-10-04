#!/usr/bin/env python3
"""Builds the simple data entry workbook: one sheet with every date of the year (one row per day) and one column per
indicator (two for ratio indicators: البسط and المقام). All visible text is Arabic. The platform's Data Management page
imports it (dmLocateSimpleSheet in index.html): columns are matched by the hidden code row, or by the Arabic header.

Usage: python3 tools/build_simple_template.py [--year 2026] [--out templates/APH_KPI_Simple_2026.xlsx]
"""
import argparse, datetime as dt, json, os, subprocess

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHEET_DATA, SHEET_KPIS, SHEET_HELP = "البيانات", "المؤشرات", "التعليمات"
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
    ws = wb.active
    ws.title = SHEET_DATA
    ws.sheet_view.rightToLeft = True
    # Row 1 (hidden): indicator codes, read by the platform so renamed headers still import.
    # Row 2: department. Row 3: indicator (Arabic). Rows 4+: one row per day of the year.
    ws.cell(row=1, column=1, value="DATE")
    ws.cell(row=3, column=1, value=DATE_HEADER)
    ws.cell(row=3, column=2, value=DAY_HEADER)
    ws.cell(row=2, column=1, value="الإدارة")
    for j, (k, part) in enumerate(cols, 3):
        ws.cell(row=1, column=j, value=k["code"] + (":" + part if part else ""))
        ws.cell(row=2, column=j, value=k["department"])
        label = k["name"] + (f" — {NUM_SUFFIX}" if part == "NUM" else f" — {DEN_SUFFIX}" if part == "DEN" else "")
        ws.cell(row=3, column=j, value=label)
    ws.row_dimensions[1].hidden = True
    ws.row_dimensions[3].height = 66
    for j in range(1, len(cols) + 3):
        c2, c3 = ws.cell(row=2, column=j), ws.cell(row=3, column=j)
        c2.font, c2.fill, c2.alignment, c2.border = F_DEP, FILL_DEP, CENTER, BORDER
        c3.font, c3.fill, c3.alignment, c3.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    for j in range(3, len(cols) + 3): ws.column_dimensions[get_column_letter(j)].width = 15

    d, r = dt.date(year, 1, 1), 4
    while d.year == year:
        a = ws.cell(row=r, column=1, value=d)
        a.number_format, a.font, a.fill, a.border, a.alignment = "dd/mm/yyyy", F_BODY, FILL_DATE, BORDER, CENTER
        b = ws.cell(row=r, column=2, value=WEEKDAYS[d.weekday()])
        b.font, b.fill, b.border, b.alignment = F_BODY, FILL_DATE, BORDER, CENTER
        for j, (k, _) in enumerate(cols, 3):
            c = ws.cell(row=r, column=j)
            c.border = BORDER
            c.protection = Protection(locked=False)
            if not entry_day(k["frequency"], d, year): c.fill = FILL_OFF
        d += dt.timedelta(days=1); r += 1
    last = r - 1

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv.error, dv.errorTitle = "اكتب رقماً فقط (صفر أو أكثر) بدون رموز أو نص.", "قيمة غير صحيحة"
    dv.prompt, dv.promptTitle = "رقم فقط. الخلايا الرمادية ليست يوم إدخال هذا المؤشر.", "إدخال"
    dv.add(f"C4:{get_column_letter(len(cols) + 2)}{last}")
    ws.add_data_validation(dv)
    ws.freeze_panes = "C4"
    ws.protection.sheet = True          # protects dates and headers from accidental edits; no password
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
        (f"1. ورقة «{SHEET_DATA}» فيها كل أيام السنة من 01/01/{year} إلى 31/12/{year}، سطر لكل يوم، وعمود لكل مؤشر.", F_BODY),
        ("2. اكتب الرقم في خلية المؤشر عند تاريخ اليوم. الخلايا البيضاء هي أيام الإدخال، والرمادية ليست يوم إدخال هذا المؤشر.", F_BODY),
        ("3. المؤشر اليومي يُدخل كل يوم، والأسبوعي يوم الأحد، والشهري أول يوم في الشهر، والربعي أول يوم في الربع.", F_BODY),
        ("4. مؤشرات النسب لها عمودان: البسط والمقام. أدخلهما، وتحسب المنصة النسبة بنفسها.", F_BODY),
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

    wb.move_sheet(SHEET_HELP, offset=-2)
    wb.active = 1
    wb.properties.title = f"قالب إدخال بيانات المؤشرات {year}"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)
    print(f"{out}: {last - 3} days x {len(cols)} columns ({len(kpis)} indicators), {os.path.getsize(out) / 1e3:.0f} KB")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--out")
    a = ap.parse_args()
    build(a.year, a.out or os.path.join(ROOT, "templates", f"APH_KPI_Simple_{a.year}.xlsx"))
