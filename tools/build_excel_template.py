# -*- coding: utf-8 -*-
"""
Builds the official Excel input template for the KPI platform (templates/APH_KPI_Data_Template_<year>.xlsx).

    python3 tools/build_excel_template.py            # 2026 template
    python3 tools/build_excel_template.py --year 2027

The indicator list comes from index.html (tools/extract_kpis.js), so the INDICATORS sheet always matches
the platform. Facility, department and category codes are defined once here and in supabase/setup.sql
(the generator rewrites the seed block in setup.sql between the MASTER DATA markers).

The workbook is only an input interface: the platform reads IMPORT_TEMPLATE (or DATA) by header name,
never by position, colour or merged cells, and re-validates every row on the server.
"""
import argparse, datetime as dt, json, os, re, shutil, subprocess, sys, tempfile, zipfile

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER_ROWS = 1000                        # rows of the master sheets that formulas and lists look at
SPARE_ROWS = 1500                         # empty DATA rows with formulas after the pre-filled calendar
CAPACITY = 10000                          # upper bound; the real count is calendar rows + SPARE_ROWS
MASTER_PASSWORD = "APH-KPI-2026"
TEST_ROWS = 0                             # --test N: small workbook (N calendar rows) for checking the formulas          # protects formulas and master sheets from accidental edits only

# ---------------------------------------------------------------- platform data (single source: index.html)
# Master data, calculation types, validation codes and KPI settings are read from index.html through
# tools/extract_kpis.js, so the workbook, the platform's import screen and setup.sql's seed never disagree.
def _platform():
    out = subprocess.run(["node", os.path.join(ROOT, "tools", "extract_kpis.js")], capture_output=True, text=True, check=True).stdout
    return json.loads(out)


PLATFORM = _platform()
TEMPLATE_VERSION = PLATFORM["templateVersion"]
DATA_START = dt.date.fromisoformat(PLATFORM["dataStartDate"])   # readings are accepted from this date onward (no end date)
YN = lambda b: "N" if b is False else "Y"
# Facility_ID, Facility_Name, Facility_Name_AR, Cluster, Region, Active, Is_Primary
FACILITIES = [(f["id"], f["name"], f["nameAr"], f.get("cluster", ""), f.get("region", ""), YN(f.get("active")), "Y" if f.get("primary") else "N")
              for f in PLATFORM["facilities"]]
# Department_ID, Department_Name, Department_Name_AR, Parent_Department_ID, Department_Type, Active
DEPARTMENTS = [(d["id"], d["name"], d["nameAr"], d.get("parentId", ""), d.get("type", ""), YN(d.get("active"))) for d in PLATFORM["departments"]]
# Category_ID, Category_Name (EN), Category_Name_AR, Parent_Category_ID
CATEGORIES = [(c["id"], c["name"], c["nameAr"], c.get("parentId", "")) for c in PLATFORM["masterCategories"]]

PERIOD_TYPES = [("DAILY", "يومي"), ("WEEKLY", "أسبوعي (يبدأ الأحد)"), ("MONTHLY", "شهري"), ("QUARTERLY", "ربع سنوي"), ("ANNUAL", "سنوي")]
FREQ_TO_PERIOD = {"daily": "DAILY", "weekly": "WEEKLY", "monthly": "MONTHLY", "quarterly": "QUARTERLY", "annual": "ANNUAL"}
DIRECTIONS = [
    ("HIGHER_IS_BETTER", "الأعلى أفضل", "higher"),
    ("LOWER_IS_BETTER", "الأقل أفضل", "lower"),
    ("TARGET_BASED", "الأقرب للمستهدف", "target"),
    ("INFORMATIONAL", "معلوماتي (بدون تقييم)", "info"),
    ("RANGE_BASED", "ضمن نطاق مثالي", "range"),
    ("CONDITIONAL", "مؤشر مشروط", "conditional"),
]
DIR_FROM_PLATFORM = {d[2]: d[0] for d in DIRECTIONS}
DIR_FROM_PLATFORM["none"] = "INFORMATIONAL"
# Calculation types: code, Arabic label, multiplier (0 = value entered directly), platform aggregation, input
CALC_TYPES = [(code, c["label"], c["mult"], c["agg"], "NUM_DEN" if c["mult"] else "VALUE") for code, c in PLATFORM["calcTypes"].items()]
UNITS = ["عدد", "%", "متوسط يومي", "ريال", "لكل 1,000", "لكل 100,000", "نسبة", "دقيقة", "ساعة", "يوم"]
STATUSES = [
    ("محقق", "القيمة حققت المستهدف"), ("تحذير", "قريبة من المستهدف (90% فأكثر من الإنجاز)"), ("غير محقق", "أقل من 90% من الإنجاز"),
    ("بلا مستهدف", "لا يوجد مستهدف للمؤشر"), ("معلوماتي", "مؤشر للعرض دون تقييم"),
    ("يُقيَّم على مستوى الفترة", "مؤشرات العدد والمجموع تُقارن بمستهدف الربع في المنصة لا بقراءة يوم واحد"),
    ("غير متاح (N/A)", "لا يمكن الحساب (مثل مقام صفر)"), ("لا توجد قيمة", "السطر لم يُعبأ بعد"),
]

# Validation codes: identical in the VALIDATION sheet and in the platform's import screen and error report.
# code, level, Arabic message, English message, suggested fix
CODES = [(code, c["level"], c["ar"], c["en"], c["fix"]) for code, c in PLATFORM["importCodes"].items()]
TARGET_BASIS = {"weekly": "WEEK", "annual": "YEAR"}


def load_kpis():
    return PLATFORM["kpis"], set(PLATFORM["derived"])


# ---------------------------------------------------------------- styles
NAVY, SKY, GREY = "0B1F33", "E0F2FE", "F1F5F9"
F_TITLE = Font(name="Arial", size=16, bold=True, color=NAVY)
F_H = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_LABEL = Font(name="Arial", size=9, bold=True, color="334155")
F_BODY = Font(name="Arial", size=10)
F_SMALL = Font(name="Arial", size=9, color="475569")
F_BOLD = Font(name="Arial", size=10, bold=True, color=NAVY)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_LABEL_IN = PatternFill("solid", fgColor="FEF3C7")
FILL_LABEL_CALC = PatternFill("solid", fgColor="E2E8F0")
FILL_IN = PatternFill("solid", fgColor="FFFBEB")
FILL_CALC = PatternFill("solid", fgColor=GREY)
FILL_SKY = PatternFill("solid", fgColor=SKY)
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
UNLOCKED = Protection(locked=False)
DATE_FMT = "dd/mm/yyyy"


def protect(ws, allow_filter=True):
    ws.protection.sheet = True
    ws.protection.password = MASTER_PASSWORD
    ws.protection.autoFilter = not allow_filter      # False = allowed
    ws.protection.sort = False
    ws.protection.formatColumns = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False


def header(ws, row, fields, label_row=None, labels=None, kinds=None):
    for i, f in enumerate(fields, 1):
        c = ws.cell(row=row, column=i, value=f)
        c.font, c.fill, c.alignment, c.border = F_H, FILL_H, CENTER, BORDER
        if label_row and labels:
            lc = ws.cell(row=label_row, column=i, value=labels[i - 1])
            lc.font, lc.alignment, lc.border = F_LABEL, CENTER, BORDER
            lc.fill = FILL_LABEL_IN if (kinds and kinds[i - 1] == "in") else FILL_LABEL_CALC


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet(wb, title, tab):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = tab
    return ws


# ---------------------------------------------------------------- DATA layout
# field, Arabic label, kind (in = input, calc = formula, key = static id), type, required
DATA_FIELDS = [
    ("Record_ID", "رقم السجل", "key", "Text", "تلقائي"),
    ("Unique_Key", "المفتاح الفريد (محسوب)", "calc", "Text", "تلقائي"),
    ("Facility_ID", "كود المنشأة *", "in", "Text (قائمة)", "إلزامي"),
    ("Facility_Name", "اسم المنشأة", "calc", "Text", "تلقائي"),
    ("Department_ID", "كود الإدارة *", "in", "Text (قائمة)", "إلزامي"),
    ("Department_Name", "اسم الإدارة", "calc", "Text", "تلقائي"),
    ("Indicator_ID", "كود المؤشر *", "in", "Text (قائمة)", "إلزامي"),
    ("Indicator_Name", "اسم المؤشر", "calc", "Text", "تلقائي"),
    ("Category_ID", "كود التصنيف", "calc", "Text", "تلقائي"),
    ("Category_Name", "التصنيف", "calc", "Text", "تلقائي"),
    ("Period_Type", "نوع الفترة *", "in", "Text (قائمة)", "إلزامي"),
    ("Measurement_Date", "تاريخ القياس *", "in", "Date", "إلزامي"),
    ("Period_Start", "بداية الفترة", "calc", "Date", "تلقائي"),
    ("Period_End", "نهاية الفترة", "calc", "Date", "تلقائي"),
    ("Numerator", "البسط", "in", "Number ≥ 0", "إلزامي لمؤشرات النسب"),
    ("Denominator", "المقام", "in", "Number > 0", "إلزامي لمؤشرات النسب"),
    ("Actual_Value", "القيمة الفعلية", "in", "Number", "إلزامي لغير النسب"),
    ("Calculated_Value", "القيمة المحسوبة", "calc", "Number", "تلقائي"),
    ("Target_Value", "المستهدف", "calc", "Number", "تلقائي (من INDICATORS)"),
    ("Unit", "الوحدة", "calc", "Text", "تلقائي"),
    ("Achievement_Percentage", "نسبة الإنجاز %", "calc", "Number", "تلقائي"),
    ("Status", "حالة المؤشر", "calc", "Text", "تلقائي"),
    ("Data_Source", "مصدر البيانات", "in", "Text (قائمة/حر)", "اختياري"),
    ("Data_Entry_By", "مدخل البيانات", "in", "Text", "اختياري"),
    ("Data_Entry_Date", "تاريخ الإدخال", "in", "Date", "اختياري"),
    ("Notes", "ملاحظات", "in", "Text ≤ 1000", "اختياري"),
    ("Last_Updated", "آخر تحديث", "in", "Date", "اختياري"),
    ("Validation_Status", "نتيجة الفحص", "calc", "Text", "تلقائي"),
]
HELPERS = ["_Ind_Row", "_Has_Value", "_Multiplier", "_Calc_Type", "_Direction", "_Fac_Row", "_Dep_Row", "_Min", "_Max", "_Green", "_Yellow"]
IMPORT_FIELDS = ["Unique_Key", "Facility_ID", "Department_ID", "Indicator_ID", "Period_Type", "Measurement_Date",
                 "Numerator", "Denominator", "Actual_Value", "Data_Source", "Data_Entry_By", "Data_Entry_Date", "Notes"]

IND_FIELDS = ["Indicator_ID", "Indicator_Name", "Arabic_Name", "English_Name", "Category_ID", "Category", "Sub_Category",
              "Department_ID", "Department", "Frequency", "Unit", "Calculation_Type", "Direction", "Target",
              "Minimum_Acceptable", "Maximum_Acceptable", "Numerator_Definition", "Denominator_Definition", "Data_Source",
              "Active", "Effective_From", "Effective_To", "Input_Mode", "Green_Threshold", "Yellow_Threshold", "Target_Basis", "Definition"]
IND_LABELS = ["كود المؤشر", "اسم المؤشر (العرض)", "الاسم العربي", "الاسم الإنجليزي", "كود التصنيف", "التصنيف", "التصنيف الفرعي",
              "كود الإدارة", "الإدارة", "الدورية", "الوحدة", "نوع الحساب", "الاتجاه", "المستهدف", "الحد الأدنى المقبول",
              "الحد الأعلى المقبول", "تعريف البسط", "تعريف المقام", "مصدر البيانات", "نشط", "ساري من", "ساري حتى",
              "طريقة الإدخال", "حد الأخضر", "حد الأصفر", "أساس المستهدف", "التعريف"]


def col_map(fields):
    return {f: get_column_letter(i) for i, f in enumerate(fields, 1)}


def build(year, out_path):
    global CAPACITY
    kpis, derived = load_kpis()
    dep_by_ar = {d[2]: d[0] for d in DEPARTMENTS}
    cat_by_ar = {c[2]: c[0] for c in CATEGORIES}
    missing = sorted({k["department"] for k in kpis if k["department"] not in dep_by_ar} | {k["category"] for k in kpis if k["category"] not in cat_by_ar})
    if missing: sys.exit("Add these to DEPARTMENTS/CATEGORIES first: " + ", ".join(missing))

    wb = Workbook()
    wb.remove(wb.active)
    ws_readme = sheet(wb, "README", NAVY)
    ws_data = sheet(wb, "DATA", "F59E0B")
    ws_ind = sheet(wb, "INDICATORS", "0EA5E9")
    ws_dep = sheet(wb, "DEPARTMENTS", "0EA5E9")
    ws_fac = sheet(wb, "FACILITIES", "0EA5E9")
    ws_lk = sheet(wb, "LOOKUPS", "64748B")
    ws_val = sheet(wb, "VALIDATION", "EF4444")
    ws_imp = sheet(wb, "IMPORT_TEMPLATE", "10B981")

    # ------------------------------------------------ LOOKUPS (title row 1, header row 2, values from row 3)
    lk_tables = [
        ("A", "Period_Type", ["Period_Type", "الوصف"], PERIOD_TYPES),
        ("D", "Frequency", ["Frequency", "الوصف"], PERIOD_TYPES),
        ("G", "Direction", ["Direction", "الوصف"], [(d[0], d[1]) for d in DIRECTIONS]),
        ("J", "Calculation_Type", ["Calculation_Type", "الوصف", "Multiplier", "Platform_Aggregation", "Value_Input"], CALC_TYPES),
        ("P", "Unit", ["Unit"], [(u,) for u in UNITS]),
        ("R", "Status", ["Status", "المعنى"], STATUSES),
        ("U", "Data_Source", ["Data_Source"], [(s,) for s in sorted({k.get("dataSource") or "" for k in kpis} - {""}) + ["إدخال يدوي", "أخرى"]]),
        ("W", "Active", ["Active", "الوصف"], [("Y", "نشط"), ("N", "غير نشط")]),
        ("Z", "Category", ["Category_ID", "Category_Name_AR", "Category_Name", "Parent_Category_ID"], [(c[0], c[2], c[1], c[3]) for c in CATEGORIES]),
        ("AE", "Department", ["Department_ID", "Department_Name_AR"], None),
        ("AH", "Facility", ["Facility_ID", "Facility_Name_AR"], None),
    ]
    lk_ranges = {}
    for start, name, heads, rows in lk_tables:
        c0 = ws_lk[start + "1"].column
        t = ws_lk.cell(row=1, column=c0, value=name)
        t.font, t.fill = F_BOLD, FILL_SKY
        for j, h in enumerate(heads):
            c = ws_lk.cell(row=2, column=c0 + j, value=h)
            c.font, c.fill, c.alignment, c.border = F_H, FILL_H, CENTER, BORDER
            ws_lk.column_dimensions[get_column_letter(c0 + j)].width = 22 if j == 0 else 30
        ws_lk.column_dimensions[get_column_letter(c0 + len(heads))].width = 3
        if rows is None:   # mirrors of the master sheets, so every list is visible in one place
            src = "DEPARTMENTS" if name == "Department" else "FACILITIES"
            for r in range(3, 203):
                for j, srccol in enumerate(["A", "C"]):
                    ws_lk.cell(row=r, column=c0 + j, value=f'=IF({src}!{srccol}{r}="","",{src}!{srccol}{r})').font = F_BODY
            continue
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = ws_lk.cell(row=3 + i, column=c0 + j, value=v)
                c.font, c.border = F_BODY, BORDER
        lk_ranges[name] = (get_column_letter(c0), get_column_letter(c0 + len(heads) - 1), 3, 2 + len(rows))
    ws_lk.freeze_panes = "A3"
    protect(ws_lk)

    def dyn(sheet_name, col):
        return f"OFFSET({sheet_name}!${col}$3,0,0,MAX(1,COUNTA({sheet_name}!${col}$1:${col}${MASTER_ROWS})-2),1)"
    names = {
        "List_Facilities": dyn("FACILITIES", "A"), "List_Departments": dyn("DEPARTMENTS", "A"),
        "List_Indicators": dyn("INDICATORS", "A"), "List_Categories": dyn("LOOKUPS", "Z"),
        "List_PeriodTypes": dyn("LOOKUPS", "A"), "List_Frequency": dyn("LOOKUPS", "D"), "List_Direction": dyn("LOOKUPS", "G"),
        "List_CalcTypes": dyn("LOOKUPS", "J"), "List_Units": dyn("LOOKUPS", "P"), "List_DataSources": dyn("LOOKUPS", "U"),
        "List_YN": dyn("LOOKUPS", "W"),
    }
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    # ------------------------------------------------ FACILITIES / DEPARTMENTS
    def master(ws, fields, labels, rows, wds, lists):
        ws.cell(row=1, column=1)
        header(ws, 2, fields, 1, labels)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = ws.cell(row=3 + i, column=j + 1, value=v)
                c.font, c.border = F_BODY, BORDER
        widths(ws, wds)
        ws.freeze_panes = "B3"
        ws.auto_filter.ref = f"A2:{get_column_letter(len(fields))}{2 + len(rows)}"
        for colname, listname in lists:
            dv = DataValidation(type="list", formula1=f"={listname}", allow_blank=True)
            dv.add(f"{col_map(fields)[colname]}3:{col_map(fields)[colname]}500")
            ws.add_data_validation(dv)
        protect(ws)

    master(ws_fac, ["Facility_ID", "Facility_Name", "Facility_Name_AR", "Cluster", "Region", "Active", "Is_Primary"],
           ["كود المنشأة", "اسم المنشأة (EN)", "اسم المنشأة", "التجمع الصحي", "المنطقة", "نشط", "المنشأة الرئيسية"],
           FACILITIES, [14, 42, 36, 20, 14, 10, 16], [("Active", "List_YN"), ("Is_Primary", "List_YN")])
    master(ws_dep, ["Department_ID", "Department_Name", "Department_Name_AR", "Parent_Department_ID", "Department_Type", "Active"],
           ["كود الإدارة", "اسم الإدارة (EN)", "اسم الإدارة", "الإدارة الأم", "النوع", "نشط"],
           DEPARTMENTS, [14, 30, 30, 20, 14, 10], [("Active", "List_YN"), ("Parent_Department_ID", "List_Departments")])

    # ------------------------------------------------ INDICATORS
    ind_rows = []
    for k in kpis:
        ct = k["calcType"]
        unit = k.get("unit") or ""
        mn, mx = k["range"]
        ind_rows.append([
            k["code"], k["name"], k["name"], k.get("englishName") or "", k.get("categoryId") or cat_by_ar[k["category"]], k["category"], k.get("group") or "",
            k.get("departmentId") or dep_by_ar[k["department"]], k["department"], FREQ_TO_PERIOD.get(k.get("frequency"), "MONTHLY"), unit, ct,
            DIR_FROM_PLATFORM.get(k.get("direction") or "none", "INFORMATIONAL"), k.get("templateTarget"), mn, mx, k.get("numeratorDefinition") or "", k.get("denominatorDefinition") or "",
            k.get("dataSource") or "", "N" if k.get("active") is False else "Y", DATA_START, None,
            "DERIVED" if k["code"] in derived else "INPUT", k.get("greenMin"), k.get("yellowMin"),
            TARGET_BASIS.get(k.get("targetBasis"), "QUARTER" if ct in ("COUNT", "SUM") else "PERIOD"), k.get("definition") or "",
        ])
    header(ws_ind, 2, IND_FIELDS, 1, IND_LABELS)
    for i, row in enumerate(ind_rows):
        for j, v in enumerate(row):
            c = ws_ind.cell(row=3 + i, column=j + 1, value=v)
            c.font, c.border = F_BODY, BORDER
            if isinstance(v, dt.date): c.number_format = DATE_FMT
    widths(ws_ind, [30, 44, 44, 34, 11, 18, 22, 12, 24, 12, 11, 18, 18, 10, 10, 10, 34, 30, 26, 8, 12, 12, 11, 10, 10, 11, 50])
    ws_ind.freeze_panes = "C3"
    ws_ind.auto_filter.ref = f"A2:{get_column_letter(len(IND_FIELDS))}{2 + len(ind_rows)}"
    ic = col_map(IND_FIELDS)
    for colname, listname in [("Frequency", "List_Frequency"), ("Unit", "List_Units"), ("Calculation_Type", "List_CalcTypes"),
                              ("Direction", "List_Direction"), ("Active", "List_YN"), ("Category_ID", "List_Categories"),
                              ("Department_ID", "List_Departments"), ("Data_Source", "List_DataSources")]:
        dv = DataValidation(type="list", formula1=f"={listname}", allow_blank=True)
        if colname == "Data_Source": dv.errorStyle = "warning"
        dv.add(f"{ic[colname]}3:{ic[colname]}1000")
        ws_ind.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"INPUT,DERIVED"', allow_blank=True); dv.add(f"{ic['Input_Mode']}3:{ic['Input_Mode']}1000"); ws_ind.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"PERIOD,QUARTER,WEEK,YEAR"', allow_blank=True); dv.add(f"{ic['Target_Basis']}3:{ic['Target_Basis']}1000"); ws_ind.add_data_validation(dv)
    for r in range(3, 3 + len(ind_rows)):
        ws_ind[f"{ic['Effective_To']}{r}"].number_format = DATE_FMT
    protect(ws_ind)

    # ------------------------------------------------ DATA
    all_fields = [f[0] for f in DATA_FIELDS] + HELPERS
    D = col_map(all_fields)
    kinds = [f[2] for f in DATA_FIELDS] + ["calc"] * len(HELPERS)
    header(ws_data, 2, all_fields, 1, [f[1] for f in DATA_FIELDS] + ["مساعد"] * len(HELPERS), kinds)
    first, last = 3, 2 + CAPACITY

    # Pre-filled calendar: one row per indicator per period of its frequency (inputs only, derived KPIs excluded).
    entry = [k for k in kpis if k.get("active") is not False and k["code"] not in derived]
    rows = []
    y0, y1 = dt.date(year, 1, 1), dt.date(year, 12, 31)
    for k in entry:
        per = FREQ_TO_PERIOD.get(k.get("frequency"), "MONTHLY")
        if per == "DAILY": dates = [y0 + dt.timedelta(days=i) for i in range((y1 - y0).days + 1)]
        elif per == "WEEKLY":
            first_sun = y0 + dt.timedelta(days=(6 - y0.weekday()) % 7)
            dates = ([y0] if first_sun != y0 else []) + [first_sun + dt.timedelta(weeks=i) for i in range(54) if first_sun + dt.timedelta(weeks=i) <= y1]
        elif per == "MONTHLY": dates = [dt.date(year, m, 1) for m in range(1, 13)]
        elif per == "QUARTERLY": dates = [dt.date(year, m, 1) for m in (1, 4, 7, 10)]
        else: dates = [y0]
        for d in dates: rows.append((d, kpis.index(k), k, per))
    rows.sort(key=lambda x: (x[0], x[1]))
    if TEST_ROWS: rows = rows[:TEST_ROWS]
    else: CAPACITY = len(rows) + SPARE_ROWS
    last = 2 + CAPACITY

    ind_rng = lambda c: f"INDICATORS!${ic[c]}$1:${ic[c]}${MASTER_ROWS}"
    for i in range(CAPACITY):
        r = first + i
        if i < len(rows):
            d, _, k, per = rows[i]
            ws_data[f"{D['Facility_ID']}{r}"] = FACILITIES[0][0]
            ws_data[f"{D['Department_ID']}{r}"] = dep_by_ar[k["department"]]
            ws_data[f"{D['Indicator_ID']}{r}"] = k["code"]
            ws_data[f"{D['Period_Type']}{r}"] = per
            ws_data[f"{D['Measurement_Date']}{r}"] = d
            ws_data[f"{D['Data_Source']}{r}"] = k.get("dataSource") or None
        ws_data[f"{D['Record_ID']}{r}"] = f"R{year % 100:02d}{i + 1:05d}"
        C = lambda f: f"{D[f]}{r}"
        IR = C("_Ind_Row")
        f = {
            "_Ind_Row": f'=IF({C("Indicator_ID")}="","",IFERROR(MATCH({C("Indicator_ID")},INDICATORS!$A$1:$A${MASTER_ROWS},0),""))',
            "_Has_Value": f'=COUNTA({C("Numerator")}:{C("Actual_Value")})>0',
            "_Multiplier": f'=IF({C("_Calc_Type")}="",0,IFERROR(INDEX(LOOKUPS!$L$1:$L${MASTER_ROWS},MATCH({C("_Calc_Type")},LOOKUPS!$J$1:$J${MASTER_ROWS},0)),0))',
            "_Calc_Type": f'=IF({IR}="","",INDEX({ind_rng("Calculation_Type")},{IR})&"")',
            "_Direction": f'=IF({IR}="","",INDEX({ind_rng("Direction")},{IR})&"")',
            "_Fac_Row": f'=IF({C("Facility_ID")}="","",IFERROR(MATCH({C("Facility_ID")},FACILITIES!$A$1:$A${MASTER_ROWS},0),""))',
            "_Dep_Row": f'=IF({C("Department_ID")}="","",IFERROR(MATCH({C("Department_ID")},DEPARTMENTS!$A$1:$A${MASTER_ROWS},0),""))',
            "_Min": f'=IF({IR}="","",IF(INDEX({ind_rng("Minimum_Acceptable")},{IR})="","",INDEX({ind_rng("Minimum_Acceptable")},{IR})))',
            "_Max": f'=IF({IR}="","",IF(INDEX({ind_rng("Maximum_Acceptable")},{IR})="","",INDEX({ind_rng("Maximum_Acceptable")},{IR})))',
            "_Green": f'=IF({IR}="","",IF(INDEX({ind_rng("Green_Threshold")},{IR})="","",INDEX({ind_rng("Green_Threshold")},{IR})))',
            "_Yellow": f'=IF({IR}="","",IF(INDEX({ind_rng("Yellow_Threshold")},{IR})="","",INDEX({ind_rng("Yellow_Threshold")},{IR})))',
            "Unique_Key": (f'=IF(OR({C("Facility_ID")}="",{C("Department_ID")}="",{C("Indicator_ID")}="",{C("Period_Type")}="",NOT(ISNUMBER({C("Period_Start")}))),"",'
                           f'UPPER({C("Facility_ID")}&"-"&{C("Indicator_ID")}&"-"&YEAR({C("Period_Start")})&RIGHT("0"&MONTH({C("Period_Start")}),2)&RIGHT("0"&DAY({C("Period_Start")}),2)'
                           f'&"-"&{C("Department_ID")}&"-"&{C("Period_Type")}))'),
            "Facility_Name": f'=IF({C("_Fac_Row")}="","",INDEX(FACILITIES!$C$1:$C${MASTER_ROWS},{C("_Fac_Row")})&"")',
            "Department_Name": f'=IF({C("_Dep_Row")}="","",INDEX(DEPARTMENTS!$C$1:$C${MASTER_ROWS},{C("_Dep_Row")})&"")',
            "Indicator_Name": f'=IF({IR}="","",INDEX({ind_rng("Arabic_Name")},{IR})&"")',
            "Category_ID": f'=IF({IR}="","",INDEX({ind_rng("Category_ID")},{IR})&"")',
            "Category_Name": f'=IF({IR}="","",INDEX({ind_rng("Category")},{IR})&"")',
            "Period_Start": (f'=IF(NOT(ISNUMBER({C("Measurement_Date")})),"",IF({C("Period_Type")}="WEEKLY",{C("Measurement_Date")}-WEEKDAY({C("Measurement_Date")},1)+1,'
                             f'IF({C("Period_Type")}="MONTHLY",DATE(YEAR({C("Measurement_Date")}),MONTH({C("Measurement_Date")}),1),'
                             f'IF({C("Period_Type")}="QUARTERLY",DATE(YEAR({C("Measurement_Date")}),INT((MONTH({C("Measurement_Date")})-1)/3)*3+1,1),'
                             f'IF({C("Period_Type")}="ANNUAL",DATE(YEAR({C("Measurement_Date")}),1,1),{C("Measurement_Date")})))))'),
            "Period_End": (f'=IF(NOT(ISNUMBER({C("Period_Start")})),"",IF({C("Period_Type")}="WEEKLY",{C("Period_Start")}+6,'
                           f'IF({C("Period_Type")}="MONTHLY",DATE(YEAR({C("Period_Start")}),MONTH({C("Period_Start")})+1,0),'
                           f'IF({C("Period_Type")}="QUARTERLY",DATE(YEAR({C("Period_Start")}),MONTH({C("Period_Start")})+3,0),'
                           f'IF({C("Period_Type")}="ANNUAL",DATE(YEAR({C("Period_Start")}),12,31),{C("Period_Start")})))))'),
            "Calculated_Value": (f'=IF(OR({IR}="",NOT({C("_Has_Value")})),"",IF({C("_Multiplier")}>0,'
                                 f'IF(AND(ISNUMBER({C("Numerator")}),ISNUMBER({C("Denominator")})),IF({C("Denominator")}>0,{C("Numerator")}/{C("Denominator")}*{C("_Multiplier")},""),""),'
                                 f'IF(ISNUMBER({C("Actual_Value")}),{C("Actual_Value")},"")))'),
            "Target_Value": f'=IF({IR}="","",IF(INDEX({ind_rng("Target")},{IR})="","",INDEX({ind_rng("Target")},{IR})))',
            "Unit": f'=IF({IR}="","",INDEX({ind_rng("Unit")},{IR})&"")',
            "Achievement_Percentage": (
                f'=IF(OR(NOT(ISNUMBER({C("Calculated_Value")})),NOT(ISNUMBER({C("Target_Value")})),{C("_Calc_Type")}="COUNT",{C("_Calc_Type")}="SUM"),"",'
                f'IF({C("_Direction")}="LOWER_IS_BETTER",IF({C("Target_Value")}=0,IF({C("Calculated_Value")}<=0,100,0),IF({C("Calculated_Value")}<={C("Target_Value")},100,{C("Target_Value")}/{C("Calculated_Value")}*100)),'
                f'IF({C("_Direction")}="HIGHER_IS_BETTER",IF({C("Target_Value")}=0,"",{C("Calculated_Value")}/{C("Target_Value")}*100),'
                f'IF({C("_Direction")}="TARGET_BASED",IF({C("Target_Value")}=0,"",MAX(0,100-ABS({C("Calculated_Value")}-{C("Target_Value")})/ABS({C("Target_Value")})*100)),""))))'),
            "Status": (
                f'=IF({IR}="","",IF(NOT({C("_Has_Value")}),"لا توجد قيمة",IF(NOT(ISNUMBER({C("Calculated_Value")})),"غير متاح (N/A)",'
                f'IF(OR({C("_Direction")}="INFORMATIONAL",{C("_Direction")}="CONDITIONAL",{C("_Direction")}="RANGE_BASED"),"معلوماتي",'
                f'IF(ISNUMBER({C("_Green")}),IF(IF({C("_Direction")}="LOWER_IS_BETTER",{C("Calculated_Value")}<={C("_Green")},{C("Calculated_Value")}>={C("_Green")}),"محقق",'
                f'IF(AND(ISNUMBER({C("_Yellow")}),IF({C("_Direction")}="LOWER_IS_BETTER",{C("Calculated_Value")}<={C("_Yellow")},{C("Calculated_Value")}>={C("_Yellow")})),"تحذير","غير محقق")),'
                f'IF(OR({C("_Calc_Type")}="COUNT",{C("_Calc_Type")}="SUM"),"يُقيَّم على مستوى الفترة",IF(NOT(ISNUMBER({C("Target_Value")})),"بلا مستهدف",'
                f'IF(NOT(ISNUMBER({C("Achievement_Percentage")})),"غير متاح (N/A)",IF({C("Achievement_Percentage")}>=IF({C("_Direction")}="TARGET_BASED",98,100),"محقق",'
                f'IF({C("Achievement_Percentage")}>=90,"تحذير","غير محقق"))))))))))'),
            "Validation_Status": f'=IF(VALIDATION!D{r}="","",VALIDATION!D{r})',
        }
        for name, formula in f.items():
            ws_data[f"{D[name]}{r}"] = formula
        for name, kind, *_ in DATA_FIELDS:
            c = ws_data[f"{D[name]}{r}"]
            c.border = BORDER
            c.font = F_BODY
            if kind == "in":
                c.fill, c.protection = FILL_IN, UNLOCKED
            else:
                c.fill = FILL_CALC
        for name in ("Measurement_Date", "Period_Start", "Period_End", "Data_Entry_Date", "Last_Updated"):
            ws_data[f"{D[name]}{r}"].number_format = DATE_FMT
        for name in ("Numerator", "Denominator", "Actual_Value", "Calculated_Value", "Target_Value"):
            ws_data[f"{D[name]}{r}"].number_format = "#,##0.####"
        ws_data[f"{D['Achievement_Percentage']}{r}"].number_format = "0.0"

    widths(ws_data, [10, 44, 11, 26, 12, 22, 30, 40, 10, 18, 12, 13, 12, 12, 11, 11, 12, 13, 11, 10, 11, 20, 22, 16, 13, 30, 13, 13] + [8] * len(HELPERS))
    for h in HELPERS: ws_data.column_dimensions[D[h]].hidden = True
    for h in ("Record_ID",): ws_data.column_dimensions[D[h]].width = 10
    ws_data.freeze_panes = f"{D['Indicator_Name']}3"
    ws_data.auto_filter.ref = f"A2:{D['Validation_Status']}{last}"
    ws_data.row_dimensions[1].height = 30
    ws_data.row_dimensions[2].height = 20

    def dv_list(ws, ref, name, warn=False, title=None, msg=None):
        dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True, showErrorMessage=True)
        if warn: dv.errorStyle = "warning"
        dv.errorTitle, dv.error = "قيمة غير موجودة في القائمة", "اختر قيمة من القائمة المنسدلة."
        if title: dv.promptTitle, dv.prompt, dv.showInputMessage = title, msg, True
        dv.add(ref); ws.add_data_validation(dv)
    rng = lambda name: f"{D[name]}{first}:{D[name]}{last}"
    dv_list(ws_data, rng("Facility_ID"), "List_Facilities", title="المنشأة", msg="اختر كود المنشأة من FACILITIES")
    dv_list(ws_data, rng("Department_ID"), "List_Departments", title="الإدارة", msg="اختر كود الإدارة من DEPARTMENTS")
    dv_list(ws_data, rng("Indicator_ID"), "List_Indicators", title="المؤشر", msg="اختر كود المؤشر من INDICATORS")
    dv_list(ws_data, rng("Period_Type"), "List_PeriodTypes", title="نوع الفترة", msg="يجب أن يطابق دورية المؤشر")
    dv_list(ws_data, rng("Data_Source"), "List_DataSources", warn=True)
    dv = DataValidation(type="date", operator="between", formula1=f"DATE({DATA_START.year},{DATA_START.month},{DATA_START.day})", formula2="TODAY()",
                        allow_blank=True, showErrorMessage=True, showInputMessage=True)
    dv.errorTitle, dv.error = "تاريخ غير مقبول", f"التاريخ يجب أن يكون من {DATA_START:%d/%m/%Y} حتى تاريخ اليوم."
    dv.promptTitle, dv.prompt = "تاريخ القياس", "DD/MM/YYYY — من 01/01/2026 حتى اليوم"
    dv.add(rng("Measurement_Date")); ws_data.add_data_validation(dv)
    for name in ("Data_Entry_Date", "Last_Updated"):
        dv = DataValidation(type="date", operator="greaterThanOrEqual", formula1="DATE(2026,1,1)", allow_blank=True, showErrorMessage=True)
        dv.errorTitle, dv.error = "تاريخ غير صالح", "أدخل تاريخاً صحيحاً."
        dv.add(rng(name)); ws_data.add_data_validation(dv)
    for name, rule, msg in [("Numerator", "AND(ISNUMBER({c}),{c}>=0)", "البسط رقم أكبر من أو يساوي صفر."),
                            ("Denominator", "AND(ISNUMBER({c}),{c}>0)", "المقام رقم أكبر من صفر."),
                            ("Actual_Value", "ISNUMBER({c})", "القيمة الفعلية رقم فقط (بدون % أو مسافات).")]:
        c = f"{D[name]}{first}"
        dv = DataValidation(type="custom", formula1=rule.format(c=c), allow_blank=True, showErrorMessage=True)
        dv.errorTitle, dv.error = "قيمة غير صالحة", msg
        dv.add(rng(name)); ws_data.add_data_validation(dv)
    dv = DataValidation(type="textLength", operator="lessThanOrEqual", formula1="1000", allow_blank=True, showErrorMessage=True)
    dv.errorTitle, dv.error = "النص طويل", "الملاحظات بحد أقصى 1000 حرف."
    dv.add(rng("Notes")); ws_data.add_data_validation(dv)

    vs = rng("Validation_Status")
    ws_data.conditional_formatting.add(vs, CellIsRule(operator="equal", formula=['"VALID"'], fill=PatternFill("solid", fgColor="D1FAE5"), font=Font(color="065F46", bold=True)))
    ws_data.conditional_formatting.add(vs, CellIsRule(operator="equal", formula=['"INVALID"'], fill=PatternFill("solid", fgColor="FEE2E2"), font=Font(color="991B1B", bold=True)))
    ws_data.conditional_formatting.add(vs, CellIsRule(operator="equal", formula=['"WARNING"'], fill=PatternFill("solid", fgColor="FEF3C7"), font=Font(color="92400E", bold=True)))
    ws_data.conditional_formatting.add(vs, CellIsRule(operator="equal", formula=['"EMPTY"'], font=Font(color="94A3B8")))
    st = rng("Status")
    ws_data.conditional_formatting.add(st, CellIsRule(operator="equal", formula=['"محقق"'], font=Font(color="047857", bold=True)))
    ws_data.conditional_formatting.add(st, CellIsRule(operator="equal", formula=['"تحذير"'], font=Font(color="B45309", bold=True)))
    ws_data.conditional_formatting.add(st, CellIsRule(operator="equal", formula=['"غير محقق"'], font=Font(color="B91C1C", bold=True)))
    # The whole input row turns red when the record is INVALID.
    inv = FormulaRule(formula=[f'${D["Validation_Status"]}{first}="INVALID"'], fill=PatternFill("solid", fgColor="FECACA"))
    for name in ("Facility_ID", "Department_ID", "Indicator_ID", "Period_Type", "Measurement_Date", "Numerator", "Denominator", "Actual_Value"):
        ws_data.conditional_formatting.add(rng(name), inv)
    for name, _, kind, typ, req in DATA_FIELDS:
        ws_data[f"{D[name]}2"].comment = Comment(f"{name}\n{typ} — {req}", "APH KPI")
    protect(ws_data)

    # ------------------------------------------------ VALIDATION (row r checks DATA row r)
    VF = ["Row_No", "Unique_Key", "Issue_Codes", "Validation_Status", "Message", "Suggested_Fix"]
    header(ws_val, 2, VF, 1, ["رقم السطر في DATA", "المفتاح الفريد", "رموز الملاحظات", "النتيجة", "أول ملاحظة", "الإجراء المقترح"])
    for i in range(CAPACITY):
        r = first + i
        X = lambda f: f"DATA!{D[f]}{r}"
        chk = [
            f'IF(OR({X("Facility_ID")}="",{X("Department_ID")}="",{X("Indicator_ID")}="",{X("Period_Type")}="",{X("Measurement_Date")}=""),"E11 ","")',
            f'IF(AND({X("Measurement_Date")}<>"",NOT(ISNUMBER({X("Measurement_Date")}))),"E02 ","")',
            f'IF(ISNUMBER({X("Measurement_Date")}),IF({X("Measurement_Date")}<DATE({DATA_START.year},{DATA_START.month},{DATA_START.day}),"E03 ",""),"")',
            f'IF(AND({X("Indicator_ID")}<>"",{X("_Ind_Row")}=""),"E04 ","")',
            f'IF({X("Department_ID")}="","",IF({X("_Dep_Row")}="","E05 ",IF(INDEX(DEPARTMENTS!$F$1:$F${MASTER_ROWS},{X("_Dep_Row")})<>"Y","E05 ","")))',
            f'IF({X("Facility_ID")}="","",IF({X("_Fac_Row")}="","E06 ",IF(INDEX(FACILITIES!$F$1:$F${MASTER_ROWS},{X("_Fac_Row")})<>"Y","E06 ","")))',
            f'IF(OR(AND({X("Numerator")}<>"",NOT(ISNUMBER({X("Numerator")}))),AND({X("Denominator")}<>"",NOT(ISNUMBER({X("Denominator")}))),AND({X("Actual_Value")}<>"",NOT(ISNUMBER({X("Actual_Value")})))),"E07 ","")',
            f'IF(AND({X("_Multiplier")}>0,ISNUMBER({X("Denominator")})),IF({X("Denominator")}=0,"E08 ",""),"")',
            f'IF({X("_Ind_Row")}="","",IF({X("_Multiplier")}>0,IF(OR(NOT(ISNUMBER({X("Numerator")})),NOT(ISNUMBER({X("Denominator")}))),"E01 ",""),IF(NOT(ISNUMBER({X("Actual_Value")})),"E01 ","")))',
            f'IF(OR(AND(ISNUMBER({X("Numerator")}),{X("Numerator")}<0),AND(ISNUMBER({X("Denominator")}),{X("Denominator")}<0),'
            f'AND(ISNUMBER({X("Calculated_Value")}),ISNUMBER({X("_Min")}),{X("Calculated_Value")}<{X("_Min")}),AND(ISNUMBER({X("Calculated_Value")}),ISNUMBER({X("_Max")}),{X("Calculated_Value")}>{X("_Max")})),"E09 ","")',
            f'IF(AND({X("Period_Type")}<>"",{X("_Ind_Row")}<>""),IF({X("Period_Type")}<>INDEX({ind_rng("Frequency")},{X("_Ind_Row")}),"E12 ",""),"")',
            f'IF({X("_Ind_Row")}="","",IF(INDEX({ind_rng("Active")},{X("_Ind_Row")})<>"Y","E13 ",""))',
            f'IF({X("_Ind_Row")}="","",IF(INDEX({ind_rng("Input_Mode")},{X("_Ind_Row")})="DERIVED","E14 ",""))',
            f'IF(AND({X("_Ind_Row")}<>"",NOT(ISNUMBER({X("Target_Value")}))),"W01 ","")',
            f'IF(AND(ISNUMBER({X("Measurement_Date")}),ISNUMBER({X("Period_Start")})),IF(AND({X("Measurement_Date")}<>{X("Period_Start")},{X("Period_Start")}>=DATE({DATA_START.year},{DATA_START.month},{DATA_START.day})),"W02 ",""),"")',
            f'IF(AND({X("_Multiplier")}>0,ISNUMBER({X("Numerator")}),ISNUMBER({X("Denominator")})),IF(AND({X("Denominator")}>0,{X("Numerator")}>{X("Denominator")}),"W03 ",""),"")',
        ]
        ws_val[f"A{r}"] = r
        ws_val[f"B{r}"] = f'=IF({X("_Has_Value")},{X("Unique_Key")},"")'
        ws_val[f"C{r}"] = f'=IF(NOT({X("_Has_Value")}),"",TRIM(' + "&".join(chk) + "))"
        ws_val[f"D{r}"] = f'=IF(NOT({X("_Has_Value")}),IF({X("Indicator_ID")}="","","EMPTY"),IF(ISNUMBER(SEARCH("E",C{r})),"INVALID",IF(C{r}<>"","WARNING","VALID")))'
        ws_val[f"E{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(LEFT(C{r},3),$J$4:$N$40,3,0),""))'
        ws_val[f"F{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(LEFT(C{r},3),$J$4:$N$40,5,0),""))'
        for col in "ABCDEF":
            ws_val[f"{col}{r}"].font = F_BODY
    widths(ws_val, [10, 44, 18, 13, 60, 50, 3, 3, 3, 8, 10, 52, 38, 52])
    ws_val.freeze_panes = "A3"
    ws_val.auto_filter.ref = f"A2:F{last}"
    vrng = f"D{first}:D{last}"
    ws_val.conditional_formatting.add(vrng, CellIsRule(operator="equal", formula=['"VALID"'], fill=PatternFill("solid", fgColor="D1FAE5"), font=Font(color="065F46", bold=True)))
    ws_val.conditional_formatting.add(vrng, CellIsRule(operator="equal", formula=['"INVALID"'], fill=PatternFill("solid", fgColor="FEE2E2"), font=Font(color="991B1B", bold=True)))
    ws_val.conditional_formatting.add(vrng, CellIsRule(operator="equal", formula=['"WARNING"'], fill=PatternFill("solid", fgColor="FEF3C7"), font=Font(color="92400E", bold=True)))
    ws_val.conditional_formatting.add(vrng, CellIsRule(operator="equal", formula=['"EMPTY"'], font=Font(color="94A3B8")))
    # Summary and code legend beside the per-row table.
    ws_val["J1"] = "ملخص الفحص"; ws_val["J1"].font = F_TITLE
    summary = [("سجلات معبأة", f'=COUNTIF(D{first}:D{last},"VALID")+COUNTIF(D{first}:D{last},"WARNING")+COUNTIF(D{first}:D{last},"INVALID")'),
               ("VALID", f'=COUNTIF(D{first}:D{last},"VALID")'), ("WARNING", f'=COUNTIF(D{first}:D{last},"WARNING")'),
               ("INVALID", f'=COUNTIF(D{first}:D{last},"INVALID")')]
    ws_val["P2"], ws_val["Q2"] = "البند", "العدد"
    for c in ("P2", "Q2"): ws_val[c].font, ws_val[c].fill, ws_val[c].alignment = F_H, FILL_H, CENTER
    for i, (lab, fml) in enumerate(summary):
        ws_val[f"P{3 + i}"], ws_val[f"Q{3 + i}"] = lab, fml
        ws_val[f"P{3 + i}"].font, ws_val[f"Q{3 + i}"].font = F_BOLD, F_BOLD
        ws_val[f"P{3 + i}"].border = ws_val[f"Q{3 + i}"].border = BORDER
    ws_val.column_dimensions["P"].width, ws_val.column_dimensions["Q"].width = 16, 10
    for j, h in enumerate(["Code", "Level", "الرسالة", "Message", "الإجراء المقترح", "عدد السجلات"]):
        c = ws_val.cell(row=3, column=10 + j, value=h)
        c.font, c.fill, c.alignment, c.border = F_H, FILL_H, CENTER, BORDER
    for i, (code, lvl, ar, en, fix) in enumerate(CODES):
        r = 4 + i
        for j, v in enumerate([code, lvl, ar, en, fix, f'=COUNTIF($C${first}:$C${last},"*{code}*")']):
            c = ws_val.cell(row=r, column=10 + j, value=v)
            c.font, c.border, c.alignment = F_BODY, BORDER, RIGHT if j in (2, 4) else LEFT
            if j == 1: c.font = Font(name="Arial", size=10, bold=True, color="B91C1C" if lvl == "ERROR" else "B45309")
    ws_val.column_dimensions["O"].width = 12
    protect(ws_val)

    # ------------------------------------------------ IMPORT_TEMPLATE (fixed column order, no formulas)
    # For records pasted from another system. Left empty, the platform reads the DATA sheet (same column names).
    header(ws_imp, 1, IMPORT_FIELDS)
    for i in range(500):
        for j, name in enumerate(IMPORT_FIELDS, 1):
            c = ws_imp.cell(row=2 + i, column=j)
            if name in ("Measurement_Date", "Data_Entry_Date"): c.number_format = DATE_FMT
    widths(ws_imp, [44, 12, 14, 30, 12, 14, 11, 11, 12, 24, 16, 14, 30])
    ws_imp.freeze_panes = "A2"

    # ------------------------------------------------ README
    build_readme(ws_readme, year, len(rows), len(entry), kpis, derived, D)

    wb.calculation.fullCalcOnLoad = True
    wb.properties.title = "APH KPI Data Template"
    wb.properties.subject = f"TEMPLATE_VERSION={TEMPLATE_VERSION}"
    wb.properties.creator = "APH KPI Platform"
    wb.active = 0
    wb.save(out_path)
    share_formulas(out_path)
    return len(rows)


# ---------------------------------------------------------------- shared formulas
# openpyxl writes every formula in full. Excel stores a column of identical formulas once ("shared formula"),
# which keeps the file small enough to open quickly on any device. Runs of cells in one column whose formulas
# are the same apart from relative row numbers become one shared formula; the calculation is unchanged.
_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)><f>([^<]*)</f>(?:<v\s*/>|<v>[^<]*</v>)?</c>')
_REF = re.compile(r"(?<![A-Za-z_\d.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def _shape(formula, row):
    """The formula with relative row numbers written as offsets from its own row (text in quotes untouched)."""
    parts = formula.split('"')
    for i in range(0, len(parts), 2):
        parts[i] = _REF.sub(lambda m: m.group(0) if m.group(3) else f"{m.group(1)}{m.group(2)}[{int(m.group(4)) - row}]", parts[i])
    return '"'.join(parts)


def _share_sheet(xml):
    cells = [(m.start(), m.end(), m.group(1), int(m.group(2)), m.group(3), m.group(4)) for m in _CELL.finditer(xml)]
    by_col = {}
    for c in cells: by_col.setdefault(c[2], []).append(c)
    replace, si = {}, 0
    for col, lst in by_col.items():
        lst.sort(key=lambda c: c[3])
        i = 0
        while i < len(lst):
            shape, j = _shape(lst[i][5], lst[i][3]), i + 1
            while j < len(lst) and lst[j][3] == lst[j - 1][3] + 1 and _shape(lst[j][5], lst[j][3]) == shape: j += 1
            if j - i >= 2:
                m = lst[i]
                replace[m[0]] = (m[1], f'<c r="{col}{m[3]}"{m[4]}><f t="shared" ref="{col}{m[3]}:{col}{lst[j - 1][3]}" si="{si}">{m[5]}</f></c>')
                for c in lst[i + 1:j]: replace[c[0]] = (c[1], f'<c r="{col}{c[3]}"{c[4]}><f t="shared" si="{si}"/></c>')
                si += 1
            i = j
    out, pos = [], 0
    for start in sorted(replace):
        end, text = replace[start]
        out.append(xml[pos:start]); out.append(text); pos = end
    out.append(xml[pos:])
    return "".join(out)


def share_formulas(path):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=os.path.dirname(path) or ".").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and b"<f>" in data:
                data = _share_sheet(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def build_readme(ws, year, n_rows, n_kpis, kpis, derived, D):
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 60
    r = 1

    def title(text):
        nonlocal r
        r += 1
        c = ws.cell(row=r, column=1, value=text)
        c.font, c.fill = Font(name="Arial", size=12, bold=True, color="FFFFFF"), FILL_H
        for col in range(2, 6): ws.cell(row=r, column=col).fill = FILL_H
        r += 1

    def line(label, text, bold=False):
        nonlocal r
        a = ws.cell(row=r, column=1, value=label); a.font, a.alignment = F_BOLD, RIGHT
        b = ws.cell(row=r, column=2, value=text); b.font, b.alignment = (F_BOLD if bold else F_BODY), RIGHT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.row_dimensions[r].height = max(18, 15 * (1 + len(str(text)) // 110))
        r += 1

    def table(heads, rows, fills=None):
        nonlocal r
        for j, h in enumerate(heads):
            c = ws.cell(row=r, column=1 + j, value=h); c.font, c.fill, c.alignment, c.border = F_H, FILL_H, CENTER, BORDER
        r += 1
        for row in rows:
            for j, v in enumerate(row):
                c = ws.cell(row=r, column=1 + j, value=v); c.font, c.border, c.alignment = F_BODY, BORDER, RIGHT
            r += 1

    c = ws.cell(row=r, column=1, value="قالب إدخال بيانات مؤشرات الأداء — القالب الرسمي للمنصة"); c.font = F_TITLE
    r += 1
    ws.cell(row=r, column=1, value="مستشفى إرادة والصحة النفسية بأبها · تجمع عسير الصحي").font = F_SMALL
    r += 1
    title("معلومات الملف")
    line("اسم الملف", f"APH_KPI_Data_Template_{year}.xlsx")
    line("Template_Version", TEMPLATE_VERSION, bold=True)
    line("الغرض", "قالب الإدخال الرسمي ومصدر البيانات الخام لمنصة مؤشرات الأداء. تُعبأ القراءات هنا ثم يُرفع الملف للمنصة، فتتحقق منه وتحفظه في قاعدة البيانات وتتحدث كل البطاقات والرسوم والتقارير تلقائياً.")
    line("تاريخ بداية البيانات", f"{DATA_START:%d/%m/%Y} — لا يوجد تاريخ نهاية؛ تقبل المنصة أي تاريخ حتى اليوم.")
    line("محتوى الملف", f"تقويم {year} معبأ مسبقاً: سطر لكل مؤشر في كل فترة حسب دوريته ({n_rows:,} سطراً لـ {n_kpis} مؤشراً)، مع {CAPACITY - n_rows:,} سطراً إضافياً فارغاً لسجلات أخرى.")
    line("الهندسة", "Excel → فحص → ربط الحقول (Mapping) → قاعدة البيانات → محرك المؤشرات → اللوحة. الملف واجهة إدخال فقط وليس قاعدة البيانات.")

    title("أوراق الملف")
    table(["الورقة", "الغرض", "", "", "من يعدّلها"], [
        ("README", "هذه التعليمات", "", "", "—"),
        ("DATA", "الورقة الأساسية للتعبئة: سطر واحد لكل قراءة (Long Format)", "", "", "مدخل البيانات (الخلايا الصفراء فقط)"),
        ("INDICATORS", "جدول المؤشرات الرئيسي (Master)", "", "", "مدير النظام"),
        ("DEPARTMENTS", "الإدارات والأقسام", "", "", "مدير النظام"),
        ("FACILITIES", "المنشآت", "", "", "مدير النظام"),
        ("LOOKUPS", "القوائم المرجعية للقوائم المنسدلة", "", "", "مدير النظام"),
        ("VALIDATION", "فحص كل سجل قبل الرفع: VALID / WARNING / INVALID", "", "", "تلقائي"),
        ("IMPORT_TEMPLATE", "أعمدة الرفع بترتيب ثابت للصق بيانات من نظام آخر؛ اتركها فارغة إذا عبأت DATA", "", "", "اختياري"),
    ])

    title("طريقة التعبئة")
    for i, s in enumerate([
        "افتح ورقة DATA. كل سطر = قراءة واحدة لمؤشر واحد في فترة واحدة. الأسطر معبأة مسبقاً بالمنشأة والإدارة والمؤشر والتاريخ.",
        "اكتب الأرقام في الخلايا الصفراء فقط: البسط والمقام لمؤشرات النسب (Calculation_Type = PERCENTAGE وما شابهها)، أو القيمة الفعلية لغيرها.",
        "لا تكتب النسبة يدوياً لمؤشرات البسط/المقام: تُحسب في عمود «القيمة المحسوبة» وتعيد المنصة حسابها.",
        "اترك السطر فارغاً إذا لم تتوفر قراءته؛ الأسطر الفارغة لا تُرفع ولا تُعد أخطاء.",
        "لإضافة قراءة غير موجودة في التقويم (منشأة أو إدارة أخرى مثلاً) استخدم الأسطر الإضافية في آخر ورقة DATA واختر القيم من القوائم.",
        "راجع عمود «نتيجة الفحص» وورقة VALIDATION، وأصلح كل سطر INVALID قبل الرفع.",
        "احفظ الملف بصيغة xlsx ثم ارفعه من المنصة.",
    ], 1):
        line(f"الخطوة {i}", s)

    title("الحقول (ورقة DATA)")
    table(["الحقل", "الوصف", "النوع", "إلزامي؟", "ملاحظات"], [
        (f, ar.replace(" *", ""), typ, req, {"in": "خلية إدخال (صفراء)", "calc": "معادلة محمية", "key": "رقم ثابت"}[kind]) for f, ar, kind, typ, req in DATA_FIELDS
    ])
    line("الحقول الإلزامية", "Facility_ID، Department_ID، Indicator_ID، Period_Type، Measurement_Date، ثم البسط والمقام (للنسب) أو Actual_Value (لغيرها).", bold=True)
    line("الحقول الاختيارية", "Data_Source، Data_Entry_By، Data_Entry_Date، Notes، Last_Updated. أما المستهدف فيُدار من صفحة المستهدفات في المنصة (معتمد وبمبرر)، وعموده هنا للعرض.")

    title("التاريخ والفترات")
    line("صيغة التاريخ", "تاريخ Excel حقيقي بعرض DD/MM/YYYY (مثال 01/01/2026). تقبل المنصة أيضاً النص YYYY-MM-DD أو DD/MM/YYYY.")
    line("نوع الفترة", "يجب أن يطابق دورية المؤشر (Frequency في INDICATORS): DAILY، WEEKLY، MONTHLY، QUARTERLY، ANNUAL.")
    line("الأسبوع", "يبدأ الأسبوع يوم الأحد. أي تاريخ داخل الأسبوع يُسجل على يوم الأحد، وأي تاريخ داخل الشهر على أول الشهر، وداخل الربع على أول الربع (تحذير W02).")
    line("التجميع", "تُدخل القراءة بدوريتها فقط، وتجمعها المنصة تلقائياً إلى أسبوعي وشهري وربعي وسنوي حتى تاريخه (YTD) حسب نوع الحساب.")

    title("المفتاح الفريد ومنع التكرار")
    line("الصيغة", "Facility_ID-Indicator_ID-YYYYMMDD-Department_ID-Period_Type (التاريخ هو بداية الفترة)")
    line("مثال", "FAC001-ED_VISITS-20260101-DEP005-DAILY")
    line("عند إعادة الرفع", "السجل ذو المفتاح الموجود يُحدَّث (UPDATE) ولا يُنشأ سجل جديد؛ الجديد يُضاف (INSERT)؛ والمطابق تماماً لا يتغير. تكرار المفتاح داخل الملف نفسه: تنبه المنصة عند الرفع ويُعتمد آخر سطر (W05).")

    title("الوحدات وأنواع الحساب")
    line("الوحدات المقبولة", "، ".join(UNITS))
    table(["Calculation_Type", "المعنى", "المضاعف", "التجميع في المنصة", "الإدخال"], [(c[0], c[1], c[2] or "—", c[3], "البسط والمقام" if c[4] == "NUM_DEN" else "القيمة الفعلية") for c in CALC_TYPES])
    table(["Direction", "المعنى", "", "", "التقييم"], [
        ("HIGHER_IS_BETTER", "الأعلى أفضل", "", "", "الإنجاز = الفعلي ÷ المستهدف × 100"),
        ("LOWER_IS_BETTER", "الأقل أفضل", "", "", "100% إذا كان الفعلي ≤ المستهدف، وإلا المستهدف ÷ الفعلي × 100"),
        ("TARGET_BASED", "الأقرب للمستهدف", "", "", "100 − نسبة الانحراف عن المستهدف"),
        ("INFORMATIONAL", "معلوماتي", "", "", "يُعرض دون تقييم"),
    ])
    line("القسمة على صفر", "لا تظهر NaN أو Infinity أو #DIV/0!. إذا تعذر الحساب تظهر «غير متاح (N/A)» مع السبب.")

    title("الرفع إلى المنصة")
    for i, s in enumerate(["ادخل المنصة ← إدارة البيانات.", "اضغط «رفع ملف Excel» واختر الملف.",
                           "تقرأ المنصة ورقة DATA (أو IMPORT_TEMPLATE إذا عُبئت) بأسماء الأعمدة لا بمواقعها، وتعرض شاشة الفحص: السجلات، الجديدة، التحديثات، المكررة، الأخطاء، التحذيرات.",
                           "أصلح الأخطاء أو نزّل «تقرير الأخطاء» (Excel) لمعرفة السطر والحقل والقيمة والإجراء المقترح.",
                           "اضغط «IMPORT & SAVE». تُحفظ السجلات الصحيحة فقط، ويُسجل الرفع كدفعة (Batch) في سجل الاستيراد مع سجل كامل للتعديلات.",
                           "تتحدث لوحة المؤشرات مباشرة لكل المستخدمين."], 1):
        line(f"{i}", s)

    title("تنبيهات الأخطاء (نفس الرموز في ورقة VALIDATION وفي المنصة)")
    table(["الرمز", "المستوى", "الرسالة", "", "الإجراء المقترح"], [(c[0], c[1], c[2], "", c[4]) for c in CODES])

    title("أمثلة")
    table(["", "السجل", "", "", "النتيجة"], [
        ("سجل صحيح", "FAC001 | DEP005 | ED_VISITS | DAILY | 15/03/2026 | — | — | 142", "", "", "VALID — يُحفظ ويُحدّث «عدد مراجعي الطوارئ»"),
        ("سجل صحيح (نسبة)", "FAC001 | DEP004 | LAB_COMPLETION_PCT | WEEKLY | 08/03/2026 | 980 | 1000 | —", "", "", "VALID — القيمة المحسوبة 98%"),
        ("سجل خاطئ", "FAC001 | DEP005 | IND999 | DAILY | 2026/15/03 | — | — | 142 مراجع", "", "", "INVALID — E04 مؤشر غير موجود، E02 تاريخ غير صالح، E07 نص في حقل رقمي"),
        ("سجل خاطئ (نسبة)", "FAC001 | DEP004 | LAB_COMPLETION_PCT | WEEKLY | 08/03/2026 | 980 | 0 | —", "", "", "INVALID — E08 المقام صفر"),
    ])

    title("الحماية")
    line("الحماية", "المعادلات والأوراق المرجعية مقفلة لمنع التعديل غير المقصود، وخلايا الإدخال صفراء ومفتوحة. الحماية في Excel للتنظيم فقط: المنصة تعامل كل ملف كمدخل غير موثوق (UNTRUSTED INPUT) وتعيد فحص كل سطر في الخادم قبل الحفظ.")
    line("كلمة مرور الأوراق المرجعية", f"{MASTER_PASSWORD} (لمدير النظام فقط، عند إضافة مؤشر أو إدارة أو منشأة). أي إضافة هنا يجب أن تُضاف أيضاً في المنصة لتُقبل.")
    line("أسماء الأعمدة", "أسماء الأعمدة الإنجليزية في الصف 2 ثابتة وتعتمد عليها المنصة: لا تغيّرها. ترتيب الأعمدة أو الألوان أو التنسيق لا يؤثر على القراءة.")
    ws.sheet_view.showGridLines = False
    ws.protection.sheet = True
    ws.protection.password = MASTER_PASSWORD


def write_master_sql():
    """Rewrites the master-data seed in supabase/setup.sql so the database uses the same codes as the template."""
    path = os.path.join(ROOT, "supabase", "setup.sql")
    sql = open(path, encoding="utf-8").read()
    q = lambda v: "null" if v in (None, "") else "'" + str(v).replace("'", "''") + "'"
    lines = ["-- BEGIN MASTER DATA (generated by tools/build_excel_template.py; same codes as the Excel template)"]
    lines.append("insert into public.facilities(facility_id, facility_name, facility_name_ar, cluster, region, active, is_primary) values")
    lines.append(",\n".join(f"    ({q(f[0])}, {q(f[1])}, {q(f[2])}, {q(f[3])}, {q(f[4])}, {str(f[5] == 'Y').lower()}, {str(f[6] == 'Y').lower()})" for f in FACILITIES))
    lines.append("on conflict (facility_id) do nothing;")
    lines.append("insert into public.departments(department_id, department_name, department_name_ar, parent_department_id, department_type, active) values")
    lines.append(",\n".join(f"    ({q(d[0])}, {q(d[1])}, {q(d[2])}, {q(d[3])}, {q(d[4])}, {str(d[5] == 'Y').lower()})" for d in DEPARTMENTS))
    lines.append("on conflict (department_id) do nothing;")
    lines.append("insert into public.categories(category_id, category_name, category_name_ar, parent_category_id, sort_order) values")
    lines.append(",\n".join(f"    ({q(c[0])}, {q(c[1])}, {q(c[2])}, {q(c[3])}, {i})" for i, c in enumerate(CATEGORIES)))
    lines.append("on conflict (category_id) do nothing;")
    lines.append("-- END MASTER DATA")
    block = "\n".join(lines)
    new, n = re.subn(r"-- BEGIN MASTER DATA.*?-- END MASTER DATA", lambda m: block, sql, flags=re.S)
    if n != 1: sys.exit("setup.sql has no MASTER DATA markers")
    if new != sql: open(path, "w", encoding="utf-8").write(new)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--out")
    ap.add_argument("--no-sql", action="store_true")
    ap.add_argument("--test", type=int, default=0)
    a = ap.parse_args()
    if a.test: TEST_ROWS, CAPACITY = a.test, a.test + 20
    out = a.out or os.path.join(ROOT, "templates", f"APH_KPI_Data_Template_{a.year}.xlsx")
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    n = build(a.year, out)
    if not a.no_sql: write_master_sql()
    print(f"{out}: {n} calendar rows, {os.path.getsize(out) / 1e6:.1f} MB")
