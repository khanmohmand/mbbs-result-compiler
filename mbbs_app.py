"""MBBS Result Compiler (Streamlit)
Module-wise result: 1Best + Viva + SPE (scaled to their weightage) + CA (scaled to its weightage) = Grand Total.
Run:  streamlit run mbbs_app.py
"""
import io
import math

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MBBS Result Compiler", page_icon="🎓", layout="wide")

KEYS = ["1Best", "Viva", "SPE", "CA"]
COMPS = ["1Best", "Viva", "SPE"]                 # components that have a Minimum
INPUT_COLS = ["Roll No."] + KEYS
PROFS = {
    "1st Prof": ["Introduction to Medical Sciences", "Cell Biology & Genetics", "Haemopoietic and Immune System",
                 "Respiratory System", "Cardiovascular System"],
    "2nd Prof": ["Locomotion and Back Integumentary System", "Urinary System", "Reproductive Systems",
                 "Endocrinology", "GIT, Nutrition and Metabolism"],
}


# ------------------------------------------------------------------ calculation
def rnd(x: float) -> int:
    """Round half up to 0 decimals (54.5 -> 55)."""
    return int(math.floor(x + 0.5 + 1e-9))


def is_blank(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() in ("", "nan", "None")


def grand_bins(total: float, pass_marks: float) -> list:
    """6 dashboard bins: below pass marks, then 5 equal bands from pass marks up to the total."""
    edges = [pass_marks + round((total - pass_marks) * k / 5) for k in range(6)]
    return [(0, pass_marks - 1)] + [(edges[k], edges[k + 1] - 1 if k < 4 else total) for k in range(5)]


def improvement_text(names: list) -> str:
    if not names:
        return ""
    return f"The student needs to improve performance in {', '.join(names)}, where marks are below the required minimum."


def head_labels(wt: dict) -> list:
    return ["Roll No.", "1Best", f"Adj. on {wt['1Best']:g}", "Viva", f"Adj. on {wt['Viva']:g}", "SPE",
            f"Adj. on {wt['SPE']:g}", "Adjusted Total", f"CA Adj. on {wt['CA']:g}", "Grand Total",
            "Marks Required to Pass", "Status"]


def compute(df: pd.DataFrame, mx: dict, mn: dict, wt: dict, pass_marks: float):
    recs, errs = [], []
    for i, row in df.iterrows():
        roll = "" if is_blank(row.get("Roll No.")) else str(row.get("Roll No.")).strip()
        vals = {}
        for k in KEYS:
            v = row.get(k)
            if is_blank(v):
                vals[k] = None
            else:
                try:
                    vals[k] = float(v)
                except ValueError:
                    errs.append(f"{roll or 'Row ' + str(i + 1)}: {k} is not a number")
                    vals[k] = None
        if not roll and all(v is None for v in vals.values()):
            continue
        if not roll:
            errs.append(f"Row {i + 1}: Roll No. missing")
            roll = f"Row {i + 1}"
        if all(vals[k] is None for k in COMPS):               # no exam marks at all -> absent
            recs.append({"roll": roll, "absent": True, "status": "Absent"})
            continue
        if vals["CA"] is None:
            errs.append(f"{roll}: CA marks missing")
        m = {k: (v or 0.0) for k, v in vals.items()}
        for k in KEYS:
            if m[k] < 0 or m[k] > mx[k]:
                errs.append(f"{roll}: {k} must be between 0 and {mx[k]:g}")
        a = {k: m[k] / mx[k] * wt[k] for k in KEYS}
        adj_total = a["1Best"] + a["Viva"] + a["SPE"]
        grand = rnd(adj_total + a["CA"])
        weak = [k for k in COMPS if m[k] < mn[k]]
        recs.append({"roll": roll, "absent": False, "m": m, "a": a, "adj": adj_total, "G": grand,
                     "status": "Pass" if grand >= pass_marks else "Fail",
                     "need": "" if grand >= pass_marks else pass_marks - grand, "weak": weak})
    return recs, errs


def out_row(r: dict) -> list:
    if r["absent"]:
        return [r["roll"]] + [""] * 10 + ["Absent"]
    f = lambda x: round(x, 2)
    m, a = r["m"], r["a"]
    return [r["roll"], m["1Best"], f(a["1Best"]), m["Viva"], f(a["Viva"]), m["SPE"], f(a["SPE"]), f(r["adj"]),
            f(a["CA"]), r["G"], r["need"], r["status"]]


# ------------------------------------------------------------------ Excel
def build_excel(recs, info, mx, mn, wt, pass_marks) -> bytes:
    """3 sheets with live formulas: Result, Dashboard, Student Reports."""
    from openpyxl.formatting.rule import FormulaRule

    total = sum(wt[k] for k in KEYS)
    bins = grand_bins(total, pass_marks)
    wb = Workbook()
    ws = wb.active
    ws.title = "Result"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    head_fill = PatternFill("solid", fgColor="DCE6FF")

    for n, line in enumerate(info, start=1):
        ws.cell(n, 1, line).font = Font(bold=True, size=13)
        ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=8)
    ws.cell(5, 1, "Formulas are live: change 1Best, Viva, SPE marks or CA marks (col M) and everything updates. "
                  "Marks below the Minimum are shown in red.").font = Font(italic=True, size=9, color="777777")
    ws.cell(6, 1, "Maximum")
    ws.cell(7, 1, "Minimum")
    maxrow = {2: mx["1Best"], 3: wt["1Best"], 4: mx["Viva"], 5: wt["Viva"], 6: mx["SPE"], 7: wt["SPE"],
              8: wt["1Best"] + wt["Viva"] + wt["SPE"], 9: wt["CA"], 10: total, 13: mx["CA"]}
    for c, v in maxrow.items():
        ws.cell(6, c, v)
    for c, v in {2: mn["1Best"], 4: mn["Viva"], 6: mn["SPE"], 10: pass_marks}.items():
        ws.cell(7, c, v)
    heads = head_labels(wt) + ["CA Marks (entered)"]
    for c, h in enumerate(heads, start=1):
        ws.cell(8, c, h)
    for r in (6, 7, 8):
        for c in range(1, 14):
            cell = ws.cell(r, c)
            cell.font, cell.border, cell.alignment, cell.fill = Font(bold=True), box, center, head_fill

    first = 9
    for i, rec in enumerate(recs, start=first):
        ab = rec["absent"]
        A = f"COUNT($B{i},$D{i},$F{i})=0"

        def g(expr):
            return f'=IF({A},"",{expr})'

        cells = {
            1: rec["roll"],
            2: None if ab else rec["m"]["1Best"], 4: None if ab else rec["m"]["Viva"],
            6: None if ab else rec["m"]["SPE"], 13: None if ab else rec["m"]["CA"],
            3: g(f"N(B{i})/B$6*C$6"), 5: g(f"N(D{i})/D$6*E$6"), 7: g(f"N(F{i})/F$6*G$6"),
            8: g(f"C{i}+E{i}+G{i}"), 9: g(f"N(M{i})/M$6*I$6"), 10: g(f"ROUND(H{i}+I{i},0)"),
            11: g(f'IF(J{i}>=J$7,"",J$7-J{i})'),
            12: f'=IF({A},"Absent",IF(J{i}>=J$7,"Pass","Fail"))',
        }
        for c, v in cells.items():
            ws.cell(i, c, v)
        for c in range(1, 14):
            cell = ws.cell(i, c)
            cell.border, cell.alignment = box, center
            if c in (3, 5, 7, 8, 9):
                cell.number_format = "0.00"
            if c in (10, 11):
                cell.font = Font(bold=True)

    last = first - 1 + len(recs)
    red = PatternFill("solid", bgColor="FDE2E2", fgColor="FDE2E2")
    if recs:
        ws.conditional_formatting.add(f"A{first}:M{last}", FormulaRule(formula=[f'$L{first}<>"Pass"'], fill=red))
        ws.conditional_formatting.add(f"L{first}:L{last}", FormulaRule(formula=[f'$L{first}="Pass"'], font=Font(bold=True, color="12805C")))
        ws.conditional_formatting.add(f"L{first}:L{last}", FormulaRule(formula=[f'$L{first}<>"Pass"'], font=Font(bold=True, color="C0262D")))
        for col in "BDF":
            ws.conditional_formatting.add(f"{col}{first}:{col}{last}", FormulaRule(
                formula=[f"AND(ISNUMBER({col}{first}),{col}{first}<{col}$7)"], font=Font(bold=True, color="C0262D")))
    for c in range(1, 14):
        ws.column_dimensions[get_column_letter(c)].width = 22 if c == 11 else 14
    ws.freeze_panes = f"B{first}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "1:8"
    ws.print_area = f"A1:M{max(last, first)}"

    add_dashboard(wb, info, bins, total, pass_marks)
    add_reports(wb, info, recs, first, last)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def add_dashboard(wb, info, bins, total, pass_marks):
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    ds = wb.create_sheet("Dashboard")
    q, p = "Result!$L$9:$L$2000", "Result!$J$9:$J$2000"
    blue, light = "2F5BEA", "EAF0FF"
    thin = Side(style="thin", color="BBBBBB")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    for c in range(1, 11):
        ds.column_dimensions[get_column_letter(c)].width = 15
    ds.merge_cells("A1:J1")
    ds["A1"] = "RESULT DASHBOARD"
    ds["A1"].font = Font(bold=True, size=18, color="FFFFFF")
    ds["A1"].fill = PatternFill("solid", fgColor=blue)
    ds["A1"].alignment = center
    ds.row_dimensions[1].height = 34
    for n, line in enumerate(info, start=2):
        ds.merge_cells(start_row=n, start_column=1, end_row=n, end_column=10)
        ds.cell(n, 1, line).font = Font(bold=True, size=13)
        ds.cell(n, 1).alignment = center

    kpis = [("Students", "=COUNTA(Result!$A$9:$A$2000)", "0"),
            ("Appeared", "=A7-C7", "0"),
            ("Absent", f'=COUNTIF({q},"Absent")', "0"),
            ("Pass", f'=COUNTIF({q},"Pass")', "0"),
            ("Fail", f'=COUNTIF({q},"Fail")', "0"),
            ("Pass %", "=IFERROR(D7/B7,0)", "0.00%"),
            (f"Average /{total:g}", f"=IFERROR(AVERAGE({p}),0)", "0.00"),
            ("Highest", f"=MAX({p})", "0"),
            ("Lowest", f"=IF(COUNT({p})=0,0,MIN({p}))", "0"),
            ("Needs Improvement", "=COUNTIF('Student Reports'!$I$27:$I$2000,\"?*\")", "0")]
    ds.row_dimensions[7].height = 30
    for c, (lab, f, fmt) in enumerate(kpis, start=1):
        a, b = ds.cell(6, c, lab), ds.cell(7, c, f)
        a.font, a.fill = Font(bold=True, size=10, color="555555"), PatternFill("solid", fgColor=light)
        a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        b.font, b.number_format = Font(bold=True, size=16, color=blue), fmt
        b.alignment = center
        a.border = b.border = box

    ds["A9"], ds["B9"] = "Marks range", "Students"
    for cell in (ds["A9"], ds["B9"]):
        cell.font, cell.fill, cell.alignment, cell.border = Font(bold=True), PatternFill("solid", fgColor=light), center, box
    for i, (a, b) in enumerate(bins, start=10):
        ds.cell(i, 1, f"{a:g}-{b:g}")
        ds.cell(i, 2, f'=COUNTIFS({p},">={a:g}",{p},"<={b:g}")')
        for c in (1, 2):
            ds.cell(i, c).alignment, ds.cell(i, c).border = center, box
    ds["A16"], ds["B16"] = "Total", "=SUM(B10:B15)"
    for c in ("A16", "B16"):
        ds[c].font, ds[c].alignment, ds[c].border = Font(bold=True), center, box
    ds["A18"] = (f"Pass = Grand Total \u2265 {pass_marks:g}. 'Needs Improvement' = students below the minimum "
                 f"in at least one component (1Best, Viva or SPE).")
    ds["A18"].font = Font(italic=True, size=9, color="777777")

    ch = BarChart()
    ch.type, ch.style = "col", 10
    ch.title = f"Grand Total distribution (out of {total:g})"
    ch.add_data(Reference(ds, min_col=2, min_row=9, max_row=15), titles_from_data=True)
    ch.set_categories(Reference(ds, min_col=1, min_row=10, max_row=15))
    ch.legend = None
    ch.series[0].graphicalProperties.solidFill = blue
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.dataLabels.showCatName = ch.dataLabels.showSerName = ch.dataLabels.showLegendKey = ch.dataLabels.showPercent = False
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
    ch.y_axis.title, ch.x_axis.title = "Students", "Grand Total"
    ch.y_axis.majorGridlines = None
    ch.width, ch.height = 18, 8.5
    ds.add_chart(ch, "D9")

    ds.page_setup.orientation = "landscape"
    ds.page_setup.paperSize = ds.PAPERSIZE_A4
    ds.sheet_properties.pageSetUpPr.fitToPage = True
    ds.page_setup.fitToWidth, ds.page_setup.fitToHeight = 1, 1
    ds.print_options.horizontalCentered = True
    ds.page_margins.left = ds.page_margins.right = 0.5
    ds.print_area = "A1:J27"
    ds.sheet_view.showGridLines = False


def add_reports(wb, info, recs, first, last):
    """Sheet 3: pick a Roll No. -> that student's graph, numbers and text; below, all students with improvement remarks."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation

    rs = wb.create_sheet("Student Reports")
    blue, light = "2F5BEA", "EAF0FF"
    thin = Side(style="thin", color="BBBBBB")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    hfill = PatternFill("solid", fgColor=light)

    for n, line in enumerate(info, start=1):
        rs.cell(n, 1, line).font = Font(bold=True, size=13)
        rs.merge_cells(start_row=n, start_column=1, end_row=n, end_column=10)
    rs["A5"] = "INDIVIDUAL STUDENT REPORT"
    rs["A5"].font = Font(bold=True, size=14, color="FFFFFF")
    rs["A5"].fill = PatternFill("solid", fgColor=blue)
    rs.merge_cells("A5:J5")
    rs["A5"].alignment = center

    rs["A7"] = "Select Roll No."
    rs["A7"].font = Font(bold=True)
    rs["C7"] = recs[0]["roll"] if recs else ""
    rs["C7"].font, rs["C7"].fill, rs["C7"].border, rs["C7"].alignment = Font(bold=True, size=12), \
        PatternFill("solid", fgColor="FFF6CC"), box, center
    rs["D7"] = "\u2190 choose a Roll No. from the drop-down list"
    rs["D7"].font = Font(italic=True, size=9, color="777777")
    dv = DataValidation(type="list", formula1=f"=Result!$A${first}:$A${max(last, first)}", allow_blank=False)
    rs.add_data_validation(dv)
    dv.add("C7")
    rs["L7"] = "=IFERROR(MATCH($C$7,Result!$A$9:$A$2000,0),0)"          # helper: position of selected student
    rs["L7"].font = Font(color="FFFFFF")
    rs["K7"] = "row"
    rs["K7"].font = Font(color="FFFFFF")

    for c, h in enumerate(["Component", "Marks Obtained", "Maximum", "Minimum", "Adjusted Marks", "Weightage", "Status"], start=1):
        cell = rs.cell(9, c, h)
        cell.font, cell.fill, cell.border, cell.alignment = Font(bold=True), hfill, box, center

    def pick(col):                       # value from Result column `col` for the selected student ("" if absent)
        return (f'IF(OR($L$7=0,INDEX(Result!$J$9:$J$2000,$L$7)=""),"",INDEX(Result!${col}$9:${col}$2000,$L$7))')

    comp_cols = {"1Best": ("B", "C", "B"), "Viva": ("D", "E", "D"), "SPE": ("F", "G", "F")}   # raw, adjusted, max/min col
    r = 10
    for name, (raw, adj, mm) in comp_cols.items():
        rs.cell(r, 1, name)
        rs.cell(r, 2, "=" + pick(raw))
        rs.cell(r, 3, f"=Result!${mm}$6")
        rs.cell(r, 4, f"=Result!${mm}$7")
        rs.cell(r, 5, "=" + pick(adj))
        rs.cell(r, 6, f"=Result!${adj}$6")
        rs.cell(r, 7, f'=IF(B{r}="","",IF(B{r}<D{r},"Below Minimum","Satisfactory"))')
        r += 1
    rs.cell(13, 1, "CA")
    rs.cell(13, 2, "=" + pick("M"))
    rs.cell(13, 3, "=Result!$M$6")
    rs.cell(13, 4, "-")
    rs.cell(13, 5, "=" + pick("I"))
    rs.cell(13, 6, "=Result!$I$6")
    rs.cell(13, 7, "-")
    rs.cell(14, 1, "Grand Total")
    rs.cell(14, 5, "=" + pick("J"))
    rs.cell(14, 6, "=Result!$J$6")
    rs.cell(14, 7, '=IF($L$7=0,"",INDEX(Result!$L$9:$L$2000,$L$7))')
    rs.cell(15, 1, "Marks Required to Pass")
    rs.cell(15, 5, "=" + pick("K"))
    for rr in range(10, 16):
        for c in range(1, 8):
            cell = rs.cell(rr, c)
            cell.border, cell.alignment = box, center
            if c == 5 and rr != 14 and rr != 15:
                cell.number_format = "0.00"
        rs.cell(rr, 1).font = Font(bold=True)
    for c in (1, 5, 6, 7):
        rs.cell(14, c).font = Font(bold=True)
    rs.conditional_formatting.add("G10:G12", FormulaRule(formula=['$G10="Below Minimum"'], font=Font(bold=True, color="C0262D")))
    rs.conditional_formatting.add("B10:B12", FormulaRule(formula=['AND(ISNUMBER($B10),$B10<$D10)'], font=Font(bold=True, color="C0262D")))
    rs.conditional_formatting.add("G14", FormulaRule(formula=['$G$14="Pass"'], font=Font(bold=True, color="12805C")))
    rs.conditional_formatting.add("G14", FormulaRule(formula=['$G$14="Fail"'], font=Font(bold=True, color="C0262D")))

    rs["A17"] = ('=IF($L$7=0,"Select a Roll No. from the list above.",IF($G$14="Absent","Roll No. "&$C$7&" was absent.",'
                 '"Roll No. "&$C$7&" obtained "&$E$14&" out of "&$F$14&" marks and is declared "&$G$14&"."'
                 '&IF($E$15<>"","  "&$E$15&" more marks were required to pass.","")'
                 '&IF(INDEX($J$27:$J$2000,$L$7)<>"","  "&INDEX($J$27:$J$2000,$L$7),"")))')
    rs.merge_cells("A17:G20")
    rs["A17"].alignment = wrap
    rs["A17"].font = Font(size=11)

    ch = BarChart()
    ch.type, ch.style = "col", 10
    ch.title = "Marks Obtained vs Minimum and Maximum"
    ch.add_data(Reference(rs, min_col=2, max_col=4, min_row=9, max_row=12), titles_from_data=True)
    ch.set_categories(Reference(rs, min_col=1, min_row=10, max_row=12))
    for s, colr in zip(ch.series, (blue, "BBBBBB", "E5484D")):
        s.graphicalProperties.solidFill = colr
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
    ch.y_axis.title = "Marks"
    ch.legend.position = "b"
    ch.width, ch.height = 13, 8
    rs.add_chart(ch, "I8")

    # ---- all students table
    rs["A24"] = "ALL STUDENTS \u2013 COMPONENT ANALYSIS (use the filter on 'Below Minimum In' to see students needing improvement)"
    rs["A24"].font = Font(bold=True, size=12, color="FFFFFF")
    rs["A24"].fill = PatternFill("solid", fgColor=blue)
    rs.merge_cells("A24:J24")
    for c, h in enumerate(["Roll No.", "1Best", "Viva", "SPE", "Adjusted Total", "CA (Adjusted)", "Grand Total", "Status",
                           "Below Minimum In", "Improvement Remark"], start=1):
        cell = rs.cell(26, c, h)
        cell.font, cell.fill, cell.border, cell.alignment = Font(bold=True), hfill, box, center
    for k, i in enumerate(range(first, last + 1)):
        t = 27 + k
        gone = f"COUNT(Result!$B{i},Result!$D{i},Result!$F{i})=0"
        vals = {1: f"=Result!A{i}",
                2: f'=IF({gone},"",Result!B{i})', 3: f'=IF({gone},"",Result!D{i})', 4: f'=IF({gone},"",Result!F{i})',
                5: f"=Result!H{i}", 6: f"=Result!I{i}", 7: f"=Result!J{i}", 8: f"=Result!L{i}",
                9: (f'=IF({gone},"",MID(IF(N(Result!B{i})<Result!$B$7,", 1Best","")&IF(N(Result!D{i})<Result!$D$7,", Viva","")'
                    f'&IF(N(Result!F{i})<Result!$F$7,", SPE",""),3,50))'),
                10: f'=IF(I{t}="","","The student needs to improve performance in "&I{t}&", where marks are below the required minimum.")'}
        for c, v in vals.items():
            cell = rs.cell(t, c, v)
            cell.border = box
            cell.alignment = wrap if c == 10 else center
            if c in (5, 6):
                cell.number_format = "0.00"
    end = max(27, 26 + len(recs))
    rs.auto_filter.ref = f"A26:J{end}"
    rs.conditional_formatting.add(f"A27:J{end}", FormulaRule(
        formula=['$I27<>""'], fill=PatternFill("solid", bgColor="FDE2E2", fgColor="FDE2E2")))
    for c, w in enumerate([22, 14, 12, 12, 16, 14, 13, 14, 20, 70], start=1):
        rs.column_dimensions[get_column_letter(c)].width = w
    rs.column_dimensions["K"].width = 4
    rs.page_setup.orientation = "landscape"
    rs.page_setup.paperSize = rs.PAPERSIZE_A4
    rs.sheet_properties.pageSetUpPr.fitToPage = True
    rs.page_setup.fitToWidth, rs.page_setup.fitToHeight = 1, 0
    rs.sheet_view.showGridLines = False


# ------------------------------------------------------------------ upload helpers
def template_bytes() -> bytes:
    buf = io.BytesIO()
    pd.DataFrame([["Roll-001", None, None, None, None]], columns=INPUT_COLS).to_excel(buf, index=False)
    return buf.getvalue()


def read_table(file, required: list) -> pd.DataFrame:
    df = pd.read_csv(file) if file.name.lower().endswith(".csv") else pd.read_excel(file)
    names = {}
    for col in df.columns:
        s = "".join(ch for ch in str(col).lower() if ch.isalnum())
        if "roll" in s:
            names[col] = "Roll No."
        elif "best" in s:
            names[col] = "1Best"
        elif "viva" in s:
            names[col] = "Viva"
        elif s.startswith("spe"):
            names[col] = "SPE"
        elif s == "ca" or "continuous" in s:
            names[col] = "CA"
    df = df.rename(columns=names)
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(missing))
    return df[[c for c in INPUT_COLS if c in df.columns]]


def attach_ca(df: pd.DataFrame, ca_df) -> pd.DataFrame:
    """CA comes from the marks sheet, or (if blank / missing there) from the separately uploaded CA sheet."""
    df = df.copy()
    if "CA" not in df.columns:
        df["CA"] = None
    if ca_df is not None:
        look = {str(r).strip(): c for r, c in zip(ca_df["Roll No."], ca_df["CA"])}
        df["CA"] = [c if not is_blank(c) else look.get(str(r).strip()) for r, c in zip(df["Roll No."], df["CA"])]
    return df


# ------------------------------------------------------------------ UI
CSS = """
<style>
.stApp{background:#ffffff}
.block-container{padding-top:4rem !important;max-width:1250px}
.banner{background:#008000;color:#fff;text-align:center;font:700 32px Georgia,'Times New Roman',serif;
        padding:14px 0 14px;margin:0 0 12px;line-height:1.25}
.banner small{display:block;font:400 13px Verdana,Arial,sans-serif;opacity:.9;margin-top:2px}
.lbl{color:#0a7d0a;font:700 17px Verdana,Arial,sans-serif;padding-top:7px}
.lbl-s{color:#0a7d0a;font:700 14px Verdana,Arial,sans-serif;padding-top:7px}
.hdr{background:#0a7d0a;color:#fff;text-align:center;font:700 14px Georgia,serif;padding:4px 0;letter-spacing:.6px}
.st-key-panel1,.st-key-panel2,.st-key-panel3,.st-key-panel4{border:1px solid #22cc22;padding:12px 16px;
        margin-bottom:10px;background:#fff}
.st-key-wbox,.st-key-wtbox{border:1px solid #0a7d0a;padding:0 0 8px;background:#fff}
.st-key-wbox .hdr,.st-key-wtbox .hdr{margin-bottom:8px}
.st-key-clearbox{background:#0a7d0a;padding:8px;margin-top:8px;text-align:center}
.st-key-clearbox div[data-testid="stButton"]{width:100%;display:flex;justify-content:center}
.st-key-clearbox button{background:#fff !important;color:#0a2a5e !important;font-weight:700;border:1px solid #333 !important;
        border-radius:2px;font-family:Georgia,serif;width:auto !important;min-width:120px;white-space:nowrap}
[data-testid="stTextInputRootElement"],[data-testid="stNumberInputContainer"]{border:1.5px solid #5f9ea0 !important;
        border-radius:4px !important;background:#fff !important}
[data-testid="stTextInputRootElement"]:focus-within,[data-testid="stNumberInputContainer"]:focus-within{
        border-color:#0a7d0a !important;box-shadow:0 0 0 2px rgba(34,204,34,.25)}
div[data-baseweb="input"],div[data-baseweb="base-input"]{border:none !important;background:#fff !important}
button[kind="primary"],.stDownloadButton button{background:#0a7d0a !important;color:#fff !important;
        border:1px solid #065d06 !important;border-radius:4px !important;font-weight:700 !important}
button[kind="primary"]:hover,.stDownloadButton button:hover{background:#12a012 !important}
[data-testid="stMetric"]{background:#f1fbf1;border:1px solid #22cc22;border-radius:6px;padding:8px 12px}
[data-testid="stMetricLabel"] p{color:#0a7d0a;font-weight:700}
[data-testid="stMetricValue"]{color:#0a4d0a}
.stRadio label p{color:#0a7d0a;font-weight:700}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)
st.markdown('<div class="banner">MBBS Result Compiler<small>Module-wise result · 1Best + Viva + SPE + CA · '
            'Auto weightage, round-off, pass marks and student reports</small></div>', unsafe_allow_html=True)


def clear_all():
    for k in ("prog", "exam", "batch"):
        st.session_state[k] = ""
    for k in KEYS:
        st.session_state["w_" + k] = None
        st.session_state["wt_" + k] = None
    for k in COMPS:
        st.session_state["mn_" + k] = None
    st.session_state["pm"] = 55.0
    st.session_state.pop("out", None)


def sync_min(k):
    """When Max is typed, Min is filled with 50% of it (the user can still edit Min afterwards)."""
    v = st.session_state.get("w_" + k)
    st.session_state["mn_" + k] = (v / 2) if v else None


def lab(text, small=False):
    return f'<div class="{"lbl-s" if small else "lbl"}">{text}</div>'


# ---- Panel 1: result information, weightage, max / min
with st.container(key="panel1"):
    left, right = st.columns([3, 2], gap="large")
    with left:
        c1, c2 = st.columns([1.2, 2])
        c1.markdown(lab("Program Name"), unsafe_allow_html=True)
        prog = c2.text_input("Program Name", key="prog", label_visibility="collapsed", placeholder="e.g. MBBS")
        c1, c2 = st.columns([1.2, 2])
        c1.markdown(lab("Exam Name"), unsafe_allow_html=True)
        exam = c2.text_input("Exam Name", key="exam", label_visibility="collapsed", placeholder="e.g. Annual Examination 2026")
        c1, c2 = st.columns([1.2, 2])
        c1.markdown(lab("Batch &amp; Year"), unsafe_allow_html=True)
        batch = c2.text_input("Batch & Year", key="batch", label_visibility="collapsed", placeholder="e.g. 5th Batch, Second Year")
        c1, c2 = st.columns([1.2, 2])
        c1.markdown(lab("Prof"), unsafe_allow_html=True)
        prof = c2.selectbox("Prof", list(PROFS), index=1, label_visibility="collapsed")
        c1, c2 = st.columns([1.2, 2])
        c1.markdown(lab("Module Name"), unsafe_allow_html=True)
        module = c2.selectbox("Module Name", PROFS[prof], key="mod_" + prof, label_visibility="collapsed")
        with st.container(key="wtbox"):
            st.markdown('<div class="hdr">WEIGHTAGE (ADJUSTED MARKS)</div>', unsafe_allow_html=True)
            wt_in = {}
            for key, label in zip(KEYS, ["1Best", "Viva", "SPE", "CA"]):
                c1, c2 = st.columns([1.2, 2])
                c1.markdown(lab(label, small=True), unsafe_allow_html=True)
                wt_in[key] = c2.number_input(f"{label} (weightage)", min_value=1.0, value=None, step=1.0, key="wt_" + key,
                                             label_visibility="collapsed", placeholder="required")
            c1, c2 = st.columns([1.2, 2])
            c1.markdown(lab("Pass Marks", small=True), unsafe_allow_html=True)
            pass_marks = c2.number_input("Pass Marks", min_value=1.0, value=55.0, step=1.0, key="pm", label_visibility="collapsed")
        if all(v is not None for v in wt_in.values()):
            st.caption(f"Grand Total = {sum(wt_in.values()):g} marks · Pass at {pass_marks:g} or more (after rounding).")
        st.caption("Weightage = marks each component is adjusted to. All boxes here are required.")
    with right:
        with st.container(key="wbox"):
            st.markdown('<div class="hdr">MAX &amp; MIN MARKS</div>', unsafe_allow_html=True)
            h1, h2, h3 = st.columns([1, 1, 1])
            h2.markdown(lab("Max", small=True), unsafe_allow_html=True)
            h3.markdown(lab("Min", small=True), unsafe_allow_html=True)
            mx_in, mn_in = {}, {}
            for key in KEYS:
                c1, c2, c3 = st.columns([1, 1, 1])
                c1.markdown(lab(key, small=True), unsafe_allow_html=True)
                mx_in[key] = c2.number_input(f"{key} (max)", min_value=1.0, value=None, step=1.0, key="w_" + key,
                                             label_visibility="collapsed", placeholder="required",
                                             on_change=sync_min, args=(key,))
                if key in COMPS:
                    mn_in[key] = c3.number_input(f"{key} (min)", min_value=0.0, value=None, step=0.5, key="mn_" + key,
                                                 label_visibility="collapsed", placeholder="required")
            st.caption("Max = marks of the exam as given by the examiner. Min fills with 50% of Max; you can change it.")
        with st.container(key="clearbox"):
            st.button("CLEAR", on_click=clear_all)

# ---- Panel 2: input method
with st.container(key="panel2"):
    m1, m2 = st.columns([1, 3])
    m1.markdown(lab("Input Method"), unsafe_allow_html=True)
    mode = m2.radio("Input method", ["Type manually", "Upload Excel"], horizontal=True, label_visibility="collapsed")

# ---- Panel 3: student data
data, ca_df = None, None
with st.container(key="panel3"):
    st.markdown('<div class="hdr">STUDENT DATA</div>', unsafe_allow_html=True)
    ca_up = st.file_uploader("CA sheet (optional): Roll No. and CA, out of CA max. Same CA is used for the module you compile.",
                             type=["xlsx", "xls", "csv"], key="ca_up")
    if ca_up is not None:
        try:
            ca_df = read_table(ca_up, ["Roll No.", "CA"])
            st.success(f"CA loaded for {len(ca_df)} students. It fills CA wherever the marks sheet has none.")
        except Exception as e:
            st.error(f"CA sheet error: {e}")
    if mode == "Type manually":
        blank = pd.DataFrame({"Roll No.": pd.Series([""] * 10, dtype="object"),
                              **{k: pd.Series([None] * 10, dtype="float64") for k in KEYS}})
        data = st.data_editor(blank, num_rows="dynamic", use_container_width=True, hide_index=True, key="grid",
                              column_config={"Roll No.": st.column_config.TextColumn("Roll No.", width="medium")})
    else:
        st.caption("Columns: Roll No., 1Best, Viva, SPE (and CA, unless the CA sheet above is used). "
                   "Absent: keep the Roll No. and leave all exam marks empty.")
        st.download_button("Download blank template", template_bytes(), "MBBS_Result_Template.xlsx")
        up = st.file_uploader("Upload module marks (Excel / CSV)", type=["xlsx", "xls", "csv"], key="marks_up")
        if up is not None:
            try:
                data = read_table(up, ["Roll No.", "1Best", "Viva", "SPE"])
                st.success(f"{len(data)} rows loaded.")
            except Exception as e:
                st.error(f"Excel error: {e}")
    go = st.button("Generate Result", type="primary")

if go:
    total = sum(v for v in wt_in.values() if v is not None)
    if not (prog.strip() and exam.strip() and batch.strip()):
        st.error("Program Name, Exam Name and Batch & Year are required.")
    elif any(v is None for v in list(mx_in.values()) + list(mn_in.values())):
        st.error("All Max and Min boxes are required.")
    elif any(v is None for v in wt_in.values()):
        st.error("All weightage boxes are required.")
    elif any(mn_in[k] > mx_in[k] for k in COMPS):
        st.error("Min marks cannot be greater than Max marks.")
    elif pass_marks > total - 5:
        st.error("Pass Marks must be at least 5 marks below the Grand Total. Please check weightage and Pass Marks.")
    elif data is None or data.empty:
        st.error("No student data found.")
    else:
        recs, errs = compute(attach_ca(data, ca_df), mx_in, mn_in, wt_in, pass_marks)
        if errs:
            st.error("\n\n".join(errs[:8]) + (f"\n\n+{len(errs) - 8} more" if len(errs) > 8 else ""))
        elif not recs:
            st.error("No student data found.")
        else:
            info = [prog.strip(), f"{prof} \u2013 {exam.strip()}", batch.strip(), f"Module: {module}"]
            st.session_state["out"] = (recs, info, dict(mx_in), dict(mn_in), dict(wt_in), pass_marks, module)

# ---- Panel 4: dashboard, result sheet, student report
if "out" in st.session_state:
    recs, info, mx, mn, wt, pm, module = st.session_state["out"]
    total = sum(wt.values())
    app = [r for r in recs if not r["absent"]]
    n_pass = sum(r["status"] == "Pass" for r in app)
    weak = [r for r in app if r["weak"]]
    with st.container(key="panel4"):
        st.markdown('<div class="hdr">DASHBOARD</div>', unsafe_allow_html=True)
        st.write("")
        k = st.columns(8)
        k[0].metric("Students", len(recs))
        k[1].metric("Appeared", len(app))
        k[2].metric("Absent", len(recs) - len(app))
        k[3].metric("Pass", n_pass)
        k[4].metric("Fail", len(app) - n_pass)
        k[5].metric("Pass %", f"{n_pass / len(app) * 100:.2f}%" if app else "0%")
        k[6].metric(f"Average /{total:g}", f"{sum(r['G'] for r in app) / len(app):.2f}" if app else "0")
        k[7].metric("Needs Improvement", len(weak))

        bins = grand_bins(total, pm)
        st.bar_chart(pd.DataFrame({"Students": [sum(a <= r["G"] <= b for r in app) for a, b in bins]},
                                  index=[f"{a:g}-{b:g}" for a, b in bins]), color="#0a7d0a")

        st.markdown('<div class="hdr">RESULT SHEET</div>', unsafe_allow_html=True)
        st.markdown("  \n".join(f"**{line}**" for line in info))
        hd = head_labels(wt)
        df = pd.DataFrame([out_row(r) for r in recs], columns=[h + " " * hd[:i].count(h) for i, h in enumerate(hd)])
        for c in df.columns[[2, 4, 6, 7, 8]]:
            df[c] = df[c].map(lambda x: f"{x:.2f}" if x != "" else "\u2014")
        st.caption(f"Maximum \u2192 1Best {mx['1Best']:g} | Viva {mx['Viva']:g} | SPE {mx['SPE']:g} | CA {mx['CA']:g} \u00b7 "
                   f"Minimum \u2192 1Best {mn['1Best']:g} | Viva {mn['Viva']:g} | SPE {mn['SPE']:g} \u00b7 "
                   f"Weightage \u2192 1Best {wt['1Best']:g} | Viva {wt['Viva']:g} | SPE {wt['SPE']:g} | CA {wt['CA']:g} \u00b7 "
                   f"Pass Marks {pm:g}")
        st.dataframe(df.style.apply(lambda r: ["background-color:#fde2e2; color:#7a1117"
                                               if r["Status"] != "Pass" else "" for _ in r], axis=1),
                     use_container_width=True, hide_index=True)

        st.markdown('<div class="hdr">STUDENTS NEEDING IMPROVEMENT</div>', unsafe_allow_html=True)
        if weak:
            st.dataframe(pd.DataFrame([{"Roll No.": r["roll"],
                                        **{c: f"{r['m'][c]:g} / min {mn[c]:g}" for c in COMPS},
                                        "Below Minimum In": ", ".join(r["weak"]),
                                        "Improvement Remark": improvement_text(r["weak"])} for r in weak]),
                         use_container_width=True, hide_index=True)
        else:
            st.info("No student is below the minimum in any component.")

        st.markdown('<div class="hdr">INDIVIDUAL STUDENT REPORT</div>', unsafe_allow_html=True)
        if app:
            pick = st.selectbox("Select Roll No.", [r["roll"] for r in app])
            s = next(r for r in app if r["roll"] == pick)
            st.bar_chart(pd.DataFrame({"Marks Obtained": [s["m"][c] for c in COMPS],
                                       "Minimum": [mn[c] for c in COMPS],
                                       "Maximum": [mx[c] for c in COMPS]}, index=COMPS),
                         color=["#2F5BEA", "#E5484D", "#BBBBBB"])
            txt = f"Roll No. {s['roll']} obtained {s['G']} out of {total:g} marks and is declared {s['status']}."
            if s["need"] != "":
                txt += f" {s['need']:g} more marks were required to pass."
            if s["weak"]:
                txt += " " + improvement_text(s["weak"])
            st.write(txt)

        st.download_button("\u2b07\ufe0f Download Excel", build_excel(recs, info, mx, mn, wt, pm),
                           f"{module.replace(' ', '_').replace(',', '').replace('&', 'and')}_Result.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")