"""Build fixture_defects.xlsx: a small workbook where every lint rule has a known target.
Run:  py -3.14 make_fixture.py <out.xlsx> [--check]
--check lints the built file through xl and exits 1 if any rule id in xl.lint.RULE_IDS did not
fire, any EXPECT_WHERE regression cell was not named, or a switched-off id (xl.lint.OFF_IDS)
fired. The planted defects for switched-off rules stay in the file as negative controls.
(Before 2026-09-08 the script only built; the assertion had to be called separately. Before
2026-09-09 EXPECT listed 24 of the 32 rule ids by hand.)"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation


def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws["A1"] = "Line"
    for i, y in enumerate(range(2024, 2030)):
        ws.cell(1, 2 + i, y)
    # row 2: revenue inputs, row 3 formula row with a hardcode plug in D3
    for i in range(6):
        ws.cell(2, 2 + i, 100 + 10 * i)
    ws["A2"] = "Revenue"
    ws["A3"] = "Growth"
    ws["B3"] = 0.05
    ws["C3"] = "=C2/B2-1"
    ws["D3"] = 0.07           # hardcode-in-formula-row (formula on both sides)
    ws["E3"] = "=E2/D2-1"
    ws["F3"] = "=F2/E2-1"
    ws["G3"] = "=G2/F2-1"
    for col in "BCDEFG":
        ws[f"{col}3"].number_format = "0.0%"
    # row 4: inconsistent formula (E4 breaks after a run of consistent copies)
    ws["A4"] = "Cost"
    ws["B4"] = "=B2*0.4"
    ws["C4"] = "=C2*0.4"
    ws["D4"] = "=D2*0.4"
    ws["E4"] = "=E2*0.5"      # inconsistent-formula
    ws["F4"] = "=F2*0.4"
    ws["G4"] = "=G2*0.4"
    # row 5: mixed percent  (B3 percent + B2 number)
    ws["A5"] = "Mixed"
    ws["B5"] = "=B3+B2"       # mixed-percent
    # rows 7-10 a block, row 11 SUM that misses row 10  -> sum-misses-adjacent
    ws["A7"] = "Opex A"; ws["B7"] = 10
    ws["A8"] = "Opex B"; ws["B8"] = 20
    ws["A9"] = "Opex C"; ws["B9"] = 30
    ws["A10"] = "Opex D"; ws["B10"] = 40
    ws["A11"] = "Total opex"; ws["B11"] = "=SUM(B7:B9)"   # misses B10
    # double counting
    ws["A12"] = "Double"; ws["B12"] = "=SUM(B7:B9)+SUM(B8:B10)"   # double-counting
    # empty coercion
    ws["A13"] = "Coerce"; ws["B13"] = "=B2*Z9"            # Z9 empty
    # number as text
    ws["A14"] = "Text num"; ws["B14"] = 5; ws["C14"] = "1.234,56"   # number-as-text
    # error constant
    ws["A15"] = "Err"; ws["B15"] = "#REF!"                # error-constant
    # column clipping: big number in a narrow column
    ws["A16"] = "Clip"; ws["B16"] = 123456789012.0
    ws["B16"].number_format = "#,##0"
    ws.column_dimensions["B"].width = 6                    # column-clipping
    # merged cells
    ws.merge_cells("A18:C18"); ws["A18"] = "Merged title"  # merged-cells
    # data validation breach
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv)
    ws["A20"] = "Flag"; ws["B20"] = "Maybe"; dv.add("B20")   # data-validation-breach
    # chart with a series pointing at an empty range and one at a missing sheet
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=2, min_row=2, max_col=7, max_row=2), from_rows=True)
    ch.add_data(Reference(ws, min_col=2, min_row=30, max_col=7, max_row=30), from_rows=True)  # empty range
    ws.add_chart(ch, "A7")                                  # object-covers-cells (over populated area)
    # broken defined name
    wb.create_sheet("Scratch")
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["BadName"] = DefinedName("BadName", attr_text="#REF!")

    # ---- 2026-09-08: Witan-parity rules ----
    # broadcast-surprise: a range used directly in arithmetic
    ws["A22"] = "Broadcast"; ws["B22"] = "=B2+B7:B10"
    # row-height-clipping: long wrapped text in a 15 pt row, 10-wide column
    ws["A23"] = "Wrapped"
    ws["C23"] = "This label is far too long to fit on one line of a ten-character column and the row height is fixed"
    ws["C23"].alignment = Alignment(wrap_text=True)
    ws.column_dimensions["C"].width = 10
    ws.row_dimensions[23].height = 15
    # aggregate-over-text: SUM over a range with a text cell
    ws["I1"] = 1; ws["I2"] = "oops"; ws["I3"] = 3; ws["J1"] = "=SUM(I1:I3)"
    # mixed-currency: $ and € in one SUM
    ws["K1"] = 10; ws["K1"].number_format = '"$"#,##0.00'
    ws["K2"] = 20; ws["K2"].number_format = '"€"#,##0.00'
    ws["K3"] = "=SUM(K1:K2)"
    # currency-date-mix: currency + date
    ws["L1"] = 100; ws["L1"].number_format = '"$"#,##0'
    ws["L2"] = dt.date(2026, 1, 31); ws["L2"].number_format = "yyyy-mm-dd"
    ws["L3"] = "=L1+L2"

    lk = wb.create_sheet("Lookups")
    for row in [["Key", "Value"], [1, "one"], [3, "three"], [2, "two"]]:
        lk.append(row)
    lk["D1"] = "=VLOOKUP(2,A2:B4,2,TRUE)"                  # unsorted-lookup
    for i, row in enumerate([["ID001", "first"], ["ID001", "second"], ["ID002", "third"]], start=1):
        lk.cell(i, 6, row[0]); lk.cell(i, 7, row[1])
    lk["D2"] = '=VLOOKUP("ID001",F1:G3,2,FALSE)'            # duplicate-lookup-keys

    # chart geometry on Scratch: 12 long category labels, 9 series, narrow chart; plus an invisible chart
    sc = wb["Scratch"]
    sc["A1"] = "Category"
    for i in range(12):
        sc.cell(2 + i, 1, f"Business unit number {i + 1} (consolidated)")
    for s in range(9):
        sc.cell(1, 2 + s, f"Series {s + 1}")
        for i in range(12):
            sc.cell(2 + i, 2 + s, (i + 1) * (s + 1))
    ch2 = LineChart()
    ch2.add_data(Reference(sc, min_col=2, min_row=1, max_col=10, max_row=13), titles_from_data=True)
    ch2.set_categories(Reference(sc, min_col=1, min_row=2, max_row=13))
    ch2.width, ch2.height = 8, 6                             # chart-axis-labels-crowded, chart-legend-crowded
    sc.add_chart(ch2, "M1")
    ch3 = LineChart()
    ch3.add_data(Reference(sc, min_col=2, min_row=1, max_col=2, max_row=13), titles_from_data=True)
    ch3.width, ch3.height = 0.1, 0.1                         # chart-invisible
    sc.add_chart(ch3, "M20")

    # ---- 2026-09-09: the eight rules that had no planted defect, plus regression cases ----
    # chart-series-broken: a series whose reference names a sheet that does not exist
    ch4 = LineChart()
    ch4.add_data(Reference(ws, min_col=2, min_row=2, max_col=7, max_row=2), from_rows=True)
    ch4.series[0].val.numRef.f = "'Gone'!$B$2:$G$2"
    # chart-series-errors: a series over the #REF! constant in B15
    ch4.add_data(Reference(ws, min_col=2, min_row=15, max_col=3, max_row=15), from_rows=True)
    sc.add_chart(ch4, "M40")
    # hidden: a hidden row and a hidden column
    ws.row_dimensions[25].hidden = True
    ws.column_dimensions["N"].hidden = True
    # iterative-calc
    wb.calculation.iterate = True
    # row-format-inconsistent: six 0.00 cells with one 0.0% in the middle of the row
    ws["A26"] = "Row fmt"
    for i, col in enumerate("BCDEFG"):
        ws[f"{col}26"] = 1.5 + i
        ws[f"{col}26"].number_format = "0.00"
    ws["D26"].number_format = "0.0%"
    # external-link + external-link-absolute: an externalLink part with an absolute file target
    from openpyxl.workbook.external_link.external import (ExternalBook, ExternalLink, ExternalSheetData,
                                                          ExternalSheetDataSet, ExternalSheetNames)
    from openpyxl.packaging.relationship import Relationship
    # Excel refuses to open a bare <externalBook/>: it needs r:id, sheetNames and a sheetDataSet
    book = ExternalBook(id="rId1",
                        sheetNames=ExternalSheetNames(sheetName=["Sheet1"]),
                        sheetDataSet=ExternalSheetDataSet(sheetData=[ExternalSheetData(sheetId=0)]))
    ext = ExternalLink(externalBook=book)
    ext.file_link = Relationship(type="externalLinkPath", Target="file:///C:/Models/source.xlsx", TargetMode="External")
    wb._external_links.append(ext)
    # cached-error: a formula whose CACHED value is an error (patched into the XML after save,
    # because openpyxl never writes cached values). B27 = 1/0 -> #DIV/0!
    ws["A27"] = "Cached err"; ws["B27"] = "=1/0"
    # regression: an array formula must be seen by every formula walker (was invisible)
    from openpyxl.worksheet.formula import ArrayFormula
    ws["A28"] = "Array"; ws["B28"] = ArrayFormula("B28:B28", "=SUM(B7:B9)")   # sum-misses-adjacent via array formula
    # regression: SUMIF is not an IF guard -> empty-cell-coercion must fire on Z10
    ws["A29"] = "Sumif"; ws["B29"] = '=SUMIF(A7:A10,"Opex A",B7:B10)/Z10'
    # regression: a sheet name with an apostrophe must not be skipped by range-reading rules
    ap = wb.create_sheet("Ana's Sheet")
    for row in [["Key", "Value"], [1, "one"], [3, "three"], [2, "two"]]:
        ap.append(row)
    ap["D1"] = "=VLOOKUP(2,'Ana''s Sheet'!A2:B4,2,TRUE)"     # unsorted-lookup on the quoted sheet
    wb.save(path)
    _patch_cached_error(path, "Model", "B27", "#DIV/0!")
    return path


def _patch_cached_error(path, sheet_title, cell, err):
    """Give one formula cell a cached error value, the way Excel would have saved it.
    openpyxl writes `<c r="B27"><f>1/0</f><v></v></c>`; Excel writes t="e" and <v>#DIV/0!</v>."""
    import re
    import shutil
    import zipfile
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    idx = wb.sheetnames.index(sheet_title) + 1
    wb.close()
    part = f"xl/worksheets/sheet{idx}.xml"
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                xml = data.decode("utf-8")
                rx = re.compile(r'<c r="%s"([^>]*)>(<f>[^<]*</f>)(?:<v></v>|<v/>|<v>[^<]*</v>)?</c>' % cell)
                m = rx.search(xml)
                assert m, f"{cell} formula cell not found in {part}"
                attrs = re.sub(r'\s+t="[^"]*"', "", m.group(1))
                xml = xml[:m.start()] + f'<c r="{cell}"{attrs} t="e">{m.group(2)}<v>{err}</v></c>' + xml[m.end():]
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def _expected():
    """Every active rule id must fire on the fixture: each lint rule has a planted defect."""
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    from xl.lint import RULE_IDS
    return list(RULE_IDS)


# Cells that must be named in a finding (regression cases for bugs fixed on 2026-09-09)
EXPECT_WHERE = [("empty-cell-coercion", "Model!B29"),     # SUMIF is not an IF guard
                ("sum-misses-adjacent", "Model!B28"),     # array formula inspected
                ("unsorted-lookup", "Ana's Sheet!D1"),    # quoted sheet name unescaped
                ("cached-error", "Model!B27"),            # real cached error, not a constant
                ("chart-series-broken", "Scratch chart3 series1.val")]


def check_fixture(report):
    """Problems: an active id that did not fire, a regression cell not named, a switched-off id
    that fired (its planted defect is still in the fixture), or an id the registry does not know."""
    from xl.lint import OFF_IDS
    got = report["summary"]["by_rule"]
    expect = _expected()
    problems = [r for r in expect if r not in got]
    wheres = {(f["rule"], f["where"]) for f in report["findings"]}
    problems += [f"{r} @ {w}" for r, w in EXPECT_WHERE if r in expect and (r, w) not in wheres]
    problems += [f"{r} fired but is switched off" for r in got if r in OFF_IDS]
    problems += [f"{r} fired but is not in RULE_IDS" for r in got
                 if r not in expect and r not in OFF_IDS and r != "rule-failed"]
    return problems, got


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = args[0] if args else "fixture_defects.xlsx"
    print(build(out))
    if "--check" in sys.argv:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
        from xl.lint import lint
        rep = lint(os.path.abspath(out), as_dict=True)
        problems, got = check_fixture(rep)
        print("rules fired:", sorted(got))
        failed = [f for f in rep["findings"] if f["rule"] == "rule-failed"]
        for f in failed:
            print("RULE FAILED:", f["where"], f["msg"])
        print("PROBLEMS:", problems)
        sys.exit(1 if (problems or failed) else 0)
