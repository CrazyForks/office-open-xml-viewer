"""Generate the (gitignored) xlsx fixture for validation-dropdown.spec.ts.

Exercises the four list data-validation shapes the viewer's dropdown panel
resolves (ECMA-376 §18.3.1.32). The spec skips when the fixture is absent.

Usage: python3 tests/visual/gen-validation-fixture.py  (from packages/xlsx,
needs openpyxl) — writes public/validation-fixture.xlsx."""
import sys
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()
main = wb.active
main.title = "Main"

# A reference column on the Main sheet (for the same-sheet range case).
main["E1"] = "North"
main["E2"] = "South"
main["E3"] = "East"
main["E4"] = "West"

# A second sheet holding values for the cross-sheet range case.
lists = wb.create_sheet("Lists")
lists["A1"] = "Apple"
lists["A2"] = "Banana"
lists["A3"] = "Cherry"

# 1) Inline quoted comma list on B2.
dv_inline = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=True)
dv_inline.add("B2")
main.add_data_validation(dv_inline)

# 2) Same-sheet range reference on B4.
dv_same = DataValidation(type="list", formula1="$E$1:$E$4", allow_blank=True)
dv_same.add("B4")
main.add_data_validation(dv_same)

# 3) Cross-sheet range reference on B6.
dv_cross = DataValidation(type="list", formula1="Lists!$A$1:$A$3", allow_blank=True)
dv_cross.add("B6")
main.add_data_validation(dv_cross)

# 4) Unresolvable (defined-name) reference on B8.
wb.defined_names.add.__self__  # no-op; keep openpyxl import tidy
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names["MyNamedList"] = DefinedName("MyNamedList", attr_text="Lists!$A$1:$A$3")
dv_named = DataValidation(type="list", formula1="MyNamedList", allow_blank=True)
dv_named.add("B8")
main.add_data_validation(dv_named)

# Labels so the cells are easy to find visually.
main["A2"] = "inline:"
main["A4"] = "same-sheet:"
main["A6"] = "cross-sheet:"
main["A8"] = "named (unresolved):"

import os
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(__file__), "..", "..", "public", "validation-fixture.xlsx"
)
wb.save(out)
print("wrote", out)
