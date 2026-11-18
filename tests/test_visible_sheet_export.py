import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "frmPdfExport.txt").read_text(encoding="utf-8")
MODULE = (ROOT / "modPdfExport.txt").read_text(encoding="utf-8")


def procedure(source, name):
    lines = source.splitlines()
    for start, line in enumerate(lines):
        header = line.strip()
        kind = None
        for visibility in ["Public", "Private", "Friend"]:
            for candidate in ["Sub", "Function"]:
                if header.startswith(visibility + " " + candidate + " " + name + "("):
                    kind = candidate
                    break
            if kind:
                break
        if not kind:
            continue
        for end in range(start + 1, len(lines)):
            if lines[end].strip() == "End " + kind:
                return "\n".join(lines[start:end + 1])
    raise AssertionError("Procedure not found: " + name)


class VisibleSheetExportTests(unittest.TestCase):
    def test_sheet_list_and_refresh_include_visible_worksheets_only(self):
        loader = procedure(FORM, "LoadSheets")
        self.assertIn("lstSheets.Clear", loader)
        self.assertIn("If ws.Visible = xlSheetVisible Then", loader)
        self.assertIn("lstSheets.AddItem ws.Name", loader)
        self.assertIn("End If", loader)
        refresh = procedure(FORM, "cmdRefresh_Click")
        self.assertIn("EnsureTargetWorkbookAvailable()", refresh)
        self.assertIn("Call LoadSheets", refresh)

    def test_value_output_values_selected_sheets_before_deleting_other_sheets(self):
        export = procedure(MODULE, "ExportWorkbookCopyAsXlsx")
        self.assertLess(export.index("sourceWb.SaveCopyAs tempPath"), export.index("Workbooks.Open"))
        self.assertLess(export.index("Call SetSheetVisibleSafely(targetWs, xlSheetVisible)"), export.index("If keepFormulas Then"))
        self.assertLess(export.index("Call ConvertSheetFormulasToValuesSafely(targetWs)"), export.index("Call DeleteUnselectedSheets(wbTemp, sheetNames)"))
        keep_branch = export.split("If keepFormulas Then", 1)[1].split("\n    Else\n", 1)[0]
        value_branch = export.split("If keepFormulas Then", 1)[1].split("\n    Else\n", 1)[1]
        self.assertIn("xlSheetVeryHidden", keep_branch)
        self.assertIn("正規化外部参照_関数保持", keep_branch)
        self.assertNotIn("DeleteUnselectedSheets", keep_branch)
        self.assertIn("ConvertSheetFormulasToValuesSafely(targetWs)", value_branch)
        self.assertIn("DeleteUnselectedSheets(wbTemp, sheetNames)", value_branch)

    def test_value_export_deletes_every_unselected_sheet_object_only_from_copy(self):
        export = procedure(MODULE, "ExportWorkbookCopyAsXlsx")
        delete = procedure(MODULE, "DeleteUnselectedSheets")
        self.assertIn("Call DeleteUnselectedSheets(wbTemp, sheetNames)", export)
        self.assertIn("For i = wbTemp.Sheets.Count To 1 Step -1", delete)
        self.assertIn("TypeOf sheetItem Is Worksheet", delete)
        self.assertIn("IsSheetNameSelected(targetWs.Name, sheetNames)", delete)
        self.assertIn("If Not keepSheet Then", delete)
        self.assertLess(delete.index("SetSheetVisibleSafely(sheetItem, xlSheetVisible)"), delete.index("sheetItem.Delete"))
        self.assertIn("sheetItem.Delete", delete)

    def test_formula_export_keeps_selected_visible_and_other_sheets_very_hidden(self):
        export = procedure(MODULE, "ExportWorkbookCopyAsXlsx")
        self.assertIn("If keepFormulas Then", export)
        self.assertIn("For Each sheetItem In wbTemp.Sheets", export)
        self.assertIn("IsSheetNameSelected(sheetItem.Name, sheetNames)", export)
        self.assertIn("SetSheetVisibleSafely(sheetItem, xlSheetVisible)", export)
        self.assertIn("SetSheetVisibleSafely(sheetItem, xlSheetVeryHidden)", export)
        self.assertIn("正規化外部参照_関数保持(wbTemp, sourceWb.Name)", export)

    def test_workbook_target_safety_code_remains_in_place(self):
        selector = procedure(MODULE, "GetTargetWorkbook")
        self.assertIn("Application.ActiveWorkbook", selector)
        self.assertNotIn("For Each", selector)
        validator = procedure(MODULE, "PdfExport_IsWorkbookAvailable")
        self.assertIn("openWb Is ThisWorkbook", validator)
        self.assertIn("openWb.IsAddin", validator)


if __name__ == "__main__":
    unittest.main()
