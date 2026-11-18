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
                return chr(10).join(lines[start:end + 1])
    raise AssertionError("Procedure not found: " + name)


class WorkbookTargetSafetyTests(unittest.TestCase):
    def test_target_selection_never_guesses_another_open_workbook(self):
        selector = procedure(MODULE, "GetTargetWorkbook")
        validator = procedure(MODULE, "PdfExport_IsWorkbookAvailable")
        self.assertIn("Application.ActiveWorkbook", selector)
        self.assertIn("PdfExport_IsWorkbookAvailable(wb)", selector)
        self.assertNotIn("For Each", selector)
        self.assertIn("For Each openWb In Application.Workbooks", validator)
        self.assertIn("openWb Is ThisWorkbook", validator)
        self.assertIn("openWb.IsAddin", validator)

    def test_form_initialization_defers_sheet_loading_until_target_is_set(self):
        initialize = procedure(FORM, "UserForm_Initialize")
        setter = procedure(FORM, "SetTargetWorkbook")
        self.assertNotIn("ActiveWorkbook", initialize)
        self.assertNotIn("LoadSheets", initialize)
        self.assertNotIn("InitializePresetControls", initialize)
        self.assertLess(setter.index("PdfExport_IsWorkbookAvailable"), setter.index("Set mTargetWb = wb"))
        self.assertLess(setter.index("Set mTargetWb = wb"), setter.index("Call LoadSheets"))
        self.assertLess(setter.index("Call LoadSheets"), setter.index("Call InitializePresetControls"))

    def test_launch_captures_workbook_before_form_creation_and_show(self):
        launch = procedure(MODULE, "PDF_エクスポート_フォーム表示")
        self.assertLess(launch.index("Set targetWb = GetTargetWorkbook()"), launch.index("Set f = New frmPdfExport"))
        self.assertLess(launch.index("Set f = New frmPdfExport"), launch.index("f.SetTargetWorkbook(targetWb)"))
        self.assertLess(launch.index("f.SetTargetWorkbook(targetWb)"), launch.index("f.Show vbModeless"))

    def test_export_print_refresh_and_presets_revalidate_the_captured_workbook(self):
        for name in ["cmdExport_Click", "cmdPrint_Click", "cmdRefresh_Click"]:
            block = procedure(FORM, name)
            self.assertIn("EnsureTargetWorkbookAvailable()", block, name)
        export = procedure(FORM, "cmdExport_Click")
        self.assertGreaterEqual(export.count("mTargetWb"), 4)
        printing = procedure(FORM, "cmdPrint_Click")
        self.assertIn("mTargetWb", printing)
        presets = procedure(FORM, "GetOrCreatePresetSheet")
        self.assertIn("EnsureTargetWorkbookAvailable", presets)
        self.assertIn("mTargetWb.Worksheets", presets)

    def test_all_output_paths_validate_and_keep_explicit_target(self):
        names = [
            "PDF_出力実行",
            "Excel_MultiSheetExport",
            "Excel_出力実行",
            "Excel_出力実行_関数保持",
            "Excel_複数シート出力実行_関数保持",
        ]
        for name in names:
            block = procedure(MODULE, name)
            self.assertIn("GetTargetWorkbook()", block, name)
            self.assertIn("PdfExport_IsWorkbookAvailable(wb)", block, name)
        lookup = procedure(MODULE, "シート取得")
        self.assertIn("GetTargetWorkbook()", lookup)
        self.assertIn("PdfExport_IsWorkbookAvailable(wb)", lookup)
        self.assertNotIn("Set wb = ActiveWorkbook", MODULE)

    def test_form_displays_the_selected_workbook_name(self):
        setter = procedure(FORM, "SetTargetWorkbook")
        self.assertIn("Me.Caption", setter)
        self.assertIn("mTargetWb.Name", setter)


if __name__ == "__main__":
    unittest.main()
