import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "frmPdfExport.txt").read_text(encoding="utf-8")


def procedure(source, name):
    lines = source.splitlines()
    for start, line in enumerate(lines):
        header = line.strip()
        for kind in ("Sub", "Function"):
            if header.startswith(("Public ", "Private ", "Friend ")) and kind + " " + name + "(" in header:
                for end in range(start + 1, len(lines)):
                    if lines[end].strip() == "End " + kind:
                        return "\n".join(lines[start:end + 1])
    raise AssertionError("Procedure not found: " + name)


class PrintModeTests(unittest.TestCase):
    def test_form_requires_and_initializes_print_mode_controls(self):
        initialize = procedure(FORM, "UserForm_Initialize")
        self.assertIn('ControlExists(Me, "optPrintSeparate")', initialize)
        self.assertIn('ControlExists(Me, "optPrintCombined")', initialize)
        self.assertIn('Controls("optPrintSeparate").GroupName = "grpPrintMode"', initialize)
        self.assertIn('Controls("optPrintCombined").GroupName = "grpPrintMode"', initialize)
        self.assertIn('Controls("optPrintSeparate").Value = True', initialize)
        self.assertIn('Controls("optPrintCombined").Value = False', initialize)

    def test_separate_mode_prints_each_listed_sheet_independently(self):
        printing = procedure(FORM, "cmdPrint_Click")
        separate = printing.split("ElseIf separateMode Then", 1)[1].split("    Else\n        printWb.Activate", 1)[0]
        self.assertIn("GetSelectedSheetName(i)", printing)
        self.assertIn("names.Add sheetName", printing)
        self.assertIn("ws.PrintOut", separate)
        self.assertNotIn("SelectedSheets.PrintOut", separate)
        self.assertIn("totalJobs = names.Count", printing)

    def test_combined_mode_checks_target_and_uses_legacy_group_print(self):
        printing = procedure(FORM, "cmdPrint_Click")
        self.assertIn("printWb.Activate", printing)
        self.assertIn("Application.ActiveWorkbook Is printWb", printing)
        self.assertIn("PrintSelectionMatchesTarget(printWb, names)", printing)
        self.assertIn("Application.ActiveWindow.SelectedSheets.PrintOut", printing)

    def test_single_sheet_page_range_is_preserved(self):
        printing = procedure(FORM, "cmdPrint_Click")
        self.assertIn("ws.PrintOut From:=fromPage, To:=toPage", printing)
        self.assertIn("If names.Count = 1 Then", printing)

    def test_print_errors_do_not_retry_and_report_partial_jobs(self):
        printing = procedure(FORM, "cmdPrint_Click")
        self.assertIn("PrintOut正常終了", printing)
        self.assertIn("printJobs < totalJobs", printing)
        self.assertIn("再送信していません", printing)

    def test_fill_and_application_state_restoration_are_verified(self):
        printing = procedure(FORM, "cmdPrint_Click")
        self.assertIn("CaptureInteriors(printWb, names)", printing)
        self.assertIn("ApplyNoFill printWb, captured", printing)
        self.assertIn("RestoreAndVerifyPrintInteriors(printWb, captured)", printing)
        self.assertIn("Application.ScreenUpdating = savedScreenUpdating", printing)
        self.assertIn("Application.EnableEvents = savedEnableEvents", printing)
        self.assertIn("Application.DisplayAlerts = savedDisplayAlerts", printing)
        self.assertIn("Private Function PrintInteriorsMatchSnapshot", FORM)


if __name__ == "__main__":
    unittest.main()
