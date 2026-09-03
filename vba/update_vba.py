#!/usr/bin/env python3
"""
Replace the VBA inside the existing .xlsm, leaving every sheet untouched.

make_xlsm.py rebuilds the workbook from the .xlsx and would wipe the email
addresses that have been entered by hand, so code changes go through here
instead.
"""

import os
import sys

import win32com.client as win32

FOLDER = (
    r"C:\Users\Faris.Alsalem\OneDrive - AL-Hussien bin Abdullah Technical University"
    r"\Faris HTU\Exam Committee"
)
XLSM = os.path.join(FOLDER, "Proctor Schedule by Date.xlsm")

HERE = os.path.dirname(os.path.abspath(__file__))
BAS_MODULE = os.path.join(HERE, "vba_reminders.bas")
BAS_SHEET = os.path.join(HERE, "vba_sheet.bas")

XL_STD_MODULE = 1

ALL_BUTTONS = [
    ("Refresh list", "RefreshReminders"),
    ("Send for selected row", "SendSelectedReminder"),
    ("Preview in Outlook", "PreviewSelectedReminder"),
    ("Send all upcoming", "SendAllUpcoming"),
    ("Check Outlook", "TestOutlookConnection"),
    ("Load send-from accounts", "RefreshSendAccounts"),
    ("Unfreeze Excel", "RestoreExcel"),
]


def main():
    if not os.path.exists(XLSM):
        sys.exit("Not found: %s" % XLSM)
    try:
        open(XLSM, "r+b").close()
    except PermissionError:
        sys.exit("The workbook is open in Excel. Close it and run this again.")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    try:
        wb = excel.Workbooks.Open(XLSM)
        proj = wb.VBProject

        # --- replace the standard module -------------------------------
        for comp in list(proj.VBComponents):
            if comp.Name == "mReminders":
                proj.VBComponents.Remove(comp)
                print("removed old mReminders")

        with open(BAS_MODULE, encoding="utf-8") as fh:
            code = fh.read()
        comp = proj.VBComponents.Add(XL_STD_MODULE)
        comp.Name = "mReminders"
        comp.CodeModule.AddFromString(code)
        print("added mReminders (%d lines)" % comp.CodeModule.CountOfLines)

        # --- replace the sheet's event code ----------------------------
        ws = wb.Worksheets("Reminders")
        cm = proj.VBComponents(ws.CodeName).CodeModule
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)
        with open(BAS_SHEET, encoding="utf-8") as fh:
            cm.AddFromString(fh.read())
        print("replaced Reminders sheet events (%d lines)" % cm.CountOfLines)

        # --- rebuild the button strip ----------------------------------
        for b in list(ws.Buttons()):
            b.Delete()
        x = ws.Range("G4").Left + 8
        y = ws.Range("G4").Top
        for caption, macro in ALL_BUTTONS:
            btn = ws.Buttons().Add(x, y, 170, 26)
            btn.Caption = caption
            btn.OnAction = macro
            btn.Font.Size = 10
            y += 30
        print("buttons: %s" % ", ".join(c for c, _ in ALL_BUTTONS))

        # a plain-language note where the user will see it
        ws.Range("A2").Value = (
            "Open Outlook first, then click a row to preview and double-click its "
            "Send cell. If anything seems stuck, press Unfreeze Excel."
        )

        # "Send from" control
        ws.Range("A3").Value = "SEND FROM"
        ws.Range("A3").Font.Bold = True
        ws.Range("A3").Font.Color = 0x8A7266
        if not str(ws.Range("B3").Value or "").strip():
            ws.Range("B3").Value = "(default account)"
        ws.Range("B3").Interior.Color = 0xF2FAF2
        ws.Range("B3").Borders.LineStyle = 1
        ws.Rows(3).RowHeight = 18

        wb.Save()
        wb.Close(SaveChanges=False)
        print("saved %s" % XLSM)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
