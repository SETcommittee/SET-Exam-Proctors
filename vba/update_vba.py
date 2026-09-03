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

# The committee mailbox, which is what reminders should go out as.
SEND_FROM_DEFAULT = "set.exam.committee@htu.edu.jo"

ALL_BUTTONS = [
    ("Refresh list", "RefreshReminders"),
    ("Send for selected row", "SendSelectedReminder"),
    ("Preview in Outlook", "PreviewSelectedReminder"),
    ("Send all upcoming", "SendAllUpcoming"),
    ("Check Outlook", "TestOutlookConnection"),
    ("Load send-from accounts", "RefreshSendAccounts"),
    ("Unfreeze Excel", "RestoreExcel"),
]


PLACEHOLDERS = [
    ("{course}", "Engineering Math"),
    ("{day}", "Monday"),
    ("{date}", "7 Sep 2026"),
    ("{longdate}", "7 September 2026"),
    ("{time}", "11:00 - 13:30"),
    ("{room}", "N-301, N-303, N-304, N-306"),
    ("{coordinator}", "Rami Hammad"),
    ("{proctors}", "a numbered list, one proctor per line"),
    ("{proctorlist}", "the same names on one line, comma separated"),
    ("{count}", "how many proctors"),
]


def build_message_sheet(wb, excel):
    """Create the Message sheet, without disturbing wording already written."""
    existing = None
    for sh in wb.Worksheets:
        if sh.Name == "Message":
            existing = sh
            break

    fresh = existing is None
    ws = existing or wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    if fresh:
        ws.Name = "Message"

    ws.Range("A1").Value = "Reminder wording"
    ws.Range("A1").Font.Size = 15
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = (
        "Edit the subject and body below like any other cell. Anything in curly "
        "brackets is swapped for that exam's details when the email is built. "
        "Leave a cell empty to fall back to the original wording."
    )
    ws.Range("A2").Font.Color = 0x8A7266

    ws.Range("A4").Value = "SUBJECT"
    ws.Range("A6").Value = "BODY"
    for cell in ("A4", "A6"):
        ws.Range(cell).Font.Bold = True
        ws.Range(cell).Font.Color = 0x8A7266
        ws.Range(cell).VerticalAlignment = -4160

    for rng in ("B4:F4", "B6:F6"):
        try:
            ws.Range(rng).UnMerge()
        except Exception:
            pass
        ws.Range(rng).Merge()
        ws.Range(rng).WrapText = True
        ws.Range(rng).VerticalAlignment = -4160
        ws.Range(rng).Interior.Color = 0xF2FAF2
        ws.Range(rng).Borders.LineStyle = 1

    ws.Rows(4).RowHeight = 32
    ws.Rows(6).RowHeight = 300
    for i, w in enumerate([12, 18, 18, 18, 18, 18], start=1):
        ws.Columns(i).ColumnWidth = w

    # Placeholder reference, to the right of the body box.
    ws.Range("H3").Value = "You can use:"
    ws.Range("H3").Font.Bold = True
    for i, (token, meaning) in enumerate(PLACEHOLDERS, start=4):
        ws.Cells(i, 8).Value = token
        ws.Cells(i, 8).Font.Name = "Consolas"
        ws.Cells(i, 8).Font.Color = 0x0E3D40
        ws.Cells(i, 9).Value = meaning
        ws.Cells(i, 9).Font.Color = 0x8A7266
    ws.Columns(8).ColumnWidth = 16
    ws.Columns(9).ColumnWidth = 46

    if fresh or not str(ws.Range("B4").Value or "").strip():
        ws.Range("B4").Value = excel.Run("DefaultSubject")
    if fresh or not str(ws.Range("B6").Value or "").strip():
        ws.Range("B6").Value = excel.Run("DefaultBody")

    for b in list(ws.Buttons()):
        b.Delete()
    btn = ws.Buttons().Add(ws.Range("H16").Left, ws.Range("H16").Top, 170, 26)
    btn.Caption = "Restore original wording"
    btn.OnAction = "ResetMessageTemplate"
    btn.Font.Size = 10

    print("Message sheet %s" % ("created" if fresh else "refreshed (wording kept)"))


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

        # Settings row: which account sends, and who is always copied.
        for cell, label in (("A3", "SEND FROM"), ("D3", "ALWAYS CC")):
            ws.Range(cell).Value = label
            ws.Range(cell).Font.Bold = True
            ws.Range(cell).Font.Color = 0x8A7266
            ws.Range(cell).HorizontalAlignment = -4152  # xlRight

        for rng in ("B3:C3", "E3:F3"):
            try:
                ws.Range(rng).UnMerge()
            except Exception:
                pass
            ws.Range(rng).Merge()
            ws.Range(rng).Interior.Color = 0xF2FAF2
            ws.Range(rng).Borders.LineStyle = 1
            ws.Range(rng).HorizontalAlignment = -4131  # xlLeft

        if not str(ws.Range("B3").Value or "").strip():
            ws.Range("B3").Value = SEND_FROM_DEFAULT
        ws.Range("E3").Value = str(ws.Range("E3").Value or "")
        ws.Rows(3).RowHeight = 19

        build_message_sheet(wb, excel)

        wb.Save()
        wb.Close(SaveChanges=False)
        print("saved %s" % XLSM)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
