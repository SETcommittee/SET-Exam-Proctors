#!/usr/bin/env python3
"""
Replace the VBA inside the existing .xlsm, leaving every sheet's data intact.

make_xlsm.py rebuilds the workbook from the .xlsx and would wipe the email
addresses and wording entered by hand, so code changes go through here.
"""

import os
import re
import sys

import win32com.client as win32

FOLDER = (
    r"C:\Users\Faris.Alsalem\OneDrive - AL-Hussien bin Abdullah Technical University"
    r"\Faris HTU\Exam Committee"
)
XLSM = os.path.join(FOLDER, "Proctor Schedule by Date.xlsm")

HERE = os.path.dirname(os.path.abspath(__file__))
BAS_MODULE = os.path.join(HERE, "vba_reminders.bas")
BAS_REM_SHEET = os.path.join(HERE, "vba_sheet.bas")
BAS_ASSIGN_SHEET = os.path.join(HERE, "vba_assign_sheet.bas")

XL_STD_MODULE = 1
SEND_FROM_DEFAULT = "set.exam.committee@htu.edu.jo"

# Same layout on both list sheets, so one builder can serve them.
HEADERS = ["Date", "Day", "Time", "Course", "Room", "Coordinator", "Proctors",
           "To (proctors)", "Cc (coordinator)", "No address for", "When",
           "Send", "Sent", "Due"]
WIDTHS = [12, 10, 14, 40, 26, 30, 44, 34, 26, 28, 11, 9, 20, 10]

REMINDER_BUTTONS = [
    ("Refresh list", "RefreshReminders"),
    ("Send due reminders", "SendDueReminders"),
    ("Send for selected row", "SendSelectedReminder"),
    ("Preview in Outlook", "PreviewSelectedReminder"),
    ("Clear sent flag", "ClearSentFlag"),
    ("Check Outlook", "TestOutlookConnection"),
    ("Load send-from accounts", "RefreshSendAccounts"),
    ("Unfreeze Excel", "RestoreExcel"),
]

ASSIGN_BUTTONS = [
    ("Refresh list", "RefreshAssignments"),
    ("Send all not yet sent", "SendDueAssignments"),
    ("Send for selected row", "SendSelectedAssignment"),
    ("Preview in Outlook", "PreviewSelectedAssignment"),
    ("Clear sent flag", "ClearSentFlag"),
    ("Check Outlook", "TestOutlookConnection"),
    ("Unfreeze Excel", "RestoreExcel"),
]

PLACEHOLDERS = [
    ("{course}", "Engineering Math"),
    ("{day}", "Monday"),
    ("{date}", "7 Sep 2026"),
    ("{longdate}", "7 September 2026"),
    ("{time}", "11:00 - 13:30"),
    ("{room}", "N-301, N-303, N-304, N-306"),
    ("{rooms}", "the same room list"),
    ("{students}", "how many students sit the exam"),
    ("{coordinator}", "Rami Hammad"),
    ("{proctorrooms}", "each proctor with their own room, one per line"),
    ("{proctors}", "the same names, without rooms"),
    ("{proctorlist}", "the names on one line, comma separated"),
    ("{count}", "how many proctors"),
]


def sheet_named(wb, name):
    for sh in wb.Worksheets:
        if sh.Name == name:
            return sh
    return None


def style_header(ws, row, headers, widths):
    for i, head in enumerate(headers, start=1):
        cell = ws.Cells(row, i)
        cell.Value = head
        cell.Font.Bold = True
        cell.Interior.Color = 0xF6F1E1
        cell.Borders(9).LineStyle = 1          # xlEdgeBottom
    for i, w in enumerate(widths, start=1):
        ws.Columns(i).ColumnWidth = w


def preview_block(ws, title, note):
    """The TO / CC / SUBJECT / BODY box both list sheets carry at the top."""
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 15
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = note
    ws.Range("A2").Font.Color = 0x8A7266

    for cell, label in (("A4", "TO"), ("A5", "CC"), ("A6", "SUBJECT"), ("A7", "BODY")):
        ws.Range(cell).Value = label
        ws.Range(cell).Font.Bold = True
        ws.Range(cell).Font.Color = 0x8A7266
        ws.Range(cell).VerticalAlignment = -4160        # xlTop

    for rng in ("B4:F4", "B5:F5", "B6:F6", "B7:F7"):
        try:
            ws.Range(rng).UnMerge()
        except Exception:
            pass
        ws.Range(rng).Merge()
        ws.Range(rng).Interior.Color = 0xFAF8F5
        ws.Range(rng).VerticalAlignment = -4160
        ws.Range(rng).WrapText = True

    ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 16
    ws.Rows(6).RowHeight = 30
    ws.Rows(7).RowHeight = 190
    ws.Rows(8).RowHeight = 8


def place_buttons(ws, buttons):
    for b in list(ws.Buttons()):
        b.Delete()
    x = ws.Range("G4").Left + 8
    y = ws.Range("G4").Top
    for caption, macro in buttons:
        btn = ws.Buttons().Add(x, y, 170, 26)
        btn.Caption = caption
        btn.OnAction = macro
        btn.Font.Size = 10
        y += 29


def build_reminders_sheet(wb):
    """Add the two new columns; leave everything else as it is."""
    ws = sheet_named(wb, "Reminders")
    if ws is None:
        print("  Reminders sheet missing - skipped")
        return
    style_header(ws, 9, HEADERS, WIDTHS)
    ws.Range("A2").Value = (
        "Open Outlook first. Click a row to preview it, then double-click its Send "
        "cell. 'Send due reminders' does every exam within two days that has not "
        "had one yet."
    )
    place_buttons(ws, REMINDER_BUTTONS)
    print("  Reminders: columns Sent + Due added, buttons rebuilt")


def build_assignments_sheet(wb):
    ws = sheet_named(wb, "Assignments")
    fresh = ws is None
    if fresh:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Assignments"

    preview_block(
        ws,
        "Assignment emails",
        "Sent once, when an exam has been staffed. The reminder two days before "
        "is on the Reminders sheet. Send-from and always-Cc are taken from the "
        "Reminders sheet.",
    )
    style_header(ws, 9, HEADERS, WIDTHS)
    place_buttons(ws, ASSIGN_BUTTONS)

    try:
        wb.Application.ActiveWindow.FreezePanes = False
        ws.Activate()
        ws.Range("A10").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    print("  Assignments sheet %s" % ("created" if fresh else "refreshed"))
    return ws


def build_message_sheet(wb, excel):
    """Add the assignment wording below the reminder wording."""
    ws = sheet_named(wb, "Message")
    if ws is None:
        print("  Message sheet missing - skipped")
        return

    ws.Range("A2").Value = (
        "Two emails, two sets of wording. Anything in curly brackets is swapped "
        "for that exam's details. Leave a cell empty to fall back to the original."
    )

    ws.Range("A4").Value = "REMINDER SUBJECT"
    ws.Range("A6").Value = "REMINDER BODY"
    ws.Range("A9").Value = "ASSIGNMENT SUBJECT"
    ws.Range("A11").Value = "ASSIGNMENT BODY"
    for cell in ("A4", "A6", "A9", "A11"):
        ws.Range(cell).Font.Bold = True
        ws.Range(cell).Font.Color = 0x8A7266
        ws.Range(cell).VerticalAlignment = -4160

    for rng in ("B9:F9", "B11:F11"):
        try:
            ws.Range(rng).UnMerge()
        except Exception:
            pass
        ws.Range(rng).Merge()
        ws.Range(rng).WrapText = True
        ws.Range(rng).VerticalAlignment = -4160
        ws.Range(rng).Interior.Color = 0xF2FAF2
        ws.Range(rng).Borders.LineStyle = 1

    ws.Rows(9).RowHeight = 32
    ws.Rows(11).RowHeight = 300
    ws.Columns(1).ColumnWidth = 20

    if not str(ws.Range("B9").Value or "").strip():
        ws.Range("B9").Value = excel.Run("DefaultAssignSubject")
    if not str(ws.Range("B11").Value or "").strip():
        ws.Range("B11").Value = excel.Run("DefaultAssignBody")

    # Placeholder reference, refreshed in place.
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

    for b in list(ws.Buttons()):
        b.Delete()
    btn = ws.Buttons().Add(ws.Range("H19").Left, ws.Range("H19").Top, 190, 26)
    btn.Caption = "Restore original wording"
    btn.OnAction = "ResetMessageTemplate"
    btn.Font.Size = 10

    print("  Message: assignment subject + body added")


def inject_vba(wb, proj):
    for comp in list(proj.VBComponents):
        if comp.Name == "mReminders":
            proj.VBComponents.Remove(comp)
            print("  removed old mReminders")

    with open(BAS_MODULE, encoding="utf-8") as fh:
        code = fh.read()
    comp = proj.VBComponents.Add(XL_STD_MODULE)
    comp.Name = "mReminders"
    comp.CodeModule.AddFromString(code)
    print("  mReminders added (%d lines)" % comp.CodeModule.CountOfLines)

    for sheet_name, bas in (("Reminders", BAS_REM_SHEET),
                            ("Assignments", BAS_ASSIGN_SHEET)):
        ws = sheet_named(wb, sheet_name)
        if ws is None:
            continue
        cm = proj.VBComponents(ws.CodeName).CodeModule
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)
        with open(bas, encoding="utf-8") as fh:
            cm.AddFromString(fh.read())
        print("  %s sheet events replaced (%d lines)" % (sheet_name, cm.CountOfLines))


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

        # Sheets first: the VBA references them by name, and the Assignments
        # sheet must exist before its event code can be attached.
        build_reminders_sheet(wb)
        build_assignments_sheet(wb)

        inject_vba(wb, wb.VBProject)

        # Needs the module present, since it calls DefaultAssignSubject/Body.
        build_message_sheet(wb, excel)

        wb.Save()
        wb.Close(SaveChanges=False)
        print("saved %s" % XLSM)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
