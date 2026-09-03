Option Explicit

' ============================================================================
'  Exam reminder emails
'  Built for the SET Exam Committee. Reads the Exam sheet, resolves each
'  person's address from the Emails sheet, and sends through Outlook.
' ============================================================================

' Set to False to send silently with no confirmation box. Leave True while you
' are still checking that the addresses on the Emails sheet are correct.
Public Const CONFIRM_BEFORE_SEND As Boolean = True

' "SEND"    - send the message immediately (after the confirmation above)
' "DISPLAY" - open the message in Outlook so you can read it and press Send
Public Const SEND_MODE As String = "SEND"

Private Const SH_EXAM As String = "Exam"
Private Const SH_MAIL As String = "Emails"
Private Const SH_REM As String = "Reminders"

Private Const FIRST_DATA_ROW As Long = 10   ' first exam row on Reminders
Private Const EXAM_FIRST_ROW As Long = 6    ' first exam row on Exam sheet

' Columns on the Reminders sheet
Private Const C_DATE As Long = 1
Private Const C_DAY As Long = 2
Private Const C_TIME As Long = 3
Private Const C_COURSE As Long = 4
Private Const C_ROOM As Long = 5
Private Const C_COORD As Long = 6
Private Const C_PROCS As Long = 7
Private Const C_TO As Long = 8          ' proctors
Private Const C_CC As Long = 9          ' coordinator
Private Const C_MISSING As Long = 10
Private Const C_STATUS As Long = 11
Public Const C_SEND As Long = 12
Private Const C_SENTLOG As Long = 13


' ---------------------------------------------------------------------------
'  Look up one person's email address on the Emails sheet
' ---------------------------------------------------------------------------
Public Function EmailFor(ByVal personName As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, nm As String
    EmailFor = ""
    personName = Trim$(personName)
    If Len(personName) = 0 Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_MAIL)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 4 To lastRow
        nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(nm, personName, vbTextCompare) = 0 Then
            EmailFor = Trim$(CStr(ws.Cells(r, 3).Value))
            Exit Function
        End If
    Next r
End Function


' ---------------------------------------------------------------------------
'  Rebuild the Reminders list from the Exam sheet
' ---------------------------------------------------------------------------
' Button target: rebuild, then report how many rows were written.
Public Sub RefreshReminders()
    Dim n As Long
    n = RefreshRemindersCore()
    MsgBox "Reminder list rebuilt." & vbCrLf & vbCrLf & _
           n & " exam(s) listed.", vbInformation, "Reminders"
End Sub


' The work itself, with no dialogs, so it can also be driven from outside.
Public Function RefreshRemindersCore() As Long
    Dim wsE As Worksheet, wsR As Worksheet
    Dim rE As Long, rR As Long, lastE As Long, c As Long
    Dim course As String, coord As String, procs As String
    Dim toList As String, ccList As String, missing As String
    Dim addr As String, nm As String
    Dim dt As Variant, startT As String, endT As String

    Set wsE = ThisWorkbook.Worksheets(SH_EXAM)
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' clear old rows but keep the header block
    If wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row >= FIRST_DATA_ROW Then
        wsR.Range(wsR.Cells(FIRST_DATA_ROW, 1), _
                  wsR.Cells(wsR.Rows.Count, C_SENTLOG)).ClearContents
        wsR.Range(wsR.Cells(FIRST_DATA_ROW, 1), _
                  wsR.Cells(wsR.Rows.Count, C_SENTLOG)).Interior.Pattern = xlNone
    End If

    lastE = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    rR = FIRST_DATA_ROW

    For rE = EXAM_FIRST_ROW To lastE
        course = Trim$(CStr(wsE.Cells(rE, 1).Value))
        If Len(course) > 0 And UCase$(course) <> "TOTAL" Then
            dt = wsE.Cells(rE, 4).Value          ' D = Date
            coord = Trim$(CStr(wsE.Cells(rE, 2).Value))

            startT = FormatTime(wsE.Cells(rE, 6).Value)
            endT = FormatTime(wsE.Cells(rE, 7).Value)

            ' gather assigned proctors from K..T
            procs = ""
            toList = ""
            ccList = ""
            missing = ""
            For c = 11 To 20
                nm = Trim$(CStr(wsE.Cells(rE, c).Value))
                If Len(nm) > 0 Then
                    If Len(procs) > 0 Then procs = procs & ", "
                    procs = procs & nm
                    addr = EmailFor(nm)
                    If Len(addr) > 0 Then
                        If InStr(1, ";" & toList, ";" & addr, vbTextCompare) = 0 Then
                            If Len(toList) > 0 Then toList = toList & "; "
                            toList = toList & addr
                        End If
                    Else
                        If Len(missing) > 0 Then missing = missing & ", "
                        missing = missing & nm
                    End If
                End If
            Next c

            ' The coordinator is copied in, not addressed directly, and is not
            ' named in the body - the message is written to the proctors.
            If Len(coord) > 0 Then
                addr = EmailFor(coord)
                If Len(addr) > 0 Then
                    ' don't Cc someone who is already a proctor on this exam
                    If InStr(1, ";" & toList, ";" & addr, vbTextCompare) = 0 Then
                        ccList = addr
                    End If
                Else
                    If Len(missing) > 0 Then missing = missing & ", "
                    missing = missing & coord & " (coordinator)"
                End If
            End If

            wsR.Cells(rR, C_DATE).Value = dt
            wsR.Cells(rR, C_DATE).NumberFormat = "dd mmm yyyy"
            wsR.Cells(rR, C_DAY).Value = wsE.Cells(rE, 5).Value
            wsR.Cells(rR, C_TIME).Value = startT & IIf(Len(endT) > 0, " - " & endT, "")
            wsR.Cells(rR, C_COURSE).Value = course
            wsR.Cells(rR, C_ROOM).Value = _
                IIf(Len(Trim$(CStr(wsE.Cells(rE, 8).Value))) = 0, "TBC", wsE.Cells(rE, 8).Value)
            wsR.Cells(rR, C_COORD).Value = coord
            wsR.Cells(rR, C_PROCS).Value = procs
            wsR.Cells(rR, C_TO).Value = toList
            wsR.Cells(rR, C_CC).Value = ccList
            wsR.Cells(rR, C_MISSING).Value = missing
            wsR.Cells(rR, C_STATUS).Value = StatusFor(dt)
            wsR.Cells(rR, C_SEND).Value = "Send"

            ' highlight anything that cannot be fully delivered
            If Len(missing) > 0 Then
                wsR.Cells(rR, C_MISSING).Interior.Color = RGB(250, 229, 227)
                wsR.Cells(rR, C_MISSING).Font.Color = RGB(166, 42, 34)
            End If

            wsR.Cells(rR, C_SEND).Interior.Color = RGB(225, 239, 239)
            wsR.Cells(rR, C_SEND).Font.Color = RGB(14, 61, 64)
            wsR.Cells(rR, C_SEND).Font.Bold = True
            wsR.Cells(rR, C_SEND).HorizontalAlignment = xlCenter

            If StatusFor(dt) = "Past" Then
                wsR.Range(wsR.Cells(rR, 1), wsR.Cells(rR, C_PROCS)).Font.Color = RGB(150, 158, 170)
            End If

            rR = rR + 1
        End If
    Next rE

    wsR.Rows(FIRST_DATA_ROW & ":" & (rR - 1)).RowHeight = 16

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    RefreshRemindersCore = rR - FIRST_DATA_ROW
End Function


' The Exam sheet holds times two different ways: some cells are text ("09:30")
' and some are real Excel times, which arrive here as a fraction of a day
' (0.4166... = 10:00). Handle both, or the numeric ones print as decimals.
Private Function FormatTime(ByVal v As Variant) As String
    Dim s As String
    FormatTime = ""
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbDecimal, vbCurrency, vbInteger, vbLong
            FormatTime = Format$(CDate(CDbl(v)), "hh:mm")
        Case vbDate
            FormatTime = Format$(v, "hh:mm")
        Case Else
            s = Trim$(CStr(v))
            If Len(s) = 0 Then Exit Function
            If IsDate(s) Then
                FormatTime = Format$(CDate(s), "hh:mm")
            Else
                FormatTime = s
            End If
    End Select
End Function


Private Function StatusFor(ByVal dt As Variant) As String
    If Not IsDate(dt) Then
        StatusFor = "No date"
    ElseIf Int(CDate(dt)) < Int(Now()) Then
        StatusFor = "Past"
    ElseIf Int(CDate(dt)) = Int(Now()) Then
        StatusFor = "Today"
    Else
        StatusFor = Int(CDate(dt)) - Int(Now()) & " day(s)"
    End If
End Function


' ---------------------------------------------------------------------------
'  Build the message for one row
' ---------------------------------------------------------------------------
Public Sub BuildMessage(ByVal r As Long, ByRef toList As String, ByRef ccList As String, _
                        ByRef subj As String, ByRef body As String)
    Dim wsR As Worksheet, parts() As String, i As Long
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    toList = Trim$(CStr(wsR.Cells(r, C_TO).Value))
    ccList = Trim$(CStr(wsR.Cells(r, C_CC).Value))

    subj = "Proctoring reminder - " & wsR.Cells(r, C_COURSE).Value & _
           ", " & wsR.Cells(r, C_DAY).Value & " " & _
           Format$(wsR.Cells(r, C_DATE).Value, "d mmm yyyy") & _
           " at " & wsR.Cells(r, C_TIME).Value

    body = "Dear colleagues," & vbCrLf & vbCrLf & _
           "This is a reminder of the exam below and your proctoring duty for it." & vbCrLf & vbCrLf & _
           "Course:       " & wsR.Cells(r, C_COURSE).Value & vbCrLf & _
           "Date:         " & wsR.Cells(r, C_DAY).Value & ", " & _
                              Format$(wsR.Cells(r, C_DATE).Value, "d mmmm yyyy") & vbCrLf & _
           "Time:         " & wsR.Cells(r, C_TIME).Value & vbCrLf & _
           "Room:         " & wsR.Cells(r, C_ROOM).Value & vbCrLf & vbCrLf & _
           "Proctors:" & vbCrLf

    parts = Split(CStr(wsR.Cells(r, C_PROCS).Value), ", ")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            body = body & "  " & (i + 1) & ". " & parts(i) & vbCrLf
        End If
    Next i

    body = body & vbCrLf & _
           "Please arrive 15 minutes before the start time." & vbCrLf & vbCrLf & _
           "If you cannot attend, reply to this message as early as you can so a" & vbCrLf & _
           "replacement can be arranged." & vbCrLf & vbCrLf & _
           "Thank you," & vbCrLf & _
           "Exam Committee" & vbCrLf & _
           "School of Engineering and Technology"
End Sub


' ---------------------------------------------------------------------------
'  Send one row
' ---------------------------------------------------------------------------
Public Function SendReminderRow(ByVal r As Long, ByVal askFirst As Boolean) As Boolean
    Dim wsR As Worksheet
    Dim toList As String, ccList As String, subj As String, body As String, missing As String
    Dim ol As Object, mail As Object
    Dim answer As VbMsgBoxResult

    SendReminderRow = False
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    If Len(Trim$(CStr(wsR.Cells(r, C_COURSE).Value))) = 0 Then Exit Function

    BuildMessage r, toList, ccList, subj, body
    missing = Trim$(CStr(wsR.Cells(r, C_MISSING).Value))

    If Len(toList) = 0 Then
        MsgBox "No email addresses for this exam." & vbCrLf & vbCrLf & _
               "Add addresses on the Emails sheet, then press Refresh list.", _
               vbExclamation, "Nothing to send"
        Exit Function
    End If

    If askFirst Then
        answer = MsgBox("Send this reminder?" & vbCrLf & vbCrLf & _
                        wsR.Cells(r, C_COURSE).Value & vbCrLf & _
                        Format$(wsR.Cells(r, C_DATE).Value, "d mmm yyyy") & "  " & _
                        wsR.Cells(r, C_TIME).Value & vbCrLf & vbCrLf & _
                        "To:" & vbCrLf & Replace(toList, "; ", vbCrLf) & vbCrLf & _
                        IIf(Len(ccList) > 0, vbCrLf & "Cc:" & vbCrLf & ccList & vbCrLf, "") & _
                        IIf(Len(missing) > 0, vbCrLf & "NO ADDRESS FOR: " & missing & vbCrLf, ""), _
                        vbYesNo + vbQuestion, "Confirm reminder")
        If answer <> vbYes Then Exit Function
    End If

    On Error GoTo NoOutlook
    Set ol = GetOutlook()
    Set mail = ol.CreateItem(0)
    With mail
        .To = toList
        If Len(ccList) > 0 Then .CC = ccList
        .Subject = subj
        .body = body
        If UCase$(SEND_MODE) = "DISPLAY" Then
            .Display
        Else
            .Send
        End If
    End With
    On Error GoTo 0

    wsR.Cells(r, C_SENTLOG).Value = "Sent " & Format$(Now(), "d mmm hh:mm")
    wsR.Cells(r, C_SENTLOG).Font.Color = RGB(28, 107, 69)
    SendReminderRow = True
    Exit Function

NoOutlook:
    MsgBox "Outlook could not be reached." & vbCrLf & vbCrLf & _
           "Open Outlook and try again." & vbCrLf & _
           "(" & Err.Description & ")", vbCritical, "Outlook not available"
End Function


Private Function GetOutlook() As Object
    On Error Resume Next
    Set GetOutlook = GetObject(, "Outlook.Application")
    If GetOutlook Is Nothing Then
        Err.Clear
        Set GetOutlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
'  Button targets
' ---------------------------------------------------------------------------
Public Sub SendSelectedReminder()
    Dim r As Long
    r = ActiveCell.Row
    If r < FIRST_DATA_ROW Then
        MsgBox "Click any exam row first, then press this button.", _
               vbInformation, "Pick a row"
        Exit Sub
    End If
    If SendReminderRow(r, CONFIRM_BEFORE_SEND) Then
        MsgBox "Reminder sent.", vbInformation, "Done"
    End If
End Sub


Public Sub PreviewSelectedReminder()
    Dim r As Long, toList As String, ccList As String, subj As String, body As String
    Dim ol As Object, mail As Object

    r = ActiveCell.Row
    If r < FIRST_DATA_ROW Then
        MsgBox "Click any exam row first, then press this button.", _
               vbInformation, "Pick a row"
        Exit Sub
    End If

    BuildMessage r, toList, ccList, subj, body
    If Len(toList) = 0 Then
        MsgBox "No addresses yet for this exam.", vbExclamation, "Nothing to preview"
        Exit Sub
    End If

    On Error GoTo NoOutlook
    Set ol = GetOutlook()
    Set mail = ol.CreateItem(0)
    With mail
        .To = toList
        If Len(ccList) > 0 Then .CC = ccList
        .Subject = subj
        .body = body
        .Display                      ' always opens, never sends
    End With
    Exit Sub

NoOutlook:
    MsgBox "Outlook could not be reached." & vbCrLf & Err.Description, _
           vbCritical, "Outlook not available"
End Sub


Public Sub SendAllUpcoming()
    Dim wsR As Worksheet, r As Long, lastRow As Long
    Dim n As Long, sent As Long, answer As VbMsgBoxResult
    Dim status As String

    Set wsR = ThisWorkbook.Worksheets(SH_REM)
    lastRow = wsR.Cells(wsR.Rows.Count, C_COURSE).End(xlUp).Row

    n = 0
    For r = FIRST_DATA_ROW To lastRow
        status = CStr(wsR.Cells(r, C_STATUS).Value)
        If status <> "Past" And Len(Trim$(CStr(wsR.Cells(r, C_TO).Value))) > 0 Then
            n = n + 1
        End If
    Next r

    If n = 0 Then
        MsgBox "Nothing upcoming to send.", vbInformation, "Reminders"
        Exit Sub
    End If

    answer = MsgBox("Send reminders for ALL " & n & " upcoming exam(s)?" & vbCrLf & vbCrLf & _
                    "This sends " & n & " separate emails.", _
                    vbYesNo + vbExclamation, "Send all upcoming")
    If answer <> vbYes Then Exit Sub

    sent = 0
    For r = FIRST_DATA_ROW To lastRow
        status = CStr(wsR.Cells(r, C_STATUS).Value)
        If status <> "Past" And Len(Trim$(CStr(wsR.Cells(r, C_TO).Value))) > 0 Then
            If SendReminderRow(r, False) Then sent = sent + 1
        End If
    Next r

    MsgBox sent & " reminder(s) sent.", vbInformation, "Done"
End Sub
