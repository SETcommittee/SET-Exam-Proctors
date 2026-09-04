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

' Wording of the email lives on this sheet so it can be edited without
' opening the code. Blank cells fall back to the built-in text below.
Private Const SH_TPL As String = "Message"
Private Const CELL_SUBJ As String = "B4"
Private Const CELL_BODY As String = "B6"

Private Const FIRST_DATA_ROW As Long = 10   ' first exam row on Reminders
Private Const EXAM_FIRST_ROW As Long = 6    ' first exam row on Exam sheet

' Room plan sheet. Every module-level declaration has to live up here in the
' declarations block - VBA ignores a Const written further down between
' procedures, and the name then reads as undefined wherever it is used.
Private Const SH_ROOMS As String = "Room Plan"
Private Const ROOMS_FIRST_ROW As Long = 6

' Cell holding the address to send from. Blank or "(default)" uses whichever
' account Outlook would normally use.
Private Const CELL_SENDFROM As String = "B3"
Private Const SENDFROM_DEFAULT As String = "(default account)"

' Addresses copied on every single reminder, separated by ; or , - for people
' who should see all of them (a chair, an assistant, the committee mailbox).
Private Const CELL_ALWAYSCC As String = "E3"

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

    ' Any exit from here on must put these back, or Excel is left looking
    ' frozen with a stale screen and dead click handlers.
    On Error GoTo Restore
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
            ccList = MergeCc(ccList, AlwaysCcList(), toList)

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

    RefreshRemindersCore = rR - FIRST_DATA_ROW

Restore:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Could not rebuild the list." & vbCrLf & vbCrLf & _
               Err.Description, vbCritical + vbSystemModal, "Reminders"
    End If
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
'  Room plan: one row per proctor, with the room they are covering
' ---------------------------------------------------------------------------
'
' The Exam sheet keeps rooms as one list per exam ("N-301, N-303, N-304") and
' proctors in separate columns. The committee's convention is positional - the
' first proctor takes the first room, the second the second, and so on - which
' is how the printed room list is laid out. This expands that into one row per
' person, and says so plainly when the two lists do not line up.

Public Sub RefreshRoomPlan()
    Dim n As Long
    n = RefreshRoomPlanCore()
    MsgBox "Room plan rebuilt." & vbCrLf & vbCrLf & _
           n & " proctor/room row(s).", vbInformation + vbSystemModal, "Room plan"
End Sub


Public Function RefreshRoomPlanCore() As Long
    Dim wsE As Worksheet, wsP As Worksheet
    Dim rE As Long, rP As Long, lastE As Long, c As Long, i As Long
    Dim course As String, nm As String, note As String
    Dim procs() As String, rooms() As String
    Dim nProc As Long, nRoom As Long, slots As Long
    Dim dt As Variant, roomTxt As String

    Set wsE = ThisWorkbook.Worksheets(SH_EXAM)
    Set wsP = ThisWorkbook.Worksheets(SH_ROOMS)

    On Error GoTo Restore
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Clear only the rows actually used. Resetting the font colour across the
    ' whole sheet is enough work to make Excel refuse the call outright.
    Dim lastP As Long
    lastP = wsP.Cells(wsP.Rows.Count, 4).End(xlUp).Row
    If lastP >= ROOMS_FIRST_ROW Then
        With wsP.Range(wsP.Cells(ROOMS_FIRST_ROW, 1), wsP.Cells(lastP, 9))
            .ClearContents
            .Interior.Pattern = xlNone
            .Font.Color = RGB(0, 0, 0)
        End With
    End If

    lastE = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    rP = ROOMS_FIRST_ROW

    For rE = EXAM_FIRST_ROW To lastE
        course = Trim$(CStr(wsE.Cells(rE, 1).Value))
        If Len(course) > 0 And UCase$(course) <> "TOTAL" Then
            dt = wsE.Cells(rE, 4).Value

            ReDim procs(0 To 9)
            nProc = 0
            For c = 11 To 20
                nm = Trim$(CStr(wsE.Cells(rE, c).Value))
                If Len(nm) > 0 Then
                    procs(nProc) = nm
                    nProc = nProc + 1
                End If
            Next c

            roomTxt = Trim$(CStr(wsE.Cells(rE, 8).Value))
            nRoom = 0
            If Len(roomTxt) > 0 Then
                rooms = Split(Replace(roomTxt, "/", ","), ",")
                For i = LBound(rooms) To UBound(rooms)
                    rooms(i) = Trim$(rooms(i))
                    If Len(rooms(i)) > 0 Then nRoom = nRoom + 1
                Next i
            End If

            ' A row for every proctor, plus a row for any room nobody is in.
            slots = nProc
            If nRoom > slots Then slots = nRoom
            If slots = 0 Then slots = 1

            For i = 0 To slots - 1
                note = ""
                wsP.Cells(rP, 1).Value = dt
                wsP.Cells(rP, 1).NumberFormat = "dd mmm yyyy"
                wsP.Cells(rP, 2).Value = wsE.Cells(rE, 5).Value
                wsP.Cells(rP, 3).Value = FormatTime(wsE.Cells(rE, 6).Value) & _
                    IIf(Len(FormatTime(wsE.Cells(rE, 7).Value)) > 0, _
                        " - " & FormatTime(wsE.Cells(rE, 7).Value), "")
                wsP.Cells(rP, 4).Value = course

                If i < nProc Then
                    wsP.Cells(rP, 5).Value = procs(i)
                Else
                    wsP.Cells(rP, 5).Value = "(nobody assigned)"
                    note = "Room has no proctor"
                End If

                If i < nRoom Then
                    wsP.Cells(rP, 6).Value = rooms(i)
                ElseIf nRoom = 0 Then
                    wsP.Cells(rP, 6).Value = "TBC"
                    note = "No room set on the Exam sheet"
                Else
                    wsP.Cells(rP, 6).Value = "TBC"
                    note = "More proctors than rooms"
                End If

                wsP.Cells(rP, 7).Value = wsE.Cells(rE, 9).Value
                wsP.Cells(rP, 8).Value = wsE.Cells(rE, 2).Value
                wsP.Cells(rP, 9).Value = note

                If Len(note) > 0 Then
                    wsP.Cells(rP, 9).Interior.Color = RGB(250, 229, 227)
                    wsP.Cells(rP, 9).Font.Color = RGB(166, 42, 34)
                    wsP.Cells(rP, 6).Font.Color = RGB(166, 42, 34)
                End If

                If IsDate(dt) Then
                    If Int(CDate(dt)) < Int(Now()) Then
                        wsP.Range(wsP.Cells(rP, 1), wsP.Cells(rP, 8)).Font.Color = _
                            RGB(150, 158, 170)
                    End If
                End If

                rP = rP + 1
            Next i
        End If
    Next rE

    RefreshRoomPlanCore = rP - ROOMS_FIRST_ROW

Restore:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Could not rebuild the room plan." & vbCrLf & vbCrLf & Err.Description, _
               vbCritical + vbSystemModal, "Room plan"
    End If
End Function


' ---------------------------------------------------------------------------
'  Build the message for one row
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
'  The wording, and the placeholders it can use
' ---------------------------------------------------------------------------

Public Function DefaultSubject() As String
    DefaultSubject = "Proctoring reminder - {course}, {day} {date} at {time}"
End Function


Public Function DefaultBody() As String
    DefaultBody = _
        "Dear colleagues," & vbCrLf & vbCrLf & _
        "This is a reminder of the exam below and your proctoring duty for it." & vbCrLf & vbCrLf & _
        "Course:       {course}" & vbCrLf & _
        "Date:         {day}, {longdate}" & vbCrLf & _
        "Time:         {time}" & vbCrLf & vbCrLf & _
        "Proctors and rooms:" & vbCrLf & _
        "{proctorrooms}" & vbCrLf & _
        "Please go to the room shown next to your name, and arrive 15 minutes" & vbCrLf & _
        "before the start time." & vbCrLf & vbCrLf & _
        "If you cannot attend, reply to this message as early as you can so a" & vbCrLf & _
        "replacement can be arranged." & vbCrLf & vbCrLf & _
        "Thank you," & vbCrLf & _
        "Exam Committee" & vbCrLf & _
        "School of Engineering and Technology"
End Function


' Read a template cell, falling back to the built-in wording when the Message
' sheet is missing or the cell has been left empty.
Private Function Template(ByVal whichCell As String, ByVal fallback As String) As String
    Dim ws As Worksheet, v As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_TPL)
    On Error GoTo 0
    If ws Is Nothing Then
        Template = fallback
        Exit Function
    End If

    v = CStr(ws.Range(whichCell).Value)
    If Len(Trim$(v)) = 0 Then v = fallback
    Template = v
End Function


' Split an exam's room text into its individual rooms.
Private Function RoomsOf(ByVal roomTxt As String, ByRef n As Long) As Variant
    Dim rooms() As String, i As Long, keep() As String, k As Long

    n = 0
    roomTxt = Trim$(roomTxt)
    If Len(roomTxt) = 0 Or UCase$(roomTxt) = "TBC" Then Exit Function

    rooms = Split(Replace(roomTxt, "/", ","), ",")
    ReDim keep(0 To UBound(rooms) - LBound(rooms))
    k = 0
    For i = LBound(rooms) To UBound(rooms)
        If Len(Trim$(rooms(i))) > 0 Then
            keep(k) = Trim$(rooms(i))
            k = k + 1
        End If
    Next i
    n = k
    RoomsOf = keep
End Function


' Each proctor with the room they are covering, using the same positional
' rule as the Room Plan sheet: first proctor takes the first room listed.
Private Function ProctorRoomBlock(ByVal r As Long) As String
    Dim wsR As Worksheet, procs() As String, rooms As Variant
    Dim i As Long, nRoom As Long, out As String, rm As String

    Set wsR = ThisWorkbook.Worksheets(SH_REM)
    procs = Split(CStr(wsR.Cells(r, C_PROCS).Value), ", ")
    rooms = RoomsOf(CStr(wsR.Cells(r, C_ROOM).Value), nRoom)

    For i = LBound(procs) To UBound(procs)
        If Len(Trim$(procs(i))) > 0 Then
            If i < nRoom Then rm = rooms(i) Else rm = "room to be confirmed"
            out = out & "  " & (i + 1) & ". " & Trim$(procs(i)) & _
                  "  -  " & rm & vbCrLf
        End If
    Next i
    ProctorRoomBlock = out
End Function


' The numbered proctor list, as its own block so {proctors} can sit anywhere.
Private Function ProctorBlock(ByVal r As Long) As String
    Dim wsR As Worksheet, parts() As String, i As Long, out As String
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    parts = Split(CStr(wsR.Cells(r, C_PROCS).Value), ", ")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            out = out & "  " & (i + 1) & ". " & Trim$(parts(i)) & vbCrLf
        End If
    Next i
    ProctorBlock = out
End Function


' Swap every {placeholder} for this row's value.
Private Function FillTemplate(ByVal tpl As String, ByVal r As Long) As String
    Dim wsR As Worksheet, s As String, d As Variant
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    s = tpl
    d = wsR.Cells(r, C_DATE).Value

    s = Replace(s, "{course}", CStr(wsR.Cells(r, C_COURSE).Value))
    s = Replace(s, "{day}", CStr(wsR.Cells(r, C_DAY).Value))
    s = Replace(s, "{date}", IIf(IsDate(d), Format$(d, "d mmm yyyy"), ""))
    s = Replace(s, "{longdate}", IIf(IsDate(d), Format$(d, "d mmmm yyyy"), ""))
    s = Replace(s, "{time}", CStr(wsR.Cells(r, C_TIME).Value))
    s = Replace(s, "{room}", CStr(wsR.Cells(r, C_ROOM).Value))
    s = Replace(s, "{coordinator}", CStr(wsR.Cells(r, C_COORD).Value))
    s = Replace(s, "{proctorlist}", CStr(wsR.Cells(r, C_PROCS).Value))
    s = Replace(s, "{proctorrooms}", ProctorRoomBlock(r))
    s = Replace(s, "{proctors}", ProctorBlock(r))
    s = Replace(s, "{rooms}", CStr(wsR.Cells(r, C_ROOM).Value))
    s = Replace(s, "{count}", CStr(UBound(Split(CStr(wsR.Cells(r, C_PROCS).Value), ", ")) + 1))

    FillTemplate = s
End Function


Public Sub BuildMessage(ByVal r As Long, ByRef toList As String, ByRef ccList As String, _
                        ByRef subj As String, ByRef body As String)
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Worksheets(SH_REM)

    toList = Trim$(CStr(wsR.Cells(r, C_TO).Value))
    ' Applied again here so a change to the always-Cc box takes effect even if
    ' the list has not been refreshed since. MergeCc will not duplicate.
    ccList = MergeCc(Trim$(CStr(wsR.Cells(r, C_CC).Value)), AlwaysCcList(), toList)

    subj = FillTemplate(Template(CELL_SUBJ, DefaultSubject()), r)
    body = FillTemplate(Template(CELL_BODY, DefaultBody()), r)
End Sub


' Button target: put the built-in wording back on the Message sheet.
Public Sub ResetMessageTemplate()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_TPL)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "The Message sheet is missing.", vbExclamation + vbSystemModal, "Template"
        Exit Sub
    End If

    If MsgBox("Replace the subject and body with the original wording?" & vbCrLf & vbCrLf & _
              "Anything you have written there will be lost.", _
              vbYesNo + vbExclamation + vbSystemModal, "Reset wording") <> vbYes Then Exit Sub

    ws.Range(CELL_SUBJ).Value = DefaultSubject()
    ws.Range(CELL_BODY).Value = DefaultBody()
    MsgBox "Original wording restored.", vbInformation + vbSystemModal, "Template"
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
                        "From: " & SenderLabel() & vbCrLf & vbCrLf & _
                        "To:" & vbCrLf & Replace(toList, "; ", vbCrLf) & vbCrLf & _
                        IIf(Len(ccList) > 0, vbCrLf & "Cc:" & vbCrLf & ccList & vbCrLf, "") & _
                        IIf(Len(missing) > 0, vbCrLf & "NO ADDRESS FOR: " & missing & vbCrLf, ""), _
                        vbYesNo + vbQuestion + vbSystemModal, "Confirm reminder")
        If answer <> vbYes Then Exit Function
    End If

    On Error GoTo NoOutlook
    Application.StatusBar = "Contacting Outlook..."
    Set ol = GetOutlook()
    If ol Is Nothing Then GoTo NoOutlook

    Set mail = ol.CreateItem(0)
    ApplySender mail, ol
    With mail
        .To = toList
        If Len(ccList) > 0 Then .CC = ccList
        .Subject = subj
        .body = body
        If UCase$(SEND_MODE) = "DISPLAY" Then
            .Display
            On Error Resume Next
            .GetInspector.Activate      ' bring it in front of Excel
            On Error GoTo NoOutlook
        Else
            .Send
        End If
    End With
    Application.StatusBar = False
    On Error GoTo 0

    wsR.Cells(r, C_SENTLOG).Value = "Sent " & Format$(Now(), "d mmm hh:mm")
    wsR.Cells(r, C_SENTLOG).Font.Color = RGB(28, 107, 69)
    SendReminderRow = True
    Exit Function

NoOutlook:
    Application.StatusBar = False
    MsgBox "Outlook could not be reached, so nothing was sent." & vbCrLf & vbCrLf & _
           "Open Outlook yourself and wait until the Inbox is showing, " & _
           "then try again. Outlook needs to be running before this button " & _
           "is used." & vbCrLf & vbCrLf & _
           "(" & Err.Description & ")", _
           vbCritical + vbSystemModal, "Outlook not available"
End Function


' Getting hold of Outlook is the fragile part of this whole workbook.
'
' If Outlook is not already open, CreateObject starts it with no window. In
' that state it sits waiting on a sign-in prompt it has nowhere to display,
' and every later call blocks forever - Excel just appears to freeze. So:
' log the MAPI session on explicitly, and if there is no Explorer window,
' open one. Then Outlook is properly awake before we hand it a message.
Private Function GetOutlook() As Object
    Dim ol As Object, ns As Object

    Set GetOutlook = Nothing

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")   ' attach to a running copy
    On Error GoTo 0

    If ol Is Nothing Then
        On Error Resume Next
        Set ol = CreateObject("Outlook.Application")
        On Error GoTo 0
    End If
    If ol Is Nothing Then Exit Function

    On Error Resume Next
    Set ns = ol.GetNamespace("MAPI")
    ns.Logon "", "", False, False
    If ol.Explorers.Count = 0 Then
        ol.Explorers.Add(ns.GetDefaultFolder(6), 0).Display   ' 6 = Inbox
    End If
    On Error GoTo 0

    Set GetOutlook = ol
End Function


' ---------------------------------------------------------------------------
'  Which address the reminders are sent from
' ---------------------------------------------------------------------------

' Apply the address chosen in B3 to one message.
'
' If it matches a configured Outlook account, send through that account. If it
' does not, treat it as an alias or shared mailbox and set it as the "from"
' instead - that works when the account has Send As rights, and Outlook will
' reject it clearly if not.
Private Sub ApplySender(ByVal mail As Object, ByVal ol As Object)
    Dim want As String, i As Long, accs As Object

    want = Trim$(CStr(ThisWorkbook.Worksheets(SH_REM).Range(CELL_SENDFROM).Value))
    If Len(want) = 0 Then Exit Sub
    If StrComp(want, SENDFROM_DEFAULT, vbTextCompare) = 0 Then Exit Sub

    On Error Resume Next
    Set accs = ol.Session.Accounts
    If Not accs Is Nothing Then
        For i = 1 To accs.Count
            If StrComp(CStr(accs.Item(i).SmtpAddress), want, vbTextCompare) = 0 Then
                Set mail.SendUsingAccount = accs.Item(i)
                Exit Sub
            End If
        Next i
    End If
    mail.SentOnBehalfOfName = want
    On Error GoTo 0
End Sub


Private Function AlwaysCcList() As String
    AlwaysCcList = Trim$(CStr(ThisWorkbook.Worksheets(SH_REM).Range(CELL_ALWAYSCC).Value))
End Function


' Fold the always-Cc addresses into a row's Cc, skipping any that are already
' there and any that are already in the To line. Running this twice on the
' same value changes nothing, so it is safe to apply on refresh and again at
' send time.
Private Function MergeCc(ByVal baseCc As String, ByVal extra As String, _
                         ByVal excludeTo As String) As String
    Dim parts() As String, i As Long, one As String
    Dim out As String, hay As String, toHay As String

    out = Trim$(baseCc)
    If Len(Trim$(extra)) = 0 Then
        MergeCc = out
        Exit Function
    End If

    toHay = ";" & Replace(Trim$(excludeTo), "; ", ";") & ";"

    parts = Split(Replace(Replace(Replace(extra, ",", ";"), vbLf, ";"), vbCr, ";"), ";")
    For i = LBound(parts) To UBound(parts)
        one = Trim$(parts(i))
        If Len(one) > 0 Then
            hay = ";" & Replace(out, "; ", ";") & ";"
            If InStr(1, hay, ";" & one & ";", vbTextCompare) = 0 _
               And InStr(1, toHay, ";" & one & ";", vbTextCompare) = 0 Then
                If Len(out) > 0 Then out = out & "; "
                out = out & one
            End If
        End If
    Next i

    MergeCc = out
End Function


Private Function SenderLabel() As String
    SenderLabel = Trim$(CStr(ThisWorkbook.Worksheets(SH_REM).Range(CELL_SENDFROM).Value))
    If Len(SenderLabel) = 0 Then SenderLabel = SENDFROM_DEFAULT
End Function


' Read the accounts out of Outlook and turn B3 into a dropdown of them.
Public Sub RefreshSendAccounts()
    Dim ol As Object, accs As Object, i As Long
    Dim list As String, smtp As String, ws As Worksheet

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SH_REM)
    Set ol = GetOutlook()
    If ol Is Nothing Then GoTo Failed

    Set accs = ol.Session.Accounts
    list = SENDFROM_DEFAULT
    For i = 1 To accs.Count
        smtp = ""
        On Error Resume Next
        smtp = CStr(accs.Item(i).SmtpAddress)
        On Error GoTo Failed
        If Len(smtp) > 0 Then list = list & "," & smtp
    Next i

    With ws.Range(CELL_SENDFROM).Validation
        .Delete
        .Add Type:=3, AlertStyle:=2, Operator:=1, Formula1:=list   ' xlValidateList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Send from"
        .InputMessage = "Pick which account the reminders are sent from."
    End With

    If Len(Trim$(CStr(ws.Range(CELL_SENDFROM).Value))) = 0 Then
        ws.Range(CELL_SENDFROM).Value = SENDFROM_DEFAULT
    End If

    MsgBox "Found " & accs.Count & " account(s) in Outlook." & vbCrLf & vbCrLf & _
           Replace(Mid$(list, Len(SENDFROM_DEFAULT) + 2), ",", vbCrLf) & vbCrLf & vbCrLf & _
           "Pick one in the 'Send from' box." & vbCrLf & vbCrLf & _
           "Only accounts added to the classic Outlook desktop app appear here. " & _
           "An address that lives only in the new Outlook app cannot be used, " & _
           "but you can still type an alias or shared mailbox into the box by hand.", _
           vbInformation + vbSystemModal, "Send from"
    Exit Sub

Failed:
    MsgBox "Could not read the Outlook accounts." & vbCrLf & vbCrLf & _
           "Open Outlook and try again." & vbCrLf & "(" & Err.Description & ")", _
           vbCritical + vbSystemModal, "Send from"
End Sub


' Safe check - creates nothing, sends nothing. Use this if a send ever seems
' to hang, to find out whether Outlook is the problem.
Public Sub TestOutlookConnection()
    Dim ol As Object, ns As Object, who As String, wins As Long

    On Error GoTo Failed
    Set ol = GetOutlook()
    If ol Is Nothing Then GoTo Failed

    Set ns = ol.GetNamespace("MAPI")
    who = CStr(ns.CurrentUser)
    wins = ol.Explorers.Count

    MsgBox "Outlook is reachable." & vbCrLf & vbCrLf & _
           "Signed in as: " & who & vbCrLf & _
           "Open windows: " & wins, _
           vbInformation + vbSystemModal, "Outlook check"
    Exit Sub

Failed:
    MsgBox "Could not reach Outlook." & vbCrLf & vbCrLf & _
           "Open Outlook yourself, wait until the Inbox is showing, " & _
           "then try again." & vbCrLf & vbCrLf & _
           "(" & Err.Description & ")", _
           vbCritical + vbSystemModal, "Outlook check"
End Sub


' If a macro ever stops halfway, Excel can be left with screen updating or
' events switched off, which looks exactly like a freeze. Run this to undo it.
Public Sub RestoreExcel()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Excel restored.", vbInformation + vbSystemModal, "Reset"
End Sub


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
    Application.StatusBar = "Contacting Outlook..."
    Set ol = GetOutlook()
    If ol Is Nothing Then GoTo NoOutlook

    Set mail = ol.CreateItem(0)
    ApplySender mail, ol
    With mail
        .To = toList
        If Len(ccList) > 0 Then .CC = ccList
        .Subject = subj
        .body = body
        .Display                      ' always opens, never sends
        On Error Resume Next
        .GetInspector.Activate        ' otherwise it can open behind Excel
        On Error GoTo NoOutlook
    End With
    Application.StatusBar = False
    Exit Sub

NoOutlook:
    Application.StatusBar = False
    MsgBox "Outlook could not be reached." & vbCrLf & vbCrLf & _
           "Open Outlook yourself and wait until the Inbox is showing, " & _
           "then try again." & vbCrLf & vbCrLf & _
           "(" & Err.Description & ")", _
           vbCritical + vbSystemModal, "Outlook not available"
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
