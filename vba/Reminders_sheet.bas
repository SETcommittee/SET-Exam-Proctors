Option Explicit

' Live preview: whenever you click an exam row, the box at the top of this
' sheet fills in with exactly what would be sent for that row.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r As Long
    Dim toList As String, ccList As String, subj As String, body As String

    If Target.Cells.Count > 1 Then Exit Sub
    r = Target.Row
    If r < 10 Then Exit Sub
    If Len(Trim$(CStr(Me.Cells(r, 4).Value))) = 0 Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo Cleanup

    BuildMessage r, toList, ccList, subj, body

    Me.Range("B4").Value = IIf(Len(toList) = 0, _
        "(no addresses yet - fill in the Emails sheet)", Replace(toList, "; ", "   "))
    Me.Range("B5").Value = IIf(Len(ccList) = 0, "(none)", ccList)
    Me.Range("B6").Value = subj
    Me.Range("B7").Value = body

Cleanup:
    Application.EnableEvents = True
End Sub


' Double-click a cell in the Send column to send that row's reminder.
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column <> C_SEND Then Exit Sub
    If Target.Row < 10 Then Exit Sub
    If Len(Trim$(CStr(Me.Cells(Target.Row, 4).Value))) = 0 Then Exit Sub

    Cancel = True
    If SendReminderRow(Target.Row, CONFIRM_BEFORE_SEND) Then
        MsgBox "Reminder sent.", vbInformation, "Done"
    End If
End Sub
