Option Explicit

' Jason Rametta

'Validates data everytime a cell value changes
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$B$3" Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(NEW_METRIC_NAME)

    If InStr(1, ws.Range("B3").Text, "'") > 0 Then
        MsgBox "No Apostrophe's Allowed in Metric Name", vbCritical, "Symbol Error"
        ws.Range("B3") = ""
    End If
    If Len(ws.Range("B3").Text) > 30 Then
        MsgBox "Metric name must not exceed 30 characters", vbCritical, "Name Length Error"
        ws.Range("B3") = ""
    End If
End Sub
