Option Explicit
'Event Listeners

' Jason Rametta

Private Sub Workbook_Open()
    StartUp
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Worksheets(SETTINGS_NAME).Range("isAutoHide") Then Call StartUp(Sh.Name)
End Sub
