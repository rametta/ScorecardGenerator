Option Explicit

Private Sub Worksheet_Activate()

    On Error Resume Next

    Const FIRST_ROW_SETTINGS = 8

    Dim settings        As Worksheet
    Dim i               As Byte
    Dim lastRowSettings As Byte
    Dim listStr         As String

    Set settings = Worksheets(SETTINGS_NAME)
    lastRowSettings = settings.Cells(settings.Rows.Count, 1).End(xlUp).Row
    listStr = ""

    For i = FIRST_ROW_SETTINGS To lastRowSettings
        listStr = listStr & "," & settings.Cells(i, 1).Value2
    Next

    With Range("metricToDelete").Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=listStr
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Incorrect Metric"
        .InputMessage = ""
        .ErrorMessage = "Please select a valid metric to delete"
        .ShowInput = True
        .ShowError = True
    End With

    Range("metricToDelete") = "Click Here"

End Sub
