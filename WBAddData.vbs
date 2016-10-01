Option Explicit

Private Sub Worksheet_Activate()

    On Error Resume Next

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

    With Range("metricToAddData").Validation
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

    Range("metricToAddData") = "Click Here"
    Worksheets("Add Data").Range("A10:E50").delete xlShiftUp

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$B$4" Or Target.Text = "Click Here" Then Exit Sub

    Dim AddData        As Worksheet
    Dim dashboard      As Worksheet
    Dim Metric         As String
    Dim hasSubMetrics  As Boolean
    Dim i              As Byte
    Dim lastRowAddData As Byte

    Set AddData = Worksheets(ADD_DATA_NAME)
    Metric = AddData.Range("metricToAddData").Value2
    Set dashboard = Worksheets(Metric)
    hasSubMetrics = True
    i = FIRST_ROW_DASHBOARD + 1
    lastRowAddData = 10

    AddData.Range("A10:E50").delete xlShiftUp

    Do While hasSubMetrics
        If dashboard.Cells(i, 1).Value2 <> "" And dashboard.Cells(i, 1).Value2 <> "Total" Then
            AddData.Cells(lastRowAddData, 2) = dashboard.Cells(i, 1).Value2
            lastRowAddData = lastRowAddData + 1
            i = i + 1
        Else
            hasSubMetrics = False
        End If
    Loop
End Sub
