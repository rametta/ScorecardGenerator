Option Explicit

Public Sub RemoveMetric()
    FastProcessing

    Dim i                As Integer
    Dim lastRowSettings  As Integer
    Dim lastRowScorecard As Integer
    Dim lastRowColors    As Integer
    Dim lastRowData      As Integer

    Dim settings         As Worksheet
    Dim delete           As Worksheet
    Dim colors           As Worksheet
    Dim data             As Worksheet
    Dim scorecard        As Worksheet

    Dim Metric           As String

    Set settings = Worksheets(SETTINGS_NAME)
    Set delete = Worksheets(DELETE_METRIC_NAME)
    Set colors = Worksheets(COLORS_NAME)
    Set data = Worksheets(DATA_NAME)
    Set scorecard = Worksheets(SCORECARD_NAME)

    lastRowSettings = settings.Cells(settings.Rows.Count, 1).End(xlUp).Row
    lastRowColors = colors.Cells(colors.Rows.Count, 1).End(xlUp).Row
    lastRowData = data.Cells(data.Rows.Count, 1).End(xlUp).Row
    lastRowScorecard = scorecard.Cells(scorecard.Rows.Count, 1).End(xlUp).Row

    Metric = delete.Range("metricToDelete").Value2

    'start deleting all the things, backwards loops so nothing gets skiped when rows shift upwards
    'color coding deletion
    For i = lastRowColors To 2 Step -1
        If colors.Cells(i, 1).Value2 = Metric Then colors.Rows(i).delete
    Next

    'data deletion
    For i = lastRowData To 2 Step -1
        If data.Cells(i, 1).Value2 = Metric Then data.Rows(i).delete
    Next

    'scorecard row deletion
    For i = lastRowScorecard To 5 Step -1
        If scorecard.Cells(i, 1).Value2 = Metric Then
            scorecard.Rows(i).delete
            Exit For
        End If
    Next

    'settings row deletion
    For i = lastRowSettings To 5 Step -1
        If settings.Cells(i, 1).Value2 = Metric Then
            settings.Cells(i, 1).delete shift:=xlUp
            Exit For
        End If
    Next

    'delete the sheet (dashboard)
    Worksheets(Metric).delete

    'done
    Application.Goto scorecard.Range("A5")
    MsgBox "Success: " & Metric & " has been deleted", vbInformation, "Success"

    NormalProcessing
End Sub
