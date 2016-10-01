Option Explicit

' Runs when Add Data button clicked on Add Data sheet
Public Sub AddData()
    FastProcessing

    Dim AddData        As Worksheet
    Dim data           As Worksheet
    Dim lastRowAddData As Integer
    Dim lastRowData    As Integer
    Dim i              As Integer

    Set AddData = Worksheets(ADD_DATA_NAME)
    Set data = Worksheets(DATA_NAME)
    lastRowAddData = AddData.Cells(255, 2).End(xlUp).Row

    If AddData.Range("metricToAddData") = "Click Here" Then Exit Sub

    For i = 10 To lastRowAddData
        lastRowData = data.Cells(data.Rows.Count, 1).End(xlUp).Row + 1
        With data.Rows(lastRowData)
            .Cells(DataColumns.Metric) = AddData.Range("metricToAddData").Value2 'metric
            .Cells(DataColumns.Sub_Metric) = AddData.Cells(i, 2).Value2          'submetric
            .Cells(DataColumns.Mnth) = AddData.Range("E7").Value2                'month
            .Cells(DataColumns.Yr) = AddData.Range("C7").Value2                  'year
            .Cells(DataColumns.Actual) = AddData.Cells(i, 4).Value2              'actual
            .Cells(DataColumns.Target) = AddData.Cells(i, 5).Value2              'target
        End With
    Next

    Application.Goto Worksheets(SCORECARD_NAME).Range("C2")
    MsgBox "Data has been added successfully!", vbInformation, "Great Success!"

    NormalProcessing
End Sub
