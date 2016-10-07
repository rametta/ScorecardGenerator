Option Explicit

' Jason Rametta

Public Sub formatRow(rng As Range, isColorMatrix As Boolean)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.color = vbBlack
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(242, 242, 242)
        .Item(1, 1).HorizontalAlignment = xlLeft
    End With

    If isColorMatrix Then
        rng.Columns("C:N").Interior.color = vbWhite
    Else
        rng.Columns("E:P").Interior.color = vbWhite
    End If
End Sub

' Used for adding those small '+' buttons to activate dashboards
Public Sub addBtn(ws As Worksheet, Name As String, targetRow As Integer, btnPixels As Byte)
    ws.Activate
    ws.Shapes.AddShape(msoShapeMathPlus, ws.Cells(targetRow, 1).Left + btnPixels, ws.Cells(targetRow, 1).Top + 5, 14.25, 14.25).Select
    With Selection
        .ShapeRange.Fill.Visible = msoTrue
        .ShapeRange.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .ShapeRange.Fill.Transparency = 0
        .ShapeRange.Fill.Solid
        .ShapeRange.Line.Visible = msoFalse
        .ShapeRange.Name = Name
        .Name = Name
        .OnAction = "ShowDashboard"
    End With
End Sub

Public Sub editConditionalFormating(ws As Worksheet, rng As Range, appliesToRng As Range)
    Dim i As Byte

    ws.Activate
    For i = 1 To rng.FormatConditions.Count
        rng.FormatConditions(i).ModifyAppliesToRange appliesToRng
    Next
End Sub

Public Function Letter(lngCol As Long) As String
    Letter = Split(Cells(1, lngCol).Address(True, False), "$")(0)
End Function

Public Function percentFormat(cellStr As String, decimalPlaces As Byte) As String
    Dim i As Byte
    Dim result As String
    
    result = IIf(decimalPlaces > 0, "text(" & cellStr & ",""0.", "text(" & cellStr & ",""0")
    For i = 1 To decimalPlaces
        result = result & "0"
    Next
    result = result & "%"")"
    percentFormat = result
End Function

' Unhides the dashboard tab when button clicked
Public Sub ShowDashboard()
On Error GoTo errHandler:
    Dim pSheet As Worksheet
    Set pSheet = Worksheets(Replace(ActiveSheet.Shapes(Application.Caller).Name, "_", " "))
    pSheet.Visible = True
    Application.Goto pSheet.Range("A1")
Exit Sub

errHandler:
    MsgBox "Error 404: Sheet Not Found" & vbNewLine & "We have our smartest monkeys working on the problem", vbCritical, "Whoops!"
End Sub

' Runs when the scorecard workbook is initially opened, and when tabs are changed
Public Sub StartUp(Optional except As String)
    On Error Resume Next
    Dim settings  As Worksheet
    Dim scorecard As Worksheet
    Dim i         As Byte
    Dim rowCount  As Byte

    Set settings = Worksheets(SETTINGS_NAME)
    Set scorecard = Worksheets(SCORECARD_NAME)
    rowCount = settings.Cells(settings.Rows.Count, 1).End(xlUp).Row

    'TODO replace for loop with foreach loop in range object to reduce variables
    For i = 2 To rowCount
        If except <> settings.Cells(i, 1) Then Worksheets(settings.Cells(i, 1).Text).Visible = False
    Next

    ' Executes if it's on startup, not on tab changes
    If except = "" Then
        settings.Range("isAutoHide") = True
        Application.Goto scorecard.Range("C2")
    End If

End Sub

Public Sub FastProcessing()
    With Application
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
End Sub

Public Sub NormalProcessing()
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub
