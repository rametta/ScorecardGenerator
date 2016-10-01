Option Explicit

' Event handler when Add New Metric is clicked
Public Sub AddMetric()
    FastProcessing

    ' DECLARATIONS
    Dim Metric As Metric

    ' Integers
    Dim lastRowScorecard     As Integer
    Dim lastRowColors        As Integer
    Dim lastRowSettings      As Integer
    Dim lastRowNewSubMetrics As Integer
    Dim i                    As Integer
    Dim r                    As Byte

    ' Worksheets
    Dim scorecard As Worksheet
    Dim colors    As Worksheet
    Dim settings  As Worksheet
    Dim newMetric As Worksheet
    Dim template  As Worksheet
    Dim dashboard As Worksheet

    ' Ranges
    Dim submetric  As Range

    ' INSTANTIATIONS

    ' Worksheets
    Set scorecard = Worksheets(SCORECARD_NAME)
    Set colors = Worksheets(COLORS_NAME)
    Set settings = Worksheets(SETTINGS_NAME)
    Set newMetric = Worksheets(NEW_METRIC_NAME)
    Set template = Worksheets(TEMPLATE_NAME)

    ' Metrics
    Set Metric = New Metric

    ' Integers
    lastRowColors = colors.Cells(colors.Rows.Count, ColorCodingColumns.Metric).End(xlUp).Row + 1
    lastRowSettings = settings.Cells(settings.Rows.Count, 1).End(xlUp).Row + 1
    lastRowScorecard = scorecard.Cells(scorecard.Rows.Count, ScorecardColumns.Metric).End(xlUp).Row + 1
    lastRowNewSubMetrics = newMetric.Cells(newMetric.Rows.Count, 4).End(xlUp).Row

    'assign values to metric object
    With Metric
        .Name = newMetric.Range("nmMetricName").Value2
        .SampleSize = newMetric.Range("nmSampleSize").Value2
        .Prime = newMetric.Range("nmPrime").Value2
        .Chart = newMetric.Range("nmChartType").Value2
        .colors(Legend.RedDisplay) = newMetric.Range("nmRedDisplay").Value2
        .colors(Legend.YellowDisplay) = newMetric.Range("nmYellowDisplay").Value2
        .colors(Legend.YelhighDisplay) = newMetric.Range("nmYelhighDisplay").Value2
        .colors(Legend.GreenDisplay) = newMetric.Range("nmGreenDisplay").Value2
        .colors(Legend.RealRedDisplay) = newMetric.Range("nmRealRedDisplay").Value2
        .colors(Legend.RealYellowDisplay) = newMetric.Range("nmRealYellowDisplay").Value2
        .colors(Legend.RealYelhighDisplay) = newMetric.Range("nmRealYelhighDisplay").Value2
        .colors(Legend.RealGreenDisplay) = newMetric.Range("nmRealGreenDisplay").Value2
        .colors(Legend.red) = newMetric.Range("nmRed").Value2
        .colors(Legend.yellow) = newMetric.Range("nmYellow").Value2
        .colors(Legend.yelhigh) = newMetric.Range("nmYelhigh").Value2
        .colors(Legend.green) = newMetric.Range("nmGreen").Value2
        .isPercentage = IIf(newMetric.Range("nmPercentage").Value2 = "No", False, True)
        .DecimalsDisplay = newMetric.Range("nmDecimals").Value2
        Set .SubMetrics = newMetric.Range("D4:D" & lastRowNewSubMetrics)
    End With

    'checks to see if all required fields are filled out
    If Not Metric.Valid Then
        MsgBox "Please fill out all required fields", vbExclamation, "Missing Field"
        NormalProcessing
        Exit Sub
    End If

    'check to see if metric name has already been taken
    For i = 2 To lastRowSettings - 1
        If settings.Cells(i, 1).Value2 = Metric.Name Then
            MsgBox "That name is already taken" & vbNewLine & "Please choose a different name", , "Metric Name Error"
            NormalProcessing
            Exit Sub
        End If
    Next

    'format Color sheet
    With colors.Rows(lastRowColors)
        .Cells(ColorCodingColumns.Metric).Formula = "=Scorecard!A$" & lastRowScorecard        'add metric name reference to color coding sheet
        .Cells(ColorCodingColumns.Yr) = Year(Date)                                            'add current year to new color coding row
        .Cells(ColorCodingColumns.red) = Metric.colors(Legend.red)                            'add value for red color
        .Cells(ColorCodingColumns.yellow) = Metric.colors(Legend.yellow)                      'add value for yellow color
        .Cells(ColorCodingColumns.yelhigh) = Metric.colors(Legend.yelhigh)                    'add value for yelhigh color
        .Cells(ColorCodingColumns.green) = Metric.colors(Legend.green)                        'add value for green color
        .Cells(ColorCodingColumns.Sample) = Metric.SampleSize                                 'add text for sample size
        .Cells(ColorCodingColumns.Prime) = Metric.Prime                                       'add text for prime
    End With

    template.Copy after:=settings                                                             'duplicate template for dashboard
    Set dashboard = Worksheets("TEMPLATE (2)")                                                'assign duplicate to object dashboard
    dashboard.Name = Metric.Name                                                              'rename duplicated template sheet name
    settings.Cells(lastRowSettings, 1) = Metric.Name                                          'add dashboard to list of sheets to hide on startup and tab change
    dashboard.Range("A1").Formula = "=Scorecard!" & FIRST_COL_SCORECARD & lastRowScorecard    'add reference to in dahsboard to metric name on scorecard

    r = FIRST_ROW_DASHBOARD + 1

    With scorecard
        .Range(COL_COLORS).EntireColumn.Hidden = False                                        'unhide color matrix
        .Cells(lastRowScorecard, ScorecardColumns.Metric) = Metric.Name                       'add metric name to scorecard
        .Range(FIRST_COL_COLORS & lastRowScorecard).Formula = "=A" & lastRowScorecard         'add metric name reference to scorecard color matrix

        'format newly added row in scorecard
        .Rows(lastRowScorecard).RowHeight = ROW_HEIGHT

        With .Rows(lastRowScorecard)
            .Columns(ScorecardColumns.Prime).FormulaArray = "=IFERROR(INDEX(Prime,MATCH($A" & lastRowScorecard & " & cYear,ColorMetric & ColorYear,0)),""-"")"   'prime array formula
            .Columns(ScorecardColumns.Sample).FormulaArray = "=IFERROR(INDEX(Sample,MATCH($A" & lastRowScorecard & " & cYear,ColorMetric & ColorYear,0)),""-"")" 'sample size array formula
            .Columns(ScorecardColumns.red).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColumns.red) & r        'red
            .Columns(ScorecardColumns.yellow).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColumns.yellow) & r  'yellow
            .Columns(ScorecardColumns.green).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColumns.green) & r    'green
        End With
    End With

    'format rows in scorecard for values and colors
    Call formatRow(scorecard.Range(FIRST_COL_SCORECARD & lastRowScorecard & ":" & LAST_COL_SCORECARD & lastRowScorecard), False)
    Call formatRow(scorecard.Range(FIRST_COL_COLORS & lastRowScorecard & ":" & LAST_COL_COLORS & lastRowScorecard), True)
    'add small blue plus button to activate dashboard on scorecard
    Call addBtn(scorecard, Metric.Name, lastRowScorecard, BUTTON_PLACEMENT_RIGHT)

    With dashboard
        'format symbols in legend on new dashboard
        With .Rows(r)
            'add formulas to color legend display
            .Columns(TemplateColumns.red).Formula = "=""" & Metric.colors(Legend.RedDisplay) & " "" & " & IIf(Metric.isPercentage, percentFormat("AP" & r, Metric.DecimalsDisplay), "AP" & r)
            .Columns(TemplateColumns.yellow).Formula = "=""" & Metric.colors(Legend.YellowDisplay) & " "" & " & IIf(Metric.isPercentage, percentFormat("AQ" & r, Metric.DecimalsDisplay), "AQ" & r) & " & "" - " & Metric.colors(Legend.YelhighDisplay) & " "" & " & IIf(Metric.isPercentage, percentFormat("AR" & r, Metric.DecimalsDisplay), "AR" & r) & " "
            .Columns(TemplateColumns.green).Formula = "=""" & Metric.colors(Legend.GreenDisplay) & " "" & " & IIf(Metric.isPercentage, percentFormat("AS" & r, Metric.DecimalsDisplay), "AS" & r) & ""

            'add submetrics to dashboard
            For i = Metric.SubMetrics.Count To 1 Step -1
                .Columns(Letter(TemplateColumns.SubMetrics) & ":" & Letter(TemplateColumns.Target)).Insert shift:=xlDown                                                                                            'insert new row
                .Columns(TemplateColumns.SubMetrics) = Metric.SubMetrics.Value2(i, 1)                                                                                                                               'add submetric name
                .Columns(TemplateColumns.PrevYearRestated).Formula = "=sumifs(actual,metric,$A$1,submetric,$A" & r & ",year,pyear)"                                                                                 'pYear
                .Columns(TemplateColumns.Target).Formula = "=sumifs(target,metric,$A$1,submetric,$A" & r & ",year,cyear)"                                                                                           'target
                .Columns(Letter(TemplateColumns.jan) & ":" & Letter(TemplateColumns.dec)).Formula = "=if(Settings!T$2=TRUE,sumifs(actual,metric,$A$1,submetric,$A" & r & ",month,C$" & r - 1 & ",year,cyear),"""")" 'jan-dec
                .Columns(TemplateColumns.ytd).Formula = "=sum(C" & r & ":N" & r & ")"                                                                                                                               'YTD
            Next
        End With

        'add formulas to newly added submetrics and totals on dashboard
        With .Rows(Metric.SubMetrics.Count + r)
            .Columns(Letter(TemplateColumns.jan) & ":" & Letter(TemplateColumns.dec)).Formula = "=if(Settings!T$2=TRUE,sum(C" & r & ":C" & Metric.SubMetrics.Count + r - 1 & "),"""")" 'total months
            .Columns(TemplateColumns.PrevYearRestated).Formula = "=sum(B" & r & ":B" & Metric.SubMetrics.Count + r - 1 & ")"                                                           'total previous year restated
            .Columns(Letter(TemplateColumns.ytd) & ":" & Letter(TemplateColumns.Target)).Formula = "=sum(O" & r & ":O" & Metric.SubMetrics.Count + r - 1 & ")"                         'total ytd and target
            .Columns(Letter(TemplateColumns.SubMetrics) & ":" & Letter(TemplateColumns.Target)).Copy                                                                                   'copy formating
        End With
        .Rows(r & ":" & Metric.SubMetrics.Count + r).Columns(Letter(TemplateColumns.SubMetrics) & ":" & Letter(TemplateColumns.Target)).PasteSpecial xlPasteFormats                    'paste formating

        'change chart type, must be explicitly activated first before changing
        .ChartObjects("metricChart").Activate
        ActiveChart.chartType = Metric.Chart
    End With

    'add references on scorecard to dashboard
    With scorecard.Rows(lastRowScorecard)
        .Columns(Letter(ScorecardColorColumns.PrevYearRestated) & ":" & Letter(ScorecardColorColumns.ytd)).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColorColumns.PrevYearRestated) & r               'add reference to all color codes
        .Columns(Letter(ScorecardColorColumns.red) & ":" & Letter(ScorecardColorColumns.green)).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColorColumns.red) & r                                       'add reference to color legend
        .Columns(Letter(ScorecardColumns.PrevYearRestated) & ":" & Letter(ScorecardColumns.Target)).Formula = "='" & Metric.Name & "'!" & Letter(TemplateColumns.PrevYearRestated) & Metric.SubMetrics.Count + r 'add reference to values
    End With

    'modify the conditional formating rules
    Call editConditionalFormating(scorecard, scorecard.Range("D" & FIRST_ROW_SCORECARD), scorecard.Range("D" & FIRST_ROW_SCORECARD & ":" & "Q" & lastRowScorecard))
    Call editConditionalFormating(dashboard, dashboard.Range("B" & Metric.SubMetrics.Count + r), dashboard.Range("B" & Metric.SubMetrics.Count + r & ":" & "O" & Metric.SubMetrics.Count + r))

    'clean up for presentation
    'hide color matrix on scorecard and new dahsboard
    scorecard.Range(COL_COLORS).EntireColumn.Hidden = True
    dashboard.Range(COL_COLORS).EntireColumn.Hidden = True

    Application.Goto scorecard.Range(Letter(ScorecardColumns.Metric) & lastRowScorecard)

    NormalProcessing
End Sub
