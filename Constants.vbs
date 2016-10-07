Option Explicit

' Jason Rametta

' Worksheet Names
Public Const SCORECARD_NAME         As String = "Scorecard"
Public Const COLORS_NAME            As String = "Color Coding"
Public Const SETTINGS_NAME          As String = "Settings"
Public Const NEW_METRIC_NAME        As String = "New Metric"
Public Const DELETE_METRIC_NAME     As String = "Delete Metric"
Public Const ADD_DATA_NAME          As String = "Add Data"
Public Const TEMPLATE_NAME          As String = "TEMPLATE"
Public Const DATA_NAME              As String = "Data"

Public Const FIRST_ROW_SCORECARD    As Byte = 5
Public Const FIRST_ROW_DASHBOARD    As Byte = 1
Public Const FIRST_ROW_SETTINGS     As Byte = 7

Public Const FIRST_COL_SCORECARD    As String = "A"
Public Const LAST_COL_SCORECARD     As String = "U"
Public Const FIRST_COL_COLORS       As String = "AA"
Public Const LAST_COL_COLORS        As String = "AS"
Public Const COL_COLORS             As String = FIRST_COL_COLORS & ":" & LAST_COL_COLORS

Public Const BUTTON_PLACEMENT_RIGHT As Byte = 200
Public Const ROW_HEIGHT             As Byte = 22

' Color Coding Legend for array
Public Enum Legend
    red = 1
    yellow
    yelhigh
    green
    RedDisplay
    YellowDisplay
    YelhighDisplay
    GreenDisplay
    RealRedDisplay
    RealYellowDisplay
    RealYelhighDisplay
    RealGreenDisplay
End Enum

' Columns
Public Enum ColorCodingColumns
    Metric = 1
    Yr
    red
    yellow
    yelhigh
    green
    Sample
    Prime
End Enum

Public Enum ScorecardColumns
    Metric = 1
    Prime
    Sample
    PrevYearRestated
    jan
    Feb
    Mar
    Apr
    May
    Jun
    Jul
    Aug
    Sep
    Oct
    Nov
    dec
    ytd
    Target
    red
    yellow
    green
End Enum

Public Enum ScorecardColorColumns
    Metric = 27
    PrevYearRestated
    jan
    Feb
    Mar
    Apr
    May
    Jun
    Jul
    Aug
    Sep
    Oct
    Nov
    dec
    ytd
    red
    yellow
    yelhigh
    green
End Enum

Public Enum TemplateColumns
    SubMetrics = 1
    PrevYearRestated
    jan
    Feb
    Mar
    Apr
    May
    Jun
    Jul
    Aug
    Sep
    Oct
    Nov
    dec
    ytd
    Target
    red = 18
    yellow
    green
End Enum

Public Enum TemplateColorColumns
    Total = 27
    PrevYearRestated
    jan
    Feb
    Mar
    Apr
    May
    Jun
    Jul
    Aug
    Sep
    Oct
    Nov
    dec
    ytd
    red
    yellow
    yelhigh
    green
End Enum

Public Enum DataColumns
    Metric = 1
    Sub_Metric
    Assumption
    Mnth
    Yr
    Actual
    Target
End Enum
