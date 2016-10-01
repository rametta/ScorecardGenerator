Option Explicit

' PRIVATES
Private mName            As String
Private mSampleSize      As String
Private mPrime           As String
Private mChart           As String
Private mColors(1 To 12) As String
Private mPercentage      As Boolean
Private mDecimalsDisplay As String

' PUBLICS
Public SubMetrics        As Range

' METHODS
Public Function Valid() As Boolean
    If Me.Name = "" Or _
    Me.colors(Legend.red) = "" Or _
    Me.colors(Legend.yellow) = "" Or _
    Me.colors(Legend.yelhigh) = "" Or _
    Me.colors(Legend.green) = "" Or _
    Me.colors(Legend.RedDisplay) = "" Or _
    Me.colors(Legend.YellowDisplay) = "" Or _
    Me.colors(Legend.YelhighDisplay) = "" Or _
    Me.colors(Legend.GreenDisplay) = "" Then
        Valid = False
    Else
        Valid = True
    End If
End Function

' GETTERS / SETTERS

' Name
Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(Value As String)
    mName = Value
End Property

' Sample Size
Public Property Get SampleSize() As String
    SampleSize = mSampleSize
End Property
Public Property Let SampleSize(Value As String)
    mSampleSize = Value
End Property

' Business Prime
Public Property Get Prime() As String
    Prime = mPrime
End Property
Public Property Let Prime(Value As String)
    mPrime = Value
End Property

' Chart Type
Public Property Get Chart() As String
    Chart = mChart
End Property
Public Property Let Chart(Value As String)
    Select Case Value
        Case "Vertical Bar"
            mChart = xlColumnClustered
        Case "Horizontal Bar"
            mChart = xlBarClustered
        Case "Line"
            mChart = xlLine
        Case "Area"
            mChart = xlArea
        Case "Pie"
            mChart = xlPie
        Case "Donut"
            mChart = xlDoughnut
        Case Else
            mChart = xlColumnClustered
    End Select
End Property

' Percentage
Public Property Get isPercentage() As Boolean
    isPercentage = mPercentage
End Property
Public Property Let isPercentage(Value As Boolean)
    mPercentage = Value
End Property

' Colors
Public Property Get colors(color As Byte) As String
    colors = mColors(color)
End Property
Public Property Let colors(color As Byte, Value As String)
    mColors(color) = Value
End Property

' Decimals Displayed
Public Property Get DecimalsDisplay() As String
    DecimalsDisplay = mDecimalsDisplay
End Property
Public Property Let DecimalsDisplay(Value As String)
    If Value = "" Then
        mDecimalsDisplay = 0
    Else
        mDecimalsDisplay = Value
    End If
End Property
