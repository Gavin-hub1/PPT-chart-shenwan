Attribute VB_Name = "ChartFormatter"
Option Explicit

Private Const REFERENCE_BLUE As Long = 11802661  ' RGB(37, 24, 180)
Private Const ORANGE As Long = 3243501      ' RGB(237, 125, 49)
Private Const LABEL_GRAY As Long = 4210752   ' RGB(64, 64, 64)

Private Const XL_CATEGORY As Long = 1
Private Const XL_VALUE As Long = 2
Private Const XL_PRIMARY As Long = 1
Private Const XL_SECONDARY As Long = 2
Private Const XL_TICK_MARK_OUTSIDE As Long = 3
Private Const XL_TICK_MARK_NONE As Long = -4142
Private Const XL_LEGEND_POSITION_BOTTOM As Long = -4107
Private Const XL_DATA_LABEL_POSITION_ABOVE As Long = 0
Private Const MsoTrue As Long = -1

' Ribbon callback and direct macro entry point.
Public Sub FormatSelectedCharts(Optional ByVal control As Object)
    Dim sel As Selection
    Dim shapes As ShapeRange
    Dim shp As Shape
    Dim handled As Long
    Dim skipped As Long

    On Error GoTo NoSelection
    Set sel = ActiveWindow.Selection
    Set shapes = sel.ShapeRange
    On Error GoTo 0

    For Each shp In shapes
        If ShapeContainsChart(shp) Then
            FormatChart shp.Chart
            handled = handled + 1
        Else
            skipped = skipped + 1
        End If
    Next shp

    ShowResult handled, skipped
    Exit Sub

NoSelection:
    MsgBox "Select one or more chart objects, then click Format Chart.", vbInformation, "SW Research Format"
End Sub

Private Function ShapeContainsChart(ByVal shp As Shape) As Boolean
    On Error Resume Next
    ShapeContainsChart = (shp.HasChart = MsoTrue)
    On Error GoTo 0
End Function

Private Sub FormatChart(ByVal cht As Chart)
    On Error Resume Next

    cht.HasTitle = False

    cht.HasLegend = True
    cht.Legend.Position = XL_LEGEND_POSITION_BOTTOM
    cht.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_GRAY

    FormatAxis cht, XL_CATEGORY, XL_PRIMARY, True, False
    FormatAxis cht, XL_VALUE, XL_PRIMARY, True, False
    FormatAxis cht, XL_VALUE, XL_SECONDARY, True, False
    RemoveGridlines cht, XL_CATEGORY, XL_PRIMARY
    RemoveGridlines cht, XL_VALUE, XL_PRIMARY
    RemoveGridlines cht, XL_VALUE, XL_SECONDARY

    ApplyReferenceSeriesStyle cht

    On Error GoTo 0
End Sub

Private Sub FormatAxis(ByVal cht As Chart, ByVal axisType As Long, ByVal axisGroup As Long, ByVal majorOutside As Boolean, ByVal minorOutside As Boolean)
    Dim ax As Axis

    On Error Resume Next
    Set ax = cht.Axes(axisType, axisGroup)
    If ax Is Nothing Then Exit Sub

    If majorOutside Then ax.MajorTickMark = XL_TICK_MARK_OUTSIDE
    If minorOutside Then
        ax.MinorTickMark = XL_TICK_MARK_OUTSIDE
    Else
        ax.MinorTickMark = XL_TICK_MARK_NONE
    End If

    ax.Format.Line.Visible = MsoTrue
    ax.Format.Line.ForeColor.RGB = REFERENCE_BLUE
    ax.Format.Line.Weight = 1
    ax.TickLabels.Font.Color = REFERENCE_BLUE
    ax.TickLabels.Font.Size = 11
    On Error GoTo 0
End Sub

Private Sub RemoveGridlines(ByVal cht As Chart, ByVal axisType As Long, ByVal axisGroup As Long)
    Dim ax As Axis

    On Error Resume Next
    Set ax = cht.Axes(axisType, axisGroup)
    If ax Is Nothing Then Exit Sub

    ax.HasMajorGridlines = False
    ax.HasMinorGridlines = False
    On Error GoTo 0
End Sub

Private Sub ApplyReferenceSeriesStyle(ByVal cht As Chart)
    Dim seriesIndex As Long
    Dim sr As Series

    On Error Resume Next
    For seriesIndex = 1 To cht.SeriesCollection.Count
        Set sr = cht.SeriesCollection(seriesIndex)

        If IsLineChart(sr.ChartType) Then
            StyleLineSeries sr
        ElseIf IsBarOrColumnChart(sr.ChartType) Or IsBarOrColumnChart(cht.ChartType) Then
            StyleColumnSeries sr
        Else
            StyleOtherSeries sr, seriesIndex
        End If
    Next seriesIndex
    On Error GoTo 0
End Sub

Private Sub StyleColumnSeries(ByVal sr As Series)
    On Error Resume Next
    sr.Format.Fill.Visible = MsoTrue
    sr.Format.Fill.ForeColor.RGB = REFERENCE_BLUE
    sr.Format.Line.Visible = MsoTrue
    sr.Format.Line.ForeColor.RGB = REFERENCE_BLUE
    sr.Format.Line.Weight = 0.75

    FormatExistingDataLabels sr
    On Error GoTo 0
End Sub

Private Sub StyleLineSeries(ByVal sr As Series)
    On Error Resume Next
    sr.Format.Line.Visible = MsoTrue
    sr.Format.Line.ForeColor.RGB = ORANGE
    sr.Format.Line.Weight = 3
    sr.MarkerStyle = -4142

    FormatExistingDataLabels sr
    On Error GoTo 0
End Sub

Private Sub StyleOtherSeries(ByVal sr As Series, ByVal seriesIndex As Long)
    On Error Resume Next
    If seriesIndex Mod 2 = 1 Then
        sr.Format.Fill.ForeColor.RGB = REFERENCE_BLUE
        sr.Format.Line.ForeColor.RGB = REFERENCE_BLUE
    Else
        sr.Format.Fill.ForeColor.RGB = ORANGE
        sr.Format.Line.ForeColor.RGB = ORANGE
    End If
    On Error GoTo 0
End Sub

Private Sub FormatExistingDataLabels(ByVal sr As Series)
    On Error Resume Next
    If sr.HasDataLabels Then
        sr.DataLabels.Position = XL_DATA_LABEL_POSITION_ABOVE
        sr.DataLabels.Font.Color = LABEL_GRAY
        sr.DataLabels.Font.Size = 11
    End If
    On Error GoTo 0
End Sub

Private Function IsBarOrColumnChart(ByVal chartType As Long) As Boolean
    Select Case chartType
        Case -4100, 51, 52, 53, 54, 55, 56, _
             57, 58, 59, 60, 61, 62
            IsBarOrColumnChart = True
        Case Else
            IsBarOrColumnChart = False
    End Select
End Function

Private Function IsLineChart(ByVal chartType As Long) As Boolean
    Select Case chartType
        Case 4, 63, 64, 65, 66, 74, 75
            IsLineChart = True
        Case Else
            IsLineChart = False
    End Select
End Function

Private Sub ShowResult(ByVal handled As Long, ByVal skipped As Long)
    If handled = 0 Then
        MsgBox "No chart objects were found in the current selection.", vbInformation, "SW Research Format"
    ElseIf skipped = 0 Then
        MsgBox "Formatted " & handled & " chart(s).", vbInformation, "SW Research Format"
    Else
        MsgBox "Formatted " & handled & " chart(s); skipped " & skipped & " non-chart object(s).", vbInformation, "SW Research Format"
    End If
End Sub
