Attribute VB_Name = "modDashboardEngine"
Option Explicit

Sub GenerateDashboard()

    Application.ScreenUpdating = False

    CreateDashboardSheet
    CreateDashboardLayout

    If ShowDashTrend Then
        CreateChart "Trend_Data", _
                    "Trend Analysis", _
                    xlLineMarkers, _
                    20, 170, 520, 260
    End If

    If ShowDashCategory Then
        CreateChart "Category_Data", _
                    "Category Analysis", _
                    xlColumnClustered, _
                    570, 170, 520, 260
    End If

    If ShowDashDistribution Then
        CreateDistributionChart
    End If

    If ShowDashAnomaly Then
        CreateChart "Anomaly_Data", _
                    "Anomaly Analysis", _
                    xlBarClustered, _
                    520, 470, 520, 260
    End If

    If ShowDashKPI Then
        CreateKPICards
    End If

    Application.ScreenUpdating = True

    MsgBox "Dashboard Generated Successfully!"

End Sub

Sub CreateDashboardSheet()

    Dim ws As Worksheet
    Dim obj As ChartObject

    On Error Resume Next
    Set ws = Worksheets("Dashboard")
    On Error GoTo 0

    If ws Is Nothing Then

        Worksheets.Add After:=Worksheets(Worksheets.count)
        ActiveSheet.Name = "Dashboard"
        Set ws = Worksheets("Dashboard")

    Else

        ws.Cells.Clear

        For Each obj In ws.ChartObjects
            obj.Delete
        Next obj

    End If

End Sub

Sub CreateDashboardLayout()

    Dim ws As Worksheet

    Set ws = Worksheets("Dashboard")

    ws.Range("A1:V2").Merge

    With ws.Range("A1:V2")

        .Value = "QDAS ANALYTICS DASHBOARD"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 20
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite

    End With

    ws.Range("A4:B6").Merge
    ws.Range("C4:D6").Merge
    ws.Range("E4:F6").Merge
    ws.Range("G4:H6").Merge
    ws.Range("I4:J6").Merge
    ws.Range("K4:L6").Merge

    ws.Range("A8:F20").Interior.Color = RGB(245, 245, 245)
    ws.Range("H8:M20").Interior.Color = RGB(245, 245, 245)

    ws.Range("A22:F34").Interior.Color = RGB(245, 245, 245)
    ws.Range("H22:M34").Interior.Color = RGB(245, 245, 245)

End Sub

Sub CreateChart(SheetName As String, _
                ChartTitle As String, _
                ChartType As XlChartType, _
                LeftPos As Double, _
                TopPos As Double, _
                ChartWidth As Double, _
                ChartHeight As Double)

    Dim wsDash As Worksheet
    Dim wsData As Worksheet

    Dim tbl As ListObject
    Dim ch As ChartObject

    Set wsDash = Worksheets("Dashboard")
    Set wsData = Worksheets(SheetName)

    If wsData.ListObjects.count = 0 Then Exit Sub

    Set tbl = wsData.ListObjects(1)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set ch = wsDash.ChartObjects.Add( _
                Left:=LeftPos, _
                Top:=TopPos, _
                Width:=ChartWidth, _
                Height:=ChartHeight)

    With ch.Chart

        .ChartType = ChartType

        .SetSourceData tbl.Range

        .HasTitle = True

        .ChartTitle.Text = ChartTitle

        .Legend.Position = xlLegendPositionBottom

        If .SeriesCollection.count > 1 Then
            On Error Resume Next
            .SeriesCollection(2).AxisGroup = xlSecondary
            On Error GoTo 0
        End If

    End With

End Sub
Sub CreateKPICards()

    Dim ws As Worksheet

    Set ws = Worksheets("Dashboard")

    With ws

        .Range("A4:B6").Value = _
            "Dataset" & vbCrLf & ActiveWorkbook.Name

        .Range("C4:D6").Value = _
            "Measures" & vbCrLf & Measures.count

        .Range("E4:F6").Value = _
            "Categories" & vbCrLf & Categories.count

        .Range("G4:H6").Value = _
            "Trend" & vbCrLf & JoinCollection(Trends)

        .Range("I4:J6").Value = _
            "Anomaly" & vbCrLf & SelectedAnomaly

        .Range("K4:L6").Value = _
            "Bins" & vbCrLf & DistributionBins

        With .Range("A4:L6")

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True

            .Font.Bold = True
            .Font.Size = 11

            .Interior.Color = RGB(221, 235, 247)

            .Borders.LineStyle = xlContinuous

        End With

    End With

End Sub
Function JoinCollection(col As Collection) As String

    Dim Item As Variant
    Dim txt As String

    txt = ""

    For Each Item In col

        txt = txt & Item & ", "

    Next Item

    If Len(txt) > 2 Then

        txt = Left(txt, Len(txt) - 2)

    End If

    JoinCollection = txt

End Function

Sub CreateDistributionChart()

    Dim wsDash As Worksheet
    Dim wsData As Worksheet

    Dim tbl As ListObject
    Dim ch As ChartObject

    Set wsDash = Worksheets("Dashboard")
    Set wsData = Worksheets("Distribution_Data")

    If wsData.ListObjects.count = 0 Then Exit Sub

    Set tbl = wsData.ListObjects(1)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set ch = wsDash.ChartObjects.Add( _
                Left:=20, _
                Top:=470, _
                Width:=520, _
                Height:=260)

    With ch.Chart

        .ChartType = xlColumnClustered

        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        .SeriesCollection.NewSeries

        .SeriesCollection(1).Name = "Frequency"

        .SeriesCollection(1).Values = tbl.ListColumns("Frequency").DataBodyRange

        .SeriesCollection(1).XValues = tbl.ListColumns("BinStart").DataBodyRange

        .HasTitle = True

        .ChartTitle.Text = "Distribution"

        .Legend.Delete

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Bin Start"

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Frequency"

    End With

End Sub
