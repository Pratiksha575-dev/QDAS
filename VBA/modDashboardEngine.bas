Attribute VB_Name = "modDashboardEngine"
Option Explicit

' ==============================================================================
' GLOBAL DESIGN TOKENS & RESPONSIVE METRIC CONSTANTS
' ==============================================================================
Private Const DASH_FONT As String = "Segoe UI"
Private Const COLOR_HEADER_BG As Long = 7949855     ' Modern Deep Blue Accent (#1F4E79)
Private Const COLOR_CARD_BG As Long = 16250871      ' Clean Tech Off-White/Light Gray (#F3F4F6)
Private Const COLOR_CARD_BORDER As Long = 14211288  ' Premium Light Accent Border (#DDDDE5)
Private Const COLOR_TEXT_DARK As Long = 2105376     ' Deep Charcoal Slate for contrast (#202020)

' Proportional Form Chart Sizing Definitions
Private Const CHART_WIDTH_PX As Double = 620
Private Const CHART_HEIGHT_PX As Double = 315
Private Const SPACING_PX As Double = 25
Private Const DASHBOARD_COLUMNS As String = "A:AF"

' ==============================================================================
' 1. MAIN GENERATOR INTERFACE
' ==============================================================================
Public Sub GenerateDashboard()
    On Error GoTo ErrorHandler
    
    ' High-Performance Active State Initialization Caching
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Phase 1: Refresh structural sheet canvas objects
    Dim wsDash As Worksheet
    Set wsDash = CreateDashboardSheet()
    
    ' Phase 2: Structural Canvas Layout Setup
    CreateDashboardLayout wsDash
    
    ' Phase 3: Instantiate Extended Executive KPI Metadata Block
    If ShowDashKPI Then
        CreateKPICards wsDash
    End If
    
    ' Phase 4: Dynamic Decoupled Chart Assembly & Selection Verification
    Dim chartCollection As Collection
    Set chartCollection = New Collection
    
    Dim chObj As ChartObject
    
    If ShowDashTrend Then
    CreateTrendCharts wsDash, chartCollection
    End If
    
    If ShowDashCategory Then
        Set chObj = CreateCategoryChart(wsDash)
        If Not chObj Is Nothing Then chartCollection.Add chObj
    End If
    
    If ShowDashDistribution Then
        Set chObj = CreateDistributionChart(wsDash)
        If Not chObj Is Nothing Then chartCollection.Add chObj
    End If
    
    If ShowDashAnomaly Then
        Set chObj = CreateAnomalyChart(wsDash)
        If Not chObj Is Nothing Then chartCollection.Add chObj
    End If
    
    ' Phase 5: Run Responsive Grid Matrix Placement Algorithm
    PositionCharts wsDash, chartCollection
    
    ' Phase 6: Uniform Visual Polish and Column Scaling
    FormatDashboard wsDash
    
    ' Re-engage Environment Processing System Threads Safely
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "QDAS Executive Responsive Dashboard Generated Successfully!", vbInformation, "System Optimization Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Critical Dashboard Compilation Failure: " & Err.Description, vbCritical, "Execution Exception Trap"
End Sub

' ==============================================================================
' 2. INITIALIZATION & WORKSPACE ARCHITECTS
' ==============================================================================
Private Function CreateDashboardSheet() As Worksheet
    Dim ws As Worksheet
    Dim ch As ChartObject
    Dim currentSheet As Worksheet
    
    Set currentSheet = ActiveSheet
    
    On Error Resume Next
    Set ws = Worksheets("Dashboard")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
        ws.Name = "Dashboard"
    Else
        ' Clear data cells and legacy charts cleanly without tearing reference maps
        ws.Cells.Clear
        ws.Cells.ClearFormats
        For Each ch In ws.ChartObjects
            ch.Delete
        Next ch
    End If
    
    ' Turn off gridlines safely via ActiveWindow focus strategy
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Return view focus seamlessly back to the origin context layer
    If Not currentSheet Is Nothing Then currentSheet.Activate
    Set CreateDashboardSheet = ws
End Function

Private Sub CreateDashboardLayout(wsDash As Worksheet)
    ' Setup Fluid Grid Header Frame spanning across column widths (A:AF)
    Dim headerRange As Range
    Set headerRange = wsDash.Range("A1:AF2")
    headerRange.Merge
    
    With headerRange
        .Value = "QUICK DATA ANALYTICS SYSTEM (QDAS) EXECUTIVE METRIC DASHBOARD"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = DASH_FONT
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = vbWhite
        .Interior.Color = COLOR_HEADER_BG
    End With
    
    wsDash.Rows(1).RowHeight = 22
    wsDash.Rows(2).RowHeight = 22
    wsDash.Rows(3).RowHeight = 15 ' Spacing row gap
End Sub

' ==============================================================================
' 3. COMPREHENSIVE CARD ARCHITECT (KPI ENGINE HOOKS)
' ==============================================================================
Private Sub CreateKPICards(wsDash As Worksheet)

    Dim wsKPI As Worksheet
    Dim tbl As ListObject

    Set wsKPI = VerifySheet("PQ_KPI")
    If wsKPI Is Nothing Then Exit Sub

    Set tbl = VerifyTable(wsKPI)
    If tbl Is Nothing Then Exit Sub

    If tbl.DataBodyRange.Rows.count = 0 Then Exit Sub

    Dim blockWidth As Integer
    blockWidth = 8

    WriteKPICard wsDash, 4, 1, 6, 8, _
        "TOTAL" & vbCrLf & _
        Format(tbl.DataBodyRange.Cells(1, 3).Value, "#,##0.00")

    WriteKPICard wsDash, 4, 9, 6, 16, _
        "AVERAGE" & vbCrLf & _
        Format(tbl.DataBodyRange.Cells(1, 4).Value, "#,##0.00")

    WriteKPICard wsDash, 4, 17, 6, 24, _
        "MINIMUM" & vbCrLf & _
        Format(tbl.DataBodyRange.Cells(1, 5).Value, "#,##0.00")

    WriteKPICard wsDash, 4, 25, 6, 32, _
        "MAXIMUM" & vbCrLf & _
        Format(tbl.DataBodyRange.Cells(1, 6).Value, "#,##0.00")

End Sub
Private Sub WriteKPICard(ws As Worksheet, rStart As Integer, cStart As Integer, rEnd As Integer, cEnd As Integer, content As String)
    Dim cardRange As Range
    Set cardRange = ws.Range(ws.Cells(rStart, cStart), ws.Cells(rEnd, cEnd))
    cardRange.Merge
    
    With cardRange
        .Value = content
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Name = DASH_FONT
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_DARK
        .Interior.Color = COLOR_CARD_BG
        
        ' Soft modern card framing borders definitions
        .Borders(xlEdgeLeft).Color = COLOR_CARD_BORDER
        .Borders(xlEdgeTop).Color = COLOR_CARD_BORDER
        .Borders(xlEdgeRight).Color = COLOR_CARD_BORDER
        .Borders(xlEdgeBottom).Color = COLOR_CARD_BORDER
    End With
End Sub

' ==============================================================================
' 4. DECOUPLED STRUCTURAL INDEPENDENT CHARTING FUNCTIONS
' ==============================================================================
Private Sub CreateTrendCharts(wsDash As Worksheet, chartCollection As Collection)

    Dim Trend As Variant
    Dim chObj As ChartObject

    For Each Trend In Trends

        Set chObj = CreateSingleTrendChart(wsDash, CStr(Trend))

        If Not chObj Is Nothing Then
            chartCollection.Add chObj
        End If

    Next Trend

End Sub
Private Function CreateSingleTrendChart(wsDash As Worksheet, _
                                        ByVal TrendType As String) As ChartObject

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim chObj As ChartObject

    Set wsData = VerifySheet("Trend_" & TrendType)

    If wsData Is Nothing Then Exit Function

    Set tbl = VerifyTable(wsData)

    If tbl Is Nothing Then Exit Function

    Set chObj = wsDash.ChartObjects.Add(10, 10, CHART_WIDTH_PX, CHART_HEIGHT_PX)

    With chObj.Chart

        .ChartType = xlLineMarkers
        .SetSourceData tbl.Range

        .HasTitle = True
        .ChartTitle.Text = tbl.ListColumns(2).Name & _
                           " Trend by " & tbl.ListColumns(1).Name

        If .SeriesCollection.count > 2 Then

            On Error Resume Next

            .SeriesCollection(2).AxisGroup = xlSecondary

            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = _
                "Secondary Multi-Measure Metrics"

            On Error GoTo 0

        End If

        ApplyBaseChartStyling _
            chObj.Chart, _
            True, _
            tbl.ListColumns(1).Name, _
            "Value"

    End With

    Set CreateSingleTrendChart = chObj

End Function

Private Function CreateCategoryChart(wsDash As Worksheet) As ChartObject
    Dim wsData As Worksheet: Set wsData = VerifySheet("Category_Data")
    If wsData Is Nothing Then Exit Function
    Dim tbl As ListObject: Set tbl = VerifyTable(wsData)
    If tbl Is Nothing Then Exit Function
    
    ' Isolate and structural filter columns targeting explicit measures to exclude background indices
    Dim axisCol As ListColumn: Set axisCol = tbl.ListColumns(1)
    Dim structuralRange As Range: Set structuralRange = axisCol.Range
    
    Dim item As Variant, ColName As String, validColumnFound As Boolean
    validColumnFound = False
    
    On Error Resume Next
    For Each item In Measures
        ColName = "Total_" & CStr(item) ' Follows Power Query standard aggregation prefixes
        Dim valCol As ListColumn
        Set valCol = tbl.ListColumns(ColName)
        If valCol Is Nothing Then Set valCol = tbl.ListColumns(CStr(item)) ' Fallback check
        
        If Not valCol Is Nothing Then
            Set structuralRange = Union(structuralRange, valCol.Range)
            validColumnFound = True
        End If
    Next item
    On Error GoTo 0
    
    If Not validColumnFound Then Set structuralRange = tbl.Range
    
    Dim chObj As ChartObject
    Set chObj = wsDash.ChartObjects.Add(10, 10, CHART_WIDTH_PX, CHART_HEIGHT_PX)
    
    With chObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData structuralRange
        .HasTitle = True
        .ChartTitle.Text = tbl.ListColumns(2).Name & _
                   " by " & axisCol.Name
        ApplyBaseChartStyling chObj.Chart, True, axisCol.Name, "Aggregated Core Volumes"
    End With
    
    Set CreateCategoryChart = chObj
End Function

Private Function CreateDistributionChart(wsDash As Worksheet) As ChartObject
    Dim wsData As Worksheet: Set wsData = VerifySheet("Distribution_Data")
    If wsData Is Nothing Then Exit Function
    Dim tbl As ListObject: Set tbl = VerifyTable(wsData)
    If tbl Is Nothing Then Exit Function
    
    ' Strict Column Selection: Isolates structural elements avoiding cross-talk metrics
    Dim binStartCol As ListColumn, freqCol As ListColumn
    On Error Resume Next
    Set binStartCol = tbl.ListColumns("BinStart")
    Set freqCol = tbl.ListColumns("Frequency")
    On Error GoTo 0
    
    If binStartCol Is Nothing Or freqCol Is Nothing Then Exit Function
    
    Dim chObj As ChartObject
    Set chObj = wsDash.ChartObjects.Add(10, 10, CHART_WIDTH_PX, CHART_HEIGHT_PX)
    
    With chObj.Chart
        .ChartType = xlColumnClustered
        
        ' Clear implicit structures from chart container area
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        Dim srs As Series
        Set srs = .SeriesCollection.NewSeries
        srs.Name = "Observation Frequency"
        srs.XValues = binStartCol.DataBodyRange
        srs.Values = freqCol.DataBodyRange
        
        ' Scale column proximity gap limits to accurately resemble a standard density histogram profile
        .ChartGroups(1).GapWidth = 3
        
        .HasTitle = True
        .ChartTitle.Text = "Calculated Value Density Distribution Histogram"
        ApplyBaseChartStyling chObj.Chart, False, "Structural Bin Intervals", "Data Frequency Distribution"
    End With
    
    Set CreateDistributionChart = chObj
End Function

Private Function CreateAnomalyChart(wsDash As Worksheet) As ChartObject
    ' Target correct destination source layer mapping options cleanly
    Dim wsData As Worksheet: Set wsData = VerifySheet("PQ_AnomalySummary")
    If wsData Is Nothing Then Set wsData = VerifySheet("Anomaly_Data")
    If wsData Is Nothing Then Exit Function
    
    Dim tbl As ListObject: Set tbl = VerifyTable(wsData)
    If tbl Is Nothing Then Exit Function
    
    Dim measureCol As ListColumn, totalOutliersCol As ListColumn
    On Error Resume Next
    Set measureCol = tbl.ListColumns("Measure")
    Set totalOutliersCol = tbl.ListColumns("TotalOutliers")
    On Error GoTo 0
    
    If measureCol Is Nothing Or totalOutliersCol Is Nothing Then Exit Function
    
    Dim chObj As ChartObject
    Set chObj = wsDash.ChartObjects.Add(10, 10, CHART_WIDTH_PX, CHART_HEIGHT_PX)
    
    With chObj.Chart
        .ChartType = xlBarClustered
        
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        Dim srs As Series
        Set srs = .SeriesCollection.NewSeries
        srs.Name = "Total Outliers Detected"
        srs.XValues = measureCol.DataBodyRange
        srs.Values = totalOutliersCol.DataBodyRange
        
        .HasTitle = True
        .ChartTitle.Text = "Statistical Deviations & Outlier Volume Summary"
        ApplyBaseChartStyling chObj.Chart, False, "Deviation Occurrences", "Evaluated Target Measures"
        
        ' Automatically build detailed data matrix summary tables downstream safely out of layout paths
        WriteAnomalyMetaSummary wsDash, tbl
    End With
    
    Set CreateAnomalyChart = chObj
End Function

' ==============================================================================
' 5. RESPONSIVE ENGINE GRID MATRIX PLACEMENT ALGORITHM
' ==============================================================================
Private Sub PositionCharts(wsDash As Worksheet, charts As Collection)
    Dim count As Integer: count = charts.count
    If count = 0 Then Exit Sub
    
    ' Dynamic matrix start index clearance baseline definition following KPI structural card frames
    Dim startRow As Double: startRow = 13
    Dim leftMarginPx As Double: leftMarginPx = wsDash.Cells(startRow, "A").Left + 15
    
    ' Calculate current full workspace column grid bounds pixel constraints fluidly
    Dim totalDashWidthPx As Double
    Dim cellIter As Range
    For Each cellIter In wsDash.Range(DASHBOARD_COLUMNS).Rows(1).Cells
        totalDashWidthPx = totalDashWidthPx + cellIter.Width
    Next cellIter
    
    Dim ch1 As ChartObject, ch2 As ChartObject, ch3 As ChartObject, ch4 As ChartObject
    
    Select Case count
        Case 1
            ' Form Arrangement: Center single standalone visualization object inside container frame matrix
            Set ch1 = charts(1)
            ch1.Left = leftMarginPx + ((totalDashWidthPx - CHART_WIDTH_PX) / 2)
            ch1.Top = wsDash.Cells(startRow, "A").Top
            
        Case 2
            ' Form Arrangement: Dynamic parallel side-by-side split row matching logic
            Set ch1 = charts(1)
            Set ch2 = charts(2)
            
            ch1.Left = leftMarginPx + 40
            ch1.Top = wsDash.Cells(startRow, "A").Top
            
            ch2.Left = ch1.Left + CHART_WIDTH_PX + SPACING_PX
            ch2.Top = ch1.Top
            
        Case 3
            ' Form Arrangement: Dynamic structural pyramid (2 Horizontal Parallel on row 1, 1 Centered on Row 2)
            Set ch1 = charts(1)
            Set ch2 = charts(2)
            Set ch3 = charts(3)
            
            ch1.Left = leftMarginPx + 40
            ch1.Top = wsDash.Cells(startRow, "A").Top
            
            ch2.Left = ch1.Left + CHART_WIDTH_PX + SPACING_PX
            ch2.Top = ch1.Top
            
            ch3.Left = leftMarginPx + ((totalDashWidthPx - CHART_WIDTH_PX) / 2)
            ch3.Top = ch1.Top + CHART_HEIGHT_PX + SPACING_PX
            
        Case 4
            ' Form Arrangement: Standard balanced matrix grid quadrant layout (2x2 Structure Allocation)
            Set ch1 = charts(1)
            Set ch2 = charts(2)
            Set ch3 = charts(3)
            Set ch4 = charts(4)
            
            ' Row Grid Generation Level 1
            ch1.Left = leftMarginPx + 40
            ch1.Top = wsDash.Cells(startRow, "A").Top
            ch2.Left = ch1.Left + CHART_WIDTH_PX + SPACING_PX
            ch2.Top = ch1.Top
            
            ' Row Grid Generation Level 2
            ch3.Left = ch1.Left
            ch3.Top = ch1.Top + CHART_HEIGHT_PX + SPACING_PX
            ch4.Left = ch2.Left
            ch4.Top = ch3.Top
    End Select
End Sub

' ==============================================================================
' 6. DASHBOARD STYLING & AUXILIARY HELPERS
' ==============================================================================
Private Sub ApplyBaseChartStyling(ch As Chart, showLegend As Boolean, xAxisTitle As String, yAxisTitle As String)
    With ch
        ' Enforce strict modern design token logic (white space backgrounds, border concealment)
        .ChartArea.Format.Fill.ForeColor.RGB = vbWhite
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = vbWhite
        .PlotArea.Format.Line.Visible = msoFalse
        
        ' Global Corporate Typography Synchronization
        With .ChartTitle
            .Font.Name = DASH_FONT
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = COLOR_TEXT_DARK
        End With
        
        ' Primary Category Axis Formatting Architecture
        If .HasAxis(xlCategory) Then
            With .Axes(xlCategory)
                .HasTitle = (xAxisTitle <> "")
                If .HasTitle Then
                    .AxisTitle.Text = xAxisTitle
                    .AxisTitle.Font.Name = DASH_FONT
                    .AxisTitle.Font.Size = 9
                    .AxisTitle.Font.Bold = False
                End If
                .TickLabels.Font.Name = DASH_FONT
                .TickLabels.Font.Size = 8.5
            End With
        End If
        
        ' Primary Operational Range Metric Axis Formatting Architecture
        If .HasAxis(xlValue) Then
            With .Axes(xlValue)
                .HasTitle = (yAxisTitle <> "")
                If .HasTitle Then
                    .AxisTitle.Text = yAxisTitle
                    .AxisTitle.Font.Name = DASH_FONT
                    .AxisTitle.Font.Size = 9
                    .AxisTitle.Font.Bold = False
                End If
                .TickLabels.Font.Name = DASH_FONT
                .TickLabels.Font.Size = 8.5
                .HasMajorGridlines = True
                .MajorGridlines.Format.Line.ForeColor.RGB = RGB(242, 242, 247) ' Soft horizontal rule accents
            End With
        End If
        
        ' Smart Legend Visibility Optimization Matrix
        .HasLegend = showLegend
        If .HasLegend Then
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Name = DASH_FONT
            .Legend.Font.Size = 8.5
            .Legend.Format.Line.Visible = msoFalse
        End If
    End With
End Sub

Private Sub FormatDashboard(wsDash As Worksheet)
    ' Fluidly auto-scale dashboard grid canvas footprint dimensions without clip blocks
    On Error Resume Next
    wsDash.Range(DASHBOARD_COLUMNS).Columns.AutoFit
    On Error GoTo 0
End Sub

Private Sub WriteAnomalyMetaSummary(wsDash As Worksheet, tbl As ListObject)
    ' Safe positional coordinate definitions to prevent asset collisions down sheet grid surface
    Dim outputRow As Long: outputRow = 38
    Dim startCol As Integer: startCol = 2 ' Sets to structural column matrix coordinate Index B
    
    Dim targetRange As Range
    Set targetRange = wsDash.Cells(outputRow, startCol)
    
    targetRange.Value = "Granular High/Low Outlier Statistical Breakdown"
    targetRange.Font.Bold = True
    targetRange.Font.Size = 11
    targetRange.Font.Name = DASH_FONT
    targetRange.Font.Color = COLOR_TEXT_DARK
    
    ' Safely clone raw detailed metadata tables downwards to the sheet grid
    On Error Resume Next
    tbl.Range.Copy
    With wsDash.Cells(outputRow + 1, startCol)
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

' ==============================================================================
' 7. HIGH-FIDELITY SINGLE PAGE EXECUTIVE PDF EXPORT ENGINE
' ==============================================================================
Public Sub ExportDashboardToPDF()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = Worksheets("Dashboard")
    On Error GoTo 0
    
    If wsDash Is Nothing Then
        MsgBox "Dashboard workspace sheet connection missing! Generate the engine view canvas first.", vbExclamation, "System Alert"
        Exit Sub
    End If
    
    ' Invoke OS native Save File Dialog windows explorer panel cleanly
    Dim fileSavePath As Variant
    fileSavePath = Application.GetSaveAsFilename( _
        InitialFileName:="QDAS_Executive_Analytics_Report", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Destination Canvas Path For PDF Publication")
        
    If fileSavePath = False Then Exit Sub ' User closed pane thread
    
    On Error GoTo ExportFailure
    
    ' Configure professional print mapping bounds (Single Page Landscape Alignment)
    With wsDash.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
    End With
    
    ' Trigger high-fidelity internal spreadsheet-to-file rendering sequence macro
    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=fileSavePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    Exit Sub

ExportFailure:
    MsgBox "PDF compilation layout engine processing failure: " & Err.Description, vbCritical, "System Core Exception"
End Sub

' ==============================================================================
' 8. ROBUST ERROR TRAPPING DATA VALIDATORS
' ==============================================================================
Private Function VerifySheet(SheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(SheetName)
    On Error GoTo 0
    Set VerifySheet = ws
End Function

Private Function VerifyTable(ws As Worksheet) As ListObject
    Dim tbl As ListObject
    If ws.ListObjects.count > 0 Then
        Set tbl = ws.ListObjects(1)
        If Not tbl.DataBodyRange Is Nothing Then
            If tbl.DataBodyRange.Rows.count > 0 Then
                Set VerifyTable = tbl
                Exit Function
            End If
        End If
    End If
    Set VerifyTable = Nothing
End Function

Private Function JoinCollection(col As Collection) As String
    Dim item As Variant
    Dim txt As String
    If col Is Nothing Then Exit Function
    
    For Each item In col
        txt = txt & item & ", "
    Next item
    
    If Len(txt) > 2 Then txt = Left(txt, Len(txt) - 2)
    JoinCollection = txt
End Function

