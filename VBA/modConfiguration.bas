Attribute VB_Name = "modConfiguration"
Option Explicit

'==========================
' Configuration Variables
'==========================

Public SelectedDate As String
Public SelectedAnomaly As String

Public DistributionMode As String
Public DistributionBins As Long

Public Measures As Collection
Public Categories As Collection
Public Trends As Collection

Public ShowDashKPI As Boolean
Public ShowDashTrend As Boolean
Public ShowDashCategory As Boolean
Public ShowDashDistribution As Boolean
Public ShowDashAnomaly As Boolean


Sub SaveConfiguration()
    
    Dim ws As Worksheet
    Dim i As Long
    Dim RowNum As Long

    Set ws = Worksheets("Config")

    ws.Cells.Clear

    ws.Range("A1").Value = "Setting"
    ws.Range("B1").Value = "Value"
    ws.Range("C1").Value = "Selected"

    RowNum = 2

    'Date Column
    ws.Cells(RowNum, 1).Value = "Date Column"
    ws.Cells(RowNum, 2).Value = frmAnalysisConfig.cmbDate.Value
    RowNum = RowNum + 1

    'Measures
    For i = 0 To frmAnalysisConfig.lstMeasures.ListCount - 1

        If frmAnalysisConfig.lstMeasures.Selected(i) Then

            ws.Cells(RowNum, 1).Value = "Measure"
            ws.Cells(RowNum, 2).Value = frmAnalysisConfig.lstMeasures.List(i)

            RowNum = RowNum + 1

        End If

    Next i

    'Categories
    For i = 0 To frmAnalysisConfig.lstCategory.ListCount - 1

        If frmAnalysisConfig.lstCategory.Selected(i) Then

            ws.Cells(RowNum, 1).Value = "Category"
            ws.Cells(RowNum, 2).Value = frmAnalysisConfig.lstCategory.List(i)

            RowNum = RowNum + 1

        End If

    Next i

    'Trend Analysis
    If frmAnalysisConfig.chkDaily.Value Then
        ws.Cells(RowNum, 1) = "Trend"
        ws.Cells(RowNum, 2) = "Daily"
        RowNum = RowNum + 1
    End If

    If frmAnalysisConfig.chkWeekly.Value Then
        ws.Cells(RowNum, 1) = "Trend"
        ws.Cells(RowNum, 2) = "Weekly"
        RowNum = RowNum + 1
    End If

    If frmAnalysisConfig.chkMonthly.Value Then
        ws.Cells(RowNum, 1) = "Trend"
        ws.Cells(RowNum, 2) = "Monthly"
        RowNum = RowNum + 1
    End If

    If frmAnalysisConfig.chkHourly.Value Then
        ws.Cells(RowNum, 1) = "Trend"
        ws.Cells(RowNum, 2) = "Hourly"
        RowNum = RowNum + 1
    End If

    If frmAnalysisConfig.chkWeekday.Value Then
        ws.Cells(RowNum, 1) = "Trend"
        ws.Cells(RowNum, 2) = "Weekday"
        RowNum = RowNum + 1
    End If

 'Distribution Mode
  ws.Cells(RowNum, 1).Value = "Distribution Mode"

  If frmAnalysisConfig.optAutoBins.Value Then
       ws.Cells(RowNum, 2).Value = "Auto"
  Else
       ws.Cells(RowNum, 2).Value = "Custom"
  End If

  RowNum = RowNum + 1

  'Distribution Bins
   ws.Cells(RowNum, 1).Value = "Distribution Bins"

   If frmAnalysisConfig.optAutoBins.Value Then
    ws.Cells(RowNum, 2).Value = 20
  Else
    ws.Cells(RowNum, 2).Value = CLng(frmAnalysisConfig.txtBins.Value)
   End If

    RowNum = RowNum + 1

    'Anomaly Method
    ws.Cells(RowNum, 1).Value = "Anomaly"

    If frmAnalysisConfig.optIQR.Value Then
        ws.Cells(RowNum, 2).Value = "IQR"
    Else
        ws.Cells(RowNum, 2).Value = "Z-Score"
    End If
    
    RowNum = RowNum + 1
'Dashboard - KPI
ws.Cells(RowNum, 1).Value = "Dashboard"
ws.Cells(RowNum, 2).Value = "KPI"
ws.Cells(RowNum, 3).Value = frmAnalysisConfig.chkDashKPI.Value
RowNum = RowNum + 1

'Dashboard - Trend
ws.Cells(RowNum, 1).Value = "Dashboard"
ws.Cells(RowNum, 2).Value = "Trend"
ws.Cells(RowNum, 3).Value = frmAnalysisConfig.chkDashTrend.Value
RowNum = RowNum + 1

'Dashboard - Category
ws.Cells(RowNum, 1).Value = "Dashboard"
ws.Cells(RowNum, 2).Value = "Category"
ws.Cells(RowNum, 3).Value = frmAnalysisConfig.chkDashCategory.Value
RowNum = RowNum + 1

'Dashboard - Distribution
ws.Cells(RowNum, 1).Value = "Dashboard"
ws.Cells(RowNum, 2).Value = "Distribution"
ws.Cells(RowNum, 3).Value = frmAnalysisConfig.chkDashDistribution.Value
RowNum = RowNum + 1

'Dashboard - Anomaly
ws.Cells(RowNum, 1).Value = "Dashboard"
ws.Cells(RowNum, 2).Value = "Anomaly"
ws.Cells(RowNum, 3).Value = frmAnalysisConfig.chkDashAnomaly.Value

End Sub

Sub ReadConfiguration()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long

    Set ws = Worksheets("Config")

    Set Measures = New Collection
    Set Categories = New Collection
    Set Trends = New Collection

    'Reset Dashboard Options
    ShowDashKPI = False
    ShowDashTrend = False
    ShowDashCategory = False
    ShowDashDistribution = False
    ShowDashAnomaly = False

    LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To LastRow

        Select Case ws.Cells(i, 1).Value

            Case "Date Column"
                SelectedDate = ws.Cells(i, 2).Value

            Case "Measure"
                Measures.Add ws.Cells(i, 2).Value

            Case "Category"
                Categories.Add ws.Cells(i, 2).Value

            Case "Trend"
                Trends.Add ws.Cells(i, 2).Value

            Case "Anomaly"
                SelectedAnomaly = ws.Cells(i, 2).Value

            Case "Distribution Mode"
                DistributionMode = ws.Cells(i, 2).Value

            Case "Distribution Bins"
                DistributionBins = CLng(ws.Cells(i, 2).Value)

            Case "Dashboard"

                Select Case ws.Cells(i, 2).Value

                    Case "KPI"
                        ShowDashKPI = CBool(ws.Cells(i, 3).Value)

                    Case "Trend"
                        ShowDashTrend = CBool(ws.Cells(i, 3).Value)

                    Case "Category"
                        ShowDashCategory = CBool(ws.Cells(i, 3).Value)

                    Case "Distribution"
                        ShowDashDistribution = CBool(ws.Cells(i, 3).Value)

                    Case "Anomaly"
                        ShowDashAnomaly = CBool(ws.Cells(i, 3).Value)

                End Select

        End Select

    Next i

End Sub

Sub TestConfiguration()

    Dim Item As Variant

    ReadConfiguration

    Debug.Print "Date : " & SelectedDate

    Debug.Print "Measures"

    For Each Item In Measures
        Debug.Print Item
    Next Item

    Debug.Print "Categories"

    For Each Item In Categories
        Debug.Print Item
    Next Item

    Debug.Print "Trend"

    For Each Item In Trends
        Debug.Print Item
    Next Item

    Debug.Print "Anomaly"

    Debug.Print SelectedAnomaly

    Debug.Print "Distribution"

    Debug.Print DistributionMode

    Debug.Print DistributionBins

    Debug.Print "Dashboard"

    Debug.Print ShowDashKPI
    Debug.Print ShowDashTrend
    Debug.Print ShowDashCategory
    Debug.Print ShowDashDistribution
    Debug.Print ShowDashAnomaly

End Sub
