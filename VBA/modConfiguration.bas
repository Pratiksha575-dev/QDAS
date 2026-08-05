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

    Set ws = GetOrCreateWorksheet("Config")

    ws.Cells.Clear

    ws.Range("A1").value = "Setting"
    ws.Range("B1").value = "Value"

    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    RowNum = 2

    '-------------------------
    ' Date Column
    '-------------------------
    ws.Cells(RowNum, 1).value = "Date Column"

    If frmAnalysisConfig.cmbDate.ListIndex <> -1 Then
        ws.Cells(RowNum, 2).value = frmAnalysisConfig.cmbDate.value
    Else
        ws.Cells(RowNum, 2).value = ""
    End If

    RowNum = RowNum + 1

    '-------------------------
    ' Measures
    '-------------------------
    For i = 0 To frmAnalysisConfig.lstMeasures.ListCount - 1

        If frmAnalysisConfig.lstMeasures.Selected(i) Then

            ws.Cells(RowNum, 1).value = "Measure"
            ws.Cells(RowNum, 2).value = frmAnalysisConfig.lstMeasures.List(i)

            RowNum = RowNum + 1

        End If

    Next i

    '-------------------------
    ' Categories
    '-------------------------
    For i = 0 To frmAnalysisConfig.lstCategory.ListCount - 1

        If frmAnalysisConfig.lstCategory.Selected(i) Then

            ws.Cells(RowNum, 1).value = "Category"
            ws.Cells(RowNum, 2).value = frmAnalysisConfig.lstCategory.List(i)

            RowNum = RowNum + 1

        End If

    Next i

    '-------------------------
    ' Trend (Fixed Rows)
    '-------------------------

    ws.Cells(RowNum, 1).value = "Trend"
    If frmAnalysisConfig.chkDaily.Enabled And frmAnalysisConfig.chkDaily.value Then
        ws.Cells(RowNum, 2).value = "Daily"
    Else
        ws.Cells(RowNum, 2).value = ""
    End If
    RowNum = RowNum + 1

    ws.Cells(RowNum, 1).value = "Trend"
    If frmAnalysisConfig.chkMonthly.Enabled And frmAnalysisConfig.chkMonthly.value Then
        ws.Cells(RowNum, 2).value = "Monthly"
    Else
        ws.Cells(RowNum, 2).value = ""
    End If
    RowNum = RowNum + 1

    ws.Cells(RowNum, 1).value = "Trend"
    If frmAnalysisConfig.chkWeekday.Enabled And frmAnalysisConfig.chkWeekday.value Then
        ws.Cells(RowNum, 2).value = "Weekday"
    Else
        ws.Cells(RowNum, 2).value = ""
    End If
    RowNum = RowNum + 1

    '-------------------------
    ' Anomaly
    '-------------------------

    ws.Cells(RowNum, 1).value = "Anomaly"

    If frmAnalysisConfig.optIQR.value Then
        ws.Cells(RowNum, 2).value = "IQR"
    Else
        ws.Cells(RowNum, 2).value = "Z-Score"
    End If

End Sub

Sub ReadConfiguration()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = Worksheets("Config")

    Set Measures = New Collection
    Set Categories = New Collection
    Set Trends = New Collection
    SelectedDate = ""
    SelectedAnomaly = ""
    

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow

        Select Case ws.Cells(i, 1).value

            Case "Date Column"
                SelectedDate = ws.Cells(i, 2).value

            Case "Measure"
                Measures.Add ws.Cells(i, 2).value

            Case "Category"
                Categories.Add ws.Cells(i, 2).value

            Case "Trend"

          If Trim(ws.Cells(i, 2).value) <> "" Then
             Trends.Add ws.Cells(i, 2).value
          End If

            Case "Anomaly"
                SelectedAnomaly = ws.Cells(i, 2).value
        End Select

    Next i

End Sub

Sub TestConfiguration()

    Dim item As Variant

    ReadConfiguration

    Debug.Print "Date : " & SelectedDate

    Debug.Print "Measures"

    For Each item In Measures
        Debug.Print item
    Next item

    Debug.Print "Categories"

    For Each item In Categories
        Debug.Print item
    Next item

    Debug.Print "Trend"

    For Each item In Trends
        Debug.Print item
    Next item

    Debug.Print "Anomaly"

    Debug.Print SelectedAnomaly

End Sub
