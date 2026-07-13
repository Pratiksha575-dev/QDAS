Attribute VB_Name = "modConfiguration"
Option Explicit

Sub SaveConfiguration()

    Dim ws As Worksheet
    Dim i As Long
    Dim RowNum As Long

    Set ws = Worksheets("Config")

    ws.Cells.Clear

    ws.Range("A1").Value = "Setting"
    ws.Range("B1").Value = "Value"

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

    'Anomaly Method
    ws.Cells(RowNum, 1).Value = "Anomaly"

    If frmAnalysisConfig.optIQR.Value Then
        ws.Cells(RowNum, 2).Value = "IQR"
    Else
        ws.Cells(RowNum, 2).Value = "Z-Score"
    End If

End Sub
