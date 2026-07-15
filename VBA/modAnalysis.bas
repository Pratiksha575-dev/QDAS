Attribute VB_Name = "modAnalysis"
Option Explicit

Sub GenerateAnalysisSummary()

    Dim wsConfig As Worksheet
    Dim wsMeta As Worksheet
    Dim wsAnalysis As Worksheet

    Dim LastRow As Long
    Dim MeasureCount As Long
    Dim CategoryCount As Long
    Dim DateColumn As String
    Dim i As Long

    Set wsConfig = Worksheets("Config")
    Set wsMeta = Worksheets("Metadata")
    Set wsAnalysis = Worksheets("Analysis")

    wsAnalysis.Cells.Clear

    wsAnalysis.Range("A1").Value = "QDAS Analysis Summary"

    wsAnalysis.Range("A3").Value = "Analysis Date"
    wsAnalysis.Range("B3").Value = Now

    wsAnalysis.Range("A4").Value = "Dataset"
    wsAnalysis.Range("B4").Value = SelectedFile

    LastRow = wsConfig.Cells(wsConfig.Rows.count, 1).End(xlUp).Row

    For i = 2 To LastRow

        Select Case wsConfig.Cells(i, 1).Value

            Case "Date Column"
                DateColumn = wsConfig.Cells(i, 2).Value

            Case "Measure"
                MeasureCount = MeasureCount + 1

            Case "Category"
                CategoryCount = CategoryCount + 1

        End Select

    Next i

    wsAnalysis.Range("A6").Value = "Selected Date Column"
    wsAnalysis.Range("B6").Value = DateColumn

    wsAnalysis.Range("A7").Value = "Measures Selected"
    wsAnalysis.Range("B7").Value = MeasureCount

    wsAnalysis.Range("A8").Value = "Categories Selected"
    wsAnalysis.Range("B8").Value = CategoryCount

End Sub
