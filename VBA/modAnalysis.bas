Attribute VB_Name = "modAnalysis"
Option Explicit

Sub GenerateAnalysisSummary()

    Dim wsConfig As Worksheet
    Dim wsMeta As Worksheet
    Dim wsAnalysis As Worksheet

    Dim lastRow As Long
    Dim MeasureCount As Long
    Dim CategoryCount As Long
    Dim DateColumn As String
    Dim i As Long

    Set wsConfig = Worksheets("Config")
    Set wsMeta = Worksheets("Metadata")
    Set wsAnalysis = Worksheets("Analysis")

    wsAnalysis.Cells.Clear

    wsAnalysis.Range("A1").value = "QDAS Analysis Summary"

    wsAnalysis.Range("A3").value = "Analysis Date"
    wsAnalysis.Range("B3").value = Now

    wsAnalysis.Range("A4").value = "Dataset"
    wsAnalysis.Range("B4").value = SelectedFile
    
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    lastRow = wsConfig.Cells(wsConfig.Rows.count, 1).End(xlUp).Row
   
    For i = 2 To lastRow

        Select Case wsConfig.Cells(i, 1).value

            Case "Date Column"
                DateColumn = wsConfig.Cells(i, 2).value

            Case "Measure"
                MeasureCount = MeasureCount + 1

            Case "Category"
                CategoryCount = CategoryCount + 1

        End Select

    Next i

    wsAnalysis.Range("A6").value = "Selected Date Column"
    wsAnalysis.Range("B6").value = DateColumn

    wsAnalysis.Range("A7").value = "Measures Selected"
    wsAnalysis.Range("B7").value = MeasureCount

    wsAnalysis.Range("A8").value = "Categories Selected"
    wsAnalysis.Range("B8").value = CategoryCount

End Sub
