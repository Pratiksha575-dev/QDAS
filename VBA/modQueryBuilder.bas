Attribute VB_Name = "modQueryBuilder"
Option Explicit

Public SelectedDate As String
Public SelectedAnomaly As String

Public Measures As Collection
Public Categories As Collection
Public Trends As Collection

Sub ReadConfiguration()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long

    Set ws = Worksheets("Config")

    Set Measures = New Collection
    Set Categories = New Collection
    Set Trends = New Collection

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

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

        End Select

    Next i

End Sub
Sub TestConfiguration()

    Dim Item As Variant

    ReadConfiguration

    Debug.Print "Date Column : " & SelectedDate

    Debug.Print "----------------"

    Debug.Print "Measures"

    For Each Item In Measures

        Debug.Print Item

    Next Item

    Debug.Print "----------------"

    Debug.Print "Categories"

    For Each Item In Categories

        Debug.Print Item

    Next Item

    Debug.Print "----------------"

    Debug.Print "Trend"

    For Each Item In Trends

        Debug.Print Item

    Next Item

    Debug.Print "----------------"

    Debug.Print "Anomaly"

    Debug.Print SelectedAnomaly

End Sub

