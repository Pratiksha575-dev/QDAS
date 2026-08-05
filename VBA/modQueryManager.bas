Attribute VB_Name = "modQueryManager"
Option Explicit

'=========================================================
' Create or Update Power Query Safely on Fresh Slates
'=========================================================
Sub CreateOrUpdateQuery(QueryName As String, MCode As String)
    Dim q As WorkbookQuery
    Dim conn As WorkbookConnection
    
    On Error Resume Next
    Set q = ThisWorkbook.Queries(QueryName)
    On Error GoTo 0

    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:=QueryName, Formula:=MCode
    Else
        q.Formula = MCode
    End If
End Sub

'=========================================================
' Update a Power Query's M code safely
'=========================================================
Sub UpdateQuery(QueryName As String, MCode As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = ThisWorkbook.Queries(QueryName)
    On Error GoTo 0
    
    If Not q Is Nothing Then
        q.Formula = MCode
    Else
        CreateOrUpdateQuery QueryName, MCode
    End If
End Sub

'=========================================================
' Creates a worksheet if it doesn't exist (Strict Target Context)
'=========================================================
Function GetOrCreateWorksheet(SheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SheetName
    End If
    Set GetOrCreateWorksheet = ws
End Function

'=========================================================
' Returns TRUE if a table already exists
'=========================================================
Function TableExists(TableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = TableName Then
                TableExists = True
                Exit Function
            End If
        Next lo
    Next ws
    TableExists = False
End Function
