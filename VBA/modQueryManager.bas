Attribute VB_Name = "modQueryManager"
Option Explicit

'=========================================
' Update a Power Query's M code
'=========================================
Sub UpdateQuery(QueryName As String, MCode As String)

    Dim q As WorkbookQuery

    Set q = ThisWorkbook.Queries(QueryName)

    q.Formula = MCode

End Sub

'=========================================
' Create or Update Power Query
'=========================================
Sub CreateOrUpdateQuery(QueryName As String, MCode As String)

    Dim q As WorkbookQuery

    On Error Resume Next
    Set q = ThisWorkbook.Queries(QueryName)
    On Error GoTo 0

    If q Is Nothing Then

        ThisWorkbook.Queries.Add _
            Name:=QueryName, _
            Formula:=MCode

    Else

        q.Formula = MCode

    End If

End Sub

'==================================================
' Creates a worksheet if it doesn't exist
'==================================================
Function GetOrCreateWorksheet(SheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateWorksheet = Worksheets(SheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then

        Set GetOrCreateWorksheet = Worksheets.Add
        GetOrCreateWorksheet.Name = SheetName

    End If

End Function

'==================================================
' Returns TRUE if a table already exists
'==================================================
Function TableExists(TableName As String) As Boolean

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In Worksheets

        For Each lo In ws.ListObjects

            If lo.Name = TableName Then

                TableExists = True
                Exit Function

            End If

        Next lo

    Next ws

End Function
