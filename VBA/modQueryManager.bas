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
