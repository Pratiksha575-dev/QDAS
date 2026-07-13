Attribute VB_Name = "modRefreshManager"
Option Explicit

'=========================================
' Refresh All Analysis Queries
'=========================================
Sub RefreshAnalysisQueries()

    Dim Queries As Variant
    Dim q As Variant

    Queries = Array( _
    "PQ_Trend", _
    "PQ_KPI", _
    "PQ_Category", _
    "PQ_Distribution")

    For Each q In Queries

        RefreshQuery CStr(q)

    Next q

End Sub

'=========================================
' Refresh One Query
'=========================================
Sub RefreshQuery(QueryName As String)

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets

        For Each lo In ws.ListObjects

            On Error Resume Next

            If lo.SourceType = xlSrcQuery Then

                If lo.QueryTable.WorkbookConnection.Name = _
                    "Query - " & QueryName Then

                    lo.QueryTable.Refresh BackgroundQuery:=False

                    Exit Sub

                End If

            End If

            On Error GoTo 0

        Next lo

    Next ws

End Sub
Sub TestRefresh()

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In Worksheets

        For Each lo In ws.ListObjects

            Debug.Print ws.Name, lo.Name

            On Error Resume Next
            Debug.Print lo.QueryTable.WorkbookConnection.Name
            On Error GoTo 0

            Debug.Print "----------------"

        Next

    Next

End Sub
