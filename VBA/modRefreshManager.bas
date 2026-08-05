Attribute VB_Name = "modRefreshManager"
Option Explicit

'=========================================================
' Refresh All Analysis Queries Safely & Synchronously
'=========================================================
Sub RefreshAnalysisQueries()

    Dim Trend As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    RefreshQuery "PQ_Trend_Daily"
    RefreshQuery "PQ_Trend_Monthly"
    RefreshQuery "PQ_Trend_Weekday"
    '-----------------------------------------
    ' Refresh Other Queries
    '-----------------------------------------
    RefreshQuery "PQ_KPI"
    RefreshQuery "PQ_Category"
    RefreshQuery "PQ_Anomaly"

End Sub

'=========================================
' Refresh One Query Safely
'=========================================
Sub RefreshQuery(QueryName As String)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim Found As Boolean

Found = False

For Each ws In ThisWorkbook.Worksheets

    For Each lo In ws.ListObjects

        On Error Resume Next

        If lo.SourceType = xlSrcQuery Then

            If lo.QueryTable.WorkbookConnection.Name = _
                "Query - " & QueryName Then

                lo.QueryTable.Refresh BackgroundQuery:=False

                Found = True
                Exit For

            End If

        End If

        On Error GoTo 0

    Next lo

    If Found Then Exit For

Next ws

End Sub
Function AnalysisConnectionsExist() As Boolean

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets

        For Each lo In ws.ListObjects

            On Error Resume Next

            If lo.SourceType = xlSrcQuery Then
                AnalysisConnectionsExist = True
                Exit Function
            End If

            On Error GoTo 0

        Next lo

    Next ws

    AnalysisConnectionsExist = False

End Function

'=========================================
' Diagnostic Environment Tester
'=========================================
Sub TestRefresh()

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects

            Debug.Print ws.Name, lo.Name

            On Error Resume Next
            Debug.Print lo.QueryTable.WorkbookConnection.Name
            On Error GoTo 0

            Debug.Print "----------------"

        Next lo
    Next ws

End Sub
Sub ListConnections()

    Dim cn As WorkbookConnection

    For Each cn In ThisWorkbook.Connections
        Debug.Print cn.Name
    Next cn

End Sub
Sub CheckPreviewLink()

    Dim lo As ListObject

    Set lo = Worksheets("PreviewData").ListObjects(1)

    Debug.Print "Table Name: "; lo.Name
    Debug.Print "Source Type: "; lo.SourceType

    On Error Resume Next
    Debug.Print "Workbook Connection: "; lo.QueryTable.WorkbookConnection.Name
    Debug.Print "Command Text: "; lo.QueryTable.CommandText
    Debug.Print "Error: "; Err.Number & " - " & Err.Description
    On Error GoTo 0

End Sub
Sub ListQueries()

    Dim q As WorkbookQuery

    For Each q In ThisWorkbook.Queries
        Debug.Print q.Name
    Next

End Sub
Sub InspectAllTables()

    Dim lo As ListObject

    For Each lo In Worksheets("PreviewData").ListObjects

        Debug.Print "----------------"
        Debug.Print lo.Name
        Debug.Print lo.Range.Address

        On Error Resume Next
        Debug.Print "Has QueryTable:", Not lo.QueryTable Is Nothing
        Debug.Print "Connection:", lo.QueryTable.WorkbookConnection.Name
        Debug.Print "Err:", Err.Number
        Err.Clear
        On Error GoTo 0

    Next

End Sub
Sub DeletePreviewTables()

    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Worksheets("PreviewData")

    Do While ws.ListObjects.count > 0
        ws.ListObjects(1).Delete
    Loop

    MsgBox "All tables deleted from PreviewData."

End Sub
Sub CheckPreview()

    Dim lo As ListObject

    Set lo = Worksheets("PreviewData").ListObjects(1)

    Debug.Print lo.Name
    Debug.Print lo.Range.Address
    Debug.Print Not lo.QueryTable Is Nothing

End Sub
