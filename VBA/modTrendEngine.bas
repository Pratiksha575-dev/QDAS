Attribute VB_Name = "modTrendEngine"
Option Explicit

'==================================================
' Builds aggregation part of M Code
'==================================================
Function BuildMeasureAggregation() As String

    Dim Item As Variant
    Dim Agg As String

    Agg = ""

    For Each Item In Measures

        Agg = Agg & _
        "{""Sum_" & Item & """, each List.Sum([" & Item & "]), type number},"

    Next Item

    If Right(Agg, 1) = "," Then
        Agg = Left(Agg, Len(Agg) - 1)
    End If

    BuildMeasureAggregation = Agg

End Function

'==================================================
' Builds Trend Query M Code
'==================================================
Function BuildTrendMCode(ByVal TrendType As String) As String

    Dim MCode As String
    Dim DateExpression As String

    Select Case TrendType

        Case "Hourly"
            DateExpression = "DateTime.Hour([" & SelectedDate & "])"

        Case "Daily"
            DateExpression = "Date.From([" & SelectedDate & "])"

        Case "Weekly"
            DateExpression = "Date.StartOfWeek(Date.From([" & SelectedDate & "]))"

        Case "Monthly"
            DateExpression = "Date.StartOfMonth(Date.From([" & SelectedDate & "]))"

        Case "Weekday"
            DateExpression = "Date.DayOfWeekName(Date.From([" & SelectedDate & "]))"

        Case Else

            MsgBox "Unknown Trend Type : " & TrendType
            Exit Function

    End Select

    MCode = ""

    MCode = MCode & "let" & vbCrLf

    MCode = MCode & _
        "    Source = #""nyc-taxidata""," & vbCrLf

    MCode = MCode & _
        "    #""Grouped Rows"" = Table.Group(" & _
        "Table.AddColumn(Source,""Trend"", each " & _
        DateExpression & "),{""Trend""},{" & _
        BuildMeasureAggregation & "})" & vbCrLf

    MCode = MCode & "in" & vbCrLf
    MCode = MCode & _
        "    #""Grouped Rows"""

    BuildTrendMCode = MCode

End Function

'==================================================
' Updates One Trend Query
'==================================================
Sub UpdateTrendQuery(QueryName As String, TrendType As String)

    UpdateQuery QueryName, BuildTrendMCode(TrendType)

End Sub

'==================================================
' Updates All Selected Trend Queries
'==================================================
Sub GenerateAllTrendQueries()

    Dim Trend As Variant

    ReadConfiguration

    If Trends.Count = 0 Then
        MsgBox "Please select at least one Trend."
        Exit Sub
    End If

    For Each Trend In Trends

        UpdateTrendQuery "PQ_Trend", CStr(Trend)

        'RefreshTrendQuery "PQ_Trend"

        Debug.Print Trend & " Updated"

    Next Trend

End Sub

Sub RefreshTrendQuery(QueryName As String)

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

