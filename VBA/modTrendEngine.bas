Attribute VB_Name = "modTrendEngine"
Option Explicit

'=========================================================
' Build Aggregation
'=========================================================
Private Function BuildMeasureAggregation() As String

    Dim MeasureName As String

    MeasureName = Measures.item(1)

    BuildMeasureAggregation = _
        "{""Measure"", each List.Sum([" & MeasureName & "]), type number}"


End Function
'=========================================================
' Build Trend M Code
'=========================================================
Public Function BuildTrendMCode(ByVal TrendType As String) As String

    Dim MCode As String
    Dim DateExpression As String
    Dim ReturnType As String
    Dim TrendColumn As String

    TrendColumn = "Trend"

    Select Case TrendType

        Case "Daily"

            DateExpression = "Date.From([" & SelectedDate & "])"
            ReturnType = "type date"


        Case "Monthly"

            DateExpression = "Date.StartOfMonth(Date.From([" & SelectedDate & "]))"
            ReturnType = "type date"

        Case "Weekday"

            DateExpression = "Date.DayOfWeekName(Date.From([" & SelectedDate & "]))"
            ReturnType = "type text"

        Case Else

            Err.Raise vbObjectError + 1001, _
                      "BuildTrendMCode", _
                      "Unknown Trend Type : " & TrendType

    End Select

    MCode = _
    "let" & vbCrLf & _
    "    Source = PQ_RawData," & vbCrLf & _
    "    AddTrend = Table.AddColumn(Source, ""Trend"", each " & DateExpression & ", " & ReturnType & ")," & vbCrLf & _
    "    Grouped = Table.Group(AddTrend,{""Trend""},{" & BuildMeasureAggregation() & "})," & vbCrLf

    Select Case TrendType

        Case "Daily", "Monthly"

            MCode = MCode & _
            "    Sorted = Table.Sort(Grouped,{{""Trend"", Order.Ascending}})" & vbCrLf & _
            "in" & vbCrLf & _
            "    Sorted"

        Case "Weekday"

            MCode = MCode & _
            "    WeekOrder = {""Monday"",""Tuesday"",""Wednesday"",""Thursday"",""Friday"",""Saturday"",""Sunday""}," & vbCrLf & _
            "    AddIndex = Table.AddColumn(Grouped,""SortOrder"", each List.PositionOf(WeekOrder,[Trend]), Int64.Type)," & vbCrLf & _
            "    Sorted = Table.Sort(AddIndex,{{""SortOrder"",Order.Ascending}})," & vbCrLf & _
            "    Final = Table.RemoveColumns(Sorted,{""SortOrder""})" & vbCrLf & _
            "in" & vbCrLf & _
            "    Final"

    End Select

    BuildTrendMCode = MCode

End Function

'=========================================================
' Update Trend Query
'=========================================================
Public Sub UpdateTrendQuery(ByVal TrendType As String)

    Dim QueryName As String

    QueryName = "PQ_Trend_" & TrendType

    UpdateQuery QueryName, BuildTrendMCode(TrendType)

End Sub

'=========================================================
' Generate Selected Trend Queries
'=========================================================
Public Sub GenerateAllTrendQueries()

    Dim Trend As Variant

    ReadConfiguration
    Debug.Print "SelectedDate = [" & SelectedDate & "]"
Debug.Print "Trend Count = " & Trends.count

    'No trend selected
    If Trends.count = 0 Then
       ClearTrendQueries

    MsgBox "No Trend Analysis selected.", vbInformation
      Exit Sub
    End If

    If Trim(SelectedDate) = "" Then

    ClearTrendQueries

    MsgBox "Trend analysis skipped because no Date column was selected.", vbInformation

    Exit Sub

End If

    For Each Trend In Trends

        UpdateTrendQuery CStr(Trend)
        Debug.Print "Generated : PQ_Trend_" & Trend

    Next Trend

End Sub
Public Function BuildEmptyTrendMCode() As String

    BuildEmptyTrendMCode = _
        "let" & vbCrLf & _
        "    Source = PQ_RawData," & vbCrLf & _
        "    EmptySource = Table.FirstN(Source,0)," & vbCrLf & _
        "    AddTrend = Table.AddColumn(EmptySource,""Trend"", each null, type text)," & vbCrLf & _
        "    AddMeasure = Table.AddColumn(AddTrend,""Measure"", each null, type number)," & vbCrLf & _
        "    Final = Table.SelectColumns(AddMeasure,{""Trend"",""Measure""})" & vbCrLf & _
        "in" & vbCrLf & _
        "    Final"

End Function
Public Sub ClearTrendQueries()

    Dim Trend As Variant

    For Each Trend In Array("Daily", "Monthly", "Weekday")

        UpdateQuery "PQ_Trend_" & Trend, BuildEmptyTrendMCode()

    Next Trend

End Sub
