Attribute VB_Name = "modDistributionEngine"
Option Explicit

Function BuildDistributionAggregation() As String

    Dim Item As Variant
    Dim Agg As String

    Agg = ""

    For Each Item In Measures

        Agg = Agg & _
        "{""Count_" & Item & """, each Table.RowCount(_), Int64.Type},"

    Next Item

    If Right(Agg, 1) = "," Then
        Agg = Left(Agg, Len(Agg) - 1)
    End If

    BuildDistributionAggregation = Agg

End Function

Function BuildDistributionMCode() As String

    Dim MCode As String
    Dim Measure As String

    ReadConfiguration

    Measure = Measures(1)

    MCode = ""

    MCode = MCode & "let" & vbCrLf

    MCode = MCode & _
    "    Source = #""nyc-taxidata""," & vbCrLf

    MCode = MCode & _
    "    MinValue = List.Min(Source[" & Measure & "])," & vbCrLf

    MCode = MCode & _
    "    MaxValue = List.Max(Source[" & Measure & "])," & vbCrLf

    MCode = MCode & _
    "    BinWidth = Number.RoundUp((MaxValue-MinValue)/20)," & vbCrLf

    MCode = MCode & _
    "    BinWidth2 = if BinWidth=0 then 1 else BinWidth," & vbCrLf

    MCode = MCode & _
    "    #""Added Bin"" = Table.AddColumn(Source,""Bin"", each Number.RoundDown(([" & _
    Measure & "]-MinValue)/BinWidth2)*BinWidth2 + MinValue)," & vbCrLf

    MCode = MCode & _
    "    #""Grouped Rows"" = Table.Group(#""Added Bin"",{""Bin""}," & _
    "{{""Frequency"", each Table.RowCount(_), Int64.Type}})," & vbCrLf

    MCode = MCode & _
    "    #""Sorted Rows"" = Table.Sort(#""Grouped Rows"",{{""Bin"", Order.Ascending}})" & vbCrLf

    MCode = MCode & _
    "in" & vbCrLf

    MCode = MCode & _
    "    #""Sorted Rows"""

    BuildDistributionMCode = MCode

End Function

Sub UpdateDistributionQuery()

    UpdateQuery "PQ_Distribution", BuildDistributionMCode()

End Sub

Sub GenerateDistributionQuery()

    ReadConfiguration

    UpdateDistributionQuery

    MsgBox "Distribution Query Generated Successfully."

End Sub

Sub TestDistributionMCode()

    Debug.Print BuildDistributionMCode()

End Sub
