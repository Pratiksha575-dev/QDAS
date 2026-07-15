Attribute VB_Name = "modDistributionEngine"
Option Explicit

'=========================================================
' Builds Distribution M Code for ONE Measure
'=========================================================
Function BuildDistributionMCode() As String

    Dim MCode As String
    Dim Measure As Variant
    Dim MeasureList As String

    ReadConfiguration

    MeasureList = ""

    For Each Measure In Measures
        MeasureList = MeasureList & """" & Measure & ""","
    Next Measure

    If Right(MeasureList, 1) = "," Then
        MeasureList = Left(MeasureList, Len(MeasureList) - 1)
    End If

    MCode = ""

    MCode = MCode & "let" & vbCrLf

    MCode = MCode & _
    "    Source = #""nyc-taxidata""," & vbCrLf

    MCode = MCode & _
    "    Measures = {" & MeasureList & "}," & vbCrLf

    MCode = MCode & _
    "    DistributionTables = List.Transform(Measures,(Measure)=>" & vbCrLf

    MCode = MCode & _
    "        let" & vbCrLf

    MCode = MCode & _
    "            MinValue = List.Min(Table.Column(Source,Measure))," & vbCrLf

    MCode = MCode & _
    "            MaxValue = List.Max(Table.Column(Source,Measure))," & vbCrLf

    MCode = MCode & _
    "            BinWidth = Number.RoundUp((MaxValue-MinValue)/" & DistributionBins & ")," & vbCrLf

    MCode = MCode & _
    "            BinWidth2 = if BinWidth=0 then 1 else BinWidth," & vbCrLf

    MCode = MCode & _
    "            AddedBin = Table.AddColumn(Source,""BinStart"", each Number.RoundDown((Record.Field(_,Measure)-MinValue)/BinWidth2)*BinWidth2+MinValue,type number)," & vbCrLf

    MCode = MCode & _
    "            Grouped = Table.Group(AddedBin,{""BinStart""},{{""Frequency"", each Table.RowCount(_), Int64.Type}})," & vbCrLf

    MCode = MCode & _
    "            AddedBinEnd = Table.AddColumn(Grouped,""BinEnd"", each [BinStart]+BinWidth2,type number)," & vbCrLf

    MCode = MCode & _
    "            AddedMeasure = Table.AddColumn(AddedBinEnd,""Measure"", each Measure,type text)" & vbCrLf

    MCode = MCode & _
    "        in" & vbCrLf

    MCode = MCode & _
    "            AddedMeasure)," & vbCrLf

    MCode = MCode & _
    "    Output = Table.Combine(DistributionTables)," & vbCrLf

    MCode = MCode & _
    "    Sorted = Table.Sort(Output,{{""Measure"",Order.Ascending},{""BinStart"",Order.Ascending}})" & vbCrLf

    MCode = MCode & _
    "in" & vbCrLf

    MCode = MCode & _
    "    Sorted"

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

Sub TestOneDistribution()

    ReadConfiguration

    Debug.Print BuildDistributionMCode()

End Sub
