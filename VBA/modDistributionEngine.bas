Attribute VB_Name = "modDistributionEngine"
Option Explicit

Function BuildDistributionMCode() As String
    Dim MCode As String
    'Dim Measure As Variant
    Dim MeasureList As String

    ReadConfiguration
    MeasureList = ""
   ' For Each Measure In Measures
       ' MeasureList = MeasureList & """" & Measure & ""","
    'Next Measure

    If Right(MeasureList, 1) = "," Then
        MeasureList = Left(MeasureList, Len(MeasureList) - 1)
    End If

    ' Dynamic fix: Points straight to the generic imported data connection
    MCode = "let" & vbCrLf & _
    "    Source = PQ_RawData," & vbCrLf & _
    "    Measures = {" & MeasureList & "}," & vbCrLf & _
    "    DistributionTables = List.Transform(Measures,(Measure)=>" & vbCrLf & _
    "        let" & vbCrLf & _
    "            Values = Table.Column(Source, Measure)," & vbCrLf & _
    "            MinValue = List.Min(Values)," & vbCrLf & _
    "            MaxValue = List.Max(Values)," & vbCrLf & _
    "            BinWidth = Number.RoundUp((MaxValue-MinValue)/" & DistributionBins & ")," & vbCrLf & _
    "            BinWidth2 = if BinWidth=0 then 1 else BinWidth," & vbCrLf & _
    "            AddedBin = Table.AddColumn(Source,""BinStart"", each Number.RoundDown((Record.Field(_,Measure)-MinValue)/BinWidth2)*BinWidth2+MinValue, type number)," & vbCrLf & _
    "            Grouped = Table.Group(AddedBin,{""BinStart""},{{""Frequency"", each Table.RowCount(_), Int64.Type}})," & vbCrLf & _
    "            AddedBinEnd = Table.AddColumn(Grouped,""BinEnd"", each [BinStart]+BinWidth2, type number)," & vbCrLf & _
    "            AddedMeasure = Table.AddColumn(AddedBinEnd,""Measure"", each Measure, type text)" & vbCrLf & _
    "        in" & vbCrLf & _
    "            AddedMeasure)," & vbCrLf & _
    "    Output = Table.Combine(DistributionTables)," & vbCrLf & _
    "    Sorted = Table.Sort(Output,{{""Measure"",Order.Ascending},{""BinStart"",Order.Ascending}})" & vbCrLf & _
    "in" & vbCrLf & _
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

