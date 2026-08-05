Attribute VB_Name = "modAnomalyEngine"
Option Explicit

Function BuildAnomalyMCode() As String
    ReadConfiguration
    Select Case SelectedAnomaly
        Case "IQR"
            BuildAnomalyMCode = BuildIQRMCode()
        Case "Z-Score"
            BuildAnomalyMCode = BuildZScoreMCode()
        Case Else
            MsgBox "Unknown Anomaly Method."
            Exit Function
    End Select
End Function

Function BuildIQRMCode() As String
    Dim Measure As Variant
    Dim MeasureList As String
    Dim MCode As String

    ReadConfiguration
    MeasureList = ""
    For Each Measure In Measures
        MeasureList = MeasureList & """" & Measure & ""","
    Next Measure

    If Len(MeasureList) > 0 Then
        MeasureList = Left(MeasureList, Len(MeasureList) - 1)
    End If
    MeasureList = "{" & MeasureList & "}"
    MCode = ""

MCode = MCode & "let" & vbCrLf
MCode = MCode & "    Source = PQ_RawData," & vbCrLf
MCode = MCode & "    Measures = " & MeasureList & "," & vbCrLf
MCode = MCode & "    Summary =" & vbCrLf
MCode = MCode & "        List.Transform(Measures,(Measure)=>" & vbCrLf
MCode = MCode & "        let" & vbCrLf
MCode = MCode & "            Values = List.RemoveNulls(Table.Column(Source,Measure))," & vbCrLf
MCode = MCode & "            CountValues = List.Count(Values)," & vbCrLf
MCode = MCode & "            Q1 = if CountValues=0 then null else List.Percentile(Values,0.25)," & vbCrLf
MCode = MCode & "            Q3 = if CountValues=0 then null else List.Percentile(Values,0.75)," & vbCrLf
MCode = MCode & "            IQR = if CountValues=0 then null else Q3-Q1," & vbCrLf
MCode = MCode & "            LowerLimit = if CountValues=0 then null else Q1-1.5*IQR," & vbCrLf
MCode = MCode & "            UpperLimit = if CountValues=0 then null else Q3+1.5*IQR," & vbCrLf
MCode = MCode & "            HighOutliers = if CountValues=0 then 0 else List.Count(List.Select(Values, each _ > UpperLimit))," & vbCrLf
MCode = MCode & "            LowOutliers = if CountValues=0 then 0 else List.Count(List.Select(Values, each _ < LowerLimit))," & vbCrLf
MCode = MCode & "            TotalOutliers = HighOutliers + LowOutliers," & vbCrLf
MCode = MCode & "            OutlierPercentage = if CountValues=0 then 0 else Number.Round((TotalOutliers/CountValues)*100,2)," & vbCrLf
MCode = MCode & "            MaxValue = if CountValues=0 then null else List.Max(Values)," & vbCrLf
MCode = MCode & "            MinValue = if CountValues=0 then null else List.Min(Values)" & vbCrLf
MCode = MCode & "        in" & vbCrLf
MCode = MCode & "            [Measure=Measure,Method=""IQR"",HighOutliers=HighOutliers,LowOutliers=LowOutliers,TotalOutliers=TotalOutliers,OutlierPercentage=OutlierPercentage,MaximumValue=MaxValue,MinimumValue=MinValue]" & vbCrLf

MCode = MCode & "    )," & vbCrLf
MCode = MCode & "    Output = Table.FromRecords(Summary, type table [Measure=text, Method=text, HighOutliers=number, LowOutliers=number, TotalOutliers=number, OutlierPercentage=number, MaximumValue=number, MinimumValue=number])" & vbCrLf
MCode = MCode & "in" & vbCrLf
MCode = MCode & "    Output"

    BuildIQRMCode = MCode
End Function

Function BuildZScoreMCode() As String
    Dim Measure As Variant
    Dim MeasureList As String
    Dim MCode As String

    ReadConfiguration
    MeasureList = ""
    For Each Measure In Measures
        MeasureList = MeasureList & """" & Measure & ""","
    Next Measure

    If Len(MeasureList) > 0 Then
        MeasureList = Left(MeasureList, Len(MeasureList) - 1)
    End If
    MeasureList = "{" & MeasureList & "}"
    MCode = ""

MCode = MCode & "let" & vbCrLf
MCode = MCode & "    Source = PQ_RawData," & vbCrLf
MCode = MCode & "    Measures = " & MeasureList & "," & vbCrLf
MCode = MCode & "    Summary =" & vbCrLf
MCode = MCode & "        List.Transform(Measures,(Measure)=>" & vbCrLf
MCode = MCode & "        let" & vbCrLf
MCode = MCode & "            Values = List.RemoveNulls(Table.Column(Source,Measure))," & vbCrLf
MCode = MCode & "            CountValues = List.Count(Values)," & vbCrLf
MCode = MCode & "            MeanValue = if CountValues=0 then null else List.Average(Values)," & vbCrLf
MCode = MCode & "            StdDev = if CountValues=0 then null else List.StandardDeviation(Values)," & vbCrLf
MCode = MCode & "            HighOutliers = if CountValues=0 or StdDev=0 then 0 else List.Count(List.Select(Values, each ((_ - MeanValue)/StdDev) > 3))," & vbCrLf
MCode = MCode & "            LowOutliers = if CountValues=0 or StdDev=0 then 0 else List.Count(List.Select(Values, each ((_ - MeanValue)/StdDev) < -3))," & vbCrLf
MCode = MCode & "            TotalOutliers = HighOutliers + LowOutliers," & vbCrLf
MCode = MCode & "            OutlierPercentage = if CountValues=0 then 0 else Number.Round((TotalOutliers/CountValues)*100,2)," & vbCrLf
MCode = MCode & "            MaxValue = if CountValues=0 then null else List.Max(Values)," & vbCrLf
MCode = MCode & "            MinValue = if CountValues=0 then null else List.Min(Values)" & vbCrLf
MCode = MCode & "        in" & vbCrLf
MCode = MCode & "            [Measure=Measure," & _
"Method=""Z-Score""," & _
"HighOutliers=HighOutliers," & _
"LowOutliers=LowOutliers," & _
"TotalOutliers=TotalOutliers," & _
"OutlierPercentage=OutlierPercentage," & _
"MaximumValue=MaxValue," & _
"MinimumValue=MinValue]" & vbCrLf

MCode = MCode & "    )," & vbCrLf

MCode = MCode & "    Output = Table.FromRecords(" & vbCrLf
MCode = MCode & "        Summary," & vbCrLf
MCode = MCode & "        type table [" & _
"Measure=text," & _
"Method=text," & _
"HighOutliers=number," & _
"LowOutliers=number," & _
"TotalOutliers=number," & _
"OutlierPercentage=number," & _
"MaximumValue=number," & _
"MinimumValue=number])" & vbCrLf

MCode = MCode & "in" & vbCrLf
MCode = MCode & "    Output"

    
    BuildZScoreMCode = MCode
End Function

Sub UpdateAnomalyQuery()
    UpdateQuery "PQ_Anomaly", BuildAnomalyMCode()
End Sub

Sub GenerateAnomalyQuery()
    ReadConfiguration
    UpdateAnomalyQuery
    MsgBox "Anomaly Query Generated Successfully."
End Sub

