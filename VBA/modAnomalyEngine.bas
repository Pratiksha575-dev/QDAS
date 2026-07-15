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

    MCode = _
"let" & vbCrLf & _
"    Source = #""nyc-taxidata""," & vbCrLf & _
"    Measures = " & MeasureList & "," & vbCrLf & _
"    Summary =" & vbCrLf & _
"        List.Transform(Measures,(Measure)=>" & vbCrLf & _
"        let" & vbCrLf & _
"            Values = Table.Column(Source,Measure)," & vbCrLf & _
"            Q1 = List.Percentile(Values,0.25)," & vbCrLf & _
"            Q3 = List.Percentile(Values,0.75)," & vbCrLf & _
"            IQR = Q3-Q1," & vbCrLf & _
"            LowerLimit = Q1-1.5*IQR," & vbCrLf & _
"            UpperLimit = Q3+1.5*IQR," & vbCrLf & _
"            HighOutliers = List.Count(List.Select(Values, each _>UpperLimit))," & vbCrLf & _
"            LowOutliers = List.Count(List.Select(Values, each _<LowerLimit))," & vbCrLf & _
"            TotalOutliers = HighOutliers+LowOutliers," & vbCrLf & _
"            OutlierPercentage = Number.Round((TotalOutliers / List.Count(Values)) * 100, 2)," & _
"            MaxValue = List.Max(Values)," & vbCrLf & _
"            MinValue = List.Min(Values)" & vbCrLf & _
"        in" & vbCrLf & _
"            [Measure=Measure,Method=""IQR"",HighOutliers=HighOutliers,LowOutliers=LowOutliers,TotalOutliers=TotalOutliers,OutlierPercentage=OutlierPercentage,MaximumValue=MaxValue,MinimumValue=MinValue]" & vbCrLf & _
"    )," & vbCrLf & _
"    Output = Table.FromRecords(Summary)" & vbCrLf & _
"in" & vbCrLf & _
"    Output"

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

    MeasureList = Left(MeasureList, Len(MeasureList) - 1)

    MeasureList = "{" & MeasureList & "}"

    MCode = _
"let" & vbCrLf & _
"    Source = #""nyc-taxidata""," & vbCrLf & _
"    Measures = " & MeasureList & "," & vbCrLf & _
"    Summary =" & vbCrLf & _
"        List.Transform(Measures,(Measure)=>" & vbCrLf & _
"        let" & vbCrLf & _
"            Values = Table.Column(Source,Measure)," & vbCrLf & _
"            MeanValue = List.Average(Values)," & vbCrLf & _
"            StdDev = List.StandardDeviation(Values)," & vbCrLf & _
"            HighOutliers = List.Count(List.Select(Values, each ((_ - MeanValue)/StdDev) > 3))," & vbCrLf & _
"            LowOutliers = List.Count(List.Select(Values, each ((_ - MeanValue)/StdDev) < -3))," & vbCrLf & _
"            TotalOutliers = HighOutliers + LowOutliers," & vbCrLf & _
"            MaxValue = List.Max(Values)," & vbCrLf & _
"            MinValue = List.Min(Values)," & vbCrLf & _
"            OutlierPercentage = Number.Round((TotalOutliers/List.Count(Values))*100,2)" & vbCrLf & _
"        in" & vbCrLf & _
"            [Measure=Measure,Method=""Z-Score"",HighOutliers=HighOutliers,LowOutliers=LowOutliers,TotalOutliers=TotalOutliers,OutlierPercentage=OutlierPercentage,MaximumValue=MaxValue,MinimumValue=MinValue]" & vbCrLf & _
"    )," & vbCrLf & _
"    Output = Table.FromRecords(Summary)" & vbCrLf & _
"in" & vbCrLf & _
"    Output"

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

Sub TestAnomalyMCode()

    Debug.Print BuildIQRMCode()

End Sub

