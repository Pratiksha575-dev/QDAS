Attribute VB_Name = "modKPIEngine"
Option Explicit

'=========================================================
' Build KPI M Code
'=========================================================
Public Function BuildKPIMCode() As String

    Dim item As Variant
    Dim MeasureList As String
    Dim MCode As String

    MeasureList = "{"

    For Each item In Measures
        MeasureList = MeasureList & """" & item & ""","
    Next item

    If Right(MeasureList, 1) = "," Then
        MeasureList = Left(MeasureList, Len(MeasureList) - 1)
    End If

    MeasureList = MeasureList & "}"
    

    MCode = ""

    MCode = MCode & "let" & vbCrLf
    MCode = MCode & "    Source = PQ_RawData," & vbCrLf
    MCode = MCode & "    Measure = " & MeasureList & "{0}," & vbCrLf
    MCode = MCode & "    Values = List.RemoveNulls(Table.Column(Source, Measure))," & vbCrLf
    MCode = MCode & "    CountValue = List.Count(Values)," & vbCrLf
    MCode = MCode & "    TotalValue = if CountValue = 0 then null else List.Sum(Values)," & vbCrLf
    MCode = MCode & "    AverageValue = if CountValue = 0 then null else List.Average(Values)," & vbCrLf
    MCode = MCode & "    MinimumValue = if CountValue = 0 then null else List.Min(Values)," & vbCrLf
    MCode = MCode & "    MaximumValue = if CountValue = 0 then null else List.Max(Values)," & vbCrLf

    MCode = MCode & "    Output = #table(" & vbCrLf
    MCode = MCode & "        {""Measure"", ""Count"", ""Total"", ""Average"", ""Minimum"", ""Maximum""}," & vbCrLf
    MCode = MCode & "        {{" & _
            "Measure," & _
            "CountValue," & _
            "TotalValue," & _
            "AverageValue," & _
            "MinimumValue," & _
            "MaximumValue}}" & vbCrLf
    MCode = MCode & "    )" & vbCrLf

    MCode = MCode & "in" & vbCrLf
    MCode = MCode & "    Output"

    BuildKPIMCode = MCode

End Function
'=========================================================
' Update KPI Query
'=========================================================
Public Sub UpdateKPIQuery()

    UpdateQuery "PQ_KPI", BuildKPIMCode()

End Sub

'=========================================================
' Generate KPI
'=========================================================
Public Sub GenerateKPIQuery()

    ReadConfiguration

    If Measures.count = 0 Then

        MsgBox "Please select at least one Measure.", vbExclamation
        Exit Sub

    End If

    UpdateKPIQuery

    MsgBox "KPI Query Generated Successfully.", vbInformation

End Sub
