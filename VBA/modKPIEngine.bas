Attribute VB_Name = "modKPIEngine"
Option Explicit

Function BuildKPIMCode() As String

    Dim MCode As String
    Dim KPIRows As String
    Dim Item As Variant

    ReadConfiguration

    KPIRows = ""

    For Each Item In Measures

        KPIRows = KPIRows & _
        "        [Measure=""" & Item & """, KPI=""Total"", Value=List.Sum(Source[" & Item & "])]," & vbCrLf

        KPIRows = KPIRows & _
        "        [Measure=""" & Item & """, KPI=""Average"", Value=List.Average(Source[" & Item & "])]," & vbCrLf

        KPIRows = KPIRows & _
        "        [Measure=""" & Item & """, KPI=""Minimum"", Value=List.Min(Source[" & Item & "])]," & vbCrLf

        KPIRows = KPIRows & _
        "        [Measure=""" & Item & """, KPI=""Maximum"", Value=List.Max(Source[" & Item & "])]," & vbCrLf

        KPIRows = KPIRows & _
        "        [Measure=""" & Item & """, KPI=""Count"", Value=List.Count(Source[" & Item & "])]," & vbCrLf

    Next Item

    'Remove the final comma
    KPIRows = Left(KPIRows, Len(KPIRows) - 3)

    MCode = ""
    MCode = MCode & "let" & vbCrLf
    MCode = MCode & "    Source = #""nyc-taxidata""," & vbCrLf
    MCode = MCode & "    KPIList = {" & vbCrLf
    MCode = MCode & KPIRows & vbCrLf
    MCode = MCode & "    }," & vbCrLf
    MCode = MCode & "    Output = Table.FromRecords(KPIList)" & vbCrLf
    MCode = MCode & "in" & vbCrLf
    MCode = MCode & "    Output"

    BuildKPIMCode = MCode

End Function
Sub TestKPIMCode()

    Debug.Print BuildKPIMCode()

End Sub

Sub UpdateKPIQuery()

    UpdateQuery "PQ_KPI", BuildKPIMCode()

End Sub

Sub GenerateKPIQuery()

    
    ReadConfiguration

    UpdateKPIQuery

    MsgBox "KPI Query Generated Successfully."

End Sub
