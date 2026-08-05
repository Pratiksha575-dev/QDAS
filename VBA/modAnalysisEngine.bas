Attribute VB_Name = "modAnalysisEngine"
Sub RunAnalysis()
Dim FirstRun As Boolean

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '----------------------------------------
    'Save & Read Configuration
    '----------------------------------------
    SaveConfiguration
    Debug.Print "1. Reading Configuration"
    ReadConfiguration

    '----------------------------------------
    'Generate Trend Queries
    '----------------------------------------
    Debug.Print "2. Trend"
     GenerateAllTrendQueries
    '----------------------------------------
    'Generate Measure Queries
    '----------------------------------------
    If Measures.count > 0 Then
        
        Debug.Print "4. KPI"
        GenerateKPIQuery
        'GenerateDistributionQuery
        Debug.Print "6. Anomaly"
        GenerateAnomalyQuery
    End If

    '----------------------------------------
    'Generate Category Query
    '----------------------------------------
    If Categories.count > 0 Then
    Debug.Print "4. Category"
        GenerateCategoryQuery
    End If

    '----------------------------------------
    '----------------------------------------
' Refresh Only If Queries Already Loaded
'----------------------------------------
If AnalysisConnectionsExist Then
    Debug.Print "7. Refresh"
    RefreshAnalysisQueries
    Debug.Print "8.Pivot refresh"
    RefreshAllPivotTables
    

Else

    MsgBox "Analysis queries created successfully." & vbCrLf & vbCrLf & _
           "Load them as 'Connection Only', create PivotTables/PivotCharts once, then future runs will refresh automatically.", vbInformation

End If

    '----------------------------------------
    'Generate Summary
    '----------------------------------------
    GenerateAnalysisSummary

CleanExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Analysis Completed Successfully!", vbInformation

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox Err.Description, vbCritical
    Exit Sub

End Sub

Sub RefreshAllPivotTables()

    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    UpdatePivotCaptions

End Sub

Public Sub UpdatePivotCaptions()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim MeasureCaption As String

    ReadConfiguration

    If Measures Is Nothing Then Exit Sub
    If Measures.count = 0 Then Exit Sub

    MeasureCaption = StrConv(Replace(Measures.item(1), "_", " "), vbProperCase)

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables

            If pt.DataFields.count > 0 Then
                pt.DataFields(1).Caption = "Sum of " & MeasureCaption
            End If

        Next pt
    Next ws

End Sub
Sub ListQueries()

    Dim q As WorkbookQuery

    For Each q In ThisWorkbook.Queries
        Debug.Print q.Name
    Next q

End Sub
Public Function IsTrendAvailable() As Boolean

    If Trends Is Nothing Then Exit Function
    If Trends.count = 0 Then Exit Function

    If Trim(Trends.item(1)) = "" Then Exit Function

    IsTrendAvailable = True

End Function
