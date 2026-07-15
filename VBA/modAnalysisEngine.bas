Attribute VB_Name = "modAnalysisEngine"
Option Explicit


Sub RunAnalysis()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '----------------------------------------
    'Save & Read Configuration
    '----------------------------------------
    SaveConfiguration
    ReadConfiguration

    '----------------------------------------
    'Generate Trend Queries
    '----------------------------------------
    If Trends.count > 0 Then
        GenerateAllTrendQueries
    End If

    '----------------------------------------
    'Generate Measure Queries
    '----------------------------------------
    If Measures.count > 0 Then
        GenerateKPIQuery
        GenerateDistributionQuery
        GenerateAnomalyQuery
    End If

    '----------------------------------------
    'Generate Category Query
    '----------------------------------------
    If Categories.count > 0 Then
        GenerateCategoryQuery
    End If

    '----------------------------------------
    'Refresh Power Queries
    '----------------------------------------
    RefreshAnalysisQueries

    '----------------------------------------
    'Generate Summary
    '----------------------------------------
    GenerateAnalysisSummary

    '----------------------------------------
    'Generate Dashboard
    '----------------------------------------
    GenerateDashboard

CleanExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Analysis Completed Successfully!", vbInformation

    Exit Sub

ErrorHandler:

    MsgBox "Error : " & Err.Description, vbCritical

    Resume CleanExit

End Sub
