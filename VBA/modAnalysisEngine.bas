Attribute VB_Name = "modAnalysisEngine"
Option Explicit

Sub RunAnalysis()

    Application.ScreenUpdating = False

    ReadConfiguration

    GenerateAllTrendQueries

    GenerateKPIQuery

    GenerateCategoryQuery
    
    GenerateDistributionQuery

    RefreshAnalysisQueries
    
    GenerateAnalysisSummary

    Application.ScreenUpdating = True

    MsgBox "Analysis Completed Successfully!"

End Sub
