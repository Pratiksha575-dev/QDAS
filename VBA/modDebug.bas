Attribute VB_Name = "modDebug"
Option Explicit

Sub ListPivotFields()

    Dim pt As PivotTable
    Dim pf As PivotField

    Set pt = Worksheets("TestPivot").PivotTables(1)

    Debug.Print "========================"

    For Each pf In pt.PivotFields

        Debug.Print "Name : " & pf.Name
        Debug.Print "SourceName : " & pf.SourceName
        Debug.Print "Caption : " & pf.Caption
        Debug.Print "----------------------"

    Next pf

End Sub
