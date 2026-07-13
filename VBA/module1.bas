Attribute VB_Name = "module1"
Sub RecordCreatePivot()
Attribute RecordCreatePivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RecordCreatePivot Macro
'

'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=8). _
        CreatePivotTable TableDestination:="RecordedPivot!R1C1", TableName:= _
        "PivotTable1", DefaultVersion:=8
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = True
        .CompactRowIndent = 1
        .VisualTotals = False
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = True
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .DisplayEmptyRow = False
        .DisplayEmptyColumn = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .DisplayImmediateItems = True
        .ViewCalculatedMembers = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = True
        .RowAxisLayout xlCompactRow
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotCache.RefreshOnFileOpen = False
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("RecordedPivot!$A$1:$C$18")
    ActiveChart.PivotLayout.PivotTable.CubeFields( _
        "[nyc-taxidata].[tpep_pickup_datetime]").AutoGroup 2, 1
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").CubeFields("[Measures].[Sum of fare_amount]"), _
        "Sum of fare_amount"
    ActiveSheet.Shapes("Chart 1").IncrementLeft -0.0004724409
    ActiveSheet.Shapes("Chart 1").IncrementTop 16
End Sub
