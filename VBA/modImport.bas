Attribute VB_Name = "ModImport"
Option Explicit

Public SelectedFile As String

'==============================================================================
' IMPORT DATASET
'==============================================================================
Sub ImportDataset()

    Dim fd As FileDialog
    Dim cleanPath As String
    Dim fullMCode As String
    Dim previewMCode As String
    Dim wsPreview As Worksheet

    On Error GoTo IngestionFailure

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Dataset"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False

        If .Show <> -1 Then Exit Sub

        SelectedFile = .SelectedItems(1)
    End With

    cleanPath = Replace(SelectedFile, "\", "\\")

    '--------------------------------------------------------------------------
    ' Raw Data Query
    '--------------------------------------------------------------------------
    fullMCode = _
"let" & vbCrLf & _
"    Source = Csv.Document(File.Contents(""" & cleanPath & """), [Delimiter="","", Columns=null, Encoding=65001, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
"    PromoteHeaders = Table.PromoteHeaders(Source,[PromoteAllScalars=true])" & vbCrLf & _
"in" & vbCrLf & _
"    PromoteHeaders"

    '--------------------------------------------------------------------------
    ' Preview Query
    '--------------------------------------------------------------------------
    previewMCode = _
    "let" & vbCrLf & _
    "    Source = PQ_RawData," & vbCrLf & _
    "    KeepFirstRows = Table.FirstN(Source, 100)" & vbCrLf & _
    "in" & vbCrLf & _
    "    KeepFirstRows"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    '--------------------------------------------------------------------------
    ' Create / Update Queries
    '--------------------------------------------------------------------------
    CreateOrUpdateQuery "PQ_RawData", fullMCode
    CreateOrUpdateQuery "PQ_PreviewData", previewMCode
    Dim cn As WorkbookConnection

For Each cn In ThisWorkbook.Connections
    Debug.Print cn.Name
Next cn
    '--------------------------------------------------------------------------
    ' Refresh Preview Table
    '--------------------------------------------------------------------------
    Set wsPreview = GetOrCreateWorksheet("PreviewData")

    If wsPreview.ListObjects.count = 0 Then

        MsgBox "PreviewData table not found." & vbCrLf & vbCrLf & _
               "Please load PQ_PreviewData to the PreviewData worksheet once." & vbCrLf & _
               "Future imports will refresh automatically.", vbInformation

    Else

        wsPreview.ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
        wsPreview.Columns.AutoFit

    End If

    '--------------------------------------------------------------------------
    ' Generate Metadata
    '--------------------------------------------------------------------------
    GenerateMetadata

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox _
"Dataset imported successfully!" & vbCrLf & vbCrLf & _
"Please detect data types before running Analysis:" & vbCrLf & _
"PQ_RawData -> Transform -> Select All -> Detect Data Type" & vbCrLf & _
"PQ_PreviewData -> Transform -> Select All -> Detect Data Type" & vbCrLf & _
"Then click 'Run Analysis'.", _
vbInformation, "Data Type Detection Required"
    Exit Sub

'--------------------------------------------------------------------------
' Error Handler
'--------------------------------------------------------------------------
IngestionFailure:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "Import failed." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical

End Sub
Sub CheckLoadDestination()

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In Worksheets
        For Each lo In ws.ListObjects
            Debug.Print ws.Name, lo.Name, lo.SourceType
        Next lo
    Next ws

End Sub
Sub ShowPreviewTables()

    Dim lo As ListObject

    For Each lo In Worksheets("PreviewData").ListObjects
        Debug.Print lo.Name, lo.Range.Address
    Next

End Sub
Option Explicit

Public Sub ExportDashboardToPDF()

    Dim wsDash As Worksheet
    Dim FilePath As Variant

    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0

    If wsDash Is Nothing Then
        MsgBox "Dashboard sheet not found.", vbExclamation, "QDAS"
        Exit Sub
    End If

    '---------------------------------------
    'Save As Dialog
    '---------------------------------------
    FilePath = Application.GetSaveAsFilename( _
                    InitialFileName:="QDAS_Report_" & Format(Date, "yyyy-mm-dd") & ".pdf", _
                    FileFilter:="PDF Files (*.pdf), *.pdf", _
                    Title:="Export Dashboard to PDF")

    If FilePath = False Then Exit Sub

    If LCase$(Right$(CStr(FilePath), 4)) <> ".pdf" Then
        FilePath = FilePath & ".pdf"
    End If

    On Error GoTo ExportError

    Application.ScreenUpdating = False

    '---------------------------------------
    'Hide Excel UI (temporary)
    '---------------------------------------
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    '---------------------------------------
    'Page Setup
    '---------------------------------------
    With wsDash.PageSetup

        '<<< Change this range if your dashboard size changes
        .PrintArea = "$B$1:$S$36"

        .Orientation = xlLandscape
        .PaperSize = xlPaperA4

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1

        .CenterHorizontally = True
        .CenterVertically = True

        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)

        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)

    End With

    '---------------------------------------
    'Export PDF
    '---------------------------------------
    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(FilePath), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    '---------------------------------------
    'Restore Excel View
    '---------------------------------------
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True

    Application.ScreenUpdating = True

    MsgBox "Dashboard exported successfully!" & vbCrLf & vbCrLf & _
           "Location:" & vbCrLf & FilePath, vbInformation, "QDAS"

    Exit Sub

'---------------------------------------
'Error Handler
'---------------------------------------
ExportError:

    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True

    Application.ScreenUpdating = True

    MsgBox "Unable to export dashboard." & vbCrLf & _
           Err.Description, vbCritical, "QDAS"

End Sub

