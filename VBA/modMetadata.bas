Attribute VB_Name = "ModMetadata"
Option Explicit

Sub CreateMetadata()

    Dim ws As Worksheet

    Set ws = Worksheets("Metadata")

    ws.Cells.Clear

    ws.Range("A1").Value = "Column Name"
    ws.Range("B1").Value = "Detected Type"
    ws.Range("C1").Value = "Suggested Role"
    ws.Range("D1").Value = "Final Role"
    ws.Range("E1").Value = "Selected"

End Sub
Sub GenerateMetadata()

    Dim wsPreview As Worksheet
    Dim wsMeta As Worksheet

    Dim LastCol As Long
    Dim LastRow As Long

    Dim c As Long
    Dim r As Long

    Dim ColName As String
    Dim DataType As String
    Dim Role As String

    Set wsPreview = Worksheets("PreviewData")
    Set wsMeta = Worksheets("Metadata")

    CreateMetadata

    LastCol = wsPreview.Cells(1, wsPreview.Columns.Count).End(xlToLeft).Column
    LastRow = wsPreview.Cells(wsPreview.Rows.Count, 1).End(xlUp).Row

    For c = 1 To LastCol

        ColName = wsPreview.Cells(1, c).Value

        DataType = DetectColumnType(wsPreview, c, LastRow)

        Role = SuggestRole(ColName, DataType)

        wsMeta.Cells(c + 1, 1).Value = ColName
        wsMeta.Cells(c + 1, 2).Value = DataType
        wsMeta.Cells(c + 1, 3).Value = Role
        wsMeta.Cells(c + 1, 4).Value = Role
        wsMeta.Cells(c + 1, 5).Value = "No"

    Next c

End Sub

Private Function DetectColumnType(ws As Worksheet, Col As Long, LastRow As Long) As String

    Dim r As Long
    Dim NumCount As Long
    Dim DateCount As Long
    Dim TextCount As Long

    Dim v

    For r = 2 To LastRow

        v = ws.Cells(r, Col).Value

        If Trim(v) <> "" Then

            If IsDate(v) Then

                DateCount = DateCount + 1

            ElseIf IsNumeric(v) Then

                NumCount = NumCount + 1

            Else

                TextCount = TextCount + 1

            End If

        End If

    Next r

    If DateCount >= NumCount And DateCount >= TextCount Then

        DetectColumnType = "Date"

    ElseIf NumCount >= TextCount Then

        DetectColumnType = "Number"

    Else

        DetectColumnType = "Text"

    End If

End Function

Private Function SuggestRole(ColumnName As String, DataType As String) As String

    Dim Col As String

    Col = LCase(Trim(ColumnName))

    '=========================
    ' DATE COLUMNS
    '=========================
    If InStr(Col, "date") > 0 _
    Or InStr(Col, "time") > 0 _
    Or InStr(Col, "year") > 0 _
    Or InStr(Col, "month") > 0 _
    Or InStr(Col, "day") > 0 Then

        SuggestRole = "Date"
        Exit Function

    End If

    '=========================
    ' CATEGORY COLUMNS
    '=========================
    If Right(Col, 2) = "id" _
    Or InStr(Col, "code") > 0 _
    Or InStr(Col, "type") > 0 _
    Or InStr(Col, "category") > 0 _
    Or InStr(Col, "city") > 0 _
    Or InStr(Col, "country") > 0 _
    Or InStr(Col, "state") > 0 _
    Or InStr(Col, "region") > 0 _
    Or InStr(Col, "vendor") > 0 _
    Or InStr(Col, "department") > 0 _
    Or InStr(Col, "gender") > 0 _
    Or InStr(Col, "payment") > 0 _
    Or InStr(Col, "location") > 0 Then

        SuggestRole = "Category"
        Exit Function

    End If

    '=========================
    ' MEASURE COLUMNS
    '=========================
    If DataType = "Number" Then

        SuggestRole = "Measure"

    Else

        SuggestRole = "Category"

    End If

End Function
Sub LoadMetadataToForm()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long

    Set ws = Worksheets("Metadata")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    With frmAnalysisConfig

        .cmbDate.Clear
        .lstMeasures.Clear
        .lstCategory.Clear

        For i = 2 To LastRow

            Select Case ws.Cells(i, 4).Value

                Case "Date"

                    .cmbDate.AddItem ws.Cells(i, 1).Value

                Case "Measure"

                    .lstMeasures.AddItem ws.Cells(i, 1).Value

                Case "Category"

                    .lstCategory.AddItem ws.Cells(i, 1).Value

            End Select

        Next i

    End With

End Sub

