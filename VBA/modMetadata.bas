Attribute VB_Name = "ModMetadata"
Option Explicit

Public Sub CreateMetadata()

    Dim ws As Worksheet

    Set ws = GetOrCreateWorksheet("Metadata")

    ws.Cells.Clear

    ws.Range("A1").value = "Column Name"
    ws.Range("B1").value = "Detected Type"
    ws.Range("C1").value = "Suggested Role"
    ws.Range("D1").value = "Final Role"
    ws.Range("E1").value = "Confidence"

    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ws.Columns("A:E").HorizontalAlignment = xlCenter
    ws.Columns("A:E").VerticalAlignment = xlCenter

End Sub

Public Sub GenerateMetadata()

    Dim wsPreview As Worksheet
    Dim wsMeta As Worksheet

    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long

    Dim DataType As String
    Dim Role As String
    Dim Confidence As Integer

    Set wsPreview = GetOrCreateWorksheet("PreviewData")

    If wsPreview.Cells(1, 1).value = "" Then
        MsgBox "PreviewData sheet is empty.", vbExclamation
        Exit Sub
    End If

    CreateMetadata
    Set wsMeta = GetOrCreateWorksheet("Metadata")

    lastCol = wsPreview.Cells(1, wsPreview.Columns.count).End(xlToLeft).Column
    lastRow = wsPreview.Cells(wsPreview.Rows.count, 1).End(xlUp).Row

    For c = 1 To lastCol

        AnalyzeColumn _
            wsPreview, _
            c, _
            lastRow, _
            DataType, _
            Role, _
            Confidence

        wsMeta.Cells(c + 1, 1).value = wsPreview.Cells(1, c).value
        wsMeta.Cells(c + 1, 2).value = DataType
        wsMeta.Cells(c + 1, 3).value = Role
        wsMeta.Cells(c + 1, 4).value = Role
        wsMeta.Cells(c + 1, 5).value = Confidence & "%"
        
    Next c

    wsMeta.Columns.AutoFit

End Sub
Public Sub AnalyzeColumn( _
        ws As Worksheet, _
        ColNum As Long, _
        lastRow As Long, _
        ByRef DataType As String, _
        ByRef Role As String, _
        ByRef Confidence As Integer)

    Dim colName As String

    Dim NumericRatio As Double
    Dim DateRatio As Double
    Dim DistinctCount As Long
    Dim TotalSamples As Long
    Dim UniqueRatio As Double

    colName = Trim(ws.Cells(1, ColNum).value)

    '-----------------------------------------
    ' Sample Statistics
    '-----------------------------------------
    NumericRatio = GetNumericRatio(ws, ColNum, lastRow)
    DateRatio = GetDateRatio(ws, ColNum, lastRow)
    DistinctCount = GetDistinctCount(ws, ColNum, lastRow)
    TotalSamples = WorksheetFunction.Min(100, lastRow - 1)

    If TotalSamples > 0 Then
        UniqueRatio = DistinctCount / TotalSamples
    Else
        UniqueRatio = 0
    End If

    '-----------------------------------------
    ' Detect Data Type
    '-----------------------------------------
    DataType = DetectDataType( _
                    colName, _
                    NumericRatio, _
                    DateRatio)

    '-----------------------------------------
    ' Suggest Role
    '-----------------------------------------
    Role = SuggestRole( _
                colName, _
                DataType, _
                NumericRatio, _
                UniqueRatio, _
                DistinctCount)

    '-----------------------------------------
    ' Confidence
    '-----------------------------------------
    Confidence = GetConfidence( _
                    DataType, _
                    Role, _
                    NumericRatio, _
                    DateRatio, _
                    UniqueRatio)

End Sub
Private Function GetNumericRatio( _
            ws As Worksheet, _
            ColNum As Long, _
            lastRow As Long) As Double

    Dim r As Long
    Dim SampleCount As Long
    Dim NumericCount As Long
    Dim MaxSample As Long
    Dim v As Variant

    MaxSample = WorksheetFunction.Min(101, lastRow)

    For r = 2 To MaxSample

        v = ws.Cells(r, ColNum).value

        If Not IsError(v) Then

            If Trim(CStr(v)) <> "" Then

                SampleCount = SampleCount + 1

                'Accept numbers stored as text also
                If IsNumeric(v) Then
                    NumericCount = NumericCount + 1
                End If

            End If

        End If

    Next r

    If SampleCount = 0 Then

        GetNumericRatio = 0

    Else

        GetNumericRatio = (NumericCount / SampleCount) * 100

    End If

End Function
Private Function GetDateRatio( _
            ws As Worksheet, _
            ColNum As Long, _
            lastRow As Long) As Double

    Dim r As Long
    Dim SampleCount As Long
    Dim DateCount As Long
    Dim MaxSample As Long
    Dim v As Variant
    Dim s As String

    MaxSample = WorksheetFunction.Min(101, lastRow)

    For r = 2 To MaxSample

        v = ws.Cells(r, ColNum).value

        If Not IsError(v) Then

            s = Trim(CStr(v))

            If s <> "" Then

                SampleCount = SampleCount + 1

                'Only check text-like values for dates.
                'Ignore plain numeric values (Excel date serials).
                If Not IsNumeric(v) Then

                    If IsDate(v) Then
                        DateCount = DateCount + 1
                    End If

                End If

            End If

        End If

    Next r

    If SampleCount = 0 Then

        GetDateRatio = 0

    Else

        GetDateRatio = (DateCount / SampleCount) * 100

    End If

End Function
Private Function GetDistinctCount( _
            ws As Worksheet, _
            ColNum As Long, _
            lastRow As Long) As Long

    Dim dict As Object
    Dim r As Long
    Dim MaxSample As Long
    Dim v As Variant
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    MaxSample = WorksheetFunction.Min(101, lastRow)

    For r = 2 To MaxSample

        v = ws.Cells(r, ColNum).value

        If Not IsError(v) Then

            key = Trim(CStr(v))

            If key <> "" Then

                If Not dict.Exists(key) Then
                    dict.Add key, 1
                End If

            End If

        End If

    Next r

    GetDistinctCount = dict.count

End Function


Sub LoadMetadataToForm()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    ' Self-heal sheet if the form initializes before metadata generation passes
    Set ws = GetOrCreateWorksheet("Metadata")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub ' Guard clause if sheet is completely empty

    With frmAnalysisConfig
        .cmbDate.Clear
        .lstMeasures.Clear
        .lstCategory.Clear

        For i = 2 To lastRow
            Select Case ws.Cells(i, 4).value
                Case "Date"
                    .cmbDate.AddItem ws.Cells(i, 1).value
                Case "Measure"
                    .lstMeasures.AddItem ws.Cells(i, 1).value
                Case "Category"
                    .lstCategory.AddItem ws.Cells(i, 1).value
            End Select
        Next i
    End With
End Sub

Private Function DetectDataType( _
            colName As String, _
            NumericRatio As Double, _
            DateRatio As Double) As String

    Dim NameLower As String

    NameLower = LCase(Trim(colName))

    '--------------------------------------
    ' Check Column Name First
    '--------------------------------------

    If InStr(NameLower, "date") > 0 _
    Or InStr(NameLower, "time") > 0 _
    Or InStr(NameLower, "year") > 0 _
    Or InStr(NameLower, "month") > 0 _
    Or InStr(NameLower, "day") > 0 _
    Or InStr(NameLower, "pickup") > 0 _
    Or InStr(NameLower, "dropoff") > 0 Then

        DetectDataType = "Date"
        Exit Function

    End If

    '--------------------------------------
    ' Detect Boolean
    '--------------------------------------

    If Left(NameLower, 3) = "is_" _
    Or Left(NameLower, 4) = "has_" _
    Or Right(NameLower, 5) = "_flag" _
    Or InStr(NameLower, "active") > 0 _
    Or InStr(NameLower, "enabled") > 0 Then

        DetectDataType = "Boolean"
        Exit Function

    End If

    '--------------------------------------
    ' Detect Date using data
    '--------------------------------------

    If DateRatio >= 80 Then
        DetectDataType = "Date"
        Exit Function
    End If

    '--------------------------------------
    ' Detect Number
    '--------------------------------------

    If NumericRatio >= 80 Then
        DetectDataType = "Number"
    Else
        DetectDataType = "Text"
    End If

End Function

Private Function SuggestRole( _
            colName As String, _
            DataType As String, _
            NumericRatio As Double, _
            UniqueRatio As Double, _
            DistinctCount As Long) As String

    Dim NameLower As String

    NameLower = LCase(Trim(colName))

    '-------------------------------
    ' Date Columns
    '-------------------------------
    If DataType = "Date" Then
        SuggestRole = "Date"
        Exit Function
    End If

    '-------------------------------
    ' Boolean Columns
    '-------------------------------
    If DataType = "Boolean" Then
        SuggestRole = "Category"
        Exit Function
    End If

    '-------------------------------
    ' Identifier Columns
    '-------------------------------
    If ContainsKeyword(NameLower, GetIDKeywords()) Then
        SuggestRole = "Category"
        Exit Function
    End If

    '-------------------------------
    ' Measure Columns
    '-------------------------------
    If ContainsKeyword(NameLower, GetMeasureKeywords()) Then
        SuggestRole = "Measure"
        Exit Function
    End If

    '-------------------------------
    ' Text Columns
    '-------------------------------
    If DataType = "Text" Then
        SuggestRole = "Category"
        Exit Function
    End If

    '-------------------------------
    ' Numeric Columns
    '-------------------------------
    If DataType = "Number" Then

        If DistinctCount <= 15 Then

            SuggestRole = "Category"

        ElseIf UniqueRatio < 0.25 Then

            SuggestRole = "Category"

        Else

            SuggestRole = "Measure"

        End If

        Exit Function

    End If

    SuggestRole = "Category"

End Function
Private Function ContainsKeyword(Text As String, Keywords As Variant) As Boolean

    Dim i As Long

    For i = LBound(Keywords) To UBound(Keywords)

        If InStr(Text, Keywords(i)) > 0 Then
            ContainsKeyword = True
            Exit Function
        End If

    Next i

End Function
Private Function GetMeasureKeywords() As Variant

    GetMeasureKeywords = Array( _
        "amount", _
        "fare", _
        "distance", _
        "price", _
        "cost", _
        "sales", _
        "profit", _
        "revenue", _
        "quantity", _
        "weight", _
        "score", _
        "rating", _
        "total", _
        "value", _
        "count", _
        "duration", _
        "time_spent", _
        "temperature", _
        "speed", _
        "volume", _
        "size")

End Function
Private Function GetIDKeywords() As Variant

    GetIDKeywords = Array( _
        "id", _
        "_id", _
        "code", _
        "key", _
        "number", _
        "no", _
        "serial", _
        "reference", _
        "identifier")

End Function
Private Function GetConfidence( _
        DataType As String, _
        Role As String, _
        NumericRatio As Double, _
        DateRatio As Double, _
        UniqueRatio As Double) As Integer

    If DataType = "Date" And DateRatio >= 80 Then
        GetConfidence = 100

    ElseIf Role = "Measure" And NumericRatio >= 95 Then
        GetConfidence = 95

    ElseIf Role = "Category" And UniqueRatio <= 0.25 Then
        GetConfidence = 90

    ElseIf DataType = "Text" Then
        GetConfidence = 90

    Else
        GetConfidence = 75
    End If

End Function


