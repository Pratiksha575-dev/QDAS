Attribute VB_Name = "modCategoryEngine"
Option Explicit

Function BuildCategoryAggregation() As String

    Dim MeasureName As String

    MeasureName = Measures.item(1)

    BuildCategoryAggregation = _
        "{""Measure"", each List.Sum([" & MeasureName & "]), type number}"

End Function

Function BuildCategoryList() As String
    Dim item As Variant
    Dim CategoryList As String

    CategoryList = ""
    For Each item In Categories
        CategoryList = CategoryList & """" & item & ""","
    Next item

    If Right(CategoryList, 1) = "," Then
        CategoryList = Left(CategoryList, Len(CategoryList) - 1)
    End If

    BuildCategoryList = CategoryList
End Function

Function BuildCategoryMCode() As String
    Dim MCode As String

    ReadConfiguration
    
    MCode = "let" & vbCrLf & _
        "    Source = PQ_RawData," & vbCrLf & _
        "    #""Grouped Rows"" = Table.Group(Source, {" & BuildCategoryList() & "}, {" & BuildCategoryAggregation() & "})," & vbCrLf & _
        "    #""Renamed Columns"" = Table.RenameColumns(#""Grouped Rows"", {{""" & Categories.item(1) & """, ""Category""}})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Renamed Columns"""

    BuildCategoryMCode = MCode
End Function

Sub UpdateCategoryQuery()
    UpdateQuery "PQ_Category", BuildCategoryMCode()
End Sub

Sub GenerateCategoryQuery()

    ReadConfiguration

    If Measures.count <> 1 Then
        MsgBox "Please select exactly one Measure for Category Analysis.", vbExclamation
        Exit Sub
    End If

    If Categories.count = 0 Then
        MsgBox "Please select at least one Category.", vbExclamation
        Exit Sub
    End If

    If Categories.count > 2 Then
        MsgBox "Please select a maximum of two Categories.", vbExclamation
        Exit Sub
    End If

    UpdateCategoryQuery

    MsgBox "Category Query Generated Successfully.", vbInformation

End Sub

