Attribute VB_Name = "modCategoryEngine"
Option Explicit

Function BuildCategoryAggregation() As String

    Dim Item As Variant
    Dim Agg As String

    Agg = ""

    For Each Item In Measures

        Agg = Agg & _
        "{""Sum_" & Item & """, each List.Sum([" & Item & "]), type number},"

    Next Item

    If Right(Agg, 1) = "," Then
        Agg = Left(Agg, Len(Agg) - 1)
    End If

    BuildCategoryAggregation = Agg

End Function
Function BuildCategoryList() As String

    Dim Item As Variant
    Dim CategoryList As String

    CategoryList = ""

    For Each Item In Categories

        CategoryList = CategoryList & """" & Item & ""","

    Next Item

    If Right(CategoryList, 1) = "," Then
        CategoryList = Left(CategoryList, Len(CategoryList) - 1)
    End If

    BuildCategoryList = CategoryList

End Function
Function BuildCategoryMCode() As String

    Dim MCode As String

    ReadConfiguration

    MCode = ""

    MCode = MCode & "let" & vbCrLf

    MCode = MCode & _
        "    Source = #""nyc-taxidata""," & vbCrLf

    MCode = MCode & _
    "    #""Grouped Rows"" = Table.Group(Source,{" & _
    BuildCategoryList & "},{" & _
    BuildCategoryAggregation & "})" & vbCrLf

    MCode = MCode & _
        "in" & vbCrLf

    MCode = MCode & _
        "    #""Grouped Rows"""

    BuildCategoryMCode = MCode

End Function

Sub UpdateCategoryQuery()

    UpdateQuery "PQ_Category", BuildCategoryMCode()

End Sub
Sub GenerateCategoryQuery()

    ReadConfiguration

    UpdateCategoryQuery

    MsgBox "Category Query Generated Successfully."

End Sub

Sub TestCategoryMCode()

    Debug.Print BuildCategoryMCode()

End Sub
