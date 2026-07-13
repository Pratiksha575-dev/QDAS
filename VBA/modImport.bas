Attribute VB_Name = "ModImport"
Option Explicit

Public SelectedFile As String

Sub ImportDataset()

    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .Title = "Select Dataset"

        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"

        .AllowMultiSelect = False

        If .Show <> -1 Then Exit Sub

        SelectedFile = .SelectedItems(1)

    End With

    MsgBox "Dataset Selected:" & vbCrLf & SelectedFile

    GenerateMetadata
    MsgBox "Metadata Generated Successfully!"

End Sub
