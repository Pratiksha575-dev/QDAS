VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnalysisConfig 
   Caption         =   "QDAS-Configure analysis"
   ClientHeight    =   12440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14390
   OleObjectBlob   =   "frmAnalysisConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAnalysisConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()

    'Load metadata
    LoadMetadataToForm
    
    If cmbDate.ListCount > 0 Then

        cmbDate.ListIndex = 0

        chkDaily.Enabled = True
        chkMonthly.Enabled = True
        chkWeekday.Enabled = True

        chkDaily.value = True
        chkMonthly.value = True
        chkWeekday.value = True

    Else

        chkDaily.Enabled = False
        chkMonthly.Enabled = False
        chkWeekday.Enabled = False

        chkDaily.value = False
        chkMonthly.value = False
        chkWeekday.value = False

    End If

    'Default selections
    optIQR.value = True


    'ListBox properties
    lstMeasures.MultiSelect = fmMultiSelectSingle
    lstCategory.MultiSelect = fmMultiSelectSingle
    
    cmdRun.Enabled = True
End Sub

Private Sub cmdRun_Click()

    'Measure validation
    If lstMeasures.ListIndex = -1 Then
        MsgBox "Please select a Measure.", vbExclamation, "Validation"
        Exit Sub
    End If

    'Category validation
    If lstCategory.ListIndex = -1 Then
        MsgBox "Please select a Category.", vbExclamation, "Validation"
        Exit Sub
    End If

    'Trend validation
    'Trend validation (only if Trend options are enabled)
If cmbDate.ListCount > 0 Then

    If (chkDaily.value Or chkMonthly.value Or chkWeekday.value) _
        And cmbDate.ListIndex = -1 Then

        MsgBox "Please select a Date column for Trend Analysis.", _
                vbExclamation, "Validation"

        Exit Sub

    End If

End If
    SaveConfiguration
    RunAnalysis

End Sub
