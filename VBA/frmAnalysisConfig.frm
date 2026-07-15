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
    End If

    'Default selections
    optIQR.Value = True

    chkDaily.Value = True
    chkWeekly.Value = False
    chkMonthly.Value = True
    chkHourly.Value = False
    chkWeekday.Value = True

    'ListBox properties
    lstMeasures.MultiSelect = fmMultiSelectMulti
    lstCategory.MultiSelect = fmMultiSelectMulti
    
    cmdRun.Enabled = True
    optAutoBins.Value = True
    txtBins.Text = "20"
    txtBins.Enabled = False
    
End Sub

Private Sub cmdRun_Click()

    SaveConfiguration

    RunAnalysis

End Sub

Private Sub optAutoBins_Click()

    txtBins.Enabled = False

End Sub

Private Sub optCustomBins_Click()

    txtBins.Enabled = True

End Sub
