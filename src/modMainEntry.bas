Attribute VB_Name = "modMainEntry"
Option Explicit

'our reference to the main form
Public fMainForm As frmMain

'the main entry
Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    loadOptions
    frmTip.forceTipToShow = False
    Unload frmSplash
    
    fMainForm.Show
    If frmTip.showAtStartup Then
      frmTip.Show
    End If
End Sub
