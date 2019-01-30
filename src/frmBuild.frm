VERSION 5.00
Begin VB.Form frmBuild 
   Caption         =   "Build code.."
   ClientHeight    =   3225
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   3840
   Begin VB.CheckBox chkTest 
      Caption         =   "Test run"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkVersion 
      Caption         =   "Version"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkStrip 
      Caption         =   "Strip debug version"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CheckBox chkParseOnly 
      Caption         =   "Parse only"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkList 
      Caption         =   "List"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtCompiledFile 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtScriptFile 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Compiled file:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Script file:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' the script file
Dim strScriptFile As String

'the user wants to parse the file only, make sure we won't trick the user in believing that we could run the code
Private Sub chkParseOnly_Click()
  setTestCheck
End Sub

'the user want's to build the code
Private Sub cmdBuild_Click()
  Dim cmd As String
  cmd = gOptions.enviroment.pathToCompiler
  If int2bool(chkVersion.Value) Then
    cmd = cmd & " -v"
  Else
    cmd = cmd & " -o " & txtCompiledFile.text
    If int2bool(chkList.Value) Then cmd = cmd & " -l "
    If int2bool(chkParseOnly.Value) Then cmd = cmd & " -p "
    If int2bool(chkStrip.Value) Then cmd = cmd & " -s "
    cmd = cmd & " -- " & txtScriptFile
  End If
  
  fMainForm.setStatus "Compiling code.."
  executeCmd cmd, True
  
  If int2bool(chkTest.Value) And chkTest.Enabled Then
    If gOptions.enviroment.hasHost Then
      Dim runCmd As String
      
      fMainForm.setStatus "Running code.."
      runCmd = gOptions.enviroment.pathToHost
      runCmd = Replace(runCmd, "[script]", txtCompiledFile.text)
      executeCmd runCmd
    Else
      MsgBox "You are missing the host, plase fix this in the options dialog"
    End If
  End If
  
  fMainForm.setReadyStatus
  
  Unload Me
End Sub

'enable/disable the test check-box
Private Sub setTestCheck()
  If Not int2bool(chkParseOnly.Value) And gOptions.enviroment.hasHost Then
    chkTest.Enabled = True
  Else
    chkTest.Enabled = False
  End If
End Sub

'disable all but the version chek-boxes
Private Sub disableAll()
  enableText txtCompiledFile, False
  enableText txtScriptFile, False
  chkList.Enabled = False
  chkParseOnly.Enabled = False
  chkStrip.Enabled = False
  chkTest.Enabled = False
End Sub

'enable all but the version chek-boxes
Private Sub enableAll()
  enableText txtCompiledFile, True
  enableText txtScriptFile, True
  chkList.Enabled = True
  chkParseOnly.Enabled = True
  chkStrip.Enabled = True
  setTestCheck
End Sub

'user changed the check-version flag
Private Sub chkVersion_Click()
  If int2bool(chkVersion.Value) Then
    disableAll
  Else
    enableAll
  End If
End Sub

'abort the build dialog
Private Sub cmdAbort_Click()
  Unload Me
End Sub

'set the standard values
Private Sub Form_Load()
  'strCriptFile = "lua.test"
  txtScriptFile.text = strScriptFile
  'txtCompiledFile = "lua.obj"
  txtCompiledFile.text = stripFileName(strScriptFile) & ".obj"
  If Not gOptions.enviroment.hasCompiler Then
    chkVersion.Enabled = False
    disableAll
    Exit Sub
  End If
  setTestCheck
End Sub

'show a build dialog, this should be used instead of the show command
Public Sub ShowBuild(script As String)
  strScriptFile = script
  Show vbModal, fMainForm
End Sub

'resize the form
Private Sub Form_Resize()
  ' is there a better way to keep the min width / height?
  If Width < 3960 Then Width = 3960
  'i couldn't find any function that limmits(min width etc) the the resize so I'll have to do it here
  Height = 3630
  
  txtScriptFile.Width = Width - 1725
  txtCompiledFile.Width = Width - 1725
  cmdAbort.Left = Width - 2760
  cmdBuild.Left = Width - 1440
End Sub

'called whenthe text has changed
'we shouln't be able to build if we don't have a output, or input file
Private Sub setBuildEnable()
  If Len(txtScriptFile.text) > 0 And Len(txtCompiledFile.text) > 0 Then
    cmdBuild.Enabled = True
  Else
    cmdBuild.Enabled = False
  End If
End Sub

'we have changed the compiled file path
Private Sub txtCompiledFile_Change()
  setBuildEnable
End Sub

'we have changed the script file path
Private Sub txtScriptFile_Change()
  setBuildEnable
End Sub
