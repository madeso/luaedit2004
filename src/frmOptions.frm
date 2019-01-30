VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin MSComDlg.CommonDialog cdl 
      Left            =   600
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab ssMainTab 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Timesaving"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "ts_min"
      Tab(0).Control(2)=   "ts_max"
      Tab(0).Control(3)=   "ts_current"
      Tab(0).Control(4)=   "ts_chkDoTimeSaving"
      Tab(0).Control(5)=   "ts_sInterval"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Enviroment"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "env_txtAppPath"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "env_cmdAppPath"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "env_txtCompPath"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "env_cmdCompPath"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "env_chkAppPath"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "env_chkCompilerPath"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Colors"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "clr_lsTypes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "clr_frmColor"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame clr_frmColor 
         Caption         =   "Color of: All"
         Height          =   3135
         Left            =   -72840
         TabIndex        =   17
         Top             =   600
         Width           =   3495
         Begin RichTextLib.RichTextBox clr_txtExample 
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   2640
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            DisableNoScroll =   -1  'True
            TextRTF         =   $"frmOptions.frx":0054
         End
         Begin VB.CheckBox clr_chkStrikeThru 
            Caption         =   "Strike-thru"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox clr_chkItalic 
            Caption         =   "Italic"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox clr_chkBold 
            Caption         =   "Bold"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox clr_chkUnderline 
            Caption         =   "Underline"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Value           =   2  'Grayed
            Width           =   2175
         End
         Begin VB.ComboBox clr_cmbColor 
            Height          =   315
            Left            =   960
            TabIndex        =   19
            Text            =   "Combo2"
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox clr_cmbSize 
            Height          =   315
            ItemData        =   "frmOptions.frx":0120
            Left            =   960
            List            =   "frmOptions.frx":0122
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "Example:"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Color"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Size"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.ListBox clr_lsTypes 
         Height          =   2595
         ItemData        =   "frmOptions.frx":0124
         Left            =   -74640
         List            =   "frmOptions.frx":013D
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   1575
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmOptions.frx":0182
         Top             =   2400
         Width           =   5415
      End
      Begin VB.CheckBox env_chkCompilerPath 
         Caption         =   "Path to compiler"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox env_chkAppPath 
         Caption         =   "Path to host application"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton env_cmdCompPath 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox env_txtCompPath 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1800
         Width           =   4815
      End
      Begin VB.CommandButton env_cmdAppPath 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox env_txtAppPath 
         Height          =   405
         Left            =   240
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   4815
      End
      Begin MSComctlLib.Slider ts_sInterval 
         Height          =   615
         Left            =   -74880
         TabIndex        =   6
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         _Version        =   393216
         Min             =   1
         Max             =   60
         SelStart        =   1
         TickFrequency   =   10
         Value           =   1
      End
      Begin VB.CheckBox ts_chkDoTimeSaving 
         Caption         =   "Do time saving"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label ts_current 
         Alignment       =   2  'Center
         Caption         =   "current"
         Height          =   255
         Left            =   -74040
         TabIndex        =   28
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label ts_max 
         Alignment       =   1  'Right Justify
         Caption         =   "max"
         Height          =   255
         Left            =   -72000
         TabIndex        =   8
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label ts_min 
         Caption         =   "min"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Interval in minutes"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'red, blue, etc
Private mColorNames(1 To MAX_COLORS) As String
'the real color
Private mColors(1 To MAX_COLORS) As Variant

Private mOptions As TOptions ' represent the current options made in this dialog
Private mSpecialColor As TColor ' represent the current color in the current dialog
Private mColorType As Integer ' represent the current color type: keyword, operator...

' store the color if it is used
' apply the color to the example text
Private Sub applyColor()
  setColorOfRichText mSpecialColor, clr_txtExample, 0, Len(clr_txtExample.text)
  'update mOptions
  
  Select Case mColorType
  Case -1
    Debug.Print "Storing nothing"
  Case 0
    Debug.Print "Storing nothing(all)"
  Case 1
    Debug.Print "Storing keywords"
    mOptions.colors.keyword = mSpecialColor
  Case 2
    Debug.Print "Storing number"
    mOptions.colors.number = mSpecialColor
  Case 3
    Debug.Print "Storing operator"
    mOptions.colors.operator = mSpecialColor
  Case 4
    Debug.Print "Storing String"
    mOptions.colors.text = mSpecialColor
  Case 5
    Debug.Print "Storing Comment"
    mOptions.colors.comment = mSpecialColor
  Case 6
    Debug.Print "Storing normal"
    mOptions.colors.normal = mSpecialColor
  Case Else
    Debug.Print "Error"
    MsgBox "Error in select statement:applyColor:frmOptions"
  End Select
  
End Sub

'store the local optrions in the global options
Private Sub doStore()
  gOptions = mOptions
End Sub

'tell the user what the current timesaving is
Private Sub ts_setCurrent()
  ts_current = "(" & CStr(ts_sInterval.selStart) & ")"
End Sub

'user clicked to change the bold setting
Private Sub clr_chkBold_Click()
  If mColorType = -1 Then
    Exit Sub
  End If
  
  If mColorType = 0 Then
    mOptions.colors.comment.bold = int2bool(clr_chkBold.Value)
    mOptions.colors.keyword.bold = int2bool(clr_chkBold.Value)
    mOptions.colors.normal.bold = int2bool(clr_chkBold.Value)
    mOptions.colors.number.bold = int2bool(clr_chkBold.Value)
    mOptions.colors.operator.bold = int2bool(clr_chkBold.Value)
    mOptions.colors.text.bold = int2bool(clr_chkBold.Value)
  Else
    mSpecialColor.bold = int2bool(clr_chkBold.Value)
  End If
  
  handleChange
  applyColor
End Sub

'user clicked to change the italic setting
Private Sub clr_chkItalic_Click()
  If mColorType = -1 Then
    Exit Sub
  End If
  
  If mColorType = 0 Then
    mOptions.colors.comment.italic = int2bool(clr_chkItalic.Value)
    mOptions.colors.keyword.italic = int2bool(clr_chkItalic.Value)
    mOptions.colors.normal.italic = int2bool(clr_chkItalic.Value)
    mOptions.colors.number.italic = int2bool(clr_chkItalic.Value)
    mOptions.colors.operator.italic = int2bool(clr_chkItalic.Value)
    mOptions.colors.text.italic = int2bool(clr_chkItalic.Value)
  Else
    mSpecialColor.italic = int2bool(clr_chkItalic.Value)
  End If
  
  handleChange
  applyColor
End Sub

'user clicked to change the strike thru setting
Private Sub clr_chkStrikeThru_Click()
  If mColorType = -1 Then
    Exit Sub
  End If
  
  If mColorType = 0 Then
    mOptions.colors.comment.StrikeThru = int2bool(clr_chkStrikeThru.Value)
    mOptions.colors.keyword.StrikeThru = int2bool(clr_chkStrikeThru.Value)
    mOptions.colors.normal.StrikeThru = int2bool(clr_chkStrikeThru.Value)
    mOptions.colors.number.StrikeThru = int2bool(clr_chkStrikeThru.Value)
    mOptions.colors.operator.StrikeThru = int2bool(clr_chkStrikeThru.Value)
    mOptions.colors.text.StrikeThru = int2bool(clr_chkStrikeThru.Value)
  Else
    mSpecialColor.StrikeThru = int2bool(clr_chkStrikeThru.Value)
  End If
  
  handleChange
  applyColor
End Sub

'user clicked to change the underline setting
Private Sub clr_chkUnderline_Click()
  If mColorType = -1 Then
    Exit Sub
  End If
  
  If mColorType = 0 Then
    mOptions.colors.comment.underline = int2bool(clr_chkUnderline.Value)
    mOptions.colors.keyword.underline = int2bool(clr_chkUnderline.Value)
    mOptions.colors.normal.underline = int2bool(clr_chkUnderline.Value)
    mOptions.colors.number.underline = int2bool(clr_chkUnderline.Value)
    mOptions.colors.operator.underline = int2bool(clr_chkUnderline.Value)
    mOptions.colors.text.underline = int2bool(clr_chkUnderline.Value)
  Else
    mSpecialColor.underline = int2bool(clr_chkUnderline.Value)
  End If
  
  handleChange
  applyColor
End Sub

'user changed the color setting
Private Sub clr_cmbColor_Change()
  If mColorType = -1 Then
    Exit Sub
  End If
  
  If mColorType = 0 Then
    mOptions.colors.comment.color = findColorFromText(clr_cmbColor.text)
    mOptions.colors.keyword.color = findColorFromText(clr_cmbColor.text)
    mOptions.colors.normal.color = findColorFromText(clr_cmbColor.text)
    mOptions.colors.number.color = findColorFromText(clr_cmbColor.text)
    mOptions.colors.operator.color = findColorFromText(clr_cmbColor.text)
    mOptions.colors.text.color = findColorFromText(clr_cmbColor.text)
  Else
    mSpecialColor.color = findColorFromText(clr_cmbColor.text)
  End If
  
  handleChange
  applyColor
End Sub

'user clicked to change the color setting
Private Sub clr_cmbColor_Click()
  clr_cmbColor_Change
End Sub

'user changed the size setting
Private Sub clr_cmbSize_Change()
  Dim str As String
  
  If mColorType = -1 Then
    Exit Sub
  End If
  
  str = clr_cmbSize.text
  If IsNumeric(str) Then
    mSpecialColor.size = CInt(str)
    handleChange
    applyColor
    'Debug.Print str
  End If
End Sub

'user clicked to change the size setting
Private Sub clr_cmbSize_Click()
  clr_cmbSize_Change
End Sub

'user changed the type to color
Private Sub clr_lsTypes_Click()
  'Debug.Print CStr(clr_lsTypes.ListIndex)
  clr_frmColor.Caption = "Color of: " & clr_lsTypes.List(clr_lsTypes.ListIndex)
  
  mColorType = -1
  
  Select Case clr_lsTypes.ListIndex
  Case 0
    Debug.Print "Filling standard"
    fillColorFormWithNothing
  Case 1
    Debug.Print "Filling keyword"
    fillColorForm mOptions.colors.keyword
  Case 2
    Debug.Print "Filling number"
    fillColorForm mOptions.colors.number
  Case 3
    Debug.Print "Filling operator"
    fillColorForm mOptions.colors.operator
  Case 4
    Debug.Print "Filling String"
    fillColorForm mOptions.colors.text
  Case 5
    Debug.Print "Filling Comment"
    fillColorForm mOptions.colors.comment
  Case 6
    Debug.Print "Filling Normal"
    fillColorForm mOptions.colors.normal
  Case Else
    MsgBox "Error in select statement:clr_lsTypes_Click:frmOptions"
  End Select
  
  mColorType = clr_lsTypes.ListIndex
  applyColor
End Sub

'select the standard color(for all)
Private Sub fillColorFormWithNothing()
  'If c = Nothing Then
  Dim all As TColor
  
  all.size = 10
  all.color = vbBlack
  all.bold = False
  all.italic = False
  all.StrikeThru = False
  all.underline = False
  
  fillColorForm all
End Sub

'find the name of the color parameter
Private Function findColorText(c As Variant) As String
  Dim i As Integer
  findColorText = mColorNames(1)
  For i = 1 To MAX_COLORS
    'apperently the Settings function only can control strings
    'converting them both to strings make sure equal are equal
    If CStr(mColors(i)) = CStr(c) Then
      findColorText = mColorNames(i)
      Exit Function
    End If
  Next i
End Function

'find the color based on the name
Private Function findColorFromText(c As String) As Variant
  Dim i As Integer
  
  findColorFromText = mColors(1)
  For i = 1 To MAX_COLORS
    If mColorNames(i) = c Then
      findColorFromText = mColors(i)
      Exit Function
    End If
  Next i
End Function

'set the colorformas the specified color
Private Sub fillColorForm(ByRef c As TColor)
  'set gui
  clr_cmbSize.text = CStr(c.size)
  clr_cmbColor.text = findColorText(c.color)
  clr_chkBold = bool2int(c.bold)
  clr_chkItalic = bool2int(c.italic)
  clr_chkStrikeThru = bool2int(c.StrikeThru)
  clr_chkUnderline = bool2int(c.underline)
  
  'set member data
  mSpecialColor = c
End Sub

'apply the color to the global setting
Private Sub cmdApply_Click()
    doStore
    cmdApply.Enabled = False
End Sub

'abort
Private Sub cmdCancel_Click()
    Unload Me
End Sub


'accept change
Private Sub cmdOK_Click()
    doStore
    Unload Me
End Sub

'change the application path
Private Sub env_cmdAppPath_Click()
  On Error Resume Next
  Debug.Print "load doc"
  With cdl
    .DialogTitle = "Select host.."
    .CancelError = True
    .filename = ""
    .Filter = "Application files (*.exe)|*.exe|All files (*.*)|*.*"
    .DefaultExt = "exe"
    .ShowOpen
    
    If Err <> MSComDlg.cdlCancel Then
      env_txtAppPath = .filename
    End If
  End With
End Sub

'change the compiler path
Private Sub env_cmdCompPath_Click()
  On Error Resume Next
  Debug.Print "load doc"
  With cdl
    .DialogTitle = "Select compiler.."
    .CancelError = True
    .filename = ""
    .Filter = "Application files (*.exe)|*.exe|All files (*.*)|*.*"
    .DefaultExt = "exe"
    .ShowOpen
    
    If Err <> MSComDlg.cdlCancel Then
      env_txtCompPath = .filename
    End If
  End With
End Sub

'set the standard values, and fill the "color collection"
Private Sub Form_Load()
  Dim i As Integer
  mOptions = gOptions
  cmdApply.Enabled = False
  
  mColorType = -1
  
  mColorNames(1) = "Red"
  mColors(1) = vbRed
  mColorNames(2) = "Black"
  mColors(2) = vbBlack
  mColorNames(3) = "Magenta"
  mColors(3) = vbMagenta
  mColorNames(4) = "Blue"
  mColors(4) = vbBlue
  mColorNames(5) = "Green"
  mColors(5) = vbGreen
  mColorNames(6) = "White"
  mColors(6) = vbWhite
  
  clr_cmbSize.Clear
  For i = MIN_FONT_SIZE To MAX_FONT_SIZE
    clr_cmbSize.AddItem (CStr(i))
    clr_cmbSize.text = CStr(i)
  Next i
  
  For i = 1 To MAX_COLORS
    clr_cmbColor.AddItem (mColorNames(i))
  Next i
  
  ts_chkDoTimeSaving = bool2int(mOptions.timesaving.useTimeSaving)
  ts_sInterval.selStart = mOptions.timesaving.interval
  ts_sInterval.Enabled = mOptions.timesaving.useTimeSaving
  
  env_chkAppPath = bool2int(mOptions.enviroment.hasHost)
  env_chkCompilerPath = bool2int(mOptions.enviroment.hasCompiler)
  env_txtCompPath = mOptions.enviroment.pathToCompiler
  env_txtAppPath = mOptions.enviroment.pathToHost
  enableText env_txtCompPath, mOptions.enviroment.hasCompiler
  env_cmdCompPath.Enabled = mOptions.enviroment.hasCompiler
  enableText env_txtAppPath, mOptions.enviroment.hasHost
  env_cmdAppPath.Enabled = mOptions.enviroment.hasHost
  
  cmdApply.Enabled = False
  
  ts_min = CStr(ts_sInterval.Min)
  ts_max = CStr(ts_sInterval.Max)
  ts_setCurrent
  
End Sub

'if something has changed
Private Sub handleChange()
  cmdApply.Enabled = True
End Sub

'user enabled/disabled the timesaving
Private Sub ts_chkDoTimeSaving_Click()
  handleChange
  mOptions.timesaving.useTimeSaving = int2bool(ts_chkDoTimeSaving)
  
  ts_sInterval.Enabled = mOptions.timesaving.useTimeSaving
End Sub

'user changed the timesaving interval
Private Sub ts_sInterval_Change()
  handleChange
  mOptions.timesaving.interval = ts_sInterval.selStart
  ts_setCurrent
End Sub

'user enabled/disabled the application/host path
Private Sub env_chkAppPath_Click()
  handleChange
  mOptions.enviroment.hasHost = int2bool(env_chkAppPath)
  enableText env_txtAppPath, mOptions.enviroment.hasHost
  env_cmdAppPath.Enabled = mOptions.enviroment.hasHost
End Sub

'user changed the compiler path
Private Sub env_chkCompilerPath_Click()
  handleChange
  mOptions.enviroment.hasCompiler = int2bool(env_chkCompilerPath)
  enableText env_txtCompPath, mOptions.enviroment.hasCompiler
  env_cmdCompPath.Enabled = mOptions.enviroment.hasCompiler
End Sub

'the appliationpath has changed
Private Sub env_txtAppPath_Change()
  handleChange
  mOptions.enviroment.pathToHost = env_txtAppPath
End Sub


'the compiler path has changed
Private Sub env_txtCompPath_Change()
  handleChange
  mOptions.enviroment.pathToCompiler = env_txtCompPath
End Sub
