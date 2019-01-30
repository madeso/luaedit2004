VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "LuaEdit2004"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6435
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run Code"
            Object.ToolTipText     =   "Run Code"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Build Code"
            Object.ToolTipText     =   "Build Code"
            ImageKey        =   "Spell Check"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            ImageKey        =   "Properties"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4665
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5715
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2004-09-13"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "21:21"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "Spell Check"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "Properties"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^W
      End
      Begin VB.Menu mmuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mmuToolsExportAs 
         Caption         =   "Export as"
         Begin VB.Menu mnuToolsExportRtf 
            Caption         =   "Rich Text Format (RTF)..."
         End
         Begin VB.Menu mnuToolsExportHtml 
            Caption         =   "Hyper Texet Markedup Language (HTML)..."
         End
         Begin VB.Menu mmuToolsExportAsUbb 
            Caption         =   "Ultimate Bulletin Bord Post (UBB post)..."
         End
      End
      Begin VB.Menu mnuToolsBuildCode 
         Caption         =   "Build code"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuToolsRunCode 
         Caption         =   "Run code"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpLuaOrg 
         Caption         =   "Lua.org"
         Shortcut        =   ^H
      End
      Begin VB.Menu mmuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpShowTipOfTheDay 
         Caption         =   "Show 'Tip of the day'"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'some needede functions
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

'set the status of the app
Public Sub setStatus(status As String)
  sbStatusBar.Panels.Item(1).text = status
End Sub

'set the line row and col
'a row value below(or equal to) zero clears the line and col
Public Sub setLineRow(row As Integer, col As Integer)
  If row >= 0 Then
    sbStatusBar.Panels.Item(2).text = "Ln " & CStr(row + 1) & ", Col " & CStr(col)
    Debug.Print "Set line row"
  Else
    sbStatusBar.Panels.Item(2).text = ""
    Debug.Print "      clear line row"
  End If
End Sub

'load the app settings
'and create the status bar
Private Sub MDIForm_Load()
    Dim index As Integer
    setStatus "Loading app.."
    Debug.Print "----------------------"
    Debug.Print ""
    Debug.Print ""
    Debug.Print ""
    Debug.Print ""
    Randomize
    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
    loadOptions
    
    sbStatusBar.Panels.Clear
    
    index = 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrText)
    sbStatusBar.Panels.Item(index).AutoSize = sbrSpring
    index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrText)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).MinWidth = 2000
    index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrCaps)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).MinWidth = 700
    sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrNum)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).MinWidth = 700
    sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrIns)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).MinWidth = 700
    sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    index = index + 1
    
    'Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrScrl)
    'sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    'sbStatusBar.Panels.Item(index).MinWidth = 700
    'sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    'index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrTime)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    index = index + 1
    
    Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrDate)
    sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    index = index + 1
    
    'Call sbStatusBar.Panels.Add(index, "sbStatusBar.Panels.Index" & CStr(index), "", sbrKana)
    'sbStatusBar.Panels.Item(index).AutoSize = sbrContents
    'sbStatusBar.Panels.Item(index).MinWidth = 700
    'sbStatusBar.Panels.Item(index).Alignment = sbrCenter
    'index = index + 1
    
    frmBrowser.StartingAddress = "http://www.lua.org/"
    
    setReadyStatus
End Sub

'create a new document
Private Sub LoadNewDoc(Optional increment As Boolean = False)
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    If increment Then
      lDocumentCount = lDocumentCount + 1
    End If
    Set frmD = New frmDocument
    frmD.setTtitle "Document " & lDocumentCount
    frmD.Show
End Sub

'load files that is dropped
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim str As String
  Dim i As Long
  'Debug.Print "File count is: " + CStr(Data.Files.Count)
  setStatus "Loading dropped files.."
  For i = 1 To Data.Files.Count
    str = Data.Files.Item(i)
    'MsgBox "Check if file is a directory, it might crash if it is"
    'This is handled in the loadTextFromFile sub
    'it will close the file and tell the user that it was an invalid file
    LoadNewDoc
    ActiveForm.loadTextFromFile str
    'Debug.Print "  " + str
  Next
  setReadyStatus
End Sub

'save the settings and kill the browser
Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
        saveOptions
        
        'If frmBrowser.Visible Then frmBrowser.Visible = False
        'let's quit the application when we say so
        'kill it all
        Unload frmBrowser
    End If
End Sub

Private Sub mmuRedo_Click()
  If Me.ActiveForm Is Nothing Then Exit Sub
  Me.ActiveForm.editRedo
End Sub

'export as ubb
Private Sub mmuToolsExportAsUbb_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.exportAsUbb
End Sub

'show the browser
Private Sub mnuHelpLuaOrg_Click()
  frmBrowser.Visible = True
End Sub

'export as html
Private Sub mnuToolsExportHtml_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.exportAsHtml
End Sub

'enable the toolbar icons
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Run Code"
            'ToDo: Add 'Run Code' button code.
            mnuToolsRunCode_Click
        Case "Build Code"
            'ToDo: Add 'Build Code' button code.
            mnuToolsBuildCode_Click
        Case "Options"
            mnuViewOptions_Click
    End Select
End Sub

'a sub to tell the user that the app is ready
Public Sub setReadyStatus()
  Dim messageType As Integer
  Dim statusMessage As String
  'A simple message of ready is kinda boring, this is much more fun :)
  messageType = Int((Rnd * 7) + 1)
  Select Case messageType
  Case 1
    statusMessage = "Ready to rock"
  Case 2
    statusMessage = "Ready for action"
  Case 3
    statusMessage = "Ready!"
  Case 4
    statusMessage = "Let's go"
  Case 5
    statusMessage = "Lights! Camera! Action!"
  Case 6
    statusMessage = "Go go go!"
  Case 7
    statusMessage = "Jalla jalla"
  'Else
  ' statusMessage = "Jalla jalla"
  End Select
  
  setStatus statusMessage
End Sub

'let's show the help
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'show the tip of the day
Private Sub mnuHelpShowTipOfTheDay_Click()
     'we really want to see the tips
    frmTip.forceTipToShow = True
    frmTip.Show
End Sub

'search in help file
Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

'show the help file
Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

'arrange icons
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

'tile vertical
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

'tile horizontal
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

'cascade windows
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

'tell the active form to run the code
Private Sub mnuToolsRunCode_Click()
    If ActiveForm Is Nothing Then Exit Sub
    ActiveForm.runCode
End Sub

'tell the active frame to build the code
Private Sub mnuToolsBuildCode_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.buildCode
End Sub

'export to rtf
Private Sub mnuToolsExportRtf_Click()
    If ActiveForm Is Nothing Then Exit Sub
    ActiveForm.exportAsRtf
End Sub

'view the options dialog
Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

'hide/show the status bar
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

'hide/show the toolbar
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

'paste it
Private Sub mnuEditPaste_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.editPaste
End Sub

'copy it
Private Sub mnuEditCopy_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.editCopy
End Sub

'cut it
Private Sub mnuEditCut_Click()
  If ActiveForm Is Nothing Then Exit Sub
  ActiveForm.editCut
End Sub

'undo
Private Sub mnuEditUndo_Click()
    If ActiveForm Is Nothing Then Exit Sub
    ActiveForm.editUndo
End Sub

'exit this app
Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub

'launch the print sub
Private Sub mnuFilePrint_Click()
    If ActiveForm Is Nothing Then Exit Sub
    setStatus "Printing.."
    ActiveForm.printDoc
    setReadyStatus
End Sub

'save the file as something
Private Sub mnuFileSaveAs_Click()
    If ActiveForm Is Nothing Then Exit Sub
    setStatus "Saving as.."
    ActiveForm.saveAsDoc
    setStatus "Ready"
End Sub

'save the file
Private Sub mnuFileSave_Click()
  If ActiveForm Is Nothing Then Exit Sub
  setStatus "Saving.."
  ActiveForm.saveDoc
  setReadyStatus
End Sub

'close the file
Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    setStatus "Closing.."
    Unload ActiveForm
    setReadyStatus
End Sub

'open a file
Private Sub mnuFileOpen_Click()
    setStatus "Loading document.."
    'If ActiveForm Is Nothing Then
    LoadNewDoc
    ActiveForm.loadDoc
    setReadyStatus
End Sub

'create a new file
Private Sub mnuFileNew_Click()
    setStatus "Creating new document.."
    LoadNewDoc True
    ActiveForm.newDoc
    setReadyStatus
End Sub
