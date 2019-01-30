VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   ClientHeight    =   5130
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6540
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   4440
      Left            =   45
      TabIndex        =   0
      Top             =   615
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   7832
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6000
      Top             =   120
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":02E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
'Dim mbDontNavigateNow As Boolean

'code from: http://www.dx21.com/VISSTUDIO/VB6/ViewItem.ASP?OID=38

'Information on this form acquired from Microsoft KB:
'http://support.microsoft.com/support/kb/articles/Q184/2/97.ASP

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = True
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False
    End If
End Function

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize
    
    Call SetTopMostWindow(Me.hwnd, True)

    'cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        'cboAddress.text = StartingAddress
        'cboAddress.AddItem cboAddress.text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
End Sub

Private Sub NavigateTo(url As String)
  On Error Resume Next
  If Len(url) > 0 Then
    'try to navigate to the starting address
    timTimer.Enabled = True
    brwWebBrowser.Navigate url
  End If
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal url As String)
    'Dim i As Integer
    'Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    'For i = 0 To cboAddress.ListCount - 1
    '    If cboAddress.List(i) = brwWebBrowser.LocationURL Then
    '        bFound = True
    '        Exit For
    '    End If
    'Next i
    'mbDontNavigateNow = True
    'If bFound Then
    '    cboAddress.RemoveItem i
    'End If
    'cboAddress.AddItem brwWebBrowser.LocationURL, 0
    'cboAddress.ListIndex = 0
    'mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    'If mbDontNavigateNow Then Exit Sub
    'timTimer.Enabled = True
    'brwWebBrowser.Navigate cboAddress.text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    'cboAddress.Width = Me.ScaleWidth - 100
    On Error Resume Next
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - 700 '(picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            'brwWebBrowser.LocationURL = "http://www.lua.org/"
            NavigateTo "http://www.lua.org/"
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub

