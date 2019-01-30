VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer tmrTimer 
      Interval        =   60000
      Left            =   3600
      Top             =   1440
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2040
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmDocument.frx":0000
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for ShellexecuteAndWai
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&

Const SPACES As String = " "
Const NUMBERS As String = "0123456789.eE"
Const CHARACTERS As String = "abcdefghijklmnopqrstuvwxyz_ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const OPERATORS As String = "+-*/,.;:?=)(&%![]"

Private Type TChange
  used As Boolean
  
  selstart As Integer
  sellen As Integer
  row As String
End Type

Const MAX_CHANGE As Integer = 10
Private ch_start As Integer
Private ch_end As Integer
Private ch_current As Integer
Private ch_buffer(1 To MAX_CHANGE) As TChange

Private bChanged As Boolean 'true if the text has changed
Private strTitle As String 'the title if it has no filename
Private bHasFileName As Boolean 'true if the document has a file associated with is
Private strFileName As String 'the filename if the bHasFileName says so
Private inputEnable As Boolean 'do we accept input
Private timeUnitspassed As Integer 'times units that has passed since we did the last timesave
Private insideColorRow As Boolean 'true if it is inside the colorRow function

'a syntax description enumeration
Enum TSyntax
  TS_KEYWORD
  TS_NUMBER
  TS_OPERATOR
  TS_TEXT
  TS_COMMENT
  TS_NORMAL
End Enum

'save as
Public Sub saveAsDoc()
  On Error Resume Next
  Debug.Print "saving as doc"
  With dlgCommonDialog
    .DialogTitle = "Save as.."
    .filename = strTitle
    .CancelError = True
    If bHasFileName Then .filename = strFileName
    .Filter = "Lua source file (*.lua)|*.lua|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .DefaultExt = "lua"
    .ShowSave
    
    If Err <> MSComDlg.cdlCancel Then
      bHasFileName = True
      strFileName = .filename
      saveTextToFile strFileName
    End If
  End With
End Sub

'load
Public Sub loadDoc()
  inputEnable = True
  On Error Resume Next
  Debug.Print "load doc"
  With dlgCommonDialog
    .DialogTitle = "Open.."
    .CancelError = True
    .filename = ""
    .Filter = "Lua source file (*.lua)|*.lua|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .DefaultExt = "lua"
    .ShowOpen
    
    If Err <> MSComDlg.cdlCancel Then
      bHasFileName = True
      strFileName = .filename
      loadTextFromFile strFileName
    Else
      setLineRow
    End If
  End With
  inputEnable = True
  addChange
End Sub


'we got the focus
Private Sub Form_GotFocus()
  'setLineRow
End Sub

'set default values
Private Sub Form_Load()
    Dim i As Integer
    Form_Resize
    
    bChanged = False
    bHasFileName = False
    strFileName = ""
    
    'disable wordwrap
    disableWordWrap
    
    ch_start = 1
    ch_current = 1
    ch_end = 1
    For i = 1 To MAX_CHANGE
      clearChange i
    Next i
    
    insideColorRow = False
End Sub

'disable the wordwrap for the rtfText
Public Sub disableWordWrap()
  SendMessageLong Me.rtfText.hwnd, EM_SETTARGETDEVICE, 0, 1
End Sub

'set the title of the the document
Public Sub setTtitle(title As String)
  If Not bHasFileName Then
    strTitle = title
  End If
  setupCaption
End Sub

'setup the document caption
Private Sub setupCaption()
  Dim strCaption As String
  
  If bHasFileName Then
    strTitle = strFileName
  End If
  
  strCaption = strTitle
  
  If bChanged Then
    strCaption = strCaption & "*"
  End If
  
  Me.Caption = strCaption
End Sub

'we lost the focus, thell the main app that it should no use our line row anymore
Private Sub Form_LostFocus()
  'fMainForm.setLineRow -1, 0
End Sub

'ask the user if he/she wants to save the doc before closing
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim response As Long
    Dim filename As String
    
    If bHasFileName Then
      filename = strFileName
    Else
      filename = strTitle
    End If
    
    Cancel = False
    If bChanged Then
        response = MsgBox("Do you want to save the changes you made to " & filename & "?", vbYesNoCancel, filename)
        Select Case response
        Case vbYes:
            saveDoc
            If bChanged Then Cancel = True
        Case vbNo
          Cancel = False
        Case vbCancel:
            Cancel = True
        End Select
    End If
End Sub

'resize
Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
    disableWordWrap
End Sub

'highlight some text
Private Sub syntaxHighlight(start As Integer, length As Integer, syntax As TSyntax)
  Select Case syntax
    Case TS_COMMENT
      setColorOfRichText gOptions.colors.comment, rtfText, start, length
    Case TS_KEYWORD
      setColorOfRichText gOptions.colors.keyword, rtfText, start, length
    Case TS_NORMAL
      setColorOfRichText gOptions.colors.normal, rtfText, start, length
    Case TS_NUMBER
      setColorOfRichText gOptions.colors.number, rtfText, start, length
    Case TS_OPERATOR
      setColorOfRichText gOptions.colors.operator, rtfText, start, length
    Case TS_TEXT
      setColorOfRichText gOptions.colors.text, rtfText, start, length
  End Select
End Sub

'returns true if the parameter is a lua keyword, false if not
Private Function isKeyword(word As String) As Boolean
  isKeyword = False
  If word = "and" Then isKeyword = True
  If word = "break" Then isKeyword = True
  If word = "do" Then isKeyword = True
  If word = "else" Then isKeyword = True
  If word = "elseif" Then isKeyword = True
  If word = "end" Then isKeyword = True
  If word = "false" Then isKeyword = True
  If word = "for" Then isKeyword = True
  If word = "function" Then isKeyword = True
  If word = "if" Then isKeyword = True
  If word = "in" Then isKeyword = True
  If word = "local" Then isKeyword = True
  If word = "nil" Then isKeyword = True
  If word = "not" Then isKeyword = True
  If word = "or" Then isKeyword = True
  If word = "repeat" Then isKeyword = True
  If word = "return" Then isKeyword = True
  If word = "then" Then isKeyword = True
  If word = "true" Then isKeyword = True
  If word = "until" Then isKeyword = True
  If word = "while" Then isKeyword = True
End Function

'color the row
Private Sub colorRow(start As Integer, length As Integer)
  Dim selstart As Integer
  Dim sellen As Integer
  Dim txt As String
  Dim currentPosition As Integer
  Dim startPosition As Integer
  Dim continue As Boolean 'continue emulator
  Dim doit As Boolean 'while this is true wee need to color more
  Dim searchPos As Integer
  Dim syntaxLength As Integer
  Dim currentChar As String
  If length <= 0 Then Exit Sub
  
  If insideColorRow Then Exit Sub
  insideColorRow = True
  
  Debug.Print "colorRow"
  'currentPosition-start
  
  'backup selection
  selstart = rtfText.selstart
  sellen = rtfText.sellength
  
  'get text
  rtfText.selstart = start
  rtfText.sellength = length
  txt = rtfText.SelText
  
  'let's tell the debuger what we are dooing
  Debug.Print "Coloring row: " & txt
  
  'syntax = TS_NORMAL
  currentPosition = start
  startPosition = start
  doit = True
  While doit
    currentChar = CharacterAt(txt, currentPosition - start)
    
    ' a string
    '34 = "
    '92 = \
    If currentChar = Chr(34) Then
      Dim ignoreNextChar As Boolean
      Dim loopSearch As Boolean
      ignoreNextChar = False
      
      startPosition = currentPosition
      currentPosition = currentPosition + 1
      loopSearch = True
      While loopSearch
        currentChar = CharacterAt(txt, currentPosition - start)
        
        If Not ignoreNextChar Then
          If currentChar = Chr(34) Then
            loopSearch = False
            currentPosition = currentPosition + 1
          ElseIf currentChar = "\" Then
            ignoreNextChar = True
          End If
        Else
          ignoreNextChar = False
        End If
        
        If loopSearch Then
          currentPosition = currentPosition + 1
          If currentPosition >= start + length Then
            doit = False
            loopSearch = False
          End If
        End If
      Wend
      
      syntaxHighlight startPosition, currentPosition - startPosition, TS_TEXT
      continue = True
    End If
    
    'color the spaces
    If characterInString(currentChar, SPACES) And Not continue Then
      startPosition = currentPosition
      searchPos = findFirstNotOf(SPACES, txt, startPosition - start)
      If searchPos = -1 Then
        'didn't find - the rest of the string are all of the same type
        syntaxLength = start + length - currentPosition
        doit = False
      Else
        syntaxLength = searchPos - currentPosition
        currentPosition = searchPos + start
      End If
      syntaxHighlight startPosition, currentPosition - startPosition, TS_NORMAL
      continue = True
    End If
    
    'color a identitifier
    If characterInString(currentChar, CHARACTERS) And Not continue Then
      startPosition = currentPosition
      searchPos = findFirstNotOf(CHARACTERS, txt, startPosition - start)
      If searchPos = -1 Then
        'didn't find - the rest of the string are all of the same type
        syntaxLength = length - (start - currentPosition)
        doit = False
      Else
        syntaxLength = searchPos - currentPosition
        currentPosition = searchPos + start
      End If
      
      rtfText.selstart = startPosition
      rtfText.sellength = currentPosition - startPosition
      
      If isKeyword(Trim(rtfText.SelText)) Then
        syntaxHighlight startPosition, currentPosition - startPosition, TS_KEYWORD
      Else
        syntaxHighlight startPosition, currentPosition - startPosition, TS_NORMAL
      End If
      
      continue = True
    End If
    
    'a number
    ' a number can't begin with a "."
    If Not currentChar = "." Then
      If characterInString(currentChar, NUMBERS) And Not continue Then
        startPosition = currentPosition
        searchPos = findFirstNotOf(NUMBERS, txt, startPosition - start)
        If searchPos = -1 Then
          'didn't find - the rest of the string are all of the same type
          syntaxLength = length - (start - currentPosition)
          doit = False
        Else
          syntaxLength = searchPos - currentPosition
          currentPosition = searchPos + start
        End If
        syntaxHighlight startPosition, currentPosition - startPosition, TS_NUMBER
        continue = True
      End If
    End If
    
    'comment
    If Mid$(txt, currentPosition + 1 - start, 2) = "--" And Not continue Then
      syntaxHighlight startPosition, start + length - startPosition, TS_COMMENT
      doit = False
      continue = True
    End If
    
    'a leading "." means a operator
    If characterInString(currentChar, OPERATORS) And Not continue Then
      startPosition = currentPosition
      searchPos = findFirstNotOf(OPERATORS, txt, startPosition - start)
      If searchPos = -1 Then
        'didn't find - the rest of the string are all of the same type
        syntaxLength = length - (start - currentPosition)
        doit = False
      Else
        syntaxLength = searchPos - currentPosition
        currentPosition = searchPos + start
      End If
      syntaxHighlight startPosition, currentPosition - startPosition, TS_OPERATOR
      continue = True
    End If
    
    'reset the continue
    If continue Then
      continue = False ' if we have continued then we don't need to move to the next character since that has the one who continued done
    Else
      currentPosition = currentPosition + 1
    End If
    
    'test for the end of the string
    If currentPosition >= start + length Then
      doit = False
    End If
  Wend
  
  'restore selection
  rtfText.selstart = selstart
  rtfText.sellength = sellen
  
  'we're leaving now
  insideColorRow = False
End Sub

'well, what is says - color the current row
Private Sub colorCurrentRow()
  Dim selstart As Integer
  Dim sellen As Integer
  Dim start As Integer
  Dim length As Integer
  Dim run As Boolean
  
  'backup
  selstart = rtfText.selstart
  sellen = rtfText.sellength
  
  'setup
  rtfText.sellength = 2
  If Not selstart = 0 Then
    rtfText.selstart = selstart - 1
    rtfText.sellength = 2
  End If
  
  'search start pos
  run = True
  With rtfText
    While run
      If .selstart = 0 Then
        run = False
      End If
      
      ' look at first sign
      If Asc(.SelText) = Asc(vbNewLine) Then
        .selstart = .selstart + 2 ' make sure we don't count the newline
        run = False
      End If
      
      If run Then
        .selstart = .selstart - 1
        .sellength = 2
      End If
    Wend
    start = .selstart
  End With
  
  'search end pos
  run = True
  With rtfText
    .selstart = selstart
    .sellength = 2
    While run
      If .selstart = Len(rtfText.text) Then
        run = False
      End If
      
      ' look at first sign
      If Asc(.SelText) = Asc(vbNewLine) Then
        '.selstart = .selstart - 1
        run = False
      End If
      
      If run Then
        .selstart = .selstart + 1
        .sellength = 2
      End If
    Wend
    length = .selstart - start
  End With
  
  'restore
  rtfText.selstart = selstart
  rtfText.sellength = sellen
  
  'color
  colorRow start, length
End Sub

' add a row
Private Sub addRow(row As String)
  Dim start As Integer
  Dim length As Integer
  
  Debug.Print "addRow"
  
  start = Len(rtfText.text)
  rtfText.selstart = start
  rtfText.SelText = row
  colorRow start, Len(row)
  
  'start = rtfText.selstart
  'rtfText.text = rtfText.text + row + vbNewLine
  'length = Len(row)
  'colorRow start, length
  
  Debug.Print CStr(start) & " : " & CStr(length) & " / " & CStr(Len(rtfText.text) - start)
  'rtfText.selstart = start + length
  'rtfText.text = rtfText.text + vbNewLine
End Sub


'this func saves a text to a file

'http://www.developer.com/net/vb/article.php/781481
'Before we begin, as a reminder you will need to add a
'reference to the Scripting Runtime library in VB6 from the
'Project|References menu. (Check the Microsoft Scripting Runtime
'reference and click OK in the Add References dialog.) The examples
'will use early bound objects rather than late bound objects using
'CreateObject, but you can use CreateObject and Variant types to
'dynamically load external libraries at runtime.

'http://c2.com/cgi/wiki?VisualBasicSurvivalGuide
'FileSystemObject? in the MS ActiveX Scripting DLL is a much saner way of dealing with files. In addition, the PropertyBag? Object can take any object and turn it into an array of bytes. You can save and reload this array and recreate the object.
'http://builder.com.com/5100-6373-1050078.html
'http://www.a1vbcode.com/vbforums/shwmessage.aspx?forumid=3&messageid=2143#bm2214:
' here's your function slightly modified using filesystemobject object
' note: make reference to Microsoft Scripting Runtime, add from vb Reference list
Private Sub saveTextToFile(filename As String)
  Dim fso As New FileSystemObject
  Dim file As TextStream
  bChanged = False
  Debug.Print "Saving to " & filename
  'On Error GoTo filfel
  
  'thw following code adds a newline to the end if the file
  'loading and saving of the file many thimes leads us to many nmewlines at the end of the file
  'this is bad and this solution(with the FileSystemObject) is the only one I've found that didn't add a newline to the file
  'Open filename For Output As #1   ' Open file for input.
  'Print #1, rtfText.text
  'Close #1   ' Close file.
  fso.CreateTextFile filename
  Set file = fso.OpenTextFile(filename, ForWriting, False)
  file.Write (rtfText.text)
  file.Close
  'Unload fso
  'I neither can do and unload or assign it nothing
  ' memoryleak?
  setupCaption
  Exit Sub
filfel:
  MsgBox "Failed to save file"
End Sub

'read a row from a file, returns a mempty string if the file is eof
Private Function readRow(file As Integer) As String
  readRow = ""
  Dim txt As String
  txt = ""
  Do While Not EOF(file)
    txt = Input(1, file)
    If Asc(txt) = 10 Then
      txt = ""
    ElseIf Asc(txt) = 13 Then
      readRow = readRow + vbNewLine
      Exit Function
    End If
    readRow = readRow + txt
  Loop
End Function

'load the text from a file, and add and colort the rows
Public Sub loadTextFromFile(filename As String)
  'Dim MyString As String
  Debug.Print "Loading from " & filename
  
  On Error GoTo filfel
  
  rtfText.Visible = False
  inputEnable = False
  rtfText.text = ""
  Open filename For Input As #1   ' Open file for input.
  Do While Not EOF(1)   ' Loop until end of file.
    'Input #1, MyString
    addRow readRow(1)
  Loop
  Close #1   ' Close file.
  inputEnable = True
  rtfText.Visible = True
  
  bChanged = False
  strTitle = filename
  strFileName = filename
  bHasFileName = True
  setupCaption
  disableWordWrap
  
  setLineRow
  Exit Sub
filfel:
  MsgBox "Failed to load file, closing window"
  Unload Me
End Sub



'save the document
Public Sub saveDoc()
  If bHasFileName Then
    saveTextToFile strFileName
  Else
    saveAsDoc
  End If
End Sub

'set the document to be as a new
Public Sub newDoc()
  Debug.Print "new doc"
  inputEnable = True
  rtfText.text = ""
  bChanged = False
  setupCaption
  addChange
End Sub

'print the doc
Public Sub printDoc()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .FLAGS = cdlPDReturnDC + cdlPDNoPageNums
        If rtfText.sellength = 0 Then
            .FLAGS = .FLAGS + cdlPDAllPages
        Else
            .FLAGS = .FLAGS + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            rtfText.SelPrint .hDC
        End If
    End With
End Sub

'cut the code
Public Sub editCut()
  Debug.Print "editCut"
  On Error GoTo errorhandler
  
  rtfText.Visible = False
  inputEnable = False
  
  Clipboard.SetText rtfText.SelText
  rtfText.SelText = vbNullString
  colorCurrentRow
  
  rtfText.Visible = True
  inputEnable = True
  Exit Sub
errorhandler:
  rtfText.Visible = True
  inputEnable = True
End Sub

'paste the code
Public Sub editPaste()
  Debug.Print "editPaste"
  On Error GoTo errorhandler
  
  rtfText.Visible = False
  inputEnable = False
  
  rtfText.SelText = Clipboard.GetText
  colorCurrentRow
  
  rtfText.Visible = True
  inputEnable = True
  Exit Sub
errorhandler:
  rtfText.Visible = True
  inputEnable = True
End Sub

'copy the code
Public Sub editCopy()
  Debug.Print "EditCpoy"
  On Error Resume Next
  Clipboard.SetText rtfText.SelText
End Sub

'unload the form
Private Sub Form_Unload(Cancel As Integer)
  fMainForm.setLineRow -1, 0
End Sub

'the text has changed
Private Sub rtfText_Change()
  If Not inputEnable Then Exit Sub
  Debug.Print "***start rtftextchange***"
  inputEnable = False
  'By disabling rtfText we looses the ugly redraw when selecting and deselcting
  'There are still some but they are minor :)
  'rtfText.Enabled = False
  rtfText.Visible = False
  'rtfText.Visible = False
  bChanged = True
  setupCaption
  colorCurrentRow
  'rtfText.Visible = True
  
  setLineRow
  addChange
  
  'rtfText.Enabled = True
  rtfText.Visible = True
  
  ' If something has been disabled, and then enabled it has lost it focus, let's give it back :)
  inputEnable = True
  
  Debug.Print "***end rtftextchange***"
End Sub

'I couldnt find any method of disabling the rich-text default double-click select behavior so I disabled this feature
'findFirstNotOfRev is a previous method, that was removed since it had no value
'Private Sub rtfText_DblClick()
 ' Dim syntax As TSyntax
  'Dim char As String
  
  'Dim selStart As Integer
  'Dim selEnd As Integer
  'Dim selLen As Integer
  
'  selEnd = 10
 ' selStart = 0
  
  'Debug.Print "Clickety Click"
  'syntax = TS_TEXT
  
  'char = CharacterAt(rtfText.text, rtfText.selStart)
  
  'If characterInString(char, OPERATORS) Then
   ' selEnd = findFirstNotOf(OPERATORS, rtfText.text, rtfText.selStart)
    'selStart = findFirstNotOfRev(OPERATORS, rtfText.text, rtfText.selStart)
  'ElseIf characterInString(char, CHARACTERS) Then
   ' selEnd = findFirstNotOf(CHARACTERS, rtfText.text, rtfText.selStart)
    'selStart = findFirstNotOfRev(CHARACTERS, rtfText.text, rtfText.selStart)
  'ElseIf characterInString(char, SPACES) Then
   ' selEnd = findFirstNotOf(SPACES, rtfText.text, rtfText.selStart)
    'selStart = findFirstNotOfRev(SPACES, rtfText.text, rtfText.selStart)
  'ElseIf characterInString(char, NUMBERS) Then
    'selEnd = findFirstNotOf(NUMBERS, rtfText.text, rtfText.selStart)
   ' selStart = findFirstNotOfRev(NUMBERS, rtfText.text, rtfText.selStart)
  'Else
 '   Exit Sub
'  End If
  
  'If selStart = -1 Then selStart = 0
  'If selEnd = -1 Then selEnd = Len(rtfText.text)
  
  'selLen = selEnd - selStart
  
  'If selLen <= 0 Then Exit Sub
  'rtfText.selStart = selStart
  'rtfText.sellength = selLen
'End Sub

'tell the mianform that we got the linerow back
Private Sub rtfText_GotFocus()
  If Not inputEnable Then Exit Sub
  setLineRow
  Debug.Print "         got focus"
End Sub

'get the current row
Private Function getCurrentRow() As String
  Dim selstart As Integer
  Dim sellen As Integer
  Dim start As Integer
  Dim length As Integer
  Dim run As Boolean
  
  'backup
  selstart = rtfText.selstart
  sellen = rtfText.sellength
  
  'setup
  rtfText.sellength = 2
  If Not selstart = 0 Then
    rtfText.selstart = selstart - 1
    rtfText.sellength = 2
  End If
  
  'search start pos
  run = True
  With rtfText
    While run
      If .selstart = 0 Then
        run = False
      End If
      
      ' look at first sign
      If run Then
        If Asc(.SelText) = Asc(vbNewLine) Then
          .selstart = .selstart + 2 ' make sure we don't count the newline
          run = False
        End If
      End If
      
      If run Then
        .selstart = .selstart - 1
        .sellength = 2
      End If
    Wend
    start = .selstart
  End With
  
  'search end pos
  run = True
  With rtfText
    .selstart = selstart
    .sellength = 2
    While run
      If .selstart = Len(rtfText.text) Then
        run = False
      End If
      
      If run Then
        ' look at first sign
        If Asc(.SelText) = Asc(vbNewLine) Then
          '.selstart = .selstart - 1
          run = False
        End If
      End If
      
      If run Then
        .selstart = .selstart + 1
        .sellength = 2
      End If
    Wend
    length = .selstart - start
  End With
  
  rtfText.selstart = start
  rtfText.sellength = length
  getCurrentRow = rtfText.SelText
  
  'restore
  rtfText.selstart = selstart
  rtfText.sellength = sellen
End Function

'set the current row
Private Sub setCurrentRow(str As String, displace As Integer)
  Dim selstart As Integer
  Dim sellen As Integer
  Dim start As Integer
  Dim length As Integer
  Dim run As Boolean
  
  Debug.Print "setCurrentRow"
  
  'backup
  selstart = rtfText.selstart
  sellen = rtfText.sellength
  
  'setup
  rtfText.sellength = 2
  If Not selstart = 0 Then
    rtfText.selstart = selstart - 1
    rtfText.sellength = 2
  End If
  
  'search start pos
  run = True
  With rtfText
    While run
      If .selstart = 0 Then
        run = False
      End If
      
      ' look at first sign
      If Asc(.SelText) = Asc(vbNewLine) Then
        .selstart = .selstart + 2 ' make sure we don't count the newline
        run = False
      End If
      
      If run Then
        .selstart = .selstart - 1
        .sellength = 2
      End If
    Wend
    start = .selstart
  End With
  
  'search end pos
  run = True
  With rtfText
    .selstart = selstart
    .sellength = 2
    While run
      If .selstart = Len(rtfText.text) Then
        run = False
      End If
      
      ' look at first sign
      If Asc(.SelText) = Asc(vbNewLine) Then
        '.selstart = .selstart - 1
        run = False
      End If
      
      If run Then
        .selstart = .selstart + 1
        .sellength = 2
      End If
    Wend
    length = .selstart - start
  End With
  
  rtfText.selstart = start
  rtfText.sellength = length
  rtfText.SelText = str
  
  'restore
  If displace >= 0 Then
    rtfText.selstart = start + displace
    rtfText.sellength = 0
  End If
End Sub

'handle insert of space when the user presses enter, and fast type
Private Sub rtfText_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim currentRow As String
   Dim splen As Integer
  ' kaycode = space
  If KeyCode = 32 Then
    Debug.Print "Space"
    If Shift = 1 Then
     
      Dim beginningSpace As String
      inputEnable = False
      rtfText.Visible = False
      
      currentRow = getCurrentRow()
      splen = findFirstNotOf(SPACES, currentRow, 0)
      currentRow = Trim(currentRow)
      'rtfText.selStart = rtfText.selStart + 2 ' put the cursor  on the new row
      If splen > 0 Then
        beginningSpace = Space$(splen)
      Else
        beginningSpace = ""
        splen = 0
      End If
      
      Debug.Print "You pressed enter"
      If currentRow = "if" Then
        setCurrentRow beginningSpace & "if  then else end", 3 + splen
        addChange
      ElseIf currentRow = "while" Then
        setCurrentRow beginningSpace & "while  do  end", 6 + splen
        addChange
      ElseIf currentRow = "do" Then
        setCurrentRow beginningSpace & "do  end", 3 + splen
        addChange
      ElseIf currentRow = "local" Then
        setCurrentRow beginningSpace & "local function () end", 15 + splen
        addChange
      ElseIf currentRow = "function" Then
        setCurrentRow beginningSpace & "function () end", 15 + splen
        addChange
      ElseIf currentRow = "repeat" Then
        setCurrentRow beginningSpace & "repeat  until", 7 + splen
        addChange
      ElseIf currentRow = "for" Then
        setCurrentRow beginningSpace & "for  do end", 4 + splen
        addChange
      ElseIf currentRow = "forin" Then
        setCurrentRow beginningSpace & "for  in do end", 4 + splen
        addChange
      End If
      colorCurrentRow
      KeyCode = 0
      
      rtfText.Visible = True
      inputEnable = True
    End If
  'user pressed enter
  ElseIf KeyCode = 13 Then
    Debug.Print "Enter"
    
    inputEnable = False
    rtfText.Visible = False
    
    Dim strCurrentRow As String
    'Dim splen As Integer
    KeyCode = 0 ' ignore enter key
    rtfText.SelText = vbNewLine 'add a newline after the cursor
    rtfText.sellength = 0
    rtfText.selstart = rtfText.selstart - 1 ' put the cursor on the previous row
    colorCurrentRow ' color it
    strCurrentRow = getCurrentRow()
    splen = findFirstNotOf(SPACES, strCurrentRow, 0)
    rtfText.selstart = rtfText.selstart + 2 ' put the cursor  on the new row
    If splen > 0 Then
      rtfText.SelText = Space$(splen)
      rtfText.sellength = 0
      'rtfText.selStart = rtfText.selStart + splen - 1
    End If
    colorCurrentRow 'color it
    
    addChange
    
    rtfText.Visible = True
    inputEnable = True
  End If
End Sub

'hide the line row
Private Sub rtfText_LostFocus()
  If Not inputEnable Then Exit Sub
  fMainForm.setLineRow -1, 0
  Debug.Print "lost focus"
End Sub

'change the line row
Private Sub rtfText_SelChange()
  If Not inputEnable Then Exit Sub
  setLineRow
End Sub

'on timer
Private Sub tmrTimer_Timer()
  If gOptions.timesaving.useTimeSaving Then
    timeUnitspassed = timeUnitspassed + 1
    If gOptions.timesaving.interval <= timeUnitspassed Then
      fMainForm.setStatus "Time saving.."
      timeUnitspassed = 0
      If bHasFileName Then
        saveTextToFile strFileName
      Else
        saveTextToFile App.Path & "\" & strTitle & ".tmp"
      End If
      fMainForm.setReadyStatus
    End If
  End If
End Sub

'get the col from a index
Private Function getColFromChar(start As Integer) As Integer
  Dim i As Integer
  Dim c As String
  i = start
  getColFromChar = 0
  While i > 0
    c = CharacterAt(rtfText.text, i)
    
    If Len(c) > 0 Then
      If Asc(c) = 10 Then
        getColFromChar = getColFromChar - 1
        Exit Function
      End If
    End If
    
    getColFromChar = getColFromChar + 1
    i = i - 1
  Wend
End Function

'setup the linerow
Private Sub setLineRow()
  If Not inputEnable Then Exit Sub
  fMainForm.setLineRow rtfText.GetLineFromChar(rtfText.selstart), getColFromChar(rtfText.selstart)
End Sub

'export as ubb
'simply save the file with a [code] in the beginning and a [/code] in the end
Public Sub exportUbbText(filename As String)
  Dim fso As New FileSystemObject
  Dim file As TextStream
  On Error GoTo filfel
  'bChanged = False
  Debug.Print "Saving to " & filename
  fso.CreateTextFile filename
  Set file = fso.OpenTextFile(filename, ForWriting, False)
  file.Write "[code]"
  file.Write (rtfText.text)
  file.Write "[/code]"
  file.Close
  'Unload fso
  'kan varken unloada eller tilldela fso nothing
  ' minnesläcka?
  Exit Sub
filfel:
  MsgBox "Failed to save file"
End Sub

'wrapper function
'get the current color
Private Function getColorAt(c As TColor, start As Integer) As String
  getColorAt = getColorOfRichtext(c, rtfText, start, 1)
End Function

'has a color changed from one to another
Private Function hasChanged(this As TColor, other As TColor) As Boolean
  If this.bold <> other.bold Then
    hasChanged = True
    Exit Function
  End If
  If this.color <> other.color Then
    hasChanged = True
    Exit Function
  End If
  If this.italic <> other.italic Then
    hasChanged = True
    Exit Function
  End If
  If this.size <> other.size Then
    hasChanged = True
    Exit Function
  End If
  If this.StrikeThru <> other.StrikeThru Then
    hasChanged = True
    Exit Function
  End If
  If this.underline <> other.underline Then
    hasChanged = True
    Exit Function
  End If
  
  hasChanged = False
End Function

'convert a bgr string expression to an rgb string expression
Private Function GetRgb(bgr As String) As String
' this could be made a lot faster but I am just another string man :)
  Dim r As String
  Dim g As String
  Dim b As String
  Dim final As String
  final = Replace(Space(6 - Len(bgr)), " ", "0") & bgr
  b = Mid(final, 1, 2)
  g = Mid(final, 3, 2)
  r = Mid(final, 5, 2)
  GetRgb = r & g & b
End Function

'output end and start tags
Private Sub handleHtmlColoring(curr As TColor, pre As TColor, stream As TextStream, startColor As Boolean, endColor As Boolean, force As Boolean)
  If Not force Then
    If Not hasChanged(curr, pre) Then
      Exit Sub
    End If
  End If

  If endColor And pre.bold Then
    stream.Write "</b>"
  End If
  
  If pre.underline And endColor Then
   stream.Write "</u>"
  End If
  
  If pre.StrikeThru And endColor Then
   stream.Write "</s>"
  End If
  
  If pre.italic And endColor Then
   stream.Write "</i>"
  End If
  
  If endColor Then
    stream.Write "</font>"
  End If
  
  'always set the fonstsize and the font color
  If startColor Then
                              'html size goes from 1 - 7, this will keep the converstion to a pizel in a reasonable size
    stream.Write "<font size=" & curr.size - MIN_FONT_SIZE & " color=""#" & GetRgb(Hex(curr.color)) & """>"
  End If
  
  If curr.italic And startColor Then
   stream.Write "<i>"
  End If
  
  If curr.StrikeThru And startColor Then
   stream.Write "<s>"
  End If
  
  If curr.underline And startColor Then
   stream.Write "<u>"
  End If
  
  If curr.bold And startColor Then
   stream.Write "<b>"
  End If
  
  'stream.Write "<font color=""" & this.color & """>"
  
  'If this.italic Then
    'stream.Write "<i>"
  'End If
  
  'stream.Write "<font size=" & other.size & ">"
  
  'this.StrikeThru And Not other.StrikeThru Then
   ' stream.Write "<s>"
  'End If
  'If this.underline And Not other.underline Then
   ' stream.Write "<u>"
  'End If
End Sub

'write html to a a text stream
'the onlyt html that is written here is the pure code, no title or copyright yadda yadda yada
Private Sub writeHtmlToTextStream(stream As TextStream)
  Dim selstart As Integer
  Dim sellen As Integer
  Dim preColor As TColor
  Dim cColor As TColor
  Dim character As String
  Dim i As Integer
  Dim l As Integer
  Dim nullColor As TColor
  
  nullColor.color = vbBlack
  nullColor.bold = False
  nullColor.italic = False
  nullColor.size = 0
  nullColor.StrikeThru = False
  nullColor.underline = False
  
  rtfText.Visible = False
  inputEnable = False
  
  selstart = rtfText.selstart
  sellen = rtfText.sellength
  i = 0
  l = Len(rtfText.text)
  Call getColorAt(cColor, i)
  handleHtmlColoring cColor, nullColor, stream, True, False, True
  preColor = cColor
    
  While i < l
    preColor = cColor
    character = getColorAt(cColor, i)
    handleHtmlColoring cColor, preColor, stream, True, True, False
    
    If Len(character) > 0 Then
    
    
      If Asc(character) = Asc("<") Or Asc(character) = Asc(">") Then
        If Asc(character) = Asc("<") Then
          stream.Write "&#60;"
        Else
          stream.Write "&#62;"
        End If
      Else
        stream.Write character
      End If
      i = i + 1
      
      
    Else 'A newline consists of two character
      stream.Write vbNewLine
      i = i + 2
    End If
    
  Wend
  handleHtmlColoring cColor, preColor, stream, False, True, True
  
  'stream.Write (rtfText.text)
  
  rtfText.selstart = selstart
  rtfText.sellength = sellen
  rtfText.Visible = True
  inputEnable = True
End Sub

'export the html, open/close the file etc
Public Sub exportHtmlText(filename As String)
  Dim fso As New FileSystemObject
  Dim file As TextStream
  Dim title As String
  
  title = strTitle
  If bHasFileName Then title = fso.GetFileName(strFileName)
  
  title = UCase(title)
  
  'On Error GoTo filfel
  
  Debug.Print "Saving to " & filename
  fso.CreateTextFile filename
  Set file = fso.OpenTextFile(filename, ForWriting, False)
  file.Write "<html><head><title>" & title & "</title></head><body>"
  file.Write "<h1><center>" & title & "</center></h1><br>"
  file.Write "<pre>"
  writeHtmlToTextStream file
  file.Write "</pre>"
  file.Write "<br><br><br><font size=""-2"" align=""right"">Created by Lua Edit 2004</font>"
  file.Write "</body></html>"
  
  file.Close
  'Unload fso
  'kan varken unloada eller tilldela fso nothing
  ' minnesläcka?
  Exit Sub
filfel:
  MsgBox "Failed to save file"
End Sub

'pick a file for htl export
Public Sub exportAsHtml()
  On Error Resume Next
  With dlgCommonDialog
    .DialogTitle = "Export html.."
    fMainForm.setStatus "Exporting to html.."
    .filename = strTitle
    .CancelError = True
    If bHasFileName Then .filename = stripFileName(strFileName)
    .Filter = "Hyper Text Markeduped Language (*.html)|*.html|Text file (*txt)|*.txt"
    .DefaultExt = "ubb"
    .ShowSave
    
    If Err <> MSComDlg.cdlCancel Then
      Debug.Print "Exporting as rtf: " & .filename
      exportHtmlText .filename
    End If
    fMainForm.setReadyStatus
  End With
End Sub

'export as a ubb
Public Sub exportAsUbb()
  On Error Resume Next
  With dlgCommonDialog
    .DialogTitle = "Export ubb.."
    fMainForm.setStatus "Exporting to ubb.."
    .filename = strTitle
    .CancelError = True
    If bHasFileName Then .filename = stripFileName(strFileName)
    .Filter = "Ultimate bulletin board (*.ubb)|*.ubb|Text file (*txt)|*.txt"
    .DefaultExt = "ubb"
    .ShowSave
    
    If Err <> MSComDlg.cdlCancel Then
      Debug.Print "Exporting as rtf: " & .filename
      exportUbbText .filename
    End If
    fMainForm.setReadyStatus
  End With
End Sub

'export as rtf, use the rtf control function
Public Sub exportAsRtf()
  On Error Resume Next
  With dlgCommonDialog
    .DialogTitle = "Export rtf.."
    fMainForm.setStatus "Exporting to rtf.."
    .filename = strTitle
    .CancelError = True
    If bHasFileName Then .filename = stripFileName(strFileName)
    .Filter = "Rich text file (*.rtf)|*.rtf"
    .DefaultExt = "rtf"
    .ShowSave
    
    If Err <> MSComDlg.cdlCancel Then
      Debug.Print "Exporting as rtf: " & .filename
      rtfText.SaveFile .filename
    End If
    fMainForm.setReadyStatus
  End With
End Sub

'run the code
Public Sub runCode()
  Dim cmd As String
  Dim dblRes As Double
  saveDoc
  If bChanged Or Not bHasFileName Then Exit Sub
  
  If Not gOptions.enviroment.hasHost Then
    MsgBox "You have't configured the enviroment, missing host Application"
    Exit Sub
  End If
  
  fMainForm.setStatus "       Running code.."
  cmd = gOptions.enviroment.pathToHost
  cmd = Replace(cmd, "[script]", strFileName) '& " >" & fMainForm.getConsoleFileName()
  executeCmd cmd
  fMainForm.setReadyStatus
End Sub

'build the code
Public Sub buildCode()
  saveDoc
  If bChanged Or Not bHasFileName Then Exit Sub
  
  If Not gOptions.enviroment.hasCompiler Then
    MsgBox "You have't configured the enviroment, missing compiler"
    Exit Sub
  End If
  
  frmBuild.ShowBuild strFileName
End Sub


'undo the changes
Public Sub editUndo()
  Debug.Print "editUndo"
  'If prevChangeIndex(prevChangeIndex(ch_current)) = nextChangeIndex(ch_start) Then Exit Sub
  If ch_current = ch_start Then Exit Sub
  'MsgBox "Fix undo call"
  
  inputEnable = False
  rtfText.Visible = False
  
  ch_current = prevChangeIndex(ch_current)
  rtfText.selstart = ch_buffer(ch_current).selstart
  rtfText.sellength = ch_buffer(ch_current).sellen
  setCurrentRow ch_buffer(ch_current).row, -1
  colorCurrentRow
  
  inputEnable = True
  rtfText.Visible = True
End Sub

'redo the changes
Public Sub editRedo()
  Debug.Print "editRedo"
  If ch_current = ch_end Then Exit Sub
  'MsgBox "Fix redo call"
  
  inputEnable = False
  rtfText.Visible = False
  
  ch_current = nextChangeIndex(ch_current)
  rtfText.selstart = ch_buffer(ch_current).selstart
  rtfText.sellength = ch_buffer(ch_current).sellen
  setCurrentRow ch_buffer(ch_current).row, -1
  colorCurrentRow
  
  inputEnable = True
  rtfText.Visible = True
End Sub

Private Sub addChange()
  'we don't need to clear the indeces since we have a end
  ch_end = ch_current
  ch_buffer(ch_current).used = True
  ch_buffer(ch_current).row = getCurrentRow
  ch_buffer(ch_current).sellen = rtfText.sellength
  ch_buffer(ch_current).selstart = rtfText.selstart
  ch_current = nextChangeIndex(ch_current)
  If ch_current = ch_start Then
    ch_start = nextChangeIndex(ch_start)
  End If
End Sub

Private Function nextChangeIndex(index As Integer) As Integer
  nextChangeIndex = index + 1
  If nextChangeIndex > MAX_CHANGE Then nextChangeIndex = 1
End Function

Private Function prevChangeIndex(index As Integer) As Integer
  prevChangeIndex = index - 1
  If prevChangeIndex <= 0 Then prevChangeIndex = MAX_CHANGE
End Function

Private Sub clearChange(index As Integer)
  ch_buffer(index).sellen = 1
  ch_buffer(index).selstart = 1
  ch_buffer(index).row = ""
  ch_buffer(index).used = False
End Sub
