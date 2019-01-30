Attribute VB_Name = "modGlobals"
Option Explicit

'for ShellexecuteAndWai
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&

'maximum distance between min and max should be 7
Public Const MIN_FONT_SIZE As Integer = 8
Public Const MAX_FONT_SIZE As Integer = 15
Public Const MAX_COLORS As Integer = 8

'for diabling wordwrap in the document
Public Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'converts a boolean to an integer
Public Function bool2int(b As Boolean) As Integer
  bool2int = 0
  If b Then bool2int = 1
End Function

'converts a integer to a boolean
Public Function int2bool(i As Integer) As Boolean
  int2bool = False
  If i = 1 Then int2bool = True
End Function

' enable/disable a textbox based on the value of the checkbox
Public Sub enableText(text As TextBox, enable As Boolean)
  If enable Then
    text.Enabled = True
    text.BackColor = &H80000005
  Else
    text.Enabled = False
    text.BackColor = &H8000000F
  End If
End Sub

'color the rich text accortding to a color
Public Sub setColorOfRichText(color As TColor, box As RichTextBox, start As Integer, length As Integer)
  Dim selStart As Integer
  Dim sellength As Integer
  ' Backup properties
  selStart = box.selStart
  sellength = box.sellength
  
  ' Select text
  box.selStart = start
  box.sellength = length
  
  ' Apply color
  box.SelBold = color.bold
  box.SelColor = color.color
  box.SelFontSize = color.size
  box.SelItalic = color.italic
  box.SelUnderline = color.underline
  box.SelStrikeThru = color.StrikeThru
  
  'Debug.Print "  coloring richtext: " & box.SelText
  
  ' Restore properties
  box.selStart = selStart
  box.sellength = sellength
End Sub

'get the color of the rich text
'color is returned trhough the color parameter and the function returns the read text
'assume the selection is backed up by the caller
Public Function getColorOfRichtext(color As TColor, box As RichTextBox, start As Integer, length As Integer) As String
  ' Select text
  box.selStart = start
  box.sellength = length
  
  ' get color
  color.bold = box.SelBold
  color.color = box.SelColor
  color.size = box.SelFontSize
  color.italic = box.SelItalic
  color.underline = box.SelUnderline
  color.StrikeThru = box.SelStrikeThru
  
  If box.sellength > 0 Then
    getColorOfRichtext = box.SelText
  Else
    getColorOfRichtext = ""
  End If
End Function

'gets the character at zero based position from a string
Public Function CharacterAt(str As String, pos As Integer) As String
  CharacterAt = Mid$(str, pos + 1, 1)
End Function

'returns true if the character is in the string, false if not
Public Function characterInString(char As String, str As String) As Boolean
  If InStr(1, str, char) = 0 Then
    characterInString = False
  Else
    characterInString = True
  End If
  If Len(char) <= 0 Then characterInString = False
End Function

'finds the first character, that is isn't in the pattern
'returns the position
Public Function findFirstNotOf(pattern As String, str As String, firstPos As Integer) As Integer
  Dim pos As Integer
  Dim length As Integer
  Dim doit As Boolean
  length = Len(str)
  
  doit = True
  For pos = firstPos To length
    If Not characterInString(CharacterAt(str, pos), pattern) Then
      findFirstNotOf = pos
      Exit Function
    End If
  Next pos
  
  findFirstNotOf = -1
End Function




'http://www.vb-helper.com/HowTo/howto_shell_wait.zip
'The ShellAndWait subroutine uses the Shell function to start the other program. It calls the OpenProcess API function to connect to the new process and then uses WaitForSingleObject to wait until the other process terminates. Note that neither the program nor the development environment can take action during this wait.
'After WaitForSingleObject returns, the ShellAndWait subroutine calls CloseHandle to close the process handle opened by OpenProcess and then exits at which point the program resumes normal execution.
Public Sub ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle)
  Dim process_id As Long
  Dim process_handle As Long

  ' Start the program.
  On Error GoTo ShellError
  process_id = Shell(program_name, window_style)
  On Error GoTo 0
  
  DoEvents

  ' Wait for the program to finish.
  ' Get the process handle.
  process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
  If process_handle <> 0 Then
  WaitForSingleObject process_handle, INFINITE
  CloseHandle process_handle
  End If

  Exit Sub

ShellError:
  MsgBox "Error starting task " & App.ProductName & vbCrLf & Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

'executes a console command
' if wait is tru then the call won't return until the cmd has finished
Public Sub executeCmd(cmd As String, Optional wait As Boolean = False)
  Dim batFileName As String
  On Error GoTo filfel

  'Shell closes automaticly and Ive found no way to print the output to an internal console
  ' or redirect the whole output to a file
  ' the bat file seems o be the best solution
  'it will be displayed until the user kills it
  batFileName = App.Path & "\consolecmd.bat"
  Open batFileName For Output As #1   ' Open file for input.
  Print #1, cmd
  Close #1   ' Close file.
  
  If wait Then
    ShellAndWait batFileName, vbNormalFocus
  Else
    Shell batFileName, vbNormalFocus
  End If
  Exit Sub
filfel:
  MsgBox "Failed to execute command"
End Sub

'removes the extention from a filename, but not the path
'this function is usefull for getting a copy of a filename but with anoter extention
Public Function stripFileName(fn As String) As String
  Dim lastPos As Integer
  lastPos = InStrRev(fn, ".")
  If lastPos = 0 Then
    stripFileName = fn
  Else
    stripFileName = Left$(fn, lastPos - 1)
  End If
End Function
