Attribute VB_Name = "modOptions"
Option Explicit

'the options
'much is self-explenatory

Type TTimeSaving
  useTimeSaving As Boolean
  interval As Integer
End Type

Type TEnviroment
  hasHost As Boolean
  pathToHost As String
  hasCompiler As Boolean
  pathToCompiler As String
End Type

'the color needs to be public
'it is used when coloring richtext boxes
Public Type TColor
'this is all the color options that I've found that is supported by the richtext
  size As Integer
  color As Variant
  underline As Boolean
  bold As Boolean
  italic As Boolean
  StrikeThru As Boolean
End Type

'one color many colors
Type TColors
  keyword As TColor
  number As TColor
  operator As TColor
  text As TColor
  comment As TColor
  normal As TColor
End Type

Type TOptions
  timesaving As TTimeSaving
  enviroment As TEnviroment
  colors As TColors
End Type

'out current options
Public gOptions As TOptions

'we use the SaveSetting and GetSetting function.
'to bad the settings is saved in the registry and not in a file
'helps keeping the settings when moving to another computer

Private Function loadTimeSaving() As TTimeSaving
  Dim ts As TTimeSaving
  ts.interval = GetSetting(App.title, "Settings", "ts_interval", 2)
  ts.useTimeSaving = GetSetting(App.title, "Settings", "ts_use", False)
  loadTimeSaving = ts
End Function

Private Function loadEnviroment() As TEnviroment
  Dim env As TEnviroment
  env.hasCompiler = GetSetting(App.title, "Settings", "env_hasCompiler", False)
  env.hasHost = GetSetting(App.title, "Settings", "env_hasHost", False)
  env.pathToCompiler = GetSetting(App.title, "Settings", "env_compPath", "")
  env.pathToHost = GetSetting(App.title, "Settings", "env_hostPath", "")
  loadEnviroment = env
End Function

Private Function loadColor(name As String) As TColor
  Dim c As TColor
  c.bold = GetSetting(App.title, "Settings", "color_" & name & "_bold", False)
  c.color = GetSetting(App.title, "Settings", "color_" & name & "_color", vbBlack)
  c.italic = GetSetting(App.title, "Settings", "color_" & name & "_italic", False)
  c.size = GetSetting(App.title, "Settings", "color_" & name & "_size", 8)
  c.StrikeThru = GetSetting(App.title, "Settings", "color_" & name & "_striketrough", False)
  c.underline = GetSetting(App.title, "Settings", "color_" & name & "_underline", False)
  loadColor = c
End Function

Private Function loadColors() As TColors
  Dim ret As TColors
  ret.keyword = loadColor("keyword")
  ret.normal = loadColor("normal")
  ret.number = loadColor("number")
  ret.operator = loadColor("operator")
  ret.comment = loadColor("comment")
  ret.text = loadColor("text")
  loadColors = ret
End Function

Public Sub loadOptions()
  gOptions.colors = loadColors
  gOptions.enviroment = loadEnviroment
  gOptions.timesaving = loadTimeSaving
End Sub

Private Sub saveTimesaving(ts As TTimeSaving)
  SaveSetting App.title, "Settings", "ts_interval", ts.interval
  SaveSetting App.title, "Settings", "ts_use", ts.useTimeSaving
End Sub

Private Sub saveEnviroment(env As TEnviroment)
  SaveSetting App.title, "Settings", "env_hasCompiler", env.hasCompiler
  SaveSetting App.title, "Settings", "env_hasHost", env.hasHost
  SaveSetting App.title, "Settings", "env_compPath", env.pathToCompiler
  SaveSetting App.title, "Settings", "env_hostPath", env.pathToHost
End Sub

Private Sub saveColor(c As TColor, name As String)
  SaveSetting App.title, "Settings", "color_" & name & "_bold", c.bold
  SaveSetting App.title, "Settings", "color_" & name & "_color", c.color
  SaveSetting App.title, "Settings", "color_" & name & "_italic", c.italic
  SaveSetting App.title, "Settings", "color_" & name & "_size", c.size
  SaveSetting App.title, "Settings", "color_" & name & "_striketrough", c.StrikeThru
  SaveSetting App.title, "Settings", "color_" & name & "_underline", c.underline
End Sub

Private Sub saveColors(c As TColors)
  saveColor c.keyword, "keyword"
  saveColor c.normal, "normal"
  saveColor c.number, "number"
  saveColor c.operator, "operator"
  saveColor c.comment, "comment"
  saveColor c.text, "text"
End Sub

Public Sub saveOptions()
  saveColors gOptions.colors
  saveEnviroment gOptions.enviroment
  saveTimesaving gOptions.timesaving
End Sub
