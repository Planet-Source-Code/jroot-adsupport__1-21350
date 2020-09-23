Attribute VB_Name = "modInternet"
Option Explicit

Dim success As Integer
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'----------------------------------------------------------
'>> BEGIN WEB NAVIGATE CONTROLS
'----------------------------------------------------------
' USEAGE:
'   site = "URL"
'   success = ShellToBrowser(Me, site, 0)
'----------------------------------------------------------

Function ShellToBrowser%(Frm As Form, ByVal URL$, ByVal WindowStyle%)
On Error Resume Next
  Dim api%
    api% = ShellExecute(Frm.hwnd, "open", URL$, "", App.Path, WindowStyle%)
  'CHECK THE RETURN VALUE
  If api% < 31 Then
    'ERROR CODE: VIEW API FOR MORE DETAILS
    MsgBox App.Title & " had a problem running your web browser. You should check that your browser is correctly installed." & "(Error" & Format$(api%) & ")", 48, "Browser Unavailable"
    ShellToBrowser% = False
  ElseIf api% = 32 Then
    'NO FILE ASSOCIATION
    MsgBox App.Title & " could not find a file association for " & URL$ & " on your system. You should check that your browser is correctly installed and associated with this type of file.", 48, "Browser Unavailable"
    ShellToBrowser% = False
  Else
    ShellToBrowser% = True
  End If
End Function

'----------------------------------------------------------
'<< END WEB NAVIGATE CONTROLS
'----------------------------------------------------------
