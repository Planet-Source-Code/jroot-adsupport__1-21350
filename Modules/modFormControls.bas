Attribute VB_Name = "modForms"
Option Explicit

Const WM_NCLBUTTONDOWN = &HA1
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'----------------------------------------------------------
'>> START FORM DRAG CODE (SYNTAX: MouseDown: "FormDrag Me")
'----------------------------------------------------------

Public Sub FormDrag(TheForm As Form)
  On Error Resume Next
  ReleaseCapture
  Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

'----------------------------------------------------------
'<< END FORM DRAG CODE
'----------------------------------------------------------

