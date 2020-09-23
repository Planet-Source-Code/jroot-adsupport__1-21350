VERSION 5.00
Begin VB.Form frmAds 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AdSupport.ucAds ucAds 
      Height          =   900
      Left            =   200
      TabIndex        =   1
      Top             =   0
      Width           =   7015
      _ExtentX        =   12700
      _ExtentY        =   1693
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00404040&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   6675
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   915
      Width           =   510
   End
   Begin VB.Label lblDragBar 
      BackColor       =   &H00400000&
      Height          =   1150
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Right click for more options."
      Top             =   0
      Width           =   200
   End
   Begin VB.Label lblHome 
      BackColor       =   &H00404040&
      Caption         =   "Visit Our Homepage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   915
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7230
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Shape shapePanel 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   0
      Top             =   900
      Width           =   7250
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuHidePanel 
         Caption         =   "Hide Panel"
      End
      Begin VB.Menu mnuShowPanel 
         Caption         =   "Show Panel"
      End
      Begin VB.Menu mnuMainSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmAds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim success As Integer
Dim site As String

Dim varHomepageURL As String
Dim varPanelVisible As String

Private SetWindowPos As New clsOnTop

Private Sub Form_Load()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> START OF CONFIGURATION VARIABLES
  '----------------------------------------------------------

    'YOUR HOMEPAGE URL
    varHomepageURL = "http://www.thespacezone.com"

  '----------------------------------------------------------
  '<< END OF CONFIGURATION VARIABLES
  '----------------------------------------------------------

  '----------------------------------------------------------
  '>> SET DEFAULT VARIABLES (NO NEED TO EDIT)
  '----------------------------------------------------------

  varPanelVisible = "true"
  
  '----------------------------------------------------------
  '<< SET DEFAULT VARIABLES (NO NEED TO EDIT)
  '----------------------------------------------------------
  
  '----------------------------------------------------------
  '>> LOAD ANYTHING ELSE THAT NEEDS TO BE LOADED
  '----------------------------------------------------------

  SetWindowPos.MakeTopMost hWnd
  
  '----------------------------------------------------------
  '<< LOAD ANYTHING ELSE THAT NEEDS TO BE LOADED
  '----------------------------------------------------------
  
End Sub

Private Sub lblAbout_Click()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> CALL THE ABOUT FUNCTION
  '----------------------------------------------------------
  
  mnuAbout_click
  
  '----------------------------------------------------------
  '<< CALL THE ABOUT FUNCTION
  '----------------------------------------------------------
  
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
  lblAbout.ForeColor = &HC0FFFF
  
  '----------------------------------------------------------
  '<< UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
  lblAbout.ForeColor = &HE0E0E0
  
  '----------------------------------------------------------
  '<< UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
End Sub

Private Sub lblDragBar_DblClick()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> UPDATE THE PANEL VIEW BASED ON CURRENT STATUS
  '----------------------------------------------------------
  
  If varPanelVisible = "true" Then
    mnuHidePanel_click
  Else
    mnuShowPanel_click
  End If
  
  '----------------------------------------------------------
  '<< UPDATE THE PANEL VIEW BASED ON CURRENT STATUS
  '----------------------------------------------------------
  
End Sub

Private Sub lblDragBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> CHECK TO SEE WHAT BUTTON WAS PRESSED
  '----------------------------------------------------------
  
  If Button = 2 Then: PopupMenu mnuMain
  If Button = 1 Then: FormDrag Me
  
  '----------------------------------------------------------
  '<< CHECK TO SEE WHAT BUTTON WAS PRESSED
  '----------------------------------------------------------

  '----------------------------------------------------------
  '>> MAKE SURE THE FORM IS VISIBLE ON THE SCREEN
  '----------------------------------------------------------
  
  If frmAds.Top < 0 Then frmAds.Top = 0
  If frmAds.Left < 0 Then frmAds.Left = 0
  If (Screen.Width - Width) < frmAds.Left Then frmAds.Left = (Screen.Width - Width)
  If (Screen.Height - Height) < frmAds.Top Then frmAds.Top = (Screen.Height - Height)
  
  '----------------------------------------------------------
  '<< MAKE SURE THE FORM IS VISIBLE ON THE SCREEN
  '----------------------------------------------------------
  
End Sub

Private Sub lblHome_Click()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> NAVIGATE TO THE HOMEPAGE
  '----------------------------------------------------------
  
  site = varHomepageURL
  success = ShellToBrowser(Me, site, 0)
  
  '----------------------------------------------------------
  '<< NAVIGATE TO THE HOMEPAGE
  '----------------------------------------------------------
  
End Sub

Private Sub lblHome_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
  lblHome.ForeColor = &HC0FFFF
  
  '----------------------------------------------------------
  '<< UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
End Sub

Private Sub lblHome_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
  lblHome.ForeColor = &HE0E0E0
  
  '----------------------------------------------------------
  '<< UPDATE THE TEXT COLOR
  '----------------------------------------------------------
  
End Sub

Private Sub mnuShowPanel_click()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> RESIZE THE FORM TO SHOW THE PANEL
  '----------------------------------------------------------
  
  frmAds.Height = "1155"
  varPanelVisible = "true"
  
  '----------------------------------------------------------
  '<< RESIZE THE FORM TO SHOW THE PANEL
  '----------------------------------------------------------
  
End Sub

Private Sub mnuAbout_click()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> NAVIGATE TO THE ABOUT URL, PLEASE DO NOT CHANGE !!!!!!
  '----------------------------------------------------------
  
  site = "h" & "t" & "t" & "p" & ":" & "/" & "/" & "w" & "w" & _
         "w" & "." & "t" & "h" & "e" & "s" & "p" & "a" & "c" & _
         "e" & "z" & "o" & "n" & "e" & "." & "c" & "o" & "m" & _
         "/" & "p" & "r" & "o" & "j" & "e" & "c" & "t" & "s" & _
         "/" & "v" & "b" & "6" & "/" & "a" & "d" & "s" & "u" & _
         "p" & "p" & "o" & "r" & "t" & "/" & "a" & "b" & "o" & _
         "u" & "t" & "." & "p" & "h" & "p"

  success = ShellToBrowser(Me, site, 0)
  
  '----------------------------------------------------------
  '<< NAVIGATE TO THE ABOUT URL, PLEASE DO NOT CHANGE !!!!!!
  '----------------------------------------------------------
  
End Sub

Private Sub mnuHidePanel_click()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> RESIZE THE FORM TO HIDE THE PANEL
  '----------------------------------------------------------
  
  frmAds.Height = "930"
  varPanelVisible = "false"
  
  '----------------------------------------------------------
  '<< RESIZE THE FORM TO SHOW THE PANEL
  '----------------------------------------------------------

End Sub

