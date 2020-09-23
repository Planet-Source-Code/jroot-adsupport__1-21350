VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl ucAds 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ScaleHeight     =   900
   ScaleWidth      =   7020
   Begin VB.Timer tmrRefresh 
      Left            =   0
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser brwAds 
      Height          =   1200
      Left            =   -30
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Ads Support Powered By: http://www.thespacezone.com"
      Top             =   -30
      Width           =   8200
      ExtentX         =   14464
      ExtentY         =   2117
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "ucAds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim varAdsLocation As String
Dim varAdsRefreshRate As Integer

'---------------------------------------------------------------------------
'>> AD DISPLAY SUPPORT: READ ME
'---------------------------------------------------------------------------
'
'  AUTHOR:  TheSpaceDude
'   EMAIL:  webmaster@thespacezone.com
'     URL:  http://www.thespacezone.com
'
' INSTALL:  1. Edit the 2 configuration variables under the sub
'              UserControl_Initialize() with your ads page, and refresh
'              rate (in milliseconds)
'
'           2. Place the user control where ever you wish the ads to be
'              displayed in your program.
'
'           3. Make a page on your website with one ad on the page. Use
'              http://www.thespacezone.com/projects/vb6/adsupport/ads.php
'              as a template.
'
'           4. Your done! Simply run your program, and everything should
'              work fine. If you have any questions/comments, please visit
'              our website and click the forums link, or you may contact
'              me direct by sending am email to webmaster@thespacezone.com
'
'   NOTES:  This program is free to use, but please give me credit for it,
'           And if you can, a link to our website! It helps us out a lot!
'
'---------------------------------------------------------------------------
'<< AD DISPLAY SUPPORT: READ ME
'---------------------------------------------------------------------------

Private Sub brwAds_DownloadComplete()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> ENABLE THE TIMER
  '----------------------------------------------------------
  
  tmrRefresh.Enabled = True
  
  '----------------------------------------------------------
  '<< ENABLE THE TIMER
  '----------------------------------------------------------
  
End Sub

Private Sub tmrRefresh_Timer()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> REFRESH THE ADS DISPLAY
  '----------------------------------------------------------
  
    If brwAds.LocationURL = varAdsLocation Then
      brwAds.Refresh
      tmrRefresh.Enabled = False
    Else
      brwAds.Navigate varAdsLocation
      tmrRefresh.Enabled = False
    End If
    
  '----------------------------------------------------------
  '<< REFRESH THE ADS DISPLAY
  '----------------------------------------------------------
  
End Sub

Private Sub UserControl_Initialize()
  On Error Resume Next
  
  '----------------------------------------------------------
  '>> START OF CONFIGURATION VARIABLES
  '----------------------------------------------------------

    'LOCATION OF THE ADS PAGE
    varAdsLocation = "http://www.thespacezone.com/projects/vb6/adsupport/ads.php"
  
    'ADS REFRESH RATE (MILLISECONDS)
    varAdsRefreshRate = "15000"
  
  '----------------------------------------------------------
  '<< END OF CONFIGURATION VARIABLES
  '----------------------------------------------------------

  '----------------------------------------------------------
  '>> SET AND ACTIVATE THE FORM 'STUFF'
  '----------------------------------------------------------
  
    tmrRefresh.Interval = varAdsRefreshRate
    brwAds.Navigate varAdsLocation
    tmrRefresh.Enabled = False

  '----------------------------------------------------------
  '<< SET AND ACTIVATE THE FORM 'STUFF'
  '----------------------------------------------------------

End Sub
