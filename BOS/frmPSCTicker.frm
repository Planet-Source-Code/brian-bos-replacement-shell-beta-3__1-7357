VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmPSCTicker 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   2415
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Loading Ticker..."
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   1140
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2400
      ExtentX         =   4233
      ExtentY         =   6879
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
End
Attribute VB_Name = "frmPSCTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetWindowPos Me.hWnd, -1, 66, Screen.Height / Screen.TwipsPerPixelY - 286, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION
WebBrowser1.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp"
End Sub

Private Sub Timer1_Timer()
WebBrowser1.Refresh
WebBrowser1.Document.body.Style.Border = "none"
WebBrowser1.Document.body.Style.backgroundcolor = "buttonface"
WebBrowser1.Document.body.Style.overflow = "hidden"
End Sub

Private Sub WebBrowser1_DownloadComplete()
Picture1.Visible = False
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WebBrowser1.Document.body.Style.Border = "none"
WebBrowser1.Document.body.Style.backgroundcolor = "buttonface"
WebBrowser1.Document.body.Style.overflow = "hidden"
End Sub

