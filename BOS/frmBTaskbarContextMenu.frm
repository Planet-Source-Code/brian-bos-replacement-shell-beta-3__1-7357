VERSION 5.00
Begin VB.Form frmTaskbarContextMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBTaskbarContextMenu.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrHide 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1980
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmTaskbarContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SetPosition(X As Single, Y As Single)
DoEvents
SetWindowPos Me.hWnd, -1, X, Y + Screen.Height / Screen.TwipsPerPixelY - Me.ScaleHeight - FrmTaskbar.ScaleHeight, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION
Debug.Print X & "," & Y
picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Top / Screen.TwipsPerPixelY, Me.Left / Screen.TwipsPerPixelX, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
Blend Me, picDesktopCapture, 80, 0, 0, Me.ScaleWidth, Me.ScaleHeight

Me.Show
Me.Refresh
End Sub

Private Sub tmrHide_Timer()
a = GetForegroundWindow
If a <> Me.hWnd Then
    Unload Me
End If
End Sub
