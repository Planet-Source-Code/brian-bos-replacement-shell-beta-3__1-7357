VERSION 5.00
Begin VB.Form frmFavoritesSubMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B65E00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFavoritesSubMenu.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   0
      Left            =   1260
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   0
   End
End
Attribute VB_Name = "frmFavoritesSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer
Dim Over(0 To 1) As Boolean

Private Sub Form_Load()
DoEvents
SetWindowPos Me.hwnd, -1, 190, Screen.Height / Screen.TwipsPerPixelY - 350, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION

picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
'AlphaBlending Me.hdc, 0, 0, 245, 295, picDesktopCapture.hdc, 0, 0, 245, 295, 100

For i = 1 To 5
    AlphaBlending Me.hdc, 245, i + 5, 5, 1, picDesktopCapture.hdc, 245, 1, 5, 1, 50 * (5 - i)
Next

For i = 1 To 5
    AlphaBlending Me.hdc, 245 + i, 5, 1, 290, picDesktopCapture.hdc, 245 + i, 5, 1, 290, 70 * i - 50
Next

Me.Show
Me.Refresh
End Sub

