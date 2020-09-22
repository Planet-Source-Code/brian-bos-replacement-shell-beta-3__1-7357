VERSION 5.00
Begin VB.Form frmShutdown 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShutdown.frx":0000
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   1620
      ScaleHeight     =   255
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   -300
      Visible         =   0   'False
      Width           =   0
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
Blend Me, picDesktopCapture, 80, 0, 0, Me.ScaleWidth, Me.ScaleHeight

Me.Show
Me.Refresh
End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
picCancel.Picture = picCancelDown.Picture
AlphaBlending picCancel.hdc, 0, 0, 100, 25, picDesktopCapture.hdc, picCancel.Left, picCancel.Top, 100, 25, 80
End Sub


Private Sub picCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
picCancel.Picture = picCancelUp.Picture
AlphaBlending picCancel.hdc, 0, 0, 100, 25, picDesktopCapture.hdc, picCancel.Left, picCancel.Top, 100, 25, 80
End Sub
