VERSION 5.00
Begin VB.Form frmRun 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRun.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3900
      Top             =   960
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   10
      Left            =   3900
      Top             =   960
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   -1500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblWebAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "C:\"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BlinkOn As Boolean, HasFocus As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If HasFocus Then
    If KeyAscii = 13 Then
        If BlinkOn Then
            Shell "start " & Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1), vbHide
            Unload Me
        Else
            Shell "start " & lblWebAddress.Caption, vbHide
            Unload Me
        End If
    Else
        If BlinkOn Then
            If KeyAscii = 8 Then
                If Len(lblWebAddress.Caption) > 1 Then
                    lblWebAddress.Caption = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 2) & "|"
                End If
            Else
                a = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1)
                lblWebAddress.Caption = a & Chr(KeyAscii) & "|"
            End If
        Else
            If KeyAscii = 8 Then
                If Len(lblWebAddress.Caption) > 0 Then
                    lblWebAddress.Caption = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1)
                End If
            Else
                lblWebAddress.Caption = lblWebAddress.Caption & Chr(KeyAscii)
            End If
        End If
    End If
End If
End Sub

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


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HasFocus Then
    If BlinkOn Then
        lblWebAddress.Caption = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1)
    End If
    BlinkOn = False
    HasFocus = False
    tmrBlink.Enabled = False
    tmrCheckFocus.Enabled = False
End If
End Sub



Private Sub lblWebAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HasFocus = True
    tmrBlink.Enabled = True
    tmrCheckFocus.Enabled = True
End Sub

Private Sub tmrBlink_Timer()

If BlinkOn Then
    lblWebAddress.Caption = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1)
Else
    lblWebAddress.Caption = lblWebAddress.Caption & "|"
End If
BlinkOn = Not BlinkOn

End Sub

Private Sub tmrCheckFocus_Timer()
If GetForegroundWindow <> Me.hwnd Then Unload Me
End Sub

