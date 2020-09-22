VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettingsSubMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imagesdown 
      Left            =   2580
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":2C78
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imageshover 
      Left            =   1500
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":58F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":8568
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList images 
      Left            =   840
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":B1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettingsSubMenu.frx":DE58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "frmSettingsSubMenu.frx":10AD0
      ScaleHeight     =   375
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   0
      Width           =   2250
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "frmSettingsSubMenu.frx":13736
      ScaleHeight     =   375
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   375
      Width           =   2250
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   0
   End
End
Attribute VB_Name = "frmSettingsSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer
Dim Over(0 To 1) As Boolean

Private Sub Form_Load()
For i = 0 To UBound(Over)
    Over(i) = False
Next

DoEvents
SetWindowPos Me.hWnd, -1, 190, Screen.Height / Screen.TwipsPerPixelY - 200, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION

picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
Blend Me, picDesktopCapture, 80, 0, 0, Me.ScaleWidth, Me.ScaleHeight
For i = 0 To picButton.Count - 1
    AlphaBlending picButton(i).hdc, 0, 0, 150, 25, picDesktopCapture.hdc, picButton(i).Left, picButton(i).Top, 150, 25, 50
Next

Me.Show
Me.Refresh
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picButton(Index).Picture = imagesdown.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 150, 25, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 150, 25, 50
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Over(Index) = False Then
    picButton(OldIndex).Picture = images.ListImages(OldIndex + 1).Picture
    AlphaBlending picButton(OldIndex).hdc, 0, 0, 150, 25, picDesktopCapture.hdc, picButton(OldIndex).Left, picButton(OldIndex).Top, 150, 25, 50
    Over(OldIndex) = False
    OldIndex = Index
    
    picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 150, 25, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 150, 25, 50
    Over(Index) = True
    
    s_Playsound "hover"
End If
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 150, 25, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 150, 25, 50
    s_Playsound ("select")
    Select Case Index
    Case 0
        frmBSettings.Show
        HideStartMenu
    Case 1
        Shell " rundll32.exe shell32.dll,Control_RunDLL ", vbNormalFocus
        HideStartMenu
    End Select
End Sub

