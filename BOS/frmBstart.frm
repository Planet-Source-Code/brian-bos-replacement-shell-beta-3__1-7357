VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBstart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmBstart.frx":0000
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   375
      Picture         =   "frmBstart.frx":24A32
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   5
      Top             =   600
      Width           =   2550
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   375
      Picture         =   "frmBstart.frx":29A74
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   4
      Top             =   1200
      Width           =   2550
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   375
      Picture         =   "frmBstart.frx":2EAB6
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   3
      Top             =   1800
      Width           =   2550
   End
   Begin MSComctlLib.ImageList imagesdown 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":33AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":38B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":3DBA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":42BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":47C48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   375
      Picture         =   "frmBstart.frx":4CC9C
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   2
      Top             =   2400
      Width           =   2550
   End
   Begin MSComctlLib.ImageList imageshover 
      Left            =   540
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":51CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":56D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":5BD86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":60DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":65E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList images 
      Left            =   1860
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":6AE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":6FED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":74F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":79F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":7EFD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   375
      Picture         =   "frmBstart.frx":84026
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   1
      Top             =   3000
      Width           =   2550
   End
   Begin VB.Timer tmrHide 
      Interval        =   1
      Left            =   960
      Top             =   1200
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3060
      ScaleHeight     =   435
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmBstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Over(0 To 4) As Boolean
Dim OldIndex As Integer

Public Sub showme()
HideSubs
SetWindowPos Me.hWnd, -1, Me.ScaleLeft, Screen.Height / Screen.TwipsPerPixelY - Me.ScaleHeight - 30, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION
Me.Left = 0

picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
Blend Me, picDesktopCapture, 80, 0, 0, Me.ScaleWidth, Me.ScaleHeight
For i = 0 To picButton.Count - 1
    AlphaBlending picButton(i).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(i).Left, picButton(i).Top, 170, 40, 80
Next

Me.Show
Me.Refresh
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Over(OldIndex) = True Then
    picButton(OldIndex).Picture = images.ListImages(OldIndex + 1).Picture
    AlphaBlending picButton(OldIndex).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(OldIndex).Left, picButton(OldIndex).Top, 170, 40, 80
    Over(OldIndex) = False
End If
HideSubs
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picButton(Index).Picture = imagesdown.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 80
    Select Case Index
        Case 0
                frmShutdownSubMenu.SetFocus
        Case 2
                frmHelpSubMenu.SetFocus
    End Select
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Over(Index) = False Then
    picButton(OldIndex).Picture = images.ListImages(OldIndex + 1).Picture
    AlphaBlending picButton(OldIndex).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(OldIndex).Left, picButton(OldIndex).Top, 170, 40, 80
    Over(OldIndex) = False
    OldIndex = Index
    picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 80
    Over(Index) = True
    Select Case Index
        Case 0
            HideSubs
            Load frmShutdownSubMenu
            SubShown(0) = True
        Case 2
            HideSubs
            Load frmHelpSubMenu
            SubShown(1) = True
        Case 3
            HideSubs
            Load frmSettingsSubMenu
            SubShown(2) = True
        Case Else
            HideSubs
    End Select
    s_Playsound "hover"
End If
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 80
Select Case Index
Case 1
    Load frmRun
    HideStartMenu
Case 4
    frmSendPopUp.Show
    HideStartMenu
End Select
    s_Playsound "select"
End Sub

Private Sub tmrHide_Timer()
a = GetForegroundWindow
b = 0
If SubShown(0) Then b = frmShutdownSubMenu.hWnd
If SubShown(1) Then b = frmHelpSubMenu.hWnd
If SubShown(2) Then b = frmSettingsSubMenu.hWnd

If a <> Me.hWnd And TaskbarOpen And a <> b Then
    
End If
End Sub
