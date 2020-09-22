VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTaskbar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17655
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1177
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPSCDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3420
      Picture         =   "B.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   375
      TabIndex        =   37
      Top             =   -1500
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picPSCUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3000
      Picture         =   "B.frx":092A
      ScaleHeight     =   450
      ScaleWidth      =   375
      TabIndex        =   36
      Top             =   -1500
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picPSCButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   990
      Picture         =   "B.frx":1254
      ScaleHeight     =   450
      ScaleWidth      =   375
      TabIndex        =   35
      Top             =   75
      Width           =   375
   End
   Begin VB.PictureBox picSeperator 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   1365
      Picture         =   "B.frx":1B7E
      ScaleHeight     =   450
      ScaleWidth      =   90
      TabIndex        =   34
      Top             =   75
      Width           =   90
   End
   Begin VB.PictureBox picProgramUp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   8460
      Picture         =   "B.frx":1E18
      ScaleHeight     =   555
      ScaleWidth      =   5235
      TabIndex        =   18
      Top             =   -1500
      Visible         =   0   'False
      Width           =   5235
   End
   Begin VB.PictureBox picProgramDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   495
      Left            =   8760
      Picture         =   "B.frx":767A
      ScaleHeight     =   495
      ScaleWidth      =   5655
      TabIndex        =   17
      Top             =   -1500
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.PictureBox picTaskbar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   13260
      Picture         =   "B.frx":CEDC
      ScaleHeight     =   450
      ScaleWidth      =   75
      TabIndex        =   33
      Top             =   -1500
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picSeperator 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   15420
      Picture         =   "B.frx":D0FE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   22
      Top             =   75
      Width           =   690
      Begin VB.PictureBox picScrollUp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   390
         Picture         =   "B.frx":D398
         ScaleHeight     =   450
         ScaleWidth      =   300
         TabIndex        =   24
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picScrollDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   90
         Picture         =   "B.frx":DAE2
         ScaleHeight     =   450
         ScaleWidth      =   300
         TabIndex        =   23
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox picSystemTray 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   16140
      Picture         =   "B.frx":E22C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   75
      Width           =   1500
      Begin VB.Image imgNetStatus 
         Height          =   240
         Left            =   120
         Picture         =   "B.frx":10596
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   420
         Picture         =   "B.frx":106E0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11:00 PM"
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
         Left            =   660
         TabIndex        =   9
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   6180
      Picture         =   "B.frx":10B2A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   29
      Tag             =   "u"
      Top             =   75
      Width           =   3750
      Begin VB.PictureBox picNetOn 
         Height          =   195
         Left            =   1380
         Picture         =   "B.frx":1638C
         ScaleHeight     =   135
         ScaleWidth      =   315
         TabIndex        =   31
         Top             =   -300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picNetOff 
         Height          =   195
         Left            =   1920
         Picture         =   "B.frx":164D6
         ScaleHeight     =   135
         ScaleWidth      =   375
         TabIndex        =   30
         Top             =   -300
         Width           =   435
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
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
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   90
         Width           =   3675
      End
   End
   Begin VB.PictureBox picScrollDownUp 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   8100
      Picture         =   "B.frx":16620
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   -1500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picScrollDownDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   13620
      Picture         =   "B.frx":16D6A
      ScaleHeight     =   450
      ScaleWidth      =   300
      TabIndex        =   27
      Top             =   -1500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picScrollUpDown 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   8040
      Picture         =   "B.frx":174B4
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   -1500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picScrollUpUp 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   6840
      Picture         =   "B.frx":17BFE
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   25
      Top             =   -1500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   -15000
      Picture         =   "B.frx":18348
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   20
      Tag             =   "u"
      Top             =   75
      Width           =   3750
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
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
         Index           =   4
         Left            =   60
         TabIndex        =   21
         Top             =   90
         Width           =   3675
      End
   End
   Begin VB.PictureBox picProgramNone 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   6840
      Picture         =   "B.frx":1DBAA
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   19
      Top             =   -1500
      Width           =   1215
   End
   Begin VB.Timer tmrCheckWindows 
      Interval        =   10
      Left            =   10500
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   -7500
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   -15000
      Picture         =   "B.frx":2340C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   14
      Tag             =   "u"
      Top             =   75
      Width           =   3750
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
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
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   3675
      End
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   -15000
      Picture         =   "B.frx":28C6E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   12
      Tag             =   "u"
      Top             =   75
      Width           =   3750
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
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
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Top             =   90
         Width           =   3675
      End
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   -15000
      Picture         =   "B.frx":2E4D0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   10
      Tag             =   "u"
      Top             =   75
      Width           =   3750
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
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
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   3675
      End
   End
   Begin VB.Timer tmrCheckOnline 
      Interval        =   100
      Left            =   11460
      Top             =   0
   End
   Begin VB.Timer tmrUpdateTime 
      Interval        =   10000
      Left            =   10020
      Top             =   0
   End
   Begin VB.PictureBox picSeperator 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   900
      Picture         =   "B.frx":33D32
      ScaleHeight     =   450
      ScaleWidth      =   90
      TabIndex        =   8
      Top             =   75
      Width           =   90
   End
   Begin VB.PictureBox picSeperator 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   6060
      Picture         =   "B.frx":33FCC
      ScaleHeight     =   450
      ScaleWidth      =   90
      TabIndex        =   7
      Top             =   75
      Width           =   90
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   1
      Left            =   10980
      Top             =   0
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11940
      Top             =   0
   End
   Begin VB.PictureBox picLocationBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1500
      Picture         =   "B.frx":34266
      ScaleHeight     =   300
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   180
      Width           =   4500
      Begin VB.Label lblWebAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "http://"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   4395
      End
   End
   Begin VB.PictureBox picTbar 
      Height          =   315
      Left            =   10380
      Picture         =   "B.frx":388F8
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTbarDown 
      Height          =   255
      Left            =   11460
      Picture         =   "B.frx":39E52
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBstart 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "B.frx":3B3AC
      ScaleHeight     =   495
      ScaleWidth      =   900
      TabIndex        =   1
      Top             =   75
      Width           =   900
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   4500
      Left            =   -18120
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   996
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   15000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastTask As Integer
Dim CurrWindow As Long
Dim i As Integer
Dim DeskHdc&, ret&
Dim BlinkOn As Boolean, HasFocus As Boolean
Public tX As Integer, tY As Integer
Dim S As Integer
Dim OldHwnd As Long
Dim ConnectionName As String
Dim NumButs As Integer
Dim ShowAddressBar As Boolean
Dim StrRecieved As String
Dim showPSCButton As Boolean
Dim PSCOpen As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If HasFocus Then
    If KeyAscii = 13 Then
        If BlinkOn Then
                ShellExecute Me.hWnd, "open", Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1), "", "", 1
        Else
                ShellExecute Me.hWnd, "open", lblWebAddress.Caption, "", "", 1
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
loadme
Winsock1.RemotePort = 1001
Winsock1.RemotePort = 1001
Winsock1.Bind 1001
End Sub

Private Sub picPSCButton_Click()
If PSCOpen Then
    PSCOpen = False
    picPSCButton.Picture = picPSCUp.Picture
    Unload frmPSCTicker
Else
    PSCOpen = True
    picPSCButton.Picture = picPSCDown.Picture
    frmPSCTicker.Show
End If
AlphaBlending picPSCButton.hdc, 0, 0, 25, 30, picDesktopCapture.hdc, picPSCButton.Left, 0, 25, 30, 80
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData StrRecieved
frmPopUpDisplay.Show
MsgStart = InStr(StrRecieved, "<!MSG!>") + 7
UNameEnd = MsgStart - 8
frmPopUpDisplay.Label1.Caption = "From: " & Left(StrRecieved, UNameEnd)
frmPopUpDisplay.Text1.Text = Right(StrRecieved, Len(StrRecieved) - MsgStart + 1)
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    frmTaskbarContextMenu.SetPosition X, Y
    s_Playsound ("open")
End If
End Sub

Private Sub Image1_DblClick()
Shell "SNDVOL32.EXE", vbNormalFocus
End Sub

Private Sub imgNetStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConnectionName = GetSetting("Bos", "BInternet", "ConnectionName", "")
If ConnectionName = "" Then
    ConnectionName = InputBox("What is the name of the connection that you use to connect to the internet?" & vbCrLf & "Hint: Look in ""Dial Up Networking"" under My Computer", "Connection Name")
    SaveSetting "Bos", "BInternet", "ConnectionName", ConnectionName
End If
If ActiveConnection Then
    Disconnect ConnectionName, True
Else
    Connect ConnectionName, True
End If
End Sub

Private Sub lblTime_DblClick()
Shell ("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index + S < List1.ListCount Then
    If Button = vbLeftButton Then
        If picProgram(Index).Tag <> "d" Then
                picProgram(Index).Picture = picProgramDown.Picture
                AlphaBlending picProgram(Index).hdc, 0, 0, 250, 30, picDesktopCapture.hdc, picProgram(Index).Left, 0, 250, 30, 80
                picProgram(Index).Tag = "d"
                s_Playsound "select"
                pSetForegroundWindow List1.ItemData(Index + S)
            Else
                s_Playsound "select"
                ShowWindow List1.ItemData(Index + S), SW_MINIMIZE
            End If
    End If
Else
    If Button = vbRightButton Then
        frmTaskbarContextMenu.SetPosition (X / Screen.TwipsPerPixelX) + lblTitle(Index).Left + picProgram(Index).Left, Y / Screen.TwipsPerPixelY
        s_Playsound "open"
    End If
End If
End Sub

Private Sub lblWebAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HasFocus = True
    tmrBlink.Enabled = True
    tmrCheckFocus.Enabled = True
    End Sub

Private Sub picBstart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
If TaskbarOpen = False Then
    ShowStartMenu
    s_Playsound "open"
Else
    HideStartMenu
End If

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


Private Sub picScrollDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picScrollDown.Picture = picScrollDownDown.Picture
AlphaBlending picScrollDown.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollDown.Left + picSeperator(2).Left + 6, 0, 20, 30, 80
End Sub

Private Sub picScrollDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picScrollDown.Picture = picScrollDownUp.Picture
AlphaBlending picScrollDown.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollDown.Left + picSeperator(2).Left + 6, 0, 20, 30, 80
If S + NumButs < List1.ListCount Then
    s_Playsound "select"
    S = S + NumButs
Else
    s_Playsound "select"
End If
End Sub

Private Sub picScrollUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picScrollUp.Picture = picScrollUpDown.Picture
AlphaBlending picScrollUp.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollUp.Left + picSeperator(2).Left + 6, 0, 20, 30, 80
End Sub

Private Sub picScrollUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picScrollUp.Picture = picScrollUpUp.Picture
AlphaBlending picScrollUp.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollUp.Left + picSeperator(2).Left + 6, 0, 20, 30, 80
If S + NumButs > List1.ListCount Then
    s_Playsound "select"
    S = S - NumButs
Else
    s_Playsound "select"
End If
End Sub

Private Sub tmrCheckOnline_Timer()
If ActiveConnection Then
    imgNetStatus.Picture = picNetOn.Picture
Else
    imgNetStatus.Picture = picNetOff.Picture
End If
End Sub

Private Sub tmrCheckWindows_Timer()
UpdateTaskbar
End Sub

Private Sub tmrUpdateTime_Timer()
UpdateTime
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

a = GetForegroundWindow
If a <> Me.hWnd Then
    If BlinkOn Then
        lblWebAddress.Caption = Left(lblWebAddress.Caption, Len(lblWebAddress.Caption) - 1)
    End If
    BlinkOn = False
    HasFocus = False
    tmrBlink.Enabled = False
    tmrCheckFocus.Enabled = False
End If
End Sub


Sub UpdateTaskbar()
fEnumWindows List1
If List1.ListCount > NumButs Then
    If picSeperator(2).Visible = False Then picSeperator(2).Visible = True
Else
    If picSeperator(2).Visible = True Then picSeperator(2).Visible = False
    S = 0
End If

For i = 0 To 4
    If i + S > List1.ListCount - 1 Then
        lblTitle(i).Caption = ""
            If picProgram(i).Tag <> "n" Then
                picProgram(i).Picture = picProgramNone.Picture
                AlphaBlending picProgram(i).hdc, 0, 0, 250, 30, picDesktopCapture.hdc, picProgram(i).Left, 0, 250, 30, 80
                picProgram(i).Tag = "n"
            End If
    Else
        a = Left(List1.List(i + S), InStr(List1.List(i + S), "**BHWND=**") - 1)
        If Len(a) < 30 Then
            lblTitle(i).Caption = a
        Else
            lblTitle(i).Caption = Left(a, 27) & "..."
        End If
            
        If GetForegroundWindow = List1.ItemData(i + S) Then
            If picProgram(i).Tag <> "d" Then
                picProgram(i).Picture = picProgramDown.Picture
                AlphaBlending picProgram(i).hdc, 0, 0, 250, 30, picDesktopCapture.hdc, picProgram(i).Left, 0, 250, 30, 80
                picProgram(i).Tag = "d"
            End If
        Else
            If picProgram(i).Tag <> "u" Then
                picProgram(i).Picture = picProgramUp.Picture
                AlphaBlending picProgram(i).hdc, 0, 0, 250, 30, picDesktopCapture.hdc, picProgram(i).Left, 0, 250, 30, 80
                picProgram(i).Tag = "u"
            End If
        End If
    End If
Next
End Sub

Public Sub loadme()

If GetSetting("Bos", "BTaskbar", "ShowAddressBar", "True") = "True" Then
    ShowAddressBar = True
Else
    ShowAddressBar = False
End If

If GetSetting("Bos", "BTaskbar", "ShowPSCButton", "True") = "True" Then
    showPSCButton = True
Else
    showPSCButton = False
End If

Select Case Screen.Width / Screen.TwipsPerPixelX
    Case 1600
        NumButs = 4 + Z
    Case 1280
        picProgram(3).Visible = False
        NumButs = 3 + Z
    Case 1024
        If ShowAddressBar = True Then
            ans = MsgBox("Your sreen size is set to 1024 X 768. Do you wish to hide the address bar so that more taskbar buttons can be shown?", vbQuestion Or vbYesNo, "Hide Address Bar?")
            If ans = vbYes Then
                ShowAddressBar = False
                SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "False"
            End If
            If ShowAddressBar Then Z = 0 Else Z = 1
        End If
        picProgram(3).Visible = False
        picProgram(2).Visible = False
        NumButs = 2 + Z
    Case 800
            SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "False"
            picProgram(3).Visible = False
            picProgram(2).Visible = False
            NumButs = 2
    Case Is < 800
        MsgBox "Your screen size is too small to run BoS. Falling back on Windows Explorer.", vbOKOnly Or vbInformation, "Screen Size Too Small"
        Shell "explorer.exe"
        End
    Case Else
        MsgBox "Your screen is set to an abnormal resiloution." & vbCrLf & "Please change your resiloution to 1600x1200, 1280x1024, or 1024x768." & vbCrLf & "Falling back on Windows Explorer."
        Shell "explorer.exe"
        End
End Select
If ShowAddressBar = False Then
    picLocationBar.Visible = False
    picSeperator(0).Visible = False
End If

If showPSCButton = False Then
    picPSCButton.Visible = False
    picSeperator(1).Left = 65
    picSeperator(3).Visible = False
    picLocationBar.Left = 75
    picSeperator(0).Left = 380
End If

fEnumWindows List1
Me.Width = Screen.Width
picSystemTray.Left = Screen.Width / Screen.TwipsPerPixelX - picSystemTray.ScaleWidth
picSeperator(2).Left = picSystemTray.Left - picSeperator(2).Width

a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
If a <> "BoS Standard" Then
    picBstart.Picture = LoadPicture(App.path & "\skins\" & a & "\StartButtonUp.bmp")
    picTbar.Picture = picBstart.Picture
    picTbarDown.Picture = LoadPicture(App.path & "\skins\" & a & "\StartButtonDown.bmp")
    picTaskbar.Picture = LoadPicture(App.path & "\skins\" & a & "\TaskbarBg.bmp")
    picSeperator(0).Picture = LoadPicture(App.path & "\skins\" & a & "\Seperator.bmp")
    picProgramUp.Picture = LoadPicture(App.path & "\skins\" & a & "\ProgramUp.bmp")
    picProgramDown.Picture = LoadPicture(App.path & "\skins\" & a & "\ProgramDown.bmp")
    picProgramNone.Picture = LoadPicture(App.path & "\skins\" & a & "\ProgramNone.bmp")
    picSystemTray.Picture = LoadPicture(App.path & "\skins\" & a & "\SystemTray.bmp")
    picScrollUpUp.Picture = LoadPicture(App.path & "\skins\" & a & "\UpArrowUp.bmp")
    picScrollUpDown.Picture = LoadPicture(App.path & "\skins\" & a & "\UpArrowDown.bmp")
    picScrollDownUp.Picture = LoadPicture(App.path & "\skins\" & a & "\DownArrowUp.bmp")
    picScrollDownDown.Picture = LoadPicture(App.path & "\skins\" & a & "\DownArrowDown.bmp")
    picPSCUp.Picture = LoadPicture(App.path & "\skins\" & a & "\PSCUp.bmp")
    picPSCDown.Picture = LoadPicture(App.path & "\skins\" & a & "\PSCDown.bmp")
    lblTitle(0).ForeColor = Val(GetSetting("BoS", "BTaskbar", "FgColor", "&H0000000&"))
    lblTime.ForeColor = Val(GetSetting("BoS", "BTaskbar", "ClockFgColor", "&H0000000&"))
End If
picScrollDown.Picture = picScrollDownUp.Picture
picScrollUp.Picture = picScrollUpUp.Picture
picPSCButton.Picture = picPSCUp.Picture

For i = 0 To lblTitle.Count - 1
lblTitle(i).ForeColor = lblTitle(0).ForeColor
Next

For i = 0 To picProgram.Count - 1
    picProgram(i).Picture = picProgramUp.Picture
Next
StretchBlt Me.hdc, 0, 5, Screen.Width / Screen.TwipsPerPixelX, 30, picTaskbar.hdc, 0, 0, 5, 30, vbSrcCopy
For i = 1 To 4
    Me.Line (0, i)-(Screen.Width / Screen.TwipsPerPixelX, i)
Next
For i = 1 To picSeperator.Count - 1
    picSeperator(i).Picture = picSeperator(0).Picture
Next
SetWindowPos Me.hWnd, -1, 0, Screen.Height / Screen.TwipsPerPixelY - 35, 0, 0, SWP_NOREPOSITION & SWP_NOSIZE
picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX + 1, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
AlphaBlending Me.hdc, 0, 5, Me.ScaleWidth, Me.ScaleHeight, picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 80
Blend picBstart, picDesktopCapture, 100, 0, 0, picBstart.Width, picBstart.Height

AlphaBlending picLocationBar.hdc, 0, 0, picLocationBar.Width, picLocationBar.Height, Me.hdc, 64, 7, picLocationBar.Width, picLocationBar.Height, 200
AlphaBlending picSystemTray.hdc, 0, 0, 100, 30, picDesktopCapture.hdc, Me.ScaleWidth - 100, 0, 100, 30, 80
AlphaBlending picScrollDown.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollDown.Left + picSeperator(2).Left + 6, 0, 20, 30, 80
AlphaBlending picScrollUp.hdc, 0, 0, 20, 30, picDesktopCapture.hdc, picScrollUp.Left + picSeperator(2).Left + 6, 0, 20, 30, 80


For i = 0 To 4
    Blend Me, picDesktopCapture, (5 - i) * 50, 140, i, Me.ScaleWidth, 1
Next


For i = 0 To picSeperator.Count - 1
    AlphaBlending picSeperator(i).hdc, 0, 0, 6, 30, picDesktopCapture.hdc, picSeperator(i).Left, 0, 6, 30, 80
Next
For i = 0 To NumButs - 1
    picProgram(i).Left = IIf(ShowAddressBar, 315, 0) + IIf(showPSCButton, 22, 0) + 75 + 255 * i
    AlphaBlending picProgram(i).hdc, 0, 0, 250, 30, picDesktopCapture.hdc, picProgram(i).Left, 0, 250, 30, 80
Next
AlphaBlending picPSCButton.hdc, 0, 0, 25, 30, picDesktopCapture.hdc, picPSCButton.Left, 0, 25, 30, 80
Load frmBDesktopIcons

For i = 0 To 4
AlphaBlending Me.hdc, i + 5, 0, 1, 5, frmBDesktopIcons.hdc, i + 5, frmBDesktopIcons.ScaleHeight - 5, 1, 5, (5 - i) * 50
Next
For i = 0 To 4
AlphaBlending Me.hdc, 5, i, 135, 1, frmBDesktopIcons.hdc, 5, frmBDesktopIcons.ScaleHeight - 5 + i, 135, 1, (5 - i) * 50
Next
BitBlt Me.hdc, 0, 0, 5, 5, frmBDesktopIcons.hdc, 0, frmBDesktopIcons.ScaleHeight - 5, vbSrcCopy

tmrUpdateTime_Timer
ConnectionName = GetSetting("Bos", "BInternet", "ConnectionName", "")
End Sub

Public Sub UpdateTime()
a = Hour(Now)
If GetSetting("Bos", "BSystemTray", "TimeFormat", "12") = 12 Then
    If a > 12 Then
        tStr = "PM"
        a = a - 12
    Else
        tStr = "AM"
    End If
    lblTime.Caption = a & ":" & Format(Minute(Now), "00") & " " & tStr
Else
    lblTime.Caption = a & ":" & Format(Minute(Now), "00")
End If
End Sub


