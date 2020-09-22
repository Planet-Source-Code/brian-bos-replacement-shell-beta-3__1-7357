VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGetSkin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get skin from internet"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmGetSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox dirSkins 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   -1000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   3060
      Width           =   1035
   End
   Begin VB.CommandButton cmdInstallSkin 
      Caption         =   "Install Selected Skin"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   3060
      Width           =   1815
   End
   Begin VB.ListBox lstSkins 
      Height          =   2595
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4515
   End
   Begin InetCtlsObjects.Inet inetGetSkin 
      Left            =   3780
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   4980
      Y1              =   2950
      Y2              =   2950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -120
      X2              =   4860
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label lblCurrentStatus 
      Caption         =   "Connecting to server..."
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   2700
      Width           =   4455
   End
End
Attribute VB_Name = "frmGetSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SkinList() As String
Dim TmpString As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdInstallSkin_Click()
Dim a As String
a = lstSkins.List(lstSkins.ListIndex)
For i = 0 To dirSkins.ListCount - 1
    If a = ExtractFileName(dirSkins.List(i)) Then
        MsgBox "This skin has already been installed.", vbInformation, "Already Installed"
        Exit Sub
    End If
Next

Me.Hide
frmDownloadSkin.InstallSkin lstSkins.List(lstSkins.ListIndex)
End Sub

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    dirSkins.path = App.path & "\skins"
    TmpString = inetGetSkin.OpenURL("http://thunder.prohosting.com/~pikared/skins/skins.txt")
    SkinList = Split(TmpString, ",")
    For i = 0 To UBound(SkinList)
        lstSkins.AddItem SkinList(i)
    Next
    lblCurrentStatus.Caption = "Please select a skin."
End Sub

Private Sub lstSkins_Click()
cmdInstallSkin.Enabled = True
End Sub

Private Sub lstSkins_DblClick()
cmdInstallSkin_Click
End Sub
