VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bos Settings"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmBSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdInstallSkin 
      Caption         =   "Install..."
      Height          =   315
      Left            =   3780
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.DirListBox dirSkins 
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   -480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.ComboBox cmbSkin 
      Height          =   315
      ItemData        =   "frmBSettings.frx":038A
      Left            =   1140
      List            =   "frmBSettings.frx":0391
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.OptionButton optTime 
      Caption         =   "24 Hour"
      Height          =   255
      Index           =   1
      Left            =   2820
      TabIndex        =   9
      Top             =   3360
      Width           =   915
   End
   Begin VB.OptionButton optTime 
      Caption         =   "12 Hour"
      Height          =   255
      Index           =   0
      Left            =   1860
      TabIndex        =   8
      Top             =   3360
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3780
      Width           =   1155
   End
   Begin VB.TextBox txtDunConnection 
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   2760
      Width           =   1875
   End
   Begin VB.CheckBox chkShowPSCButton 
      Caption         =   "Show Planet Source Code Button"
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Top             =   1380
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   3780
      Width           =   1155
   End
   Begin VB.CheckBox chkShowWebAddress 
      Caption         =   "Show web address box on taskbar"
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Skin:"
      Height          =   255
      Left            =   660
      TabIndex        =   10
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   60
      Picture         =   "frmBSettings.frx":03A3
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Clock Format:"
      Height          =   195
      Left            =   780
      TabIndex        =   7
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   60
      Picture         =   "frmBSettings.frx":2B45
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Configure how BoS looks and behaves"
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   180
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   60
      Picture         =   "frmBSettings.frx":2F87
      Top             =   0
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   4740
      Y1              =   490
      Y2              =   490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4740
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dial Up Networking Connection:"
      Height          =   435
      Left            =   660
      TabIndex        =   4
      Top             =   2700
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   60
      Picture         =   "frmBSettings.frx":5729
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   60
      Picture         =   "frmBSettings.frx":7ECB
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmBSettings.frx":8795
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmBSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TaskbarRefresh As Boolean, IconRefresh As Boolean, ShowAddress As Boolean, ShowPSC As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdInstallSkin_Click()
Dim tFilePath As String
Dim NewFileName As String
Dim NewFolder As String
Dim ExecString As String
frmGetSkin.Show
'If ExtractPath(cdBrowse.FileName) <> App.path Then
'    FileCopy cdBrowse.FileName, App.path & "\" & ExtractFileName(cdBrowse.FileName)
'    tFilePath = App.path & "\" & ExtractFileName(cdBrowse.FileName)
'Else
'    tFilePath = cdBrowse.FileName
'End If
'NewFileName = Left(tFilePath, Len(tFilePath) - 4) & ".zip"
'If FileExists(NewFileName) Then Exit Sub
'Name tFilePath As NewFileName
'NewFolder = App.path & "\skins\" & Left(ExtractFileName(NewFileName), Len(ExtractFileName(NewFileName)) - 4)
'If FileExists(NewFolder) Then Exit Sub
'MkDir NewFolder
'FileCopy NewFileName, NewFolder & "\" & ExtractFileName(NewFileName)
'ExecString = App.path & "\" & "unzip.exe " & NewFolder & "\" & ExtractFileName(NewFileName) & " -d " & NewFolder
'Debug.Print ExecString
'Shell ExecString
'Kill NewFileName
'Kill NewFolder & "\" & ExtractFileName(NewFileName)
End Sub

Private Sub cmdOK_Click()
Me.Hide
TaskbarRefresh = False
IconRefresh = False

SaveSetting "Bos", "BInternet", "ConnectionName", txtDunConnection.Text
If chkShowWebAddress.Value = 1 Then
    SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "True"
    If ShowAddress = False Then TaskbarRefresh = True
Else
    SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "False"
    If ShowAddress = True Then TaskbarRefresh = True
End If

If chkShowPSCButton.Value = 1 Then
    SaveSetting "Bos", "BTaskbar", "ShowPscButton", "True"
    If ShowPSC = False Then TaskbarRefresh = True
Else
    SaveSetting "Bos", "BTaskbar", "ShowPscButton", "False"
    If ShowPSC = True Then TaskbarRefresh = True
End If

a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
If a <> cmbSkin.List(cmbSkin.ListIndex) Then
    SaveSetting "BoS", "BInterface", "Skin", cmbSkin.List(cmbSkin.ListIndex)
    ChangeSkin cmbSkin.List(cmbSkin.ListIndex)
    TaskbarRefresh = True
    IconRefresh = True
End If

If optTime(0).Value = True Then
    SaveSetting "Bos", "BSystemTray", "TimeFormat", "12"
Else
    SaveSetting "Bos", "BSystemTray", "TimeFormat", "24"
End If
If IconRefresh Then
    Unload frmBDesktopIcons
    DoEvents
    frmBDesktopIcons.Show
End If

If TaskbarRefresh Then
    Unload FrmTaskbar
    DoEvents
    FrmTaskbar.Show
End If


FrmTaskbar.UpdateTime
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
dirSkins.path = App.path & "\skins"
For i = 0 To dirSkins.ListCount - 1
    cmbSkin.AddItem ExtractFileName(dirSkins.List(i))
Next
a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
For i = 0 To cmbSkin.ListCount - 1
    If a = cmbSkin.List(i) Then
        cmbSkin.ListIndex = i
        Exit For
    End If
Next

txtDunConnection.Text = GetSetting("Bos", "BInternet", "ConnectionName", "")
If GetSetting("Bos", "BSystemTray", "TimeFormat", "12") = "12" Then
    optTime(0).Value = True
Else
    optTime(1).Value = True
End If

If GetSetting("Bos", "BTaskbar", "ShowAddressBar", "True") = "True" Then
    chkShowWebAddress.Value = 1
    ShowAddress = True
Else
    chkShowWebAddress.Value = 0
    ShowAddress = False
End If

If GetSetting("Bos", "BTaskbar", "ShowPSCButton", "True") = "True" Then
    chkShowPSCButton.Value = 1
    ShowPSC = True
Else
    chkShowPSCButton.Value = 0
    ShowPSC = False
End If
End Sub

Sub ChangeSkin(SkinName As String)
Dim TmpString As String
Dim tmpArray() As String
If SkinName = "BoS Standard" Then
    SaveSetting "BoS", "BDesktopIcons", "BgColor", "&H00000000&"
    SaveSetting "BoS", "BDesktopIcons", "FgColor", "&H00FFFFFF&"
    SaveSetting "BoS", "BDesktopIcons", "ShadowColor", "&H00000000&"
    SaveSetting "BoS", "BTaskbar", "FgColor", "&H00000000&"
    SaveSetting "BoS", "BTaskbar", "ClockFgColor", "&H00000000&"
Else
    
    ' Load the icon background color
        Open App.path & "\skins\" & SkinName & "\IconBGColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "BgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the icon foreground color
        Open App.path & "\skins\" & SkinName & "\IconFGColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "FgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the icon shadow color
        Open App.path & "\skins\" & SkinName & "\IconShadowColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "ShadowColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the taskbar button foreground color
        Open App.path & "\skins\" & SkinName & "\TaskbarFgColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BTaskbar", "FgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the clock foreground color
        Open App.path & "\skins\" & SkinName & "\ClockFgColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BTaskbar", "ClockFgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
End If
End Sub

