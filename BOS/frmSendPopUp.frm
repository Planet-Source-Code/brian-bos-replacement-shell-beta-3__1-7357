VERSION 5.00
Begin VB.Form frmSendPopUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PopUpMsg!"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmSendPopUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3555
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   315
         Left            =   2340
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Sending PopUpMsg!"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   3555
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2235
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   420
      Width           =   3435
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   2700
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmSendPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then
    MsgBox "Please enter a message!", vbInformation, "Enter a Message"
    Exit Sub
End If

If Text2.Text = "" Then
    MsgBox "Please enter a computer name!", vbInformation, "Enter computer name"
    Exit Sub
End If

Command3.Visible = False
Command4.Visible = False
Label2.Caption = "Sending PopUpMsg!..."
Picture1.Visible = True
Me.Refresh
FrmTaskbar.Winsock1.RemoteHost = Text2.Text
On Error GoTo err
FrmTaskbar.Winsock1.SendData strGetMachineName & "<!MSG!>" & Text1.Text
Unload Me

err:
If err.Number = 10014 Then
    Label2.Caption = "Can not find computer: " & Text2.Text
    err.Clear
    Command4.Visible = True
ElseIf err.Number = 0 Then
    Label2.Caption = "PopUpMsg! Sent to computer " & Text2.Text
    Command3.Visible = True
Else
    Label2.Caption = "Error: " & err.Description
    err.Clear
    Command3.Visible = True
End If

End Sub

Private Sub Command2_Click()
frmPopUpWatcher.Show
Unload Me
End Sub

Private Sub Command3_Click()
frmPopUpWatcher.Show
Unload Me
End Sub

Private Sub Command4_Click()
Picture1.Visible = False
End Sub

