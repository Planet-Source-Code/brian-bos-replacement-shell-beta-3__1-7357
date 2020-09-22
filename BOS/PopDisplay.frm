VERSION 5.00
Begin VB.Form frmPopUpDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PopUpMsg!"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "PopDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2235
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   300
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "From: "
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3375
   End
End
Attribute VB_Name = "frmPopUpDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Beep
End Sub
