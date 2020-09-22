VERSION 5.00
Begin VB.Form frmBPrograms 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   0
      Width           =   4035
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   3600
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmBPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer

Private Sub Form_Load()
Dir1.path = "C:\windows\start menu\programs"
File1.path = "C:\windows\start menu\programs"
For i = 1 To Dir1.ListCount + File1.ListCount - 2
    Load picItem(i)
    picItem(i).Visible = True
    picItem(i).Top = 20 * i
Next

For i = 0 To Dir1.ListCount - 1
    picItem(i).Print ExtractFileName(Dir1.List(i)) & " >"
    picItem(i).Tag = ExtractFileName(Dir1.List(i)) & " >"
Next

For i = 0 To File1.ListCount - 1
    picItem(i + Dir1.ListCount - 1).Print Left(File1.List(i), Len(File1.List(i)) - 4)
    picItem(i + Dir1.ListCount - 1).Tag = Left(File1.List(i), Len(File1.List(i)) - 4)
Next
Me.Height = picItem.Count * 20 * Screen.TwipsPerPixelY + 10
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picItem(Index).BackColor = vbHighlight
picItem(Index).Cls
picItem(Index).ForeColor = vbHighlightText
picItem(Index).Print picItem(Index).Tag

If Index <> OldIndex Then
    picItem(OldIndex).BackColor = vbButtonFace
    picItem(OldIndex).Cls
    picItem(OldIndex).ForeColor = vbButtonText
    picItem(OldIndex).Print picItem(OldIndex).Tag
OldIndex = Index
End If


End Sub
