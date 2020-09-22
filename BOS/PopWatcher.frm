VERSION 5.00
Begin VB.Form frmPopUpWatcher 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   95
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPopUpWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim StrRecieved As String
Dim MsgStart, UNameEnd As Integer

Private Sub FlatBttn1_Click()
frmSendPopUp.Show
Me.Hide
End Sub

