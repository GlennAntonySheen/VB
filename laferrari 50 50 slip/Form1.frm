VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6645
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   15
      Picture         =   "Form1.frx":23676
      ScaleHeight     =   6675
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      Height          =   7695
      Left            =   6240
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   -480
      Width           =   150
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      Picture         =   "Form1.frx":31841
      ScaleHeight     =   6735
      ScaleWidth      =   12015
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mdown As Boolean
Dim xx As Integer, yy As Integer

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdown = True
xx = X

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If mdown = True Then
    If Picture1.Width + (X - xx) > 0 Then
      If Picture2.Width - 150 >= Picture1.Width + (X - xx) Then
        Text1.Left = (X - xx) + Text1.Left
        
        Picture1.Width = Picture1.Width + (X - xx)
      End If
    End If
  End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdown = False
End Sub


