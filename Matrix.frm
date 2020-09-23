VERSION 5.00
Begin VB.Form form10 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   Icon            =   "Matrix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command1 
      Caption         =   "END"
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   105
      Top             =   90
   End
End
Attribute VB_Name = "form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type bollar
    x As Long
    y As Long
    Vel As Long
    Storlek As Long
    Farg As Byte
    Bok As Byte
End Type
Dim boll(120) As bollar
Const Max As Long = 120
Dim i As Integer
Dim j As Integer
Const MaxVel As Long = 160
Const MinVel As Long = 160
Const MaxStorlek As Long = 12
Private Sub Command1_Click()
End
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

Tillverka
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
Randomize
Me.BackColor = vbBlack


For i = 0 To Max
    boll(i).y = boll(i).y + boll(i).Vel
    If boll(i).y > Me.Height Then
    boll(i).y = -15
    boll(i).x = Me.Width * Rnd + 1
    boll(i).Storlek = Int(MaxStorlek * Rnd)
    boll(i).Vel = MaxVel * Rnd + MinVel
    boll(i).Farg = Int(Rnd * 255) + 1
    End If
    boll(i).Bok = 1 * Rnd
    Me.Line (boll(i).x, boll(i).y)-((boll(i).x), (boll(i).y)), vbBlack, B
    form10.ForeColor = RGB(25, boll(i).Farg, 25)
    Print boll(i).Bok
Next i


Me.Refresh
End Sub
Private Sub Tillverka()
Randomize
For i = 0 To Max
    boll(i).x = Me.Width * Rnd + 1
    boll(i).y = Me.Height * Rnd + 1
    boll(i).Storlek = MaxStorlek * Rnd + 3
    boll(i).Vel = MaxVel * Rnd + MinVel
    boll(i).Farg = Int(Rnd * 255) + 1
    boll(i).Bok = 1 * Rnd
Next i
End Sub
