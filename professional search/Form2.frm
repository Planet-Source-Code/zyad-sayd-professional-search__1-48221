VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2712
   LinkTopic       =   "Form2"
   ScaleHeight     =   900
   ScaleWidth      =   2712
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2280
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   900
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   0
      Width           =   2700
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Load Form1
Form1.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Load Form1
Form1.Show
Unload Me
End Sub
