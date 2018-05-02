VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "height"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "area"
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim h, r, a, rad, area As Single
Const pi As Single = 3.142


Private Sub Command1_Click()
r = h / 2
rad = r * 0.001763889
a = pi * rad ^ 2
area = Round(a, 2)
MsgBox ("The area is" & area)
End Sub


Private Sub Command2_Click()
h = InputBox("enter the height")
Shape1.Height = h
End Sub

