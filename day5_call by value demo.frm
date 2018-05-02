VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click"
      Height          =   1095
      Left            =   7080
      TabIndex        =   0
      Top             =   3480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer
a = 5
MsgBox "Value before is " & a
Call demo(a)
MsgBox "Value after is " & a
End Sub

Private Sub demo(ByVal b As Integer)
b = b + 1
MsgBox "Value inside is " & b
End Sub

