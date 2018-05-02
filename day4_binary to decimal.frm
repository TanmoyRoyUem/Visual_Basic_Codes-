VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enter a binary number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   0
      Top             =   3000
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, i, r, n As Integer
a = InputBox("Enter the number")
b = a
Do While (b)
r = b Mod 10
If r > 1 Then
Status = False
Exit Do
Else
Status = True
End If
b = b / 10
Loop
If Status = False Then
MsgBox "Not a binary number"
Else
MsgBox "Valid"
b = a
n = 0
i = 0
Do While (b)
r = b Mod 10
n = n + r * (2 ^ i)
b = b / 10
i = i + 1
Loop
MsgBox "the decimal number is " & n
End If
End Sub

