VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   8040
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "average"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "chemistry"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "physics"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Maths"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Single

Private Sub Command1_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
average = (a + b + c) / 3
Label5.Caption = average
If average > 75 Then
Text4.Text = "Good boy!"
ElseIf average > 65 Then
Text4.Text = "B"
ElseIf average > 55 Then
Text4.Text = "C"
ElseIf average > 45 Then
Text4.Text = "D"
Else
Text4.Text = "E"
End If
End Sub
