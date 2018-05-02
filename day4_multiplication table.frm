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
      Caption         =   "Generate multiplication table"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      TabIndex        =   2
      Top             =   5040
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   11520
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter the number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, m As Integer
Dim s As String
a = Val(Text1.Text)
b = 1
Do While (b <= 10)
m = a * b
s = a & "*" & b & "=" & m
List1.AddItem s
b = b + 1
Loop
End Sub

