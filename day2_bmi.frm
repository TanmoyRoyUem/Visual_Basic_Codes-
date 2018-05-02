VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE BMI"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Height          =   855
      Left            =   1560
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Weight (in kg)"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Height (in m)"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bmi, h, w As Integer

Private Sub Command1_Click()
h = Val(Text1.Text)
w = Val(Text2.Text)
bmi = w / (h * h)
Label3.Caption = bmi
Select Case bmi
Case 30 To 100
Label4.Caption = "Obese"
Label4.BackColor = vbRed
Case 25 To 29.9
Label4.Caption = "Overweight"
Label4.BackColor = vbAmber
Case 18.6 To 24.9
Label4.Caption = "Normal"
Label4.BackColor = vbGreen
Case 0 To 18.5
Label4.Caption = "Underweight"
Label4.BackColor = vbRed
Case Else
Label4.Caption = "Wrong choice"
End Select
End Sub
