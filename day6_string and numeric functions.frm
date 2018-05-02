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
   Begin VB.Frame Frame2 
      Caption         =   "Numeric Functions"
      Height          =   2415
      Left            =   480
      TabIndex        =   16
      Top             =   6120
      Width           =   15375
      Begin VB.CommandButton Command18 
         Caption         =   "Rnd"
         Height          =   495
         Left            =   11640
         TabIndex        =   24
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Abs"
         Height          =   495
         Left            =   7560
         TabIndex        =   23
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Sqr"
         Height          =   615
         Left            =   4320
         TabIndex        =   22
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Round"
         Height          =   615
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Log"
         Height          =   615
         Left            =   11640
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Exp"
         Height          =   615
         Left            =   7680
         TabIndex        =   19
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Fix"
         Height          =   615
         Left            =   4200
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Int"
         Height          =   615
         Left            =   1080
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "String functions"
      Height          =   3255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   15495
      Begin VB.CommandButton Command10 
         Caption         =   "Instr"
         Height          =   615
         Left            =   12960
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Mid"
         Height          =   615
         Left            =   9480
         TabIndex        =   14
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Right"
         Height          =   615
         Left            =   6360
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Left"
         Height          =   615
         Left            =   3600
         TabIndex        =   12
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Trim"
         Height          =   615
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Rtrim"
         Height          =   615
         Left            =   12960
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ltrim"
         Height          =   615
         Left            =   9840
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ucase"
         Height          =   615
         Left            =   6240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Lcase"
         Height          =   615
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Len"
         Height          =   615
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   8760
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Output"
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Input"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Value"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Len(Text1.Text)
End Sub

Private Sub Command10_Click()
Text2.Text = InStr(Text1.Text, "roy")
End Sub

Private Sub Command11_Click()
Text2.Text = Int(Text1.Text)
End Sub

Private Sub Command12_Click()
Text2.Text = Fix(Text1.Text)
End Sub

Private Sub Command13_Click()
Text2.Text = Exp(Text1.Text)
End Sub

Private Sub Command14_Click()
Text2.Text = Log(Text1.Text)
End Sub

Private Sub Command15_Click()
Text2.Text = Round(Text1.Text, 2)
End Sub

Private Sub Command16_Click()
Text2.Text = Sqr(Text1.Text)
End Sub

Private Sub Command17_Click()
Text2.Text = Abs(Text1.Text)
End Sub

Private Sub Command18_Click()
Randomize
Text2.Text = Rnd()
End Sub

Private Sub Command2_Click()
Text2.Text = LCase(Text1.Text)
End Sub

Private Sub Command3_Click()
Text2.Text = UCase(Text1.Text)
End Sub

Private Sub Command4_Click()
Text2.Text = LTrim(Text1.Text)
End Sub

Private Sub Command5_Click()
Text2.Text = RTrim(Text1.Text)
End Sub

Private Sub Command6_Click()
Text2.Text = Trim(Text1.Text)
End Sub

Private Sub Command7_Click()
Text2.Text = Left(Text1.Text, 2)
End Sub

Private Sub Command8_Click()
Text2.Text = Right(Text1.Text, 2)
End Sub

Private Sub Command9_Click()
Text2.Text = Mid(Text1.Text, 2, 4)
End Sub
