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
   Begin VB.CommandButton Command5 
      Caption         =   "Clear lists"
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Multiplication"
      Height          =   975
      Left            =   11760
      TabIndex        =   7
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Addition"
      Height          =   975
      Left            =   7920
      TabIndex        =   6
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Transpose"
      Height          =   975
      Left            =   4320
      TabIndex        =   5
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert elements"
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   5040
      Width           =   2415
   End
   Begin VB.ListBox List4 
      Height          =   2790
      Left            =   11880
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin VB.ListBox List3 
      Height          =   2790
      Left            =   8400
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, c, i, j As Integer
Dim str As String
Dim a(), b(), d(), e() As Integer
Private Sub Command1_Click()
r = InputBox("Enter number of rows")
c = InputBox("Enter number of columns")
ReDim a(r - 1, c - 1)
For i = 0 To r - 1
str = ""
For j = 0 To c - 1
a(i, j) = InputBox("Enter the element")
str = str & a(i, j) & " "
Next j
List1.AddItem str
Next i
End Sub

Private Sub Command2_Click()
ReDim b(c - 1, r - 1)
For i = 0 To c - 1
str = ""
For j = 0 To r - 1
b(i, j) = a(j, i)
str = str & b(i, j) & " "
Next j
List2.AddItem str
Next i
End Sub

Private Sub Command3_Click()
If Not r = c Then
MsgBox ("Number of rows and columns must be equal, buddy!")
Else
ReDim d(r - 1, c - 1)
For i = 0 To r - 1
str = ""
For j = 0 To c - 1
d(i, j) = Val(a(i, j)) + Val(b(i, j))
str = str & d(i, j) & " "
Next j
List3.AddItem str
Next i
End If
End Sub

Private Sub Command4_Click()
Dim sum, k As Integer
ReDim e(r - 1, r - 1)
For i = 0 To r - 1
str = ""
For j = 0 To r - 1
sum = 0
For k = 0 To c - 1
sum = sum + (Val(a(i, k)) * Val(b(k, j)))
Next k
e(i, j) = sum
str = str & e(i, j) & " "
Next j
List4.AddItem str
Next i
End Sub

Private Sub Command5_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
End Sub
