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
      Caption         =   "Exit"
      Height          =   855
      Left            =   11160
      TabIndex        =   6
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sort (Descending order)"
      Height          =   1095
      Left            =   11160
      TabIndex        =   5
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sort (Ascending order)"
      Height          =   1095
      Left            =   11040
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   855
      Left            =   11160
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   735
      Left            =   11040
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   6300
      Left            =   5880
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, i, j As Integer
Dim arr() As Integer

Private Sub Command1_Click()
List1.Clear
List2.Clear
n = InputBox("Enter size of array:")
ReDim arr(n - 1)
For i = 0 To n - 1
arr(i) = InputBox("Enter element:")
List1.AddItem arr(i)
Next
End Sub

Private Sub Command2_Click()
e = InputBox("Enter the element to be searched: ")
For i = 0 To n - 1
If arr(i) = e Then
MsgBox "element found at position " & i + 1
flag = 1
Exit For
End If
Next i
If flag = 0 Then
MsgBox "element not found"
End If
End Sub

Private Sub Command3_Click()
List2.Clear
For i = 0 To n - 2
For j = 0 To n - i - 2
If arr(j) > arr(j + 1) Then
temp = arr(j)
arr(j) = arr(j + 1)
arr(j + 1) = temp
End If
Next j
Next i
For i = 0 To n - 1
List2.AddItem arr(i)
Next i
End Sub

Private Sub Command4_Click()
List2.Clear
For i = 0 To n - 2
For j = 0 To n - i - 2
If arr(j) < arr(j + 1) Then
temp = arr(j)
arr(j) = arr(j + 1)
arr(j + 1) = temp
End If
Next j
Next i
For i = 0 To n - 1
List2.AddItem arr(i)
Next i
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
