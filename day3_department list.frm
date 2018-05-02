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
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   7320
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   3000
      TabIndex        =   2
      Top             =   3960
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   7200
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Enter department name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   3105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ind As Integer
Private Sub Command1_Click()
list1.AddItem (Text1.Text)
Text1.Text = ""
End Sub

Private Sub Command2_Click()
ind = list1.ListIndex
list1.RemoveItem ind
End Sub

Private Sub Command3_Click()
list1.Clear
End Sub

Private Sub Form_Load()
Command1.Enabled = False
End Sub

Private Sub Text1_Change()
Command1.Enabled = (Len(Text1.Text) > 0)
End Sub
