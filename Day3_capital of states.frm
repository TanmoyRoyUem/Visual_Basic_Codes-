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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   7800
      TabIndex        =   3
      Top             =   3600
      Width           =   5415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Capital of the state"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Choose the state"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text = "West Bengal" Then
Text1.Text = "Kolkata"
ElseIf Combo1.Text = "Bihar" Then
Text1.Text = "Patna"
ElseIf Combo1.Text = "Orissa" Then
Text1.Text = "Bhubaneswar"
ElseIf Combo1.Text = "Jharkhand" Then
Text1.Text = "Ranchi"
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "West Bengal"
Combo1.AddItem "Bihar"
Combo1.AddItem "Orissa"
Combo1.AddItem "Jharkhand"
End Sub
