VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change color"
      Height          =   1215
      Left            =   5400
      TabIndex        =   1
      Top             =   6600
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   10560
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CommonDialog1.ShowColor
Shape1.FillColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
List1.AddItem "Rectangle"
List1.AddItem "Square"
List1.AddItem "Oval"
List1.AddItem "Circle"
List1.AddItem "Rounded Rectangle"
List1.AddItem "Rounded Square"
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
Case 0
Shape1.Shape = 0
Case 1
Shape1.Shape = 1
Case 2
Shape1.Shape = 2
Case 3
Shape1.Shape = 3
Case 4
Shape1.Shape = 4
Case 5
Shape1.Shape = 5
End Select
End Sub
