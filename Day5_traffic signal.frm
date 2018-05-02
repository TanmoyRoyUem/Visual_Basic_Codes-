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
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   12360
      Top             =   2880
   End
   Begin VB.Label Label3 
      Height          =   1215
      Left            =   6120
      TabIndex        =   2
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   6000
      TabIndex        =   1
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   6495
      Left            =   1800
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Static state As Integer
Select Case state
Case 0
Shape2.FillColor = vbRed
Shape3.FillColor = vbWhite
Shape4.FillColor = vbWhite
Label1.Caption = "STOP"
Label1.ForeColor = vbRed
Timer1.Interval = 5000
state = 1
Case 1
Shape2.FillColor = vbWhite
Shape3.FillColor = vbYellow
Shape4.FillColor = vbWhite
Label2.Caption = "WAIT"
Label1.Caption = ""
Label2.ForeColor = vbYellow
Timer1.Interval = 5000
state = 2
Case 2
Shape2.FillColor = vbWhite
Shape3.FillColor = vbWhite
Shape4.FillColor = vbGreen
Label3.Caption = "GO"
Label2.Caption = ""
Label1.Caption = ""
Label3.ForeColor = vbGreen
Timer1.Interval = 5000
End Select
End Sub
