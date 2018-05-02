VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7320
      TabIndex        =   8
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   11280
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Non Veg"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Veg"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "PIZZA HUT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As String

Private Sub Command1_Click()
MsgBox ("You have ordered a" + a + b + "pizza")
End Sub

Private Sub Option1_Click()
a = " small "
End Sub

Private Sub Option2_Click()
a = " medium "
End Sub

Private Sub Option3_Click()
a = " large "
End Sub

Private Sub Option4_Click()
b = " veg "
End Sub

Private Sub Option5_Click()
b = " non veg "
End Sub
