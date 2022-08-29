VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "isDebug = true"
   ClientHeight    =   6030
   ClientLeft      =   1335
   ClientTop       =   2415
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   10140
   Begin VB.CommandButton Command11 
      Caption         =   "ВЫКЛ"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Cancel          =   -1  'True
      Caption         =   "Муфта ВКЛ"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Обороты ДВС"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Open K7"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Open K6"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open K5"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open K4"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open K3"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open K2"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open K1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ручное / Авто управление"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   ROn B0, 128
   'gnДатчик(15).Data = Not gnДатчик(15).Data
End Sub

Sub openK1()
   ROn A1, 4
End Sub
Sub openK2()
   ROn A1, 8
End Sub
Sub openK3()
   ROn A1, 16
End Sub
Sub openK4()
   ROn A1, 32
End Sub
Sub openK5()
   ROn A1, 64
End Sub
Sub openK6()
   ROn A1, 128
End Sub
Sub openK7()
   ROn A0, 2
End Sub

Private Sub Command10_Click()
    ' Муфта ВКЛ
    ROn B1, 16
End Sub

Private Sub Command11_Click()
    ' Муфта ВЫКЛ
    ROff B1, Not (16)
End Sub

Private Sub Command2_Click()
    openK1
End Sub

Private Sub Command3_Click()
    openK2
End Sub

Private Sub Command4_Click()
    openK3
End Sub

Private Sub Command5_Click()
openK4
End Sub

Private Sub Command6_Click()
openK5
End Sub

Private Sub Command7_Click()
openK6
End Sub

Private Sub Command8_Click()
openK7
End Sub

Private Sub Command9_Click()
   ggACL8113(14) = Val(Text1)
End Sub
