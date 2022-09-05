VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "isDebug = true"
   ClientHeight    =   6030
   ClientLeft      =   8865
   ClientTop       =   5715
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   10140
   Begin VB.CommandButton Command15 
      Caption         =   "100"
      Height          =   330
      Left            =   2475
      TabIndex        =   80
      Top             =   3015
      Width           =   645
   End
   Begin VB.CommandButton Command14 
      Caption         =   "0"
      Height          =   330
      Left            =   1755
      TabIndex        =   79
      Top             =   3015
      Width           =   600
   End
   Begin VB.CommandButton Command13 
      Caption         =   "200"
      Height          =   330
      Left            =   3195
      TabIndex        =   78
      Top             =   3015
      Width           =   510
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ñòîï"
      Height          =   330
      Left            =   2925
      TabIndex        =   77
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ÐÕ"
      Height          =   330
      Left            =   3870
      TabIndex        =   76
      Top             =   720
      Width           =   960
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   16
      Left            =   8055
      TabIndex        =   73
      Text            =   "0,8"
      Top             =   4590
      Width           =   690
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   16
      Left            =   8820
      TabIndex        =   72
      Top             =   4590
      Width           =   510
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   16
      Left            =   7425
      TabIndex        =   71
      Top             =   4590
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   15
      Left            =   7425
      TabIndex        =   57
      Top             =   4230
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   14
      Left            =   2925
      TabIndex        =   56
      Top             =   315
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   13
      Left            =   4005
      TabIndex        =   55
      Top             =   2610
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   12
      Left            =   7965
      TabIndex        =   54
      Top             =   810
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   11
      Left            =   7965
      TabIndex        =   53
      Top             =   450
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   10
      Left            =   7965
      TabIndex        =   52
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   9
      Left            =   2700
      TabIndex        =   51
      Top             =   4995
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   8
      Left            =   2700
      TabIndex        =   50
      Top             =   4635
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   7
      Left            =   1755
      TabIndex        =   49
      Top             =   3690
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   6
      Left            =   1755
      TabIndex        =   48
      Top             =   2610
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   5
      Left            =   2970
      TabIndex        =   47
      Top             =   1890
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   4
      Left            =   2970
      TabIndex        =   46
      Top             =   1530
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   3
      Left            =   7425
      TabIndex        =   45
      Top             =   1890
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   2
      Left            =   7425
      TabIndex        =   44
      Top             =   1530
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   1
      Left            =   7290
      TabIndex        =   43
      Top             =   5400
      Width           =   600
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   15
      Left            =   8820
      TabIndex        =   42
      Top             =   4230
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   14
      Left            =   4320
      TabIndex        =   41
      Top             =   315
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   13
      Left            =   5400
      TabIndex        =   40
      Top             =   2610
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   12
      Left            =   9360
      TabIndex        =   39
      Top             =   810
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   11
      Left            =   9360
      TabIndex        =   38
      Top             =   450
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   10
      Left            =   9360
      TabIndex        =   37
      Top             =   90
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   9
      Left            =   4230
      TabIndex        =   36
      Top             =   4995
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   8
      Left            =   4230
      TabIndex        =   35
      Top             =   4635
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   7
      Left            =   3195
      TabIndex        =   34
      Top             =   3690
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   6
      Left            =   3195
      TabIndex        =   33
      Top             =   2610
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   5
      Left            =   4410
      TabIndex        =   32
      Top             =   1890
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   4
      Left            =   4410
      TabIndex        =   31
      Top             =   1530
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   3
      Left            =   8865
      TabIndex        =   30
      Top             =   1890
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   2
      Left            =   8865
      TabIndex        =   29
      Top             =   1530
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   1
      Left            =   8820
      TabIndex        =   28
      Top             =   5400
      Width           =   510
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   15
      Left            =   8055
      TabIndex        =   27
      Text            =   "0,8"
      Top             =   4230
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   14
      Left            =   3555
      TabIndex        =   26
      Text            =   "0,8"
      Top             =   315
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   13
      Left            =   4635
      TabIndex        =   25
      Text            =   "0,8"
      Top             =   2610
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   12
      Left            =   8595
      TabIndex        =   24
      Text            =   "0,8"
      Top             =   810
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   11
      Left            =   8595
      TabIndex        =   23
      Text            =   "0,8"
      Top             =   450
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   10
      Left            =   8595
      TabIndex        =   22
      Text            =   "0,8"
      Top             =   90
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   9
      Left            =   3420
      TabIndex        =   21
      Text            =   "2"
      Top             =   4995
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   8
      Left            =   3420
      TabIndex        =   20
      Text            =   "1,7"
      Top             =   4635
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   7
      Left            =   2430
      TabIndex        =   19
      Text            =   "3"
      Top             =   3690
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   6
      Left            =   2430
      TabIndex        =   18
      Text            =   "0,8"
      Top             =   2610
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   5
      Left            =   3600
      TabIndex        =   17
      Text            =   "0,6"
      Top             =   1890
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   4
      Left            =   3600
      TabIndex        =   16
      Text            =   "2"
      Top             =   1530
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   3
      Left            =   8100
      TabIndex        =   15
      Text            =   "0,8"
      Top             =   1890
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   2
      Left            =   8100
      TabIndex        =   14
      Text            =   "1,9"
      Top             =   1530
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   1
      Left            =   8010
      TabIndex        =   13
      Text            =   "0"
      Top             =   5400
      Width           =   690
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   0
      Left            =   7290
      TabIndex        =   12
      Top             =   5040
      Width           =   600
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   0
      Left            =   8820
      TabIndex        =   11
      Top             =   5040
      Width           =   510
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   0
      Left            =   8010
      TabIndex        =   10
      Text            =   "0"
      Top             =   5040
      Width           =   690
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ÂÛÊË"
      Height          =   330
      Left            =   4905
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Cancel          =   -1  'True
      Caption         =   "Ìóôòà ÂÊË"
      Height          =   330
      Left            =   4905
      TabIndex        =   8
      Top             =   315
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
      Caption         =   "Ðó÷íîå / Àâòî óïðàâëåíèå"
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Íàïðÿæåíèå"
      Height          =   195
      Index           =   14
      Left            =   6300
      TabIndex        =   75
      Top             =   4680
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "ÄÄ4"
      Height          =   195
      Index           =   13
      Left            =   6840
      TabIndex        =   74
      Top             =   4275
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Îáîðîòû ÄÂÑ"
      Height          =   195
      Index           =   12
      Left            =   2925
      TabIndex        =   70
      Top             =   45
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÒ4 Âûõîä êîìïðåññîðà"
      Height          =   210
      Index           =   11
      Left            =   4005
      TabIndex        =   69
      Top             =   2340
      Width           =   1995
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "ÄÒ3"
      Height          =   195
      Index           =   10
      Left            =   7380
      TabIndex        =   68
      Top             =   855
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÒ2.1"
      Height          =   195
      Index           =   9
      Left            =   7335
      TabIndex        =   67
      Top             =   540
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÒ2"
      Height          =   195
      Index           =   8
      Left            =   7335
      TabIndex        =   66
      Top             =   180
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÒ1.1 Âõîä"
      Height          =   195
      Index           =   7
      Left            =   1710
      TabIndex        =   65
      Top             =   5040
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÒ1 Âûõîä"
      Height          =   195
      Index           =   6
      Left            =   1710
      TabIndex        =   64
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÄ8 Àêê"
      Height          =   195
      Index           =   5
      Left            =   1755
      TabIndex        =   63
      Top             =   3465
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÄÄ6 Êîìïðåññîð"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   62
      Top             =   2340
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "ÄÄ2.2"
      Height          =   240
      Index           =   3
      Left            =   1845
      TabIndex        =   61
      Top             =   1935
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "ÄÄ2.1 Àâòî"
      Height          =   240
      Index           =   2
      Left            =   1845
      TabIndex        =   60
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ÄÄ1.2"
      Height          =   240
      Index           =   1
      Left            =   6255
      TabIndex        =   59
      Top             =   1890
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "ÄÄ1.1 Âõîä"
      Height          =   240
      Index           =   0
      Left            =   6255
      TabIndex        =   58
      Top             =   1530
      Width           =   1140
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_minus_Click(Index As Integer)
   Dim d As Double
   If (Index = 4) Then
      d = 0.001
   Else
      d = 0.1
   End If
    Text_ggACL8113(Index).Text = CDbl(Text_ggACL8113(Index).Text) - d
    ggACL8113(Index) = CDbl(Text_ggACL8113(Index).Text)
End Sub

Private Sub Command_plus_Click(Index As Integer)
   Dim d As Double
   If (Index = 4) Then
      d = 0.001
   Else
      d = 0.1
   End If
    Text_ggACL8113(Index).Text = CDbl(Text_ggACL8113(Index).Text) + d
    ggACL8113(Index) = CDbl(Text_ggACL8113(Index).Text)
End Sub

Private Sub Command1_Click()
   If isHandControl Then
      ROn B0, 128
   Else
      ROff B0, Not 128
   End If
End Sub

Sub openK1()
   ROn A1, 4 'Îòêðûòü Ê1
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
   ROn A1, 128    'Îòêðûòü ÊÝ6
End Sub
Sub openK7()
   ROn A0, 2
End Sub

Private Sub Command10_Click()
    ' Ìóôòà ÂÊË
    ROn B1, 16
End Sub

Private Sub Command11_Click()
    ' Ìóôòà ÂÛÊË
    ROff B1, Not (16)
End Sub

Private Sub Command12_Click()
    Text_ggACL8113(14) = 0
    updateState
End Sub

Private Sub Command13_Click()
    Text_ggACL8113(6) = 3.3
    updateState
End Sub

Private Sub Command14_Click()
    Text_ggACL8113(6) = 0.8
    updateState
End Sub

Private Sub Command15_Click()
    Text_ggACL8113(6) = 2
    updateState
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
    Text_ggACL8113(14) = 3
    updateState
End Sub

Private Sub Form_Load()
    updateState
End Sub


Private Sub updateState()
    Dim Index As Integer
    For Index = 0 To 16
        ggACL8113(Index) = CDbl(Text_ggACL8113(Index).Text)
    Next
End Sub


Private Sub Label1_Click(Index As Integer)

End Sub
