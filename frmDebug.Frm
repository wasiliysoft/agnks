VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "isDebug = true"
   ClientHeight    =   6030
   ClientLeft      =   8865
   ClientTop       =   5715
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   8955
   Begin VB.CommandButton Command20 
      Caption         =   "Пуст. разн. давл."
      Height          =   330
      Left            =   2745
      TabIndex        =   85
      Top             =   3105
      Width           =   1950
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Выровн. Д.с магистр."
      Height          =   330
      Left            =   2745
      TabIndex        =   84
      Top             =   2745
      Width           =   1950
   End
   Begin VB.CommandButton Command13 
      Caption         =   "200"
      Height          =   330
      Left            =   4185
      TabIndex        =   83
      Top             =   4050
      Width           =   510
   End
   Begin VB.CommandButton Command14 
      Caption         =   "0"
      Height          =   330
      Left            =   2745
      TabIndex        =   82
      Top             =   4050
      Width           =   600
   End
   Begin VB.CommandButton Command15 
      Caption         =   "100"
      Height          =   330
      Left            =   3420
      TabIndex        =   81
      Top             =   4050
      Width           =   645
   End
   Begin VB.CommandButton Command18 
      Caption         =   "200"
      Height          =   330
      Left            =   4185
      TabIndex        =   80
      Top             =   4950
      Width           =   510
   End
   Begin VB.CommandButton Command17 
      Caption         =   "0"
      Height          =   330
      Left            =   2745
      TabIndex        =   79
      Top             =   4950
      Width           =   600
   End
   Begin VB.CommandButton Command16 
      Caption         =   "100"
      Height          =   330
      Left            =   3420
      TabIndex        =   78
      Top             =   4950
      Width           =   645
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Стоп"
      Height          =   330
      Left            =   2925
      TabIndex        =   77
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton Command9 
      Caption         =   "РХ"
      Height          =   330
      Left            =   3870
      TabIndex        =   76
      Top             =   720
      Width           =   960
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   16
      Left            =   7470
      TabIndex        =   73
      Text            =   "0"
      Top             =   4275
      Width           =   690
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   16
      Left            =   8235
      TabIndex        =   72
      Top             =   4275
      Width           =   510
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   16
      Left            =   6840
      TabIndex        =   71
      Top             =   4275
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   15
      Left            =   6840
      TabIndex        =   57
      Top             =   3825
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
      Left            =   6840
      TabIndex        =   55
      Top             =   3060
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   12
      Left            =   6840
      TabIndex        =   54
      Top             =   2430
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   11
      Left            =   6840
      TabIndex        =   53
      Top             =   2025
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   10
      Left            =   6840
      TabIndex        =   52
      Top             =   1665
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   9
      Left            =   6840
      TabIndex        =   51
      Top             =   1260
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   8
      Left            =   2745
      TabIndex        =   50
      Top             =   5490
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   7
      Left            =   2745
      TabIndex        =   49
      Top             =   4545
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   6
      Left            =   2745
      TabIndex        =   48
      Top             =   3690
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   5
      Left            =   2745
      TabIndex        =   47
      Top             =   2385
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   4
      Left            =   2745
      TabIndex        =   46
      Top             =   2025
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   3
      Left            =   2745
      TabIndex        =   45
      Top             =   1620
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   2
      Left            =   2745
      TabIndex        =   44
      Top             =   1260
      Width           =   600
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   1
      Left            =   6840
      TabIndex        =   43
      Top             =   5400
      Width           =   600
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   15
      Left            =   8235
      TabIndex        =   42
      Top             =   3825
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
      Left            =   8235
      TabIndex        =   40
      Top             =   3060
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   12
      Left            =   8235
      TabIndex        =   39
      Top             =   2430
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   11
      Left            =   8235
      TabIndex        =   38
      Top             =   2025
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   10
      Left            =   8235
      TabIndex        =   37
      Top             =   1665
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   9
      Left            =   8235
      TabIndex        =   36
      Top             =   1260
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   8
      Left            =   4185
      TabIndex        =   35
      Top             =   5490
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   7
      Left            =   4185
      TabIndex        =   34
      Top             =   4545
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   6
      Left            =   4185
      TabIndex        =   33
      Top             =   3690
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   5
      Left            =   4185
      TabIndex        =   32
      Top             =   2385
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   4
      Left            =   4185
      TabIndex        =   31
      Top             =   2025
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   3
      Left            =   4185
      TabIndex        =   30
      Top             =   1620
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   2
      Left            =   4185
      TabIndex        =   29
      Top             =   1260
      Width           =   510
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   1
      Left            =   8235
      TabIndex        =   28
      Top             =   5400
      Width           =   510
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   15
      Left            =   7470
      TabIndex        =   27
      Text            =   "0,94"
      Top             =   3825
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
      Left            =   7470
      TabIndex        =   25
      Text            =   "2,57"
      Top             =   3060
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   12
      Left            =   7470
      TabIndex        =   24
      Text            =   "3,9"
      Top             =   2430
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   11
      Left            =   7470
      TabIndex        =   23
      Text            =   "3,8"
      Top             =   2025
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   10
      Left            =   7470
      TabIndex        =   22
      Text            =   "3,4"
      Top             =   1665
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   9
      Left            =   7470
      TabIndex        =   21
      Text            =   "2"
      Top             =   1260
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   8
      Left            =   3420
      TabIndex        =   20
      Text            =   "1,7"
      Top             =   5490
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   7
      Left            =   3420
      TabIndex        =   19
      Text            =   "3"
      Top             =   4545
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   6
      Left            =   3420
      TabIndex        =   18
      Text            =   "0,6"
      Top             =   3690
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   5
      Left            =   3420
      TabIndex        =   17
      Text            =   "3"
      Top             =   2385
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   4
      Left            =   3420
      TabIndex        =   16
      Text            =   "2"
      Top             =   2025
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   3
      Left            =   3420
      TabIndex        =   15
      Text            =   "0,9"
      Top             =   1620
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   2
      Left            =   3420
      TabIndex        =   14
      Text            =   "1,9"
      Top             =   1260
      Width           =   690
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   1
      Left            =   7470
      TabIndex        =   13
      Text            =   "0"
      Top             =   5400
      Width           =   690
   End
   Begin VB.CommandButton Command_minus 
      Caption         =   "-"
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   12
      Top             =   5040
      Width           =   600
   End
   Begin VB.CommandButton Command_plus 
      Caption         =   "+"
      Height          =   330
      Index           =   0
      Left            =   8235
      TabIndex        =   11
      Top             =   5040
      Width           =   510
   End
   Begin VB.TextBox Text_ggACL8113 
      Height          =   330
      Index           =   0
      Left            =   7470
      TabIndex        =   10
      Text            =   "0"
      Top             =   5040
      Width           =   690
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ВЫКЛ"
      Height          =   330
      Left            =   4905
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Cancel          =   -1  'True
      Caption         =   "Муфта ВКЛ"
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
      Caption         =   "Ручное / Авто управление"
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Напряжение"
      Height          =   195
      Index           =   14
      Left            =   5445
      TabIndex        =   75
      Top             =   4365
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "ДД4"
      Height          =   195
      Index           =   13
      Left            =   5445
      TabIndex        =   74
      Top             =   3870
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Обороты ДВС"
      Height          =   195
      Index           =   12
      Left            =   2925
      TabIndex        =   70
      Top             =   45
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДТ4 Вых компр"
      Height          =   195
      Index           =   11
      Left            =   5445
      TabIndex        =   69
      Top             =   3150
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "ДТ3"
      Height          =   195
      Index           =   10
      Left            =   5445
      TabIndex        =   68
      Top             =   2475
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДТ2.1"
      Height          =   195
      Index           =   9
      Left            =   5445
      TabIndex        =   67
      Top             =   2115
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДТ2"
      Height          =   195
      Index           =   8
      Left            =   5445
      TabIndex        =   66
      Top             =   1755
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДТ1.1 Вход"
      Height          =   195
      Index           =   7
      Left            =   5445
      TabIndex        =   65
      Top             =   1305
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДТ1 Выход"
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   64
      Top             =   5580
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДД8 Акк"
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   63
      Top             =   4590
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ДД6 Компр"
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   62
      Top             =   3825
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "ДД2.2 Магистр"
      Height          =   240
      Index           =   3
      Left            =   1440
      TabIndex        =   61
      Top             =   2475
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "ДД2.1 Авто"
      Height          =   240
      Index           =   2
      Left            =   1440
      TabIndex        =   60
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ДД1.2"
      Height          =   240
      Index           =   1
      Left            =   1440
      TabIndex        =   59
      Top             =   1665
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "ДД1.1 Вход"
      Height          =   240
      Index           =   0
      Left            =   1440
      TabIndex        =   58
      Top             =   1305
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
    Text_ggACL8113(Index).Text = CDbl(Text_ggACL8113(Index).Text) - 0.1
    ggACL8113(Index) = CDbl(Text_ggACL8113(Index).Text)
End Sub

Private Sub Command_plus_Click(Index As Integer)
    Text_ggACL8113(Index).Text = CDbl(Text_ggACL8113(Index).Text) + 0.1
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
   ROn A1, 4 'Открыть К1
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
   ROn A1, 128    'Открыть КЭ6
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

Private Sub Command12_Click()
    Text_ggACL8113(14) = 0
    updateState
End Sub

Private Sub Command13_Click()
    Text_ggACL8113(6) = 3.4
    updateState
End Sub

Private Sub Command14_Click()
    Text_ggACL8113(6) = 0.6
    updateState
End Sub

Private Sub Command15_Click()
    Text_ggACL8113(6) = 2.1
    updateState
End Sub

Private Sub Command16_Click()
    Text_ggACL8113(7) = 2.1
    updateState
End Sub

Private Sub Command17_Click()
    Text_ggACL8113(7) = 0.6
    updateState
End Sub

Private Sub Command18_Click()
    Text_ggACL8113(7) = 3.4
    updateState
End Sub

Private Sub Command19_Click()
    Text_ggACL8113(4) = CDbl(Text_ggACL8113(5))
    updateState
End Sub

Private Sub Command2_Click()
    openK1
End Sub

Private Sub Command20_Click()
    Text_ggACL8113(4) = CDbl(Text_ggACL8113(5)) - 1
    If (CDbl(Text_ggACL8113(4)) < 0.6) Then
        Text_ggACL8113(4) = 0.6
    End If
    updateState
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


