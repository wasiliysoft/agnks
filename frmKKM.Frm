VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmKKM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   2445
   ClientTop       =   810
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel sspKKM 
      Height          =   6135
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   5925
      _Version        =   65536
      _ExtentX        =   10451
      _ExtentY        =   10821
      _StockProps     =   15
      Caption         =   "ОСНОВНЫЕ  КАССОВЫЕ  ОПЕРАЦИИ "
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   6
      Begin VB.Frame Frame1 
         Caption         =   "СВОЙСТВА ККМ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   180
         TabIndex        =   23
         Top             =   3800
         Width           =   5640
         Begin VB.Label lblErrorKKM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1000
            TabIndex        =   31
            Top             =   1230
            Width           =   4500
         End
         Begin VB.Label lblStatusKKM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1000
            TabIndex        =   30
            Top             =   930
            Width           =   4500
         End
         Begin VB.Label lblTimeKKM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1000
            TabIndex        =   29
            Top             =   630
            Width           =   4500
         End
         Begin VB.Label lblDateKKM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1000
            TabIndex        =   28
            Top             =   315
            Width           =   4500
         End
         Begin VB.Label Label4 
            Caption         =   "Ошибки :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   27
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Режим :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   26
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Время :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   25
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Дата :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   24
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdExitKKM 
         Cancel          =   -1  'True
         Caption         =   "ВЫХОД"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4905
         TabIndex        =   22
         Top             =   5400
         Width           =   930
      End
      Begin VB.TextBox txtKKM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   450
         Width           =   2760
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   645
         Index           =   6
         Left            =   2610
         TabIndex        =   34
         Top             =   5400
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ОТЧЕТ БЕЗ ГАШЕНИЯ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   645
         Index           =   5
         Left            =   180
         TabIndex        =   33
         Top             =   5400
         Width           =   2400
         _Version        =   65536
         _ExtentX        =   4233
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ОТЧЕТ С ГАШЕНИЕМ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   645
         Index           =   4
         Left            =   3105
         TabIndex        =   32
         Top             =   3150
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ДЕНЕЖНЫЙ ЯЩИК"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   15
         Left            =   1980
         TabIndex        =   21
         Top             =   2460
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   650
         Index           =   3
         Left            =   3105
         TabIndex        =   20
         Top             =   1130
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ПРОДОЛЖИТЬ ПЕЧАТЬ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   650
         Index           =   2
         Left            =   3105
         TabIndex        =   19
         Top             =   2480
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ОТРЕЗАТЬ ЧЕК"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   650
         Index           =   1
         Left            =   3105
         TabIndex        =   18
         Top             =   1800
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "АНУЛИРОВАТЬ ЧЕК"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdKKM 
         Height          =   650
         Index           =   0
         Left            =   3105
         TabIndex        =   17
         Top             =   450
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "ПРОДАЖА НА СУММУ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   14
         Left            =   1980
         TabIndex        =   16
         Top             =   3060
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   13
         Left            =   1380
         TabIndex        =   15
         Top             =   3060
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "С"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   12
         Left            =   1980
         TabIndex        =   14
         Top             =   1860
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   11
         Left            =   1980
         TabIndex        =   13
         Top             =   1260
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "x"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   3060
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   1260
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   2
         Left            =   780
         TabIndex        =   9
         Top             =   1260
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   3
         Left            =   1380
         TabIndex        =   8
         Top             =   1260
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   1860
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   5
         Left            =   780
         TabIndex        =   6
         Top             =   1860
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   6
         Left            =   1380
         TabIndex        =   5
         Top             =   1860
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   7
         Left            =   180
         TabIndex        =   4
         Top             =   2460
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   8
         Left            =   780
         TabIndex        =   3
         Top             =   2460
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   9
         Left            =   1380
         TabIndex        =   2
         Top             =   2460
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCalc 
         Height          =   600
         Index           =   10
         Left            =   780
         TabIndex        =   1
         Top             =   3060
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmKKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strKKM          As String    'строковое значение выражения
Dim dKKM            As Double    'числовое значение
Dim bpoint          As Boolean    'флаг точки
Dim dA              As Double    'первое число
Dim dB              As Double    'второе число
Dim dC              As Double    'результат
Dim Process         As Byte    'действие:(+ 10);(* 20);(- 30)(= 100);
'Калькулятор выполняет два основных действия сложение и умножение
Private Sub cmdCalc_Click(Index As Integer)
    Select Case Index
        Case 10    'если точки еще не было, то поднимаем флаг и добавляем справа точку
            If bpoint = False Then
                bpoint = True
                frmKKM.txtKKM.Text = strKKM & ","
                strKKM = frmKKM.txtKKM.Text
            End If
        Case 11    'умножение
            dB = dKKM
            Select Case Process
                Case 10
                    dKKM = dA + dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 20
                    dKKM = dA * dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 30
                    dKKM = dA - dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case Else
            End Select

            dA = dKKM
            Process = 20
            '    dKKM = 0
            bpoint = False
            strKKM = "0"
        Case 12    'сложение
            dB = dKKM
            Select Case Process
                Case 10
                    dKKM = dA + dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 20
                    dKKM = dA * dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 30
                    dKKM = dA - dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case Else
            End Select

            dA = dKKM
            Process = 10
            '    dKKM = 0
            bpoint = False
            strKKM = "0"
        Case 13    'сброс
            dA = 0
            Process = 0
            dKKM = 0
            bpoint = False
            strKKM = "0"
            frmKKM.txtKKM.Text = "0"
        Case 14    'равно - выполнение процесса
            dB = dKKM
            Select Case Process
                Case 10
                    dKKM = dA + dB
                Case 20
                    dKKM = dA * dB
                Case 30
                    dKKM = dA - dB
                Case Else
            End Select
            '    dA = 0
            '    dB = 0
            Process = 0
            bpoint = False
            strKKM = "0"
            frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
        Case 15
            dB = dKKM
            Select Case Process
                Case 10
                    dKKM = dA + dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 20
                    dKKM = dA * dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case 30
                    dKKM = dA - dB
                    frmKKM.txtKKM.Text = Format(dKKM, "###0.00")
                Case Else
            End Select

            dA = dKKM
            Process = 30
            '    dKKM = 0
            bpoint = False
            strKKM = "0"
        Case Else
            'If bpoint = False Then
            '    dKKM = dKKM * 10 + Index
            '    frmKKM.txtKKM.Text = CStr(dKKM)
            '    strKKM = frmKKM.txtKKM.Text
            'Else
            If strKKM = "0" Then
                frmKKM.txtKKM.Text = CStr(Index)
            Else
                frmKKM.txtKKM.Text = strKKM & CStr(Index)
            End If
            dKKM = CDbl(frmKKM.txtKKM.Text)
            strKKM = frmKKM.txtKKM.Text
            'End If
    End Select

End Sub





Private Sub cmdExitKKM_Click()
    dA = 0
    dB = 0
    dC = 0
    Process = 0
    dKKM = 0
    bpoint = False
    strKKM = "0"
    frmKKM.txtKKM.Text = "0"

    frmKKM.Hide
End Sub

Private Sub cmdKKM_Click(Index As Integer)
    Dim i           As Integer
    Dim t           As Date


    On Error GoTo err

    Select Case Index
        Case 0    'продажа
            'проверка правильного ввода денег
            If (frmKKM.txtKKM.Visible = True) And (CDbl(frmKKM.txtKKM.Text) > 0) Then
                Dim dblPrice As Double: dblPrice = agnksConfig.Price
                Dim dblSumma As Double: dblSumma = CDbl(frmKKM.txtKKM.Text)
                ' Количество с точностью до 1 знака после запятой,
                ' см. как рассчитывается первоначальная основная сумма
                Dim dblQuantity As Double: dblQuantity = CDbl(Format(dblSumma / dblPrice, "0.0"))
                Dim summ1 As Double: summ1 = CDbl(Format(dblQuantity * dblPrice, "0.00"))
                
                If (vbYes = MsgBox("Количество:  " & dblQuantity & Chr(13) _
                            & "Цена:  " & dblPrice & Chr(13) _
                            & "Сумма:  " & summ1 & Chr(13) & Chr(13) _
                            & "Продолжить?", vbYesNo + vbQuestion + vbDefaultButton2, "ККМ")) Then
                
                    'Оформление чека продажи
                    DrvFR.Password = pass
                    DrvFR.Quantity = dblQuantity
                    DrvFR.Price = dblPrice
                    DrvFR.Department = 1
                    DrvFR.Tax1 = 1
                    DrvFR.Tax2 = 0
                    DrvFR.Tax3 = 0
                    DrvFR.Tax4 = 0
                    DrvFR.StringForPrinting = "Газ природный"
                    DrvFR.Sale
                    'Закрытие чека с печатью(отрезка по умалчанию выставляется в таблице свойств ККМ)
                    DrvFR.Password = pass
                    DrvFR.CheckSubTotal 'Подводим итог чека, т.е. в Summ1 заносим все Sale
                    DrvFR.Password = pass
                    DrvFR.Summ1 = summ1
                    DrvFR.Summ2 = 0
                    DrvFR.Summ3 = 0
                    DrvFR.Summ4 = 0
                    DrvFR.DiscountOnCheck = 0 'Скидок нет
                    DrvFR.Tax1 = 1
                    DrvFR.Tax2 = 0
                    DrvFR.Tax3 = 0
                    DrvFR.Tax4 = 0
                    DrvFR.StringForPrinting = "===================================="
                    DrvFR.CloseCheck
                    giErrorKKM = DrvFR.ResultCode
                    gsErrorKKM = DrvFR.ResultCodeDescription

                End If
            End If

        Case 1    'продолжить печать после заправки бумаги
            DrvFR.Password = pass
            '    DrvFR.ContinuePrint 'Переводит ККМ из подрежима 3 в другой
            DrvFR.CancelCheck
        Case 2    'отрезать чек
            DrvFR.Password = pass
            DrvFR.CutType = True
            DrvFR.CutCheck

        Case 3    'ПРОДОЛЖИТЬ ПЕЧАТЬ 'возврат продажи
            DrvFR.Password = pass
            DrvFR.ContinuePrint    'Переводит ККМ из подрежима 3 в другой

            '    'проверка правильного ввода денег
            '    If (frmKKM.txtKKM.Visible = True) And (CDbl(frmKKM.txtKKM.Text) > 0) Then
            '        'Оформление чека продажи
            '        DrvFR.Password = pass
            '        DrvFR.Quantity = CDbl(frmKKM.txtKKM.Text)
            '        DrvFR.price = 1
            '        DrvFR.Department = 1
            '        DrvFR.Tax1 = 1
            '        DrvFR.Tax2 = 0
            '        DrvFR.Tax3 = 0
            '        DrvFR.Tax4 = 0
            '        DrvFR.StringForPrinting = "Газ природный"
            '        DrvFR.returnSale
            '        'Закрытие чека с печатью(отрезка по умалчанию выставляется в таблице свойств ККМ)
            '        DrvFR.Password = pass
            '        DrvFR.CheckSubTotal 'Подводим итог чека, т.е. в Summ1 заносим все Sale
            '        DrvFR.Password = pass
            '        DrvFR.Summ1 = CDbl(frmKKM.txtKKM.Text)
            '        DrvFR.Summ2 = 0
            '        DrvFR.Summ3 = 0
            '        DrvFR.Summ4 = 0
            '        DrvFR.DiscountOnCheck = 0 'Скидок нет
            '        DrvFR.Tax1 = 1
            '        DrvFR.Tax2 = 0
            '        DrvFR.Tax3 = 0
            '        DrvFR.Tax4 = 0
            '        DrvFR.StringForPrinting = "===================================="
            '        DrvFR.CloseCheck
            '        giErrorKKM = DrvFR.ResultCode
            '        gsErrorKKM = DrvFR.ResultCodeDescription
            ''        If gsРежимККМ <> "" Then
            ''           str = gsРежимККМ & "_подрежим " & CStr(giПодРежимККМ) & ":Ошибки " & CStr(giErrorKKM) & "_" & gsErrorKKM
            ''        Else
            ''           str = CStr(giErrorKKM) & "_" & gsErrorKKM
            ''        End If
            '        'Если процедура печати безошибочна, то снимаем флаг печати иначе оставляем флаг
            '        If (giErrorKKM = 0) Then ''And glOperator <> 0 Then
            ''            BaseZaprAdd CDbl(frmKKM.txtKKM.Text), 0, 1, glOperator, 4, 2, 2, 0, 0 'Перед печатью чека запись в базе:на сумму указанную в чеке
            ''           ZaprRS.AddNew
            ''           ZaprRS("DATE") = Now
            ''           ZaprRS("GAS") = CDbl(frmKKM.txtKKM.Text)
            ''           ZaprRS("NUMBER") = 10
            ''           ZaprRS("PRICE") = 1
            ''           ZaprRS("OPERATOR") = glOperator
            ''           ZaprRS("TYPEZAPR") = 4 'TYPEZAPR чек по возврату продажи без отпуска газа
            ''           ZaprRS("TYPEFINISH") = 2 'заправка на сумму
            ''           ZaprDN.Update
            '        Else
            '        '!!!надо бы проанализировать ошибку и принять соответствующие действия!!!
            '        End If
            '    End If
        Case 4
            DrvFR.Password = pass
            DrvFR.DrawerNumber = 0
            DrvFR.OpenDrawer
        Case 5
            DrvFR.GetECRStatus
            DrvFR.Password = txtKKM.Text    'pass
            DrvFR.PrintReportWithCleaning

            '    If DrvFR.ResultCode = 0 Then
            '        t = DrvFR.Time
            '        GetLocalTime sys_sync
            '        With sys_sync
            '           .wHour = Hour(t)
            '           .wMinute = Minute(t)
            '           .wSecond = Second(t)
            '        End With
            '        SetLocalTime sys_sync
            '    End If
            StatusKKM
        Case 6
            DrvFR.Password = txtKKM.Text    'pass
            DrvFR.PrintReportWithoutCleaning
            StatusKKM
        Case Else

    End Select

err:

End Sub

Private Sub Form_Load()
    Dim i           As Integer

    pass = 1
    dA = 0
    dB = 0
    dC = 0
    Process = 0
    dKKM = 0
    strKKM = "0"
    bpoint = False

    With frmKKM

        .Left = 2385    '6000
        .Top = 450    '1700
        .Height = 6300
        .Width = 6000
        .sspKKM.Left = 50
        .sspKKM.Top = 50
        .sspKKM.Height = 6200
        .sspKKM.Width = 5900

        '    For i = 0 To 10
        '      .cmdCalc(i).Height = 600
        '      .cmdCalc(i).Width = 600
        '    Next i
        '    For i = 1 To 7 Step 3
        '      .cmdCalc(i).Left = 200
        '    Next i
        '    For i = 2 To 8 Step 3
        '      .cmdCalc(i).Left = 200 + .cmdCalc(1).Width
        '    Next i
        '    For i = 3 To 9 Step 3
        '      .cmdCalc(i).Left = 200 + .cmdCalc(1).Width * 2
        '    Next i
        '    For i = 1 To 3
        '      .cmdCalc(i).Top = 1200
        '    Next i
        '    For i = 4 To 6
        '      .cmdCalc(i).Top = 1200 + .cmdCalc(1).Width
        '    Next i
        '    For i = 7 To 9
        '      .cmdCalc(i).Top = 1200 + .cmdCalc(1).Width * 2
        '    Next i
        '
        '    .cmdCalc(0).Left = .cmdCalc(1).Left
        '    .cmdCalc(0).Width = .cmdCalc(1).Width * 2
        '    .cmdCalc(0).Top = .cmdCalc(7).Top + .cmdCalc(1).Height
        '    .cmdCalc(10).Left = .cmdCalc(3).Left
        '    .cmdCalc(10).Width = .cmdCalc(1).Width
        '    .cmdCalc(10).Top = .cmdCalc(0).Top
    End With

End Sub



