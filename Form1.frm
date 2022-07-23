VERSION 4.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "АГНКС   БИ-40  ""МЕТАН"""
   ClientHeight    =   5745
   ClientLeft      =   -675
   ClientTop       =   165
   ClientWidth     =   9450
   ControlBox      =   0   'False
   Height          =   6150
   KeyPreview      =   -1  'True
   Left            =   -735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9450
   Top             =   -180
   Visible         =   0   'False
   Width           =   9570
   Begin VB.CommandButton cmdDanger 
      BackColor       =   &H000000FF&
      Caption         =   "АВАРИЯ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   4395
      TabIndex        =   177
      Top             =   2490
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Timer Timer_Газ 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8292
      Top             =   108
   End
   Begin VB.Timer Timer_ДВС 
      Interval        =   75
      Left            =   7920
      Top             =   96
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9000
      Top             =   1020
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      _Version        =   65536
      _ExtentX        =   16642
      _ExtentY        =   10186
      _StockProps     =   15
      Caption         =   "О программе"
      TabsPerRow      =   5
      Tab             =   2
      TabOrientation  =   0
      Tabs            =   5
      Style           =   0
      TabMaxWidth     =   0
      TabHeight       =   529
      TabCaption(0)   =   "Дискретные"
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      TabCaption(1)   =   "Аналоговые"
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      TabCaption(2)   =   "О программе"
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1(2)"
      TabCaption(3)   =   "Схема"
      Tab(3).ControlCount=   3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(1)=   "tmrTablo"
      Tab(3).Control(2)=   "tmrMotor"
      TabCaption(4)   =   "Журнал"
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(4)"
      Begin VB.Timer tmrMotor 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   -69270
         Top             =   150
      End
      Begin VB.Timer tmrTablo 
         Interval        =   500
         Left            =   -68100
         Top             =   180
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   3
         Left            =   -75000
         TabIndex        =   141
         Top             =   360
         Width           =   9345
         Begin Threed.SSPanel SSPanel1 
            Height          =   5340
            Left            =   0
            TabIndex        =   142
            Top             =   75
            Width           =   9405
            _Version        =   65536
            _ExtentX        =   16595
            _ExtentY        =   9419
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            BevelOuter      =   0
            BevelInner      =   1
            Begin VB.TextBox txtKg 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FF00&
               Height          =   285
               Left            =   5640
               TabIndex        =   185
               Text            =   "9999.99"
               Top             =   4800
               Width           =   855
            End
            Begin VB.TextBox txtTime 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   288
               Left            =   5295
               TabIndex        =   178
               Text            =   "0"
               Top             =   3855
               Width           =   720
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   840
               Left            =   120
               TabIndex        =   143
               Top             =   3000
               Width           =   9135
               _Version        =   65536
               _ExtentX        =   16108
               _ExtentY        =   1482
               _StockProps     =   15
               BackColor       =   12632256
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   2
               BorderWidth     =   2
               BevelOuter      =   1
               BevelInner      =   1
               Autosize        =   3
               Begin VB.Label ОкноСообщений 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   660
                  Left            =   90
                  TabIndex        =   188
                  Top             =   90
                  Width           =   8955
               End
            End
            Begin Threed.SSPanel ЗаправленоГаза 
               Height          =   915
               Left            =   3750
               TabIndex        =   144
               Top             =   4170
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3196
               _ExtentY        =   1609
               _StockProps     =   15
               Caption         =   "0"
               ForeColor       =   16776960
               BackColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   27
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   8
               BorderWidth     =   4
               BevelOuter      =   1
            End
            Begin Threed.SSPanel Отсек_ДВС 
               Height          =   1455
               Left            =   1515
               TabIndex        =   145
               Top             =   690
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   2561
               _StockProps     =   15
               ForeColor       =   16711680
               BackColor       =   12632256
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               Begin Threed.SSPanel ОборотыДВС 
                  Height          =   375
                  Left            =   255
                  TabIndex        =   146
                  Top             =   120
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1503
                  _ExtentY        =   656
                  _StockProps     =   15
                  Caption         =   "0"
                  ForeColor       =   65280
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   4
                  BevelOuter      =   1
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Двигатель"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   204
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Index           =   0
                  Left            =   132
                  TabIndex        =   147
                  Top             =   1116
                  Width           =   1092
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   0
                  Left            =   450
                  Picture         =   "Form1.frx":0000
                  Top             =   585
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   1
                  Left            =   450
                  Picture         =   "Form1.frx":03A2
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   2
                  Left            =   450
                  Picture         =   "Form1.frx":0744
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   3
                  Left            =   450
                  Picture         =   "Form1.frx":0AE6
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   4
                  Left            =   450
                  Picture         =   "Form1.frx":0E88
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   5
                  Left            =   450
                  Picture         =   "Form1.frx":122A
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image Температура_ДВС 
                  Height          =   480
                  Left            =   120
                  Picture         =   "Form1.frx":15CC
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   300
               End
            End
            Begin Threed.SSPanel Отсек_компр 
               Height          =   1455
               Left            =   3000
               TabIndex        =   148
               Top             =   690
               Width           =   1470
               _Version        =   65536
               _ExtentX        =   2582
               _ExtentY        =   2561
               _StockProps     =   15
               ForeColor       =   16711680
               BackColor       =   12632256
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               Begin Threed.SSPanel Р_выход_компр 
                  Height          =   375
                  Left            =   420
                  TabIndex        =   149
                  Top             =   150
                  Width           =   675
                  _Version        =   65536
                  _ExtentX        =   1185
                  _ExtentY        =   656
                  _StockProps     =   15
                  Caption         =   "0"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   4
                  BevelOuter      =   1
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Компрессор"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   204
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Index           =   1
                  Left            =   132
                  TabIndex        =   150
                  Top             =   1092
                  Width           =   1212
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   0
                  Left            =   375
                  Picture         =   "Form1.frx":17CE
                  Top             =   570
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   1
                  Left            =   375
                  Picture         =   "Form1.frx":1D50
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   2
                  Left            =   375
                  Picture         =   "Form1.frx":22D2
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   3
                  Left            =   375
                  Picture         =   "Form1.frx":2854
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   4
                  Left            =   375
                  Picture         =   "Form1.frx":2DD6
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   5
                  Left            =   375
                  Picture         =   "Form1.frx":3358
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   900
               End
            End
            Begin Threed.SSPanel Панель_Авто 
               Height          =   1755
               Left            =   7230
               TabIndex        =   151
               Top             =   870
               Visible         =   0   'False
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   3090
               _StockProps     =   15
               BackColor       =   12632256
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               Begin Threed.SSPanel Р_автобаллон 
                  Height          =   375
                  Left            =   780
                  TabIndex        =   152
                  Top             =   105
                  Width           =   960
                  _Version        =   65536
                  _ExtentX        =   1693
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "154"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   4
                  BevelOuter      =   1
               End
               Begin Threed.SSPanel Автобаллон 
                  Height          =   1515
                  Left            =   120
                  TabIndex        =   153
                  Top             =   120
                  Width           =   390
                  _Version        =   65536
                  _ExtentX        =   699
                  _ExtentY        =   2667
                  _StockProps     =   15
                  Caption         =   "SSPanel7"
                  ForeColor       =   16711680
                  BackColor       =   12632256
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   3
                  BevelOuter      =   1
                  FloodType       =   4
                  FloodColor      =   16776960
               End
               Begin Threed.SSCommand cmdStop 
                  Height          =   1170
                  Left            =   510
                  TabIndex        =   180
                  Top             =   495
                  Width           =   1485
                  _Version        =   65536
                  _ExtentX        =   2619
                  _ExtentY        =   2064
                  _StockProps     =   78
                  Caption         =   "STOP"
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   7
                  Font3D          =   4
                  Picture         =   "Form1.frx":38DA
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   1755
               Index           =   2
               Left            =   5430
               TabIndex        =   154
               Top             =   885
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2836
               _ExtentY        =   3090
               _StockProps     =   15
               BackColor       =   12632256
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               Begin Threed.SSPanel Р_аккумулятор 
                  Height          =   375
                  Left            =   615
                  TabIndex        =   155
                  Top             =   930
                  Width           =   840
                  _Version        =   65536
                  _ExtentX        =   1482
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "178"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   4
                  BevelOuter      =   1
               End
               Begin Threed.SSPanel Аккумулятор 
                  Height          =   1485
                  Left            =   120
                  TabIndex        =   156
                  Top             =   150
                  Width           =   390
                  _Version        =   65536
                  _ExtentX        =   677
                  _ExtentY        =   2625
                  _StockProps     =   15
                  Caption         =   "SSPanel7"
                  ForeColor       =   16711680
                  BackColor       =   12632256
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   3
                  BevelOuter      =   1
                  FloodType       =   4
                  FloodColor      =   16776960
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   492
                  Index           =   0
                  Left            =   636
                  Shape           =   4  'Rounded Rectangle
                  Top             =   348
                  Width           =   132
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   492
                  Index           =   1
                  Left            =   816
                  Shape           =   4  'Rounded Rectangle
                  Top             =   348
                  Width           =   132
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   492
                  Index           =   2
                  Left            =   996
                  Shape           =   4  'Rounded Rectangle
                  Top             =   348
                  Width           =   132
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   492
                  Index           =   3
                  Left            =   1176
                  Shape           =   4  'Rounded Rectangle
                  Top             =   348
                  Width           =   132
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   492
                  Index           =   4
                  Left            =   1356
                  Shape           =   4  'Rounded Rectangle
                  Top             =   348
                  Width           =   132
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   696
                  X2              =   696
                  Y1              =   348
                  Y2              =   168
               End
               Begin VB.Line Line2 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   576
                  X2              =   1416
                  Y1              =   168
                  Y2              =   168
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   876
                  X2              =   876
                  Y1              =   348
                  Y2              =   168
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   1056
                  X2              =   1056
                  Y1              =   348
                  Y2              =   168
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   1236
                  X2              =   1236
                  Y1              =   348
                  Y2              =   168
               End
               Begin VB.Line Line6 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   1416
                  X2              =   1416
                  Y1              =   348
                  Y2              =   168
               End
            End
            Begin Threed.SSPanel Наработка_ДВС 
               Height          =   330
               Left            =   1545
               TabIndex        =   157
               Top             =   285
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1503
               _ExtentY        =   572
               _StockProps     =   15
               Caption         =   "12999"
               ForeColor       =   65280
               BackColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   4
               BevelOuter      =   1
            End
            Begin Threed.SSPanel Наработка_компр 
               Height          =   330
               Left            =   3090
               TabIndex        =   158
               Top             =   315
               Visible         =   0   'False
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1503
               _ExtentY        =   572
               _StockProps     =   15
               Caption         =   "12999"
               ForeColor       =   65280
               BackColor       =   8421504
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   4
               BevelOuter      =   1
            End
            Begin Threed.SSCommand cmdKKM 
               Height          =   495
               Left            =   7200
               TabIndex        =   190
               Top             =   2520
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
               _ExtentY        =   873
               _StockProps     =   78
               Caption         =   "KKM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   5
               Font3D          =   1
            End
            Begin VB.Label Label5 
               Caption         =   "кг"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6480
               TabIndex        =   186
               Top             =   4800
               Width           =   255
            End
            Begin VB.Label Label3 
               Caption         =   "мин."
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   6120
               TabIndex        =   183
               Top             =   3960
               Width           =   492
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   15
               X1              =   4350
               X2              =   4035
               Y1              =   2820
               Y2              =   2805
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   14
               X1              =   4050
               X2              =   4050
               Y1              =   2595
               Y2              =   2820
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   13
               X1              =   5175
               X2              =   4590
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   7
               X1              =   5148
               X2              =   4428
               Y1              =   2004
               Y2              =   2016
            End
            Begin VB.Image Image3 
               Height          =   435
               Left            =   4425
               Picture         =   "Form1.frx":64AC
               Stretch         =   -1  'True
               Top             =   2130
               Width           =   375
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   4350
               Picture         =   "Form1.frx":6D76
               Top             =   2445
               Width           =   480
            End
            Begin MSCommLib.MSComm MSComm1 
               Left            =   3360
               Top             =   4200
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               DTREnable       =   -1  'True
               ParitySetting   =   1
            End
            Begin VB.Label Label9 
               Caption         =   "Время заправки :"
               Height          =   225
               Left            =   3585
               TabIndex        =   179
               Top             =   3900
               Width           =   1650
            End
            Begin VB.Label Р_вход_АГНКС 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5.7"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   264
               Left            =   144
               TabIndex        =   159
               Top             =   2676
               Width           =   492
            End
            Begin VB.Line lnZar 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Visible         =   0   'False
               X1              =   8385
               X2              =   8595
               Y1              =   225
               Y2              =   240
            End
            Begin VB.Image imgZaryad 
               Height          =   480
               Left            =   7995
               Picture         =   "Form1.frx":7640
               Top             =   120
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image imgAkkum 
               Height          =   480
               Index           =   0
               Left            =   8535
               Picture         =   "Form1.frx":794A
               Top             =   120
               Width           =   480
            End
            Begin VB.Label lblV 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Caption         =   "24 В"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8175
               TabIndex        =   176
               Top             =   600
               Width           =   690
            End
            Begin VB.Image КЭ2 
               Height          =   480
               Index           =   0
               Left            =   870
               Picture         =   "Form1.frx":7C54
               Top             =   1785
               Width           =   480
            End
            Begin VB.Image КЭ2 
               Height          =   480
               Index           =   1
               Left            =   870
               Picture         =   "Form1.frx":7F5E
               Top             =   1785
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   6
               X1              =   1116
               X2              =   1116
               Y1              =   1656
               Y2              =   2556
            End
            Begin VB.Image КЭ1 
               Height          =   480
               Index           =   0
               Left            =   600
               Picture         =   "Form1.frx":8268
               Top             =   2355
               Width           =   480
            End
            Begin VB.Image КЭ1 
               Height          =   480
               Index           =   1
               Left            =   600
               Picture         =   "Form1.frx":8572
               Top             =   2355
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ5 
               Height          =   480
               Index           =   0
               Left            =   7230
               Picture         =   "Form1.frx":887C
               Top             =   405
               Width           =   480
            End
            Begin VB.Image КЭ6 
               Height          =   480
               Index           =   0
               Left            =   5520
               Picture         =   "Form1.frx":8B86
               Top             =   405
               Width           =   480
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   552
               Index           =   1
               Left            =   4656
               Shape           =   4  'Rounded Rectangle
               Top             =   948
               Width           =   156
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   552
               Index           =   0
               Left            =   4872
               Shape           =   4  'Rounded Rectangle
               Top             =   960
               Visible         =   0   'False
               Width           =   156
            End
            Begin VB.Image КЭ3 
               Height          =   480
               Index           =   0
               Left            =   4950
               Picture         =   "Form1.frx":8E90
               Top             =   1485
               Width           =   480
            End
            Begin VB.Image КЭ4 
               Height          =   480
               Index           =   0
               Left            =   270
               Picture         =   "Form1.frx":919A
               Top             =   1785
               Width           =   480
            End
            Begin VB.Image КЭ7 
               Height          =   480
               Index           =   0
               Left            =   270
               Picture         =   "Form1.frx":94A4
               Top             =   765
               Width           =   480
            End
            Begin VB.Image Факел 
               Height          =   480
               Index           =   1
               Left            =   555
               Picture         =   "Form1.frx":97AE
               Top             =   180
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00FFFF00&
               Height          =   60
               Index           =   0
               Left            =   200
               Shape           =   3  'Circle
               Top             =   2560
               Visible         =   0   'False
               Width           =   60
            End
            Begin Threed.SSCommand SSCmdStart 
               Height          =   1272
               Left            =   6732
               TabIndex        =   175
               Top             =   3912
               Width           =   2532
               _Version        =   65536
               _ExtentX        =   4466
               _ExtentY        =   2244
               _StockProps     =   78
               Caption         =   "Пуск АГНКС"
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BevelWidth      =   8
               Font3D          =   2
               Picture         =   "Form1.frx":9AB8
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   1272
               Index           =   0
               Left            =   1572
               TabIndex        =   174
               Top             =   3900
               Width           =   1872
               _Version        =   65536
               _ExtentX        =   3302
               _ExtentY        =   2244
               _StockProps     =   78
               Caption         =   "АГНКС"
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   10
               Font3D          =   2
               Picture         =   "Form1.frx":9AD4
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   1272
               Index           =   1
               Left            =   132
               TabIndex        =   173
               Top             =   3900
               Width           =   1392
               _Version        =   65536
               _ExtentX        =   2455
               _ExtentY        =   2244
               _StockProps     =   78
               Caption         =   "ДВС"
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   10
               Font3D          =   2
               Picture         =   "Form1.frx":9F26
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   " Нм3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   555
               Left            =   5595
               TabIndex        =   172
               Top             =   4290
               Width           =   1080
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   3
               X1              =   7464
               X2              =   1140
               Y1              =   240
               Y2              =   228
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   4
               X1              =   5196
               X2              =   4440
               Y1              =   1104
               Y2              =   1104
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   9
               X1              =   1095
               X2              =   555
               Y1              =   1410
               Y2              =   1410
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   10
               X1              =   1770
               X2              =   1770
               Y1              =   2115
               Y2              =   2583
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "час"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   2424
               TabIndex        =   171
               Top             =   348
               Width           =   372
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "час"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   1
               Left            =   3984
               TabIndex        =   170
               Top             =   336
               Visible         =   0   'False
               Width           =   372
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ3"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   4545
               TabIndex        =   169
               Top             =   1650
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ4"
               ForeColor       =   &H00FF0000&
               Height          =   192
               Index           =   3
               Left            =   132
               TabIndex        =   168
               Top             =   1608
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ7"
               ForeColor       =   &H00FF0000&
               Height          =   192
               Index           =   4
               Left            =   120
               TabIndex        =   167
               Top             =   600
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ5"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   5
               Left            =   6780
               TabIndex        =   166
               Top             =   315
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ6"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   6
               Left            =   6075
               TabIndex        =   165
               Top             =   330
               Width           =   300
            End
            Begin VB.Label Т_после_детандера 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "+17"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   5310
               TabIndex        =   164
               Top             =   2685
               Width           =   450
            End
            Begin VB.Label Т_газ_на_входе 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "+17"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   3390
               TabIndex        =   163
               Top             =   2685
               Width           =   450
            End
            Begin VB.Image Термометр 
               Height          =   240
               Index           =   0
               Left            =   3525
               Picture         =   "Form1.frx":A378
               Top             =   2250
               Width           =   150
            End
            Begin VB.Image Термометр 
               Height          =   240
               Index           =   1
               Left            =   4890
               Picture         =   "Form1.frx":A47A
               Top             =   2595
               Width           =   150
            End
            Begin VB.Label Р_вход_компр 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Давление!"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   2190
               TabIndex        =   162
               Top             =   2685
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ2"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   675
               TabIndex        =   161
               Top             =   1605
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "КЭ1"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   1110
               TabIndex        =   160
               Top             =   2760
               Width           =   300
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   0
               X1              =   1116
               X2              =   1116
               Y1              =   1392
               Y2              =   228
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00FFFF00&
               Height          =   60
               Index           =   1
               Left            =   300
               Shape           =   3  'Circle
               Top             =   2560
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.Shape Муфта 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               Height          =   252
               Left            =   2736
               Top             =   1308
               Width           =   672
            End
            Begin VB.Image КЭ4 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "Form1.frx":A57C
               Top             =   1785
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ7 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "Form1.frx":A886
               Top             =   765
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ3 
               Height          =   480
               Index           =   1
               Left            =   4950
               Picture         =   "Form1.frx":AB90
               Top             =   1485
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ6 
               Height          =   480
               Index           =   1
               Left            =   5520
               Picture         =   "Form1.frx":AE9A
               Top             =   405
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ5 
               Height          =   480
               Index           =   1
               Left            =   7230
               Picture         =   "Form1.frx":B1A4
               Top             =   405
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image Факел 
               Height          =   480
               Index           =   0
               Left            =   1110
               Picture         =   "Form1.frx":B4AE
               Top             =   1380
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   2
               X1              =   504
               X2              =   504
               Y1              =   2556
               Y2              =   468
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   8
               X1              =   4020
               X2              =   285
               Y1              =   2595
               Y2              =   2595
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   5
               X1              =   5190
               X2              =   5190
               Y1              =   2895
               Y2              =   270
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   1
               X1              =   5760
               X2              =   5760
               Y1              =   864
               Y2              =   252
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   11
               X1              =   7476
               X2              =   7476
               Y1              =   1008
               Y2              =   240
            End
            Begin VB.Line Line7 
               BorderColor     =   &H000080FF&
               BorderWidth     =   6
               Index           =   16
               X1              =   4605
               X2              =   4605
               Y1              =   2025
               Y2              =   2925
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   4
         Left            =   -75000
         TabIndex        =   132
         Top             =   360
         Width           =   9345
         Begin VB.ListBox lstStat 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   3765
            Index           =   3
            ItemData        =   "Form1.frx":B7B8
            Left            =   7080
            List            =   "Form1.frx":B7BA
            TabIndex        =   136
            Top             =   600
            Width           =   2055
         End
         Begin VB.ListBox lstStat 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   3765
            Index           =   2
            ItemData        =   "Form1.frx":B7BC
            Left            =   4800
            List            =   "Form1.frx":B7BE
            TabIndex        =   135
            Top             =   600
            Width           =   2175
         End
         Begin VB.ListBox lstStat 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   3765
            Index           =   1
            ItemData        =   "Form1.frx":B7C0
            Left            =   2400
            List            =   "Form1.frx":B7C2
            TabIndex        =   134
            Top             =   600
            Width           =   2295
         End
         Begin VB.ListBox lstStat 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   3765
            Index           =   0
            ItemData        =   "Form1.frx":B7C4
            Left            =   120
            List            =   "Form1.frx":B7C6
            TabIndex        =   133
            Top             =   600
            Width           =   2175
         End
         Begin Threed.SSCommand ssStat 
            Height          =   855
            Left            =   120
            TabIndex        =   187
            Top             =   4440
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   1508
            _StockProps     =   78
            Caption         =   "Статистика"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   5
            Font3D          =   1
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            Caption         =   "Годы"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   7320
            TabIndex        =   140
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            Caption         =   "За год"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   139
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            Caption         =   "За месяц"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   138
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "За день"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   780
            TabIndex        =   137
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   2
         Left            =   0
         TabIndex        =   131
         Top             =   360
         Width           =   9345
         Begin VB.Timer Timer2 
            Interval        =   500
            Left            =   765
            Top             =   1530
         End
         Begin VB.TextBox txtTimeDate 
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   675
            TabIndex        =   189
            Text            =   "12:01:02"
            Top             =   585
            Width           =   1410
         End
         Begin Threed.SSCommand SSExit 
            Height          =   1725
            Left            =   2580
            TabIndex        =   184
            Top             =   3510
            Width           =   4530
            _Version        =   65536
            _ExtentX        =   7990
            _ExtentY        =   3043
            _StockProps     =   78
            Caption         =   "ВЫХОД"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   7
            Font3D          =   3
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Данный программный продукт разработан лабораторией автоматизации производства Управления ""ЭНЕРГОГАЗРЕМОНТ"""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   2985
            TabIndex        =   181
            Top             =   2085
            Width           =   3885
         End
         Begin VB.Image Image1 
            Height          =   1365
            Left            =   4455
            Picture         =   "Form1.frx":B7C8
            Stretch         =   -1  'True
            Top             =   225
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   1
         Left            =   -75000
         TabIndex        =   98
         Top             =   360
         Width           =   9345
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   130
            Text            =   "Text2"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   129
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   128
            Text            =   "Text2"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   127
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   126
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   125
            Text            =   "Text2"
            Top             =   1920
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   124
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   123
            Text            =   "Text2"
            Top             =   2280
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   122
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   121
            Text            =   "Text2"
            Top             =   2640
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   120
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   119
            Text            =   "Text2"
            Top             =   3000
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   2400
            TabIndex        =   118
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   117
            Text            =   "Text2"
            Top             =   3360
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   7320
            TabIndex        =   116
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   8
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   840
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   7320
            TabIndex        =   114
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   9
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   113
            Text            =   "Text2"
            Top             =   1200
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   7320
            TabIndex        =   112
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   10
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   111
            Text            =   "Text2"
            Top             =   1560
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   7320
            TabIndex        =   110
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   11
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   109
            Text            =   "Text2"
            Top             =   1920
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   7320
            TabIndex        =   108
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   12
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   107
            Text            =   "Text2"
            Top             =   2280
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   7320
            TabIndex        =   106
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   13
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   105
            Text            =   "Text2"
            Top             =   2640
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   7320
            TabIndex        =   104
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   14
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   103
            Text            =   "Text2"
            Top             =   3000
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   7320
            TabIndex        =   102
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   15
            Left            =   5160
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   101
            Text            =   "Text2"
            Top             =   3360
            Width           =   2050
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   100
            Text            =   "Text2"
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   99
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblPC 
            Caption         =   "Label5"
            Height          =   285
            Left            =   6495
            TabIndex        =   182
            Top             =   4890
            Width           =   2715
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   0
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   9345
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   47
            Left            =   7200
            TabIndex        =   97
            Top             =   4365
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   47
            Left            =   6840
            TabIndex        =   96
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   46
            Left            =   7200
            TabIndex        =   95
            Top             =   4005
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   46
            Left            =   6840
            TabIndex        =   94
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   45
            Left            =   7200
            TabIndex        =   93
            Top             =   3645
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   45
            Left            =   6840
            TabIndex        =   92
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   44
            Left            =   7200
            TabIndex        =   91
            Top             =   3285
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   44
            Left            =   6840
            TabIndex        =   90
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   43
            Left            =   7200
            TabIndex        =   89
            Top             =   2925
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   43
            Left            =   6840
            TabIndex        =   88
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   42
            Left            =   7200
            TabIndex        =   87
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   42
            Left            =   6840
            TabIndex        =   86
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   41
            Left            =   7200
            TabIndex        =   85
            Top             =   2205
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   41
            Left            =   6840
            TabIndex        =   84
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   40
            Left            =   7200
            TabIndex        =   83
            Top             =   1845
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   40
            Left            =   6840
            TabIndex        =   82
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   39
            Left            =   7200
            TabIndex        =   81
            Top             =   1485
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   39
            Left            =   6840
            TabIndex        =   80
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   38
            Left            =   7200
            TabIndex        =   79
            Top             =   1125
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   38
            Left            =   6840
            TabIndex        =   78
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   37
            Left            =   7200
            TabIndex        =   77
            Top             =   765
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   37
            Left            =   6840
            TabIndex        =   76
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   36
            Left            =   7200
            TabIndex        =   75
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   36
            Left            =   6840
            TabIndex        =   74
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   35
            Left            =   5040
            TabIndex        =   73
            Top             =   4365
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   35
            Left            =   4680
            TabIndex        =   72
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   34
            Left            =   5040
            TabIndex        =   71
            Top             =   4005
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   34
            Left            =   4680
            TabIndex        =   70
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   33
            Left            =   5040
            TabIndex        =   69
            Top             =   3645
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   33
            Left            =   4680
            TabIndex        =   68
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   32
            Left            =   5040
            TabIndex        =   67
            Top             =   3285
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   32
            Left            =   4680
            TabIndex        =   66
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   31
            Left            =   5040
            TabIndex        =   65
            Top             =   2925
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   31
            Left            =   4680
            TabIndex        =   64
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   30
            Left            =   5040
            TabIndex        =   63
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   30
            Left            =   4680
            TabIndex        =   62
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   29
            Left            =   5040
            TabIndex        =   61
            Top             =   2205
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   29
            Left            =   4680
            TabIndex        =   60
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   28
            Left            =   5040
            TabIndex        =   59
            Top             =   1845
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   28
            Left            =   4680
            TabIndex        =   58
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   27
            Left            =   5040
            TabIndex        =   57
            Top             =   1485
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   27
            Left            =   4680
            TabIndex        =   56
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   26
            Left            =   5040
            TabIndex        =   55
            Top             =   1125
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   26
            Left            =   4680
            TabIndex        =   54
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   25
            Left            =   5040
            TabIndex        =   53
            Top             =   765
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   25
            Left            =   4680
            TabIndex        =   52
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   24
            Left            =   5040
            TabIndex        =   51
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   24
            Left            =   4680
            TabIndex        =   50
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   23
            Left            =   2640
            TabIndex        =   49
            Top             =   4365
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   23
            Left            =   2280
            TabIndex        =   48
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   22
            Left            =   2640
            TabIndex        =   47
            Top             =   4005
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   22
            Left            =   2280
            TabIndex        =   46
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   21
            Left            =   2640
            TabIndex        =   45
            Top             =   3645
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   21
            Left            =   2280
            TabIndex        =   44
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   20
            Left            =   2640
            TabIndex        =   43
            Top             =   3285
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   20
            Left            =   2280
            TabIndex        =   42
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   19
            Left            =   2640
            TabIndex        =   41
            Top             =   2925
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   19
            Left            =   2280
            TabIndex        =   40
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   18
            Left            =   2640
            TabIndex        =   39
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   18
            Left            =   2280
            TabIndex        =   38
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   17
            Left            =   2640
            TabIndex        =   37
            Top             =   2205
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   17
            Left            =   2280
            TabIndex        =   36
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   16
            Left            =   2640
            TabIndex        =   35
            Top             =   1845
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   16
            Left            =   2280
            TabIndex        =   34
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   15
            Left            =   2640
            TabIndex        =   33
            Top             =   1485
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   15
            Left            =   2280
            TabIndex        =   32
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   14
            Left            =   2640
            TabIndex        =   31
            Top             =   1125
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   14
            Left            =   2280
            TabIndex        =   30
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   13
            Left            =   2640
            TabIndex        =   29
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   13
            Left            =   2280
            TabIndex        =   28
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   12
            Left            =   2640
            TabIndex        =   27
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   12
            Left            =   2280
            TabIndex        =   26
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   11
            Left            =   480
            TabIndex        =   25
            Top             =   4365
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   24
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   10
            Left            =   480
            TabIndex        =   23
            Top             =   4005
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   22
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   9
            Left            =   480
            TabIndex        =   21
            Top             =   3645
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   20
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   19
            Top             =   3285
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   18
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   17
            Top             =   2925
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   16
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   15
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   13
            Top             =   2205
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   11
            Top             =   1845
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   9
            Top             =   1485
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   7
            Top             =   1125
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   5
            Top             =   765
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Р на входе АГНКС"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   3
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      Index           =   12
      X1              =   540
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit





Private Sub CmdЗаправка_Click(Index As Integer)
Dim t As Integer
 'Если идет заправка аккумуляторов
 If gbAkkum = True Then
   gsMsg = "Пистолет вставлен ?"
   frmЗапрос.Show 0
   gbFrmShow = True
End If
      giStage = 2
End Sub




Private Sub cmdDanger_Click()
   frmStart.cmdDanger.Visible = False
      'закрыть все КЭМы
   ROff A1, 1
   'Стоп ДВС, открыть КЭМ4
    giStage2 = 0
     giStage = 0 'Переход на этап Исходное Состояние
     giStage1 = 0
     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "Пуск АГНКС"
     gbDVSStopping = True

    gbStopAGNKS = False
        
End Sub

Private Sub cmdKKM_Click()
StatusKKM
frmKKM.lblErrorKKM.Caption = gsErrorKKM ' = Drvfr.ResultCodeDescription
frmKKM.lblStatusKKM.Caption = gsРежимККМ '= Drvfr.ECRModeDescription
frmKKM.Show 1
    


End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s, s1 As String
      cmdStop.Enabled = False
        'Закрыть пистолет
            'Закрыть КЭ5
            ROff A1, 191
            gbDontStat = False 'Можно работать с диском
        gdTime = GetTimeCounter(2)
        
            'Заполнить статистику по заправке
'          giRealCountZ = giRealCountZ + 1
'          gdaStat1(0).IR1 = giRealCountZ
          
'          gdРасход1 = gdИР2
'          gdaStat1(giRealCountZ).IR2 = gdРасход1
'          gdaStat1(giRealCountZ).IR1 = gdИР1
'          gdaStat1(giRealCountZ).dt = Now
 '         gdaStat1(giRealCountZ).Motor = GMC + MotorCount
'          GMC = gdaStat1(giRealCountZ).Motor
'          MotorCount = 0
           '<<<<Прекратить считать расход>>>>
                      StatRS.AddNew
  
           StatRS("DATA") = Now
           StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
             StopOutput (2)
           StatRS("GAZ_IR1") = gdИР1
           StatRS("MOTO") = GMC + MotorCount
           GMC = GMC + MotorCount
            MotorCount = 0
   If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
     Verify
   End If
       
           StatRS.Update
     
     s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
     frmStart.lstStat(0).AddItem s
           
          gDateRec = Now
            gbЗаправка = False
           
If gbOnlyAkk = True Then
            'Закрыть КЭ6
           ROff A1, 127
           frmStart.SSCmdStart.Enabled = True
           gbAkkum = True
          giStage = 1  'Переход на Этап Предпуска
          giStage1 = 0
          giStage2 = 0
Else
           'ЗАПРАВЛЯЕМ АККУМУЛЯТОРЫ
           'Открыть КЭ6
           ROn A1, 128
End If
             
      'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
           frmStart.SSCmdStart.Enabled = True
           gbAkkum = True
       

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As String
Dim s1 As String

'Для ввода поправочного коэффициента
Dim descr As Integer
Dim sPath As String
Dim rec As pswd

'реакция на ctrl+alt+home
On Error Resume Next
  If ((KeyCode = vbKeyHome) And (Shift = 6)) Then
    s = InputBox("Введите пароль", "DANGER")
    If (s = Password) Then
      s = InputBox("Введите поправочный коэффициент", "DANGER")
      If (CDbl(s) > 0) And (CDbl(s) <= 10) Then
       gdK = CDbl(s)
        sPath = "C:\Winnt\dll32.dll"
        descr = FreeFile
        Open sPath For Random As descr Len = Len(rec)
          rec.pwd = Password
          rec.PC = gdK
          Put #descr, 1, rec
          MsgBox ("Коэффициент введен ")
         Close #descr
      End If
     Else
        MsgBox "Пароль не верный ", vbCritical
    End If
  ElseIf ((KeyCode = vbKeyEnd) And (Shift = 6)) Then
     s = InputBox("Введите пароль", "DANGER")
    If (s = Password) Then
      s = InputBox("Введите новый пароль", "DANGER")
      If (Len(s) > 0) And (Len(s) <= 7) Then
        s1 = InputBox("Повторите новый пароль", "DANGER")
         If (s = s1) Then
         
            Password = s
            sPath = "C:\Winnt\dll32.dll"
            descr = FreeFile
            Open sPath For Random As descr Len = Len(rec)
              rec.pwd = Password
              rec.PC = gdK
              Put #descr, 1, rec
              MsgBox ("Пароль введен ")
             Close #descr
          Else
            MsgBox "Пароль не верный ", vbCritical
          End If
      End If
     Else
        MsgBox "Пароль не верный ", vbCritical
    End If
  End If

End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim s, s1 As String
'ConnectKKM
  Left = 10
  Top = 700
  MaxId = 0
  InitAGNKS
 
  FileHandle = FreeFile
  ' Получить путь программы
  ' Иначе получается каталог Бейсика
  s = App.Path & "\data.txt"
  Open s For Input Access Read As FileHandle
  Seek #FileHandle, 1
  'Ввод пояснений о датчиках из файла для обоих плат
  For i = 0 To 47
    Line Input #FileHandle, gnДатчик(i).Note
    Label2(i).Caption = gnДатчик(i).Note
  Next i
  For i = 0 To 15
    Line Input #FileHandle, s
    Text2(i).Text = s
  Next i

   Close #FileHandle
 frmStart.SSTab1.Tab = 3
   'Показать главную форму
   Show

 Timer1.Interval = 500
 Timer1.Enabled = True
 
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim k As Integer
Dim i As Integer
Dim t As MyRecType
Dim s As String
Dim temp2 As MyRecType

    If tmrMotor.Enabled = True Then
    'Не забудь проверку на пустую БД
'      gdaStat1(k).Motor = GMC + MotorCount
'      GMC = gdaStat1(k).Motor
'      MotorCount = 0
    End If
    If gbDontStat = True Then
        StatRS.AddNew
        StatRS("DATA") = Now
        StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
        StatRS("GAZ_IR1") = gdИР1
        StatRS("MOTO") = GMC + MotorCount
        GMC = GMC + MotorCount
        MotorCount = 0
        StatRS.Update
        s = Format(Now, "hh:mm:ss") + "     " + Format((gdРасход1 / gdPlot), "###0.00")
        frmStart.lstStat(0).AddItem s
        
        gDateRec = Now
        gbDontStat = False 'Можно работать с диском
      Else
         Set SelectRS = StatDB.OpenRecordset("select MAX(DATA) from stat ")
          temp2.dt = SelectRS(0)
         s = Module4.Convert_Date(Str(Month(temp2.dt)) & "/" & Day(temp2.dt) & "/" & Year(temp2.dt) & " " & Hour(temp2.dt) & ":" & Minute(temp2.dt) & ":" & Second(temp2.dt))
          Set SelectRS = StatDB.OpenRecordset("SELECT * From stat WHERE stat.data=" & s)
         SelectRS.Edit
         SelectRS("MOTO") = GMC + MotorCount
        SelectRS.Update
     End If
    
  ' StatRS.Close
  ' StatDB.Close
  ' StatWS.Close
 
   DIO_DriverClose 'Выгрузить драйвер для DIO48
   ISO813_DriverClose
   frmStart.MSComm1.PortOpen = False
   Unload frmЗапрос
   Unload frmStat
    Unload frmSt
End Sub










Private Sub Label1_Click(Index As Integer)
'Dim Maska As Integer
'Dim rez As Long
'Dim i As Integer
'Dim Temp As Integer
' Щелкаем реле
'   Maska = 1
    ' Для порта A0
'  If (Index >= 0 And Index < 8) Then
'   For i = 1 To Index
'     Maska = Maska * 2
'   Next i
'     Temp = gn48DIO(0) 'считываем состояние порта A0
'     Temp = Temp Xor Maska
     
     '!!!!Для отработки
     'rez = W_48DIO_DO(A0, Temp)
'     If gnДатчик(Index).Data = 0 Then
'       ROn A0, Maska
'     Else
'       ROff A0, Maska Xor 255
'     End If
     
'     gn48DIO(0) = Temp
     
     
     ' Для порта A1
'   ElseIf (Index > 23 And Index < 32) Then
'     For i = 1 To Index - 24
'     Maska = Maska * 2
'   Next i
'     Temp = gn48DIO(3) 'считываем состояние порта A1
'     Temp = Temp Xor Maska
     
     '!!!Для отработки
     'rez = W_48DIO_DO(A1, Temp)
'     If gnДатчик(Index).Data = 0 Then
'       ROn A1, Maska
'     Else
'       ROff A1, Maska Xor 255
'     End If

'     gn48DIO(3) = Temp
     
'   End If


End Sub

Private Sub lstStat_Click(Index As Integer)
'Dim i As Integer
' If (Index = 1) Then
'  i = lstStat(1).ListIndex
'  frmStat.txtStat(0).Text = gdaStat2(i + 1).IR1
'  frmStat.txtStat(1).Text = gdaStat2(i + 1).IR2
'  frmStat.Show 0
' End If
'
' If (Index = 0) Then
'  i = lstStat(0).ListIndex
'  frmStat.txtStat(0).Text = Format(gdИР1, "0.00")
'  frmStat.txtStat(1).Text = Format(gdaStat1(i + 1).IR2, "0.00")
'  frmStat.Show 0
' End If

End Sub

Private Sub SSCmdStart_Click()
Dim t As Integer
If gbCmdStart = True Then
   gbCmdStart = False
   SSCmdStart.Caption = "ЗАПРАВКА"
   giStage = 1 'Переход на этап ПредПуск()
   giStage2 = 0
   giStage1 = 0
   'Открыть КЭ1
     ROn A1, 4

Else
      'Если идет заправка аккумуляторов
      cmdStop.Enabled = True
      If gbAkkum = True Then
        gsMsg = "Пистолет вставлен ?"
        frmЗапрос.Show 0
        gbFrmShow = True
      End If
         giStage = 2
         SSCmdStart.Enabled = False
End If

End Sub



Private Sub SSCommand1_Click()
Dim j As Integer
   j = ExitWindowsEx(2, 0)
End Sub

Private Sub SSCommand2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s As String
If giStage = 2 Then
  StopOutput (2)
End If

Select Case Index
Case 1
 'Если открыт КЭМ5 - закрыть
 SSCommand2(1).Enabled = False
    If gnДатчик(30).Data = 1 Then
      ROff A1, 191
    End If
        ROn A1, 2

        ROff A1, 0 'Закрыть все КЭМы
     giStage2 = 0
     giStage = 0 'Переход на этап ИсхСост
     giStage1 = 0
     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "Пуск АГНКС"
     gbDVSStopping = True
  If gbDontStat = True Then
           StatRS.AddNew
           StatRS("DATA") = Now
           StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
           StatRS("GAZ_IR1") = gdИР1
           StatRS("MOTO") = GMC + MotorCount
           GMC = GMC + MotorCount
           MotorCount = 0
              If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
                 Verify
              End If
          
           StatRS.Update
           s = Format(Now, "hh:mm:ss") + "        " + Format(gdРасход1 / gdPlot, "###0.00")
           frmStart.lstStat(0).AddItem s
           
            gDateRec = Now 'StatRS("DATA")
           gbDontStat = False 'Можно работать с диском
    End If

     'ОстановДВС = "Двигатель остановлен !!!"
Case 0
 SSCommand2(0).Enabled = False

 'Если открыт КЭМ5 - закрыть
    If gnДатчик(30).Data = 1 Then
      ROff A1, 191
    End If
        ROn A1, 2

        ROff A1, 0 'Закрыть все КЭМы
     giStage2 = 0
     giStage = 0 'Переход на этап ИсхСост
     giStage1 = 0
     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "Пуск АГНКС"
     gbDVSStopping = True
             'frmStart.Timer2.Enabled = False
     'ОстановДВС = "Двигатель остановлен !!!"
   frmStart.cmdDanger.Visible = True
 If gbDontStat = True Then
                 StatRS.AddNew
  
           StatRS("DATA") = Now
           StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
 
           StatRS("GAZ_IR1") = gdИР1
           StatRS("MOTO") = GMC + MotorCount
           GMC = GMC + MotorCount
            MotorCount = 0
   If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
                 Verify
              End If
            
           StatRS.Update
     s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
     frmStart.lstStat(0).AddItem s
           
            gDateRec = Now
       gbDontStat = False 'Можно работать с диском
    End If


   ОкноСообщений.Caption = ОстановАГНКС()
End Select

End Sub

Private Sub SSExit_Click()
Dim j As Integer
Dim k As Integer
Dim i As Integer
Dim t As MyRecType

    'Выгрузить драйвер для DIO48
Dim s As String
    If tmrMotor.Enabled = True Then
    'Не забудь проверку на пустую БД
'      gdaStat1(k).Motor = GMC + MotorCount
'      GMC = gdaStat1(k).Motor
'      MotorCount = 0
    End If
    If gbDontStat = True Then
        StatRS.AddNew
        StatRS("DATA") = Now
        StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
        StatRS("GAZ_IR1") = gdИР1
        StatRS("MOTO") = GMC + MotorCount
        GMC = GMC + MotorCount
        MotorCount = 0
        StatRS.Update
        s = Format(Now, "hh:mm:ss") + "     " + Format((gdРасход1 / gdPlot), "###0.00")
        frmStart.lstStat(0).AddItem s
        
        gDateRec = Now
        gbDontStat = False 'Можно работать с диском
    End If
    
'   StatRS.Close
'   StatDB.Close
'   StatWS.Close
 
   
   If frmStart.MSComm1.PortOpen = True Then
     frmStart.MSComm1.PortOpen = False
   End If
   tmrTablo.Enabled = False
   Unload frmЗапрос
   Unload frmStat
   j = ExitWindowsEx(1, 0)
   End
   

End Sub

Private Sub ssStat_Click()
    frmSt.Calendar1.Value = Now
    frmSt.Show 0
End Sub

Private Sub SSTab1_DblClick()
    frmSt.Show 0
End Sub

Private Sub Timer_Газ_Timer()

'Управление "движением газа"
Dim i As Integer
Dim n As Integer


    n = 211
    
    Select Case CN
        Case 0
            For i = 0 To n Step 3
                Shape4(i).Visible = True
            Next i
        
            For i = 1 To n Step 3
                Shape4(i).Visible = False
            Next i
            
            For i = 2 To n Step 3
                Shape4(i).Visible = False
            Next i
        
            CN = 1
        Case 1
            For i = 0 To n Step 3
                Shape4(i).Visible = False
            Next i
        
            For i = 1 To n Step 3
                Shape4(i).Visible = True
            Next i
            
            For i = 2 To n Step 3
                Shape4(i).Visible = False
            Next i
        
            CN = 2

        Case 2
            For i = 0 To n Step 3
                Shape4(i).Visible = False
            Next i
        
            For i = 1 To n Step 3
                Shape4(i).Visible = False
            Next i
            
            For i = 2 To n Step 3
                Shape4(i).Visible = True
            Next i
        
            CN = 0
            
    End Select
End Sub

Private Sub Timer_ДВС_Timer()

    Dim i As Integer
    
'Отображение работы ДВС, компрессора, детандера
    If ОборотыДВС.Caption > 50 Then
      tmrMotor.Enabled = True 'Считать моторесурс
        For i = 0 To 5
            If ДВС(i).Visible Then
                ДВС(i).Visible = False
                If Муфта.BackColor = &HFF& Then
                    Компрессор(i).Visible = False
                End If
                If i < 5 Then
                    ДВС(i + 1).Visible = True
                    If Муфта.BackColor = &HFF& Then
                        Компрессор(i + 1).Visible = True
                    End If
                    Exit For
                Else
                    ДВС(0).Visible = True
                    If Муфта.BackColor = &HFF& Then
                        Компрессор(0).Visible = True
                    End If
                End If
            End If
        Next i
    Else
      tmrMotor.Enabled = False 'Перестать считать моторесурс
    End If
'Отображение "открытия" и "закрытия" кранов КЭ1...КЭ7, а также "факела"
    If КЭ1(1).Visible Then
        КЭ1(0).Visible = False
    Else
        КЭ1(0).Visible = True
    End If

    If КЭ2(1).Visible Then
        КЭ2(0).Visible = False
        Факел(0).Visible = True
    Else
        Факел(0).Visible = False
        КЭ2(0).Visible = True
    End If
    
    If КЭ3(1).Visible Then
        КЭ3(0).Visible = False
    Else
        КЭ3(0).Visible = True
    End If
    
    If КЭ4(1).Visible Then
        КЭ4(0).Visible = False
    Else
        КЭ4(0).Visible = True
    End If

    If КЭ5(1).Visible Then
        КЭ5(0).Visible = False
    Else
        КЭ5(0).Visible = True
    End If

    If КЭ6(1).Visible Then
        КЭ6(0).Visible = False
    Else
        КЭ6(0).Visible = True
    End If

    If КЭ7(1).Visible Then
        КЭ7(0).Visible = False
        Факел(1).Visible = True
    Else
        Факел(1).Visible = False
        КЭ7(0).Visible = True
    End If
    
'Отображение заправки автобаллона
    If КЭ5(1).Visible Then
        Панель_Авто.Visible = True
        If (100 * (Р_автобаллон / 200) >= 100) Then
           Автобаллон.FloodPercent = 100
        Else
           Автобаллон.FloodPercent = 100 * (Р_автобаллон / 200)
        End If
    Else
        Панель_Авто.Visible = False
    End If
    
'Отображение работы аккумулятора
        If (100 * (Р_аккумулятор / 200) >= 100) Then
           Аккумулятор.FloodPercent = 100
        Else
           Аккумулятор.FloodPercent = 100 * (Р_аккумулятор / 200)
        End If

    
End Sub




Private Sub SSPanel4_DblClick()
  DVSEmul = Not (DVSEmul)
End Sub


Private Sub Timer1_Timer()
 Dim k, f As Integer
 Dim Dv, Akk, t As Integer
 Dim Temp As Double
 Dim s As String
 Dim s1 As String
 Dim ErrDat As Boolean
   ErrDat = False
   s = ""
    'Для отладки !!!!! (отключить)
   
   ОпросПлат

   Обработка_1
   ' Заполнение результатми с платы 48DIO
   
   'Управление изображением
   ShowPict
   
   
   'Работа с диском
   'Если произошла еще заправка после последней записи на диск или наступил другой день(месяц)
   'И разрешена проверка
   If ((giRealCountZ > giCountZ) Or _
   ((Day(gDateRec) < Day(Date)) Or (Month(gDateRec) < Month(Date)) Or (Year(gDateRec) < Year(Date)))) _
      And (gbDontStat = False) Then
      giErrDisk = Verify
   End If
   
   
   
   For k = 0 To 47
     If gnДатчик(k).Data = 0 Then
        Label1(k).BackColor = &HFF00&
     Else
        Label1(k).BackColor = &HFF
     End If
   Next k
   
   
   glCounter = glCounter + 1
   For k = 2 To 16
     If gnDif(k) = -1 Then
       sum(k) = -1
     ElseIf sum(k) = -1 Then
      sum(k) = -1
     Else
       sum(k) = sum(k) + gnDif(k)
     End If
   Next k
   
  If glCounter >= glAver Then  'Если счетчик дошел, то усредняем
   For k = 2 To 16
    
     If sum(k) = -1 Then
        Text2(k - 1).ForeColor = &HFF
        Text1(k - 1).Text = "Не исправен"
        s = "Не исправен"
        sum(k) = 0
     Else
       sum(k) = sum(k) / glCounter
       Text2(k - 1).ForeColor = &H80000012
       ' Проверка на ДД1.1 и ДД1.2 - для них другой шаблон
           If (k = 1) Or (k = 2) Then
             Text1(k - 1).Text = Format(sum(k), "##0.000")
             s = Text1(k - 1).Text
           Else
             Text1(k - 1).Text = Format(sum(k), "##0.000")
             s = Text1(k - 1).Text
           End If
     End If
     
     'Для чистового вывода
    Select Case k
     Case 2
       s = Format(sum(k) / 0.0981, "##0.0")
       Р_вход_АГНКС.Caption = s
     Case 6
       s = Format(sum(k) / 0.0981, "##0.0")
       Р_выход_компр.Caption = s
     Case 7
       s = Format(sum(k) / 0.0981, "##0.0")
       Р_аккумулятор.Caption = s
     Case 8
      s = Format(sum(k), "#0.0")
       Т_после_детандера.Caption = s
     Case 9
      s = Format(sum(k), "#0.0")
       Т_газ_на_входе.Caption = s
     Case 4
       s = Format(sum(k) / 0.0981, "##0.0")
       Р_автобаллон.Caption = s
     Case 14
       s = Format((sum(k) \ 100) * 100, "###0")
       ОборотыДВС.Caption = s
    End Select
    
     sum(k) = 0
    Next k
    glCounter = 0
   End If
   

    
  
  
    lblV.Caption = Format(gnDif(giChanel), "00.0" & " В")
    
    Наработка_ДВС.Caption = Format((GMC + MotorCount) / 60, "00")

    
    'Выводим расход на заправку одной машины
    If (gdРасход1 < 0) Then
       gdРасход1 = 0
    End If
   s = Format((gdРасход1 / gdPlot), "0.0")
   ЗаправленоГаза.Caption = s
   
   'KKM
   If (frmKKM.txtKKM.Visible = True) Then
   Else
       frmKKM.txtKKM.Text = Format((CDbl(s) * gdPrice), "##0.00")
   End If
   'KKM
   
   s = Format(gdРасход1, "0.00")
   txtKg.Text = s
   
   'Выводить в минутах
   s = Format(gdTime / 60, "0")
   txtTime.Text = s
 'Проверка датчиков
 ErrDat = False
 If (gnDif(2) = -1) Or (gnDif(3) = -1) Or (gnDif(4) = -1) Or (gnDif(5) = -1) Or _
 (gnDif(6) = -1) Or (gnDif(7) = -1) Then
    ErrDat = True
 End If
 
 If (gnДатчик(15).Data = 0) Then
   gbHandControl = True
 Else
   gbHandControl = False
 End If
 
 'Если ручное управление
 If (gbHandControl = True) Or (ErrDat = True) Then
 'Если перешли на ручное управление
 ОкноСообщений.BackColor = &HFF
 ОкноСообщений.ForeColor = &HFFFF&
 ОкноСообщений.Caption = "Ручное управление !!! - программа не управляет процессами !"
    If ErrDat = True Then
    ОкноСообщений.Caption = "Неисправны датчики !!! - программа не управляет процессами !"
    End If
 Else
  ОкноСообщений.BackColor = &HE0E0E0
  ОкноСообщений.ForeColor = &HFF0000

 'Если на этапе Заправка заглох ДВС, то  на ИсхСост
   
 If (gnDif(14) < 100) And (giStage = 2) And (gbOnlyAkk = False) Then
    giDVS = giDVS + 1
 Else
    giDVS = 0
 End If

   
   If (giDVS > 5) Then
     giStage2 = 0
     giStage = 0 'Переход на этап ИсхСост
     giStage1 = 0
     giDVS = 0
     gbAkkum = False
     gbRunDVS = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "ПУСК АГНКС"
     gbDVSStopping = True
      'frmStart.Timer2.Enabled = False
    'Закрыть все Кэм
     ROff A1, 0

     ROn A1, 6
          
   End If
    
   
   Select Case giStage
    Case 0:
     '<<<Заправка>>> 1 Этап
     ОкноСообщений.Caption = ИсхСост
         f = DoEvents
    Case 1:
     '<<<Заправка>>> 2 Этап
     ОкноСообщений.Caption = ПредПуск
         f = DoEvents
    Case 2:
     '<<<Заправка>>> 3 Этап
     ОкноСообщений.Caption = Заправка
         f = DoEvents
    Case 3:
     'Аварийное состояние
     ОкноСообщений.Caption = Danger
         f = DoEvents
   End Select
 End If
 
    'Проверка аварийных датчиков
    s = ""
    s1 = ""
    s1 = Verify_Damage
    If s1 <> "" Then
     ОкноСообщений.BackColor = &HFF
     ОкноСообщений.ForeColor = &HFFFF&
     s = ОкноСообщений.Caption + " " + s1
     ОкноСообщений.Caption = s
    Else
     ОкноСообщений.BackColor = &HE0E0E0
     ОкноСообщений.ForeColor = &HFF0000
    End If
    
End Sub






Private Sub Timer2_Timer()
frmStart.txtTimeDate = Format(Time, "h:m:s")
End Sub

Private Sub tmrMotor_Timer()
Dim t As MyRecType
Dim i As Integer
Dim k As Integer

  MotorCount = MotorCount + 1
  
  
End Sub


Private Sub tmrTablo_Timer()
Dim i As Integer
Dim b(1 To 6) As Integer ' временный массив для хранения разрядов
Dim Stroka As String ' символьная строка для обрабоки на табло
Dim Dlina As Integer ' длина символьной переменной
Dim rasrad(1 To 6) As Integer ' массив элементов табло
Dim Obraz As Single 'численное представление введенной строки
Dim IntObraz As Single ' целая часть Obraz
Dim DrObraz As Integer  ' дробная часть образа
Dim Tis As Integer 'число тысяч в строке
Dim Sot As Integer ' число сотен в строке
Dim Des As Integer ' число десятков в строке
Dim Ed As Integer  ' число единиц в строке
Dim DrDes As Integer 'число десятых долей в строке
Dim DrSot As Integer ' число сотых долей в строке
Dim NulSumma As Integer
Dim Posylka As String

    Stroka = ЗаправленоГаза.Caption
    Dlina = Len(Stroka)
    Obraz = CDbl(Stroka)
If (Obraz < 9999.99) Then
    IntObraz = Int(Obraz)
    DrObraz = (Obraz - IntObraz) * 100
    Tis = Int(IntObraz / 1000)
    rasrad(6) = giaTableDecoder(Tis)
    b(6) = Tis
    Sot = Int((IntObraz - (Tis * 1000)) / 100)
    rasrad(5) = giaTableDecoder(Sot)
    b(5) = Sot
    Des = Int((IntObraz - (Tis * 1000) - (Sot * 100)) / 10)
    rasrad(4) = giaTableDecoder(Des)
    b(4) = Des
    Ed = Int((IntObraz - (Tis * 1000) - (Sot * 100) - (Des * 10)))
    rasrad(3) = giaTableDecoder(Ed)
    b(3) = Ed
    DrDes = Int(DrObraz / 10)
    rasrad(2) = giaTableDecoder(DrDes)
    b(2) = DrDes
    DrSot = Int(DrObraz - (DrDes * 10))
    rasrad(1) = giaTableDecoder(DrSot)
    b(1) = DrSot
For i = 6 To 3 Step -1
    NulSumma = NulSumma + b(i)
    If NulSumma = 0 Then rasrad(i) = giaTableDecoder(10)
Next i
If NulSumma = 0 Then rasrad(3) = giaTableDecoder(0)
If (Obraz - IntObraz) = 0 Then rasrad(2) = giaTableDecoder(10)
If b(1) = 0 Then rasrad(1) = giaTableDecoder(10)
For i = 1 To 6
    Posylka = Posylka + Chr(rasrad(i))
Next i

MSComm1.Output = Posylka
End If
End Sub


