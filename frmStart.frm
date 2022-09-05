VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "АГНКС   БИ-40  ""МЕТАН"""
   ClientHeight    =   7395
   ClientLeft      =   3555
   ClientTop       =   2505
   ClientWidth     =   9855
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   493
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Дискретные"
      TabPicture(0)   =   "frmStart.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Аналоговые"
      TabPicture(1)   =   "frmStart.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "О программе"
      TabPicture(2)   =   "frmStart.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSExit"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label16"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtTimeDate"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Image1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Схема"
      TabPicture(3)   =   "frmStart.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Журнал"
      TabPicture(4)   =   "frmStart.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdOpenStatForm"
      Tab(4).Control(1)=   "lstStat(0)"
      Tab(4).Control(2)=   "lstStat(1)"
      Tab(4).Control(3)=   "lstStat(2)"
      Tab(4).Control(4)=   "lstStat(3)"
      Tab(4).Control(5)=   "cmdUpdateStat"
      Tab(4).Control(6)=   "lblStat(0)"
      Tab(4).Control(7)=   "lblStat(1)"
      Tab(4).Control(8)=   "lblStat(2)"
      Tab(4).Control(9)=   "lblStat(3)"
      Tab(4).ControlCount=   10
      Begin VB.CommandButton cmdOpenStatForm 
         Caption         =   "Статистика"
         Height          =   690
         Left            =   -74775
         TabIndex        =   189
         Top             =   6525
         Width           =   2130
      End
      Begin VB.ListBox lstStat 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   5520
         Index           =   0
         ItemData        =   "frmStart.frx":008C
         Left            =   -74820
         List            =   "frmStart.frx":008E
         TabIndex        =   184
         Top             =   810
         Width           =   2200
      End
      Begin VB.ListBox lstStat 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   5520
         Index           =   1
         ItemData        =   "frmStart.frx":0090
         Left            =   -72420
         List            =   "frmStart.frx":0092
         TabIndex        =   183
         Top             =   810
         Width           =   2200
      End
      Begin VB.ListBox lstStat 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   5520
         Index           =   2
         ItemData        =   "frmStart.frx":0094
         Left            =   -70020
         List            =   "frmStart.frx":0096
         TabIndex        =   182
         Top             =   810
         Width           =   2200
      End
      Begin VB.ListBox lstStat 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   5520
         Index           =   3
         ItemData        =   "frmStart.frx":0098
         Left            =   -67620
         List            =   "frmStart.frx":009A
         TabIndex        =   181
         Top             =   810
         Width           =   2200
      End
      Begin VB.CommandButton cmdUpdateStat 
         Caption         =   "Обновить журнал"
         Height          =   690
         Left            =   -67620
         TabIndex        =   180
         Top             =   6525
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "---"
         Height          =   7110
         Index           =   3
         Left            =   0
         TabIndex        =   131
         Top             =   315
         Width           =   9795
         Begin Threed.SSCommand cmdDanger 
            Height          =   2310
            Left            =   3690
            TabIndex        =   161
            Top             =   4680
            Visible         =   0   'False
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   4075
            _StockProps     =   78
            Caption         =   "АВАРИЯ"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            BevelWidth      =   4
            Font3D          =   2
            Picture         =   "frmStart.frx":009C
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   2400
            Left            =   3735
            TabIndex        =   171
            Top             =   4590
            Width           =   3165
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Н/м. куб. в минуту"
               Height          =   195
               Left            =   180
               TabIndex        =   190
               Top             =   1125
               Width           =   1410
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Время заправки"
               Height          =   195
               Left            =   180
               TabIndex        =   179
               Top             =   1485
               Width           =   1260
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ср. скорость заправки"
               Height          =   195
               Left            =   180
               TabIndex        =   178
               Top             =   900
               Width           =   1755
            End
            Begin VB.Label Label_Avg_Speed_Car 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "150"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   330
               Left            =   2070
               TabIndex        =   177
               Top             =   990
               Width           =   960
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Расчетно осталось"
               Height          =   195
               Left            =   180
               TabIndex        =   176
               Top             =   1935
               Width           =   1455
            End
            Begin VB.Label Label_Avg_Left_Time_Car 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-- : -- : -- "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   330
               Left            =   2070
               TabIndex        =   175
               Top             =   1890
               Width           =   960
            End
            Begin VB.Label txtTime 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "00:30:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   330
               Left            =   2070
               TabIndex        =   174
               Top             =   1440
               Width           =   960
            End
            Begin VB.Label txtKg 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5.7"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   330
               Left            =   2070
               TabIndex        =   173
               Top             =   270
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Заправленно, кг."
               Height          =   195
               Left            =   180
               TabIndex        =   172
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Табло заправки"
            Height          =   2400
            Left            =   90
            TabIndex        =   162
            Top             =   4590
            Width           =   3570
            Begin Threed.SSCommand cmdKKM 
               Height          =   750
               Left            =   2655
               TabIndex        =   163
               Top             =   1485
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   1323
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
               BevelWidth      =   4
               Font3D          =   1
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Цена руб."
               Height          =   195
               Left            =   135
               TabIndex        =   170
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Label_Price 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5.70"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   510
               Left            =   90
               TabIndex        =   169
               Top             =   540
               Width           =   1230
            End
            Begin VB.Label Label_Summa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1500.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   27
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   780
               Left            =   90
               TabIndex        =   168
               Top             =   1485
               Width           =   2490
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Сумма руб."
               Height          =   195
               Left            =   135
               TabIndex        =   167
               Top             =   1215
               Width           =   855
            End
            Begin VB.Label ЗаправленоГаза 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   510
               Left            =   1350
               TabIndex        =   166
               Top             =   540
               Width           =   1230
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Н / м"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2700
               TabIndex        =   165
               Top             =   675
               Width           =   525
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "3"
               Height          =   195
               Left            =   3195
               TabIndex        =   164
               Top             =   585
               Width           =   90
            End
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   3450
            Left            =   45
            TabIndex        =   132
            Top             =   90
            Width           =   9720
            _Version        =   65536
            _ExtentX        =   17145
            _ExtentY        =   6085
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
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   6750
               Top             =   2835
            End
            Begin VB.Timer Timer_ДВС 
               Interval        =   75
               Left            =   8430
               Top             =   2835
            End
            Begin VB.Timer Timer2 
               Interval        =   500
               Left            =   7230
               Top             =   2835
            End
            Begin VB.Timer tmrMotor 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   8790
               Top             =   2835
            End
            Begin Threed.SSPanel Отсек_ДВС 
               Height          =   1500
               Left            =   1800
               TabIndex        =   133
               Top             =   870
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2646
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
               BevelWidth      =   2
               BorderWidth     =   0
               Begin Threed.SSPanel ОборотыДВС 
                  Height          =   375
                  Left            =   135
                  TabIndex        =   134
                  Top             =   90
                  Width           =   810
                  _Version        =   65536
                  _ExtentX        =   1429
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "0"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   11.99
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   2
                  BevelOuter      =   1
                  Font3D          =   1
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Двиг."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Index           =   0
                  Left            =   315
                  TabIndex        =   135
                  Top             =   1170
                  Width           =   510
               End
               Begin VB.Image Температура_ДВС 
                  Height          =   480
                  Left            =   360
                  Picture         =   "frmStart.frx":00B8
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   300
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   0
                  Left            =   225
                  Picture         =   "frmStart.frx":02BA
                  Top             =   540
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   1
                  Left            =   225
                  Picture         =   "frmStart.frx":065C
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   2
                  Left            =   225
                  Picture         =   "frmStart.frx":09FE
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   3
                  Left            =   225
                  Picture         =   "frmStart.frx":0DA0
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   4
                  Left            =   225
                  Picture         =   "frmStart.frx":1142
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ДВС 
                  Height          =   600
                  Index           =   5
                  Left            =   225
                  Picture         =   "frmStart.frx":14E4
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
            End
            Begin Threed.SSPanel Отсек_компр 
               Height          =   1500
               Left            =   3105
               TabIndex        =   136
               Top             =   870
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   2646
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
               BevelWidth      =   2
               BorderWidth     =   0
               Begin Threed.SSPanel Р_выход_компр 
                  Height          =   375
                  Left            =   135
                  TabIndex        =   137
                  Top             =   90
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "0"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   11.99
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   2
                  BevelOuter      =   1
                  Font3D          =   1
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Компр."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Index           =   1
                  Left            =   315
                  TabIndex        =   138
                  Top             =   1170
                  Width           =   660
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   0
                  Left            =   180
                  Picture         =   "frmStart.frx":1886
                  Top             =   540
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   1
                  Left            =   180
                  Picture         =   "frmStart.frx":1E08
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   2
                  Left            =   180
                  Picture         =   "frmStart.frx":238A
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   3
                  Left            =   180
                  Picture         =   "frmStart.frx":290C
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   4
                  Left            =   180
                  Picture         =   "frmStart.frx":2E8E
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image Компрессор 
                  Height          =   600
                  Index           =   5
                  Left            =   180
                  Picture         =   "frmStart.frx":3410
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
            End
            Begin Threed.SSPanel Панель_Авто 
               Height          =   1890
               Left            =   7230
               TabIndex        =   139
               Top             =   870
               Visible         =   0   'False
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   3334
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
               BorderWidth     =   0
               Begin Threed.SSPanel Р_автобаллон 
                  Height          =   375
                  Left            =   630
                  TabIndex        =   140
                  Top             =   90
                  Width           =   1320
                  _Version        =   65536
                  _ExtentX        =   2328
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "154"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   11.99
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   2
                  BevelOuter      =   1
                  Font3D          =   1
               End
               Begin Threed.SSPanel Автобаллон 
                  Height          =   1725
                  Left            =   90
                  TabIndex        =   141
                  Top             =   90
                  Width           =   420
                  _Version        =   65536
                  _ExtentX        =   741
                  _ExtentY        =   3043
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
                  BevelWidth      =   2
                  BevelOuter      =   1
                  FloodType       =   4
                  FloodColor      =   16776960
               End
               Begin Threed.SSCommand cmdStopCarRefueling 
                  Height          =   1260
                  Left            =   630
                  TabIndex        =   155
                  Top             =   540
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   2222
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
                  Font3D          =   4
                  AutoSize        =   1
                  Picture         =   "frmStart.frx":3992
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   1890
               Index           =   2
               Left            =   5490
               TabIndex        =   142
               Top             =   870
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   3334
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
               BorderWidth     =   0
               Begin Threed.SSPanel Р_аккумулятор 
                  Height          =   375
                  Left            =   585
                  TabIndex        =   143
                  Top             =   90
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   661
                  _StockProps     =   15
                  Caption         =   "178"
                  ForeColor       =   16776960
                  BackColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   11.99
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelWidth      =   2
                  BevelOuter      =   1
                  Font3D          =   1
               End
               Begin Threed.SSPanel Аккумулятор 
                  Height          =   1725
                  Left            =   90
                  TabIndex        =   144
                  Top             =   90
                  Width           =   420
                  _Version        =   65536
                  _ExtentX        =   741
                  _ExtentY        =   3043
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
                  BevelWidth      =   2
                  BevelOuter      =   1
                  FloodType       =   4
                  FloodColor      =   16776960
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   1125
                  Index           =   4
                  Left            =   1305
                  Shape           =   4  'Rounded Rectangle
                  Top             =   540
                  Width           =   135
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   1125
                  Index           =   3
                  Left            =   1125
                  Shape           =   4  'Rounded Rectangle
                  Top             =   540
                  Width           =   135
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   1125
                  Index           =   2
                  Left            =   945
                  Shape           =   4  'Rounded Rectangle
                  Top             =   540
                  Width           =   135
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   1125
                  Index           =   1
                  Left            =   765
                  Shape           =   4  'Rounded Rectangle
                  Top             =   540
                  Width           =   135
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFF00&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   1125
                  Index           =   0
                  Left            =   585
                  Shape           =   4  'Rounded Rectangle
                  Top             =   540
                  Width           =   135
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   630
                  X2              =   630
                  Y1              =   1755
                  Y2              =   1575
               End
               Begin VB.Line Line2 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   405
                  X2              =   1350
                  Y1              =   1755
                  Y2              =   1755
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   810
                  X2              =   810
                  Y1              =   1755
                  Y2              =   1575
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   990
                  X2              =   990
                  Y1              =   1755
                  Y2              =   1575
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   1170
                  X2              =   1170
                  Y1              =   1755
                  Y2              =   1575
               End
               Begin VB.Line Line6 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   1350
                  X2              =   1350
                  Y1              =   1755
                  Y2              =   1575
               End
            End
            Begin VB.Label Наработка_ДВС 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Caption         =   "Наработка ДВС 999999 ч."
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   1305
               TabIndex        =   196
               Top             =   540
               Width           =   2130
            End
            Begin VB.Image Термометр 
               Height          =   240
               Index           =   1
               Left            =   5220
               Picture         =   "frmStart.frx":6564
               Top             =   3015
               Width           =   150
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   13
               X1              =   5175
               X2              =   4680
               Y1              =   2985
               Y2              =   2970
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   7
               X1              =   5145
               X2              =   4275
               Y1              =   2010
               Y2              =   2025
            End
            Begin VB.Image Image3 
               Height          =   435
               Left            =   4500
               Picture         =   "frmStart.frx":6666
               Stretch         =   -1  'True
               Top             =   2115
               Width           =   375
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   4455
               Picture         =   "frmStart.frx":6F30
               Top             =   2430
               Width           =   480
            End
            Begin VB.Label Р_вход_АГНКС 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5.7"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   135
               TabIndex        =   145
               Top             =   2970
               Width           =   495
            End
            Begin VB.Image КЭ2 
               Height          =   480
               Index           =   0
               Left            =   1170
               Picture         =   "frmStart.frx":77FA
               Top             =   1935
               Width           =   480
            End
            Begin VB.Image КЭ2 
               Height          =   480
               Index           =   1
               Left            =   1170
               Picture         =   "frmStart.frx":7B04
               Top             =   1935
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   6
               X1              =   1395
               X2              =   1410
               Y1              =   1620
               Y2              =   2760
            End
            Begin VB.Image КЭ1 
               Height          =   480
               Index           =   0
               Left            =   765
               Picture         =   "frmStart.frx":7E0E
               Top             =   2565
               Width           =   480
            End
            Begin VB.Image КЭ1 
               Height          =   480
               Index           =   1
               Left            =   720
               Picture         =   "frmStart.frx":8118
               Top             =   2565
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ5 
               Height          =   480
               Index           =   0
               Left            =   7230
               Picture         =   "frmStart.frx":8422
               Top             =   400
               Width           =   480
            End
            Begin VB.Image КЭ6 
               Height          =   480
               Index           =   0
               Left            =   5520
               Picture         =   "frmStart.frx":872C
               Top             =   400
               Width           =   480
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   555
               Index           =   1
               Left            =   4635
               Shape           =   4  'Rounded Rectangle
               Top             =   945
               Width           =   150
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   555
               Index           =   0
               Left            =   4815
               Shape           =   4  'Rounded Rectangle
               Top             =   945
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Image КЭ3 
               Height          =   480
               Index           =   0
               Left            =   4950
               Picture         =   "frmStart.frx":8A36
               Top             =   1485
               Width           =   480
            End
            Begin VB.Image КЭ4 
               Height          =   480
               Index           =   0
               Left            =   270
               Picture         =   "frmStart.frx":8D40
               Top             =   1935
               Width           =   480
            End
            Begin VB.Image КЭ7 
               Height          =   480
               Index           =   0
               Left            =   270
               Picture         =   "frmStart.frx":904A
               Top             =   765
               Width           =   480
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
               X1              =   5175
               X2              =   4320
               Y1              =   1125
               Y2              =   1125
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   9
               X1              =   1095
               X2              =   555
               Y1              =   1395
               Y2              =   1395
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   10
               X1              =   2340
               X2              =   2340
               Y1              =   2115
               Y2              =   2790
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ3"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   4635
               TabIndex        =   154
               Top             =   1620
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ4"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   153
               Top             =   1710
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ7"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   152
               Top             =   540
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ5"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Index           =   5
               Left            =   7740
               TabIndex        =   151
               Top             =   540
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ6"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   6
               Left            =   6075
               TabIndex        =   150
               Top             =   540
               Width           =   300
            End
            Begin VB.Label Т_после_детандера 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "+17"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   5445
               TabIndex        =   149
               Top             =   2970
               Width           =   585
            End
            Begin VB.Image Термометр 
               Height          =   240
               Index           =   0
               Left            =   3690
               Picture         =   "frmStart.frx":9354
               Top             =   3015
               Width           =   150
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ2"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   990
               TabIndex        =   147
               Top             =   1710
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "КЭ1"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   855
               TabIndex        =   146
               Top             =   3060
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
               Height          =   240
               Index           =   1
               Left            =   180
               Shape           =   3  'Circle
               Top             =   2700
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Shape Муфта 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   2700
               Top             =   1575
               Width           =   675
            End
            Begin VB.Image КЭ4 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "frmStart.frx":9456
               Top             =   1935
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ7 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "frmStart.frx":9760
               Top             =   765
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ3 
               Height          =   480
               Index           =   1
               Left            =   4950
               Picture         =   "frmStart.frx":9A6A
               Top             =   1485
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ6 
               Height          =   480
               Index           =   1
               Left            =   5520
               Picture         =   "frmStart.frx":9D74
               Top             =   400
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image КЭ5 
               Height          =   480
               Index           =   1
               Left            =   7230
               Picture         =   "frmStart.frx":A07E
               Top             =   400
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image Факел 
               Height          =   480
               Index           =   0
               Left            =   1395
               Picture         =   "frmStart.frx":A388
               Top             =   1350
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   2
               X1              =   495
               X2              =   510
               Y1              =   2745
               Y2              =   465
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   8
               X1              =   4530
               X2              =   225
               Y1              =   2805
               Y2              =   2790
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   6
               Index           =   5
               X1              =   5175
               X2              =   5190
               Y1              =   2970
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
               X1              =   4680
               X2              =   4680
               Y1              =   2025
               Y2              =   2925
            End
            Begin VB.Label Т_газ_на_входе 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "+17"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   315
               Left            =   3915
               TabIndex        =   148
               Top             =   2970
               Width           =   585
            End
            Begin VB.Image Факел 
               Height          =   480
               Index           =   1
               Left            =   495
               Picture         =   "frmStart.frx":A692
               Top             =   180
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin Threed.SSCommand SSCmdStart 
            Height          =   915
            Left            =   6975
            TabIndex        =   156
            Top             =   6075
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   1614
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
            BevelWidth      =   4
            Font3D          =   2
            Picture         =   "frmStart.frx":A99C
         End
         Begin Threed.SSCommand cmdSTOP 
            Height          =   1185
            Index           =   0
            Left            =   8145
            TabIndex        =   157
            Top             =   4680
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   2090
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
            BevelWidth      =   4
            Font3D          =   2
            Picture         =   "frmStart.frx":A9B8
         End
         Begin Threed.SSCommand cmdSTOP 
            Height          =   1185
            Index           =   1
            Left            =   7020
            TabIndex        =   158
            Top             =   4680
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   2090
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
            BevelWidth      =   4
            Font3D          =   2
            Picture         =   "frmStart.frx":AE0A
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   840
            Left            =   90
            TabIndex        =   159
            Top             =   3600
            Width           =   9630
            _Version        =   65536
            _ExtentX        =   16986
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
               Caption         =   "Загрузка программы..."
               BeginProperty Font 
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
               TabIndex        =   160
               Top             =   90
               Width           =   9450
            End
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
            Left            =   5130
            TabIndex        =   195
            Top             =   3825
            Width           =   2715
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5355
         Index           =   0
         Left            =   -75000
         TabIndex        =   1
         Top             =   315
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
      Begin Threed.SSCommand SSExit 
         Height          =   1725
         Left            =   -74820
         TabIndex        =   191
         Top             =   5490
         Width           =   9390
         _Version        =   65536
         _ExtentX        =   16563
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
      Begin VB.Label Label16 
         Caption         =   "Текущее время"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74775
         TabIndex        =   194
         Top             =   2115
         Width           =   2190
      End
      Begin VB.Label txtTimeDate 
         Caption         =   "05 сеитября 2022    00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72480
         TabIndex        =   193
         Top             =   2115
         Width           =   4980
      End
      Begin VB.Image Image1 
         Height          =   1365
         Left            =   -74730
         Picture         =   "frmStart.frx":B25C
         Stretch         =   -1  'True
         Top             =   450
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Данный программный продукт разработан лабораторией автоматизации производства Управления ""ЭНЕРГОГАЗРЕМОНТ"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -73650
         TabIndex        =   192
         Top             =   990
         Width           =   6090
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "За сегодня"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -74475
         TabIndex        =   188
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "За месяц"
         BeginProperty Font 
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
         Left            =   -72435
         TabIndex        =   187
         Top             =   450
         Width           =   2205
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "За год"
         BeginProperty Font 
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
         Left            =   -70005
         TabIndex        =   186
         Top             =   450
         Width           =   2205
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "По годам"
         BeginProperty Font 
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
         Left            =   -67665
         TabIndex        =   185
         Top             =   450
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdDanger_Click()
    frmStart.cmdDanger.Visible = False
    ROff A1, 1  'закрыть К 1-6
    'Стоп ДВС, открыть КЭМ4 'TODO проверить комментарий
    toStage_0
    gbStopAGNKS = False
End Sub

Private Sub cmdKKM_Click()
    StatusKKM
    frmKKM.txtKKM.Text = frmStart.Label_Summa.Caption
    frmKKM.lblErrorKKM.Caption = gsErrorKKM    ' = Drvfr.ResultCodeDescription
    frmKKM.lblStatusKKM.Caption = gsРежимККМ    '= Drvfr.ECRModeDescription
    frmKKM.Show vbModal
End Sub

Private Sub cmdOpenStatForm_Click()
    frmSt.Show vbModeless
End Sub

Private Sub cmdStopCarRefueling_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s1       As String
    ROff A1, 191 'Закрыть К5 (пистолет)
    gbDontStat = False         'Можно работать с диском
    gdTime = GetTimeCounter_2
    StopOutput (2)
    StatRS_Insert
    If gbOnlyAkk = True Then
        ROff A1, 127 'Закрыть К6 (Акк)
        frmStart.SSCmdStart.Enabled = True
        gbAkkum = True
        giStage = 1  'Переход на этап ПредПуск()
        giStage1 = 0
        giStage2 = 0
    Else
        'ЗАПРАВЛЯЕМ АККУМУЛЯТОРЫ
        ROn A1, 128     'Открыть КЭ6
    End If

    'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
    frmStart.SSCmdStart.Enabled = True
    gbAkkum = True
End Sub

Private Sub cmdUpdateStat_Click()
   frmStart.MousePointer = vbHourglass
   load_statistic_from_DB
   frmStart.MousePointer = vbArrow
End Sub

' TODO вынести в модуль IO_File
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'реакция на ctrl+alt+home
    'On Error Resume Next
    If ((KeyCode = vbKeyHome) And (Shift = 6)) Then
      setting_gdK
    ElseIf ((KeyCode = vbKeyEnd) And (Shift = 6)) Then
      update_gdK_pass
    End If

End Sub

Private Sub Form_Load()
   Left = 10
   Top = 700
   frmStart.SSTab1.Tab = 3
   ОкноСообщений.BackColor = &HE0E0E0
   ОкноСообщений.ForeColor = &HFF0000
   ОкноСообщений.Caption = "Загрузка программы..."
   Show
   DoEvents
   InitAGNKS
   If isDebug Then
      frmDebug.Show vbModeless
      frmDbDebug.Show
   End If
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


Private Sub SSCmdStart_Click()
    If gbCmdStart = True Then
        gbCmdStart = False
        giStage = 1  'Переход на этап ПредПуск()
        giStage2 = 0
        giStage1 = 0
        ROn A1, 4 'Открыть К1
    Else
        'Если идет заправка аккумуляторов
        If gbAkkum = True Then
            frmЗапрос.Show vbModeless
            gbFrmShow = True
        End If
        giStage = 2
        SSCmdStart.Enabled = False
    End If

End Sub


Private Sub cmdSTOP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      If giStage = 2 Then
         StopOutput (2)
      End If

      ROff A1, 0    'Закрыть К 1-6, ВЫКЛ Реле 2
      ROn A1, 2 ' Стоп ДВС
      toStage_0
      If gbDontStat = True Then
         StatRS_Insert
         gbDontStat = False    'Можно работать с диском
      End If

      Select Case Index
         Case 1 ' Нажата Стоп ДВС
            cmdSTOP(1).Enabled = False
         Case 0 ' Нажата Стоп АГНКС
            cmdSTOP(0).Enabled = False
            'frmStart.Timer2.Enabled = False
            cmdDanger.Visible = True
            ОкноСообщений.Caption = ОстановАГНКС()
      End Select
End Sub

Private Sub SSExit_Click()
    ' TODO Запрос на подтверждение выхода
    ' TODO Возможно тут достаточно проверить Car или giStage
    If gbDontStat = True Then
        StatRS_Insert
        Debug.Print "Сохранено состояние заправки"
    Else
       ' Перед выходом сохраняем моточасы
        saveGMC_in_DB
        Debug.Print "Моточасы сохранены"
    End If
   ' FIXME корректно закрыть базу
   '   StatRS.Close
   '   StatDB.Close
   '   StatWS.Close
   
    DIO_DriverClose
    ISO813_DriverClose
   'TODO Закрыть подключение к ККМ
    ExitWindowsEx 1, 0
    End
End Sub


Private Sub Timer_ДВС_Timer()

    Dim i           As Integer

    'Отображение работы ДВС, компрессора, детандера
    If ОборотыДВС.Caption > 50 Then
        tmrMotor.Enabled = True    'Считать моторесурс
        For i = 0 To 5
            If ДВС(i).Visible Then
                ДВС(i).Visible = False
                If isClutchOn Then
                    Компрессор(i).Visible = False
                End If
                If i < 5 Then
                    ДВС(i + 1).Visible = True
                    If isClutchOn Then
                        Компрессор(i + 1).Visible = True
                    End If
                    Exit For
                Else
                    ДВС(0).Visible = True
                    If isClutchOn Then
                        Компрессор(0).Visible = True
                    End If
                End If
            End If
        Next i
    Else
        tmrMotor.Enabled = False    'Перестать считать моторесурс
    End If


End Sub


' 500 мсек
Private Sub Timer1_Timer()
    Dim i           As Integer
    Dim Dv, Akk, t  As Integer
    Dim Temp        As Double
    Dim s1          As String
    Dim bPSensorOsOk      As Boolean
    Dim v           As Double ' Объем заправленного газа
    Dim s           As String
    nTimer1Counter = nTimer1Counter + 1
    ОпросПлат
    
    For i = 0 To 47
        If gnДатчик(i).Data = 0 Then
            Label1(i).BackColor = &HFF00&
        Else
            Label1(i).BackColor = &HFF
        End If
    Next i

    ' Цикл для суммирования аналоговых значений
    glCounter = glCounter + 1
    For i = 2 To 16
        If gnDif(i) = -1 Then
            sum(i) = -1
        ElseIf sum(i) = -1 Then
            sum(i) = -1
        Else
            sum(i) = sum(i) + gnDif(i)
        End If
    Next i

    If glCounter >= glAver Then    'Если счетчик дошел, то усредняем
        For i = 2 To 16
            If sum(i) = -1 Then
                sum(i) = 0
                Text2(i - 1).ForeColor = &HFF
                Text1(i - 1).Text = "Не исправен"
            Else
                sum(i) = sum(i) / glCounter
                Text2(i - 1).ForeColor = &H80000012
                Text1(i - 1).Text = Format(sum(i), "##0.000")
            End If

            'Для чистового вывода
            Select Case i
                Case 2: Р_вход_АГНКС.Caption = Format(sum(i) / 0.0981, "##0.0")
                Case 6: Р_выход_компр.Caption = Format(sum(i) / 0.0981, "##0.0")
                Case 7
                  Р_аккумулятор.Caption = Format(sum(i) / 0.0981, "##0.0")
                  Аккумулятор.FloodPercent = getP_As_Percent(sum(i))
                Case 8: Т_после_детандера.Caption = Format(sum(i), "#0.0")
                Case 9: Т_газ_на_входе.Caption = Format(sum(i), "#0.0")
                Case 4
                  Р_автобаллон.Caption = Format(sum(i) / 0.0981, "##0.0")
                  Автобаллон.FloodPercent = getP_As_Percent(sum(i))
                Case 14: ОборотыДВС.Caption = Format((sum(i) \ 100) * 100, "###0")
            End Select
            sum(i) = 0
        Next i
        glCounter = 0
    End If

    ShowPict 'Управление изображением

    Наработка_ДВС = "Наработка ДВС " & Format((GMC + MotorCount) / 60, "00") & " ч."
    Панель_Авто.Visible = k5_isOpen

    'Выводим расход на заправку одной машины
    If (gdРасход1 < 0) Then gdРасход1 = 0

    txtKg.Caption = Format(gdРасход1, "0.00")

    v = Round(gdРасход1 / gdPlot, 1) ' Округление до десятых
    ЗаправленоГаза.Caption = Format(v, "0.0")
    Label_Summa.Caption = Format(v * gdPrice, "##0.00")
    txtTime.Caption = formatSecToHHMMSS(gdTime) ' Время заправки

    Label_Avg_Speed_Car = Format(getAvgRefuelingSpeed / gdPlot, "0.00")
    Label_Avg_Left_Time_Car = formatSecToHHMMSS(getLeftRefuelingTime)

    bPSensorOsOk = isAll_PSecnsor_OK
    'Если ручное управление или неисправны датчики
    If (isHandControl) Or (bPSensorOsOk = False) Then
        'Если перешли на ручное управление
        ОкноСообщений.BackColor = &HFF
        ОкноСообщений.ForeColor = &HFFFF&
        ОкноСообщений.Caption = "Ручное управление !!! - программа не управляет процессами !"
        If bPSensorOsOk = False Then
            ОкноСообщений.Caption = "Неисправны датчики !!! - программа не управляет процессами !"
        End If
    Else
        ОкноСообщений.BackColor = &HE0E0E0
        ОкноСообщений.ForeColor = &HFF0000

        Select Case giStage
            Case 0:
                '<<<Заправка>>> 1 Этап
                ОкноСообщений.Caption = ИсхСост
                DoEvents
            Case 1:
                '<<<Заправка>>> 2 Этап
                ОкноСообщений.Caption = ПредПуск
                DoEvents
            Case 2:
                '<<<Заправка>>> 3 Этап
                ОкноСообщений.Caption = Заправка
                DoEvents
            Case 3:
                'Аварийное состояние
                ОкноСообщений.Caption = Danger
                DoEvents
        End Select
    End If

    'Проверка аварийных датчиков
    s1 = ""
    s1 = Verify_Damage
    If s1 <> "" Then
        ОкноСообщений.BackColor = &HFF
        ОкноСообщений.ForeColor = &HFFFF&
        ОкноСообщений.Caption = ОкноСообщений.Caption + " " + s1
    Else
        ОкноСообщений.BackColor = &HE0E0E0
        ОкноСообщений.ForeColor = &HFF0000
    End If

   If (gbCmdStart) Then
      frmStart.SSCmdStart.Caption = "Пуск АГНКС"
   Else
      frmStart.SSCmdStart.Caption = "ЗАПРАВКА"
   End If
End Sub

Private Sub Timer2_Timer()
    txtTimeDate = Format(Now, "dd.mmmm.yyyy    hh:nn:ss")
End Sub

Private Sub tmrMotor_Timer()
    MotorCount = MotorCount + 1
End Sub



Public Sub ShowPict()
    With frmStart
        'Статус кранов
        .КЭ1(0).Visible = Not (k1_isOpen)
        .КЭ1(1).Visible = k1_isOpen

        .КЭ2(0).Visible = Not (k2_isOpen)
        .КЭ2(1).Visible = k2_isOpen
        .Факел(0).Visible = k2_isOpen

        .КЭ3(0).Visible = Not (k3_isOpen)
        .КЭ3(1).Visible = k3_isOpen

        .КЭ4(0).Visible = Not (k4_isOpen)
        .КЭ4(1).Visible = k4_isOpen

        .КЭ5(0).Visible = Not (k5_isOpen)
        .КЭ5(1).Visible = k5_isOpen

        .КЭ6(0).Visible = Not (k6_isOpen)
        .КЭ6(1).Visible = k6_isOpen
 
        .КЭ7(0).Visible = Not (k7_isOpen)
        .КЭ7(1).Visible = k7_isOpen
        .Факел(1).Visible = k7_isOpen

        ' TODO переделать на Visible
        If isClutchOn Then
            .Муфта.BackColor = &HFF&
        Else
            .Муфта.BackColor = &HC0C0C0
        End If

        'Повышение температуры охлаждения ДВС
        'If gnДатчик(33).Data = 1 Then
        'Else
        'End If

        'Пожар в отсеке ДВС
        'If gnДатчик(45).Data = 1 Then
        'Else
        'End If

        'Пожар в технологическом отсеке
        'If isFireTech Then
        'Else
        'End If

        'Газ в отсеке ДВС 10%
        'If (gnДатчик(41).Data = 1) Then
        'Else
        'End If
        'Газ в отсеке ДВС 20%
        'If (gnДатчик(42).Data = 1) Then
        'Else
        'End If

        'Газ в технологическом отсеке 10%
        'If (gnДатчик(43).Data = 1) Then
        'Else
        'End If

        'Газ в технологическом отсеке 20%
        'If (gnДатчик(44).Data = 1) Then
        'Else
        'End If
    End With
End Sub
