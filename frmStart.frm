VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����   ��-40  ""�����"""
   ClientHeight    =   7395
   ClientLeft      =   3555
   ClientTop       =   2505
   ClientWidth     =   9855
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   493
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   45
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
      TabCaption(0)   =   "����������"
      TabPicture(0)   =   "frmStart.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "����������"
      TabPicture(1)   =   "frmStart.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "���������"
      TabPicture(2)   =   "frmStart.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdUpdatePassword"
      Tab(2).Control(1)=   "cmdUpdateGMC"
      Tab(2).Control(2)=   "cmdUpdatePrice"
      Tab(2).Control(3)=   "cmdUpdatePlot"
      Tab(2).Control(4)=   "cmdUpdatePC"
      Tab(2).Control(5)=   "SSExit"
      Tab(2).Control(6)=   "lblAppVersion"
      Tab(2).Control(7)=   "Label20"
      Tab(2).Control(8)=   "lbl_gnPlot"
      Tab(2).Control(9)=   "Label19"
      Tab(2).Control(10)=   "Label18"
      Tab(2).Control(11)=   "Label17"
      Tab(2).Control(12)=   "txtTimeDate"
      Tab(2).Control(13)=   "Label4"
      Tab(2).Control(14)=   "Label10"
      Tab(2).Control(15)=   "lblPC"
      Tab(2).Control(16)=   "Label16"
      Tab(2).Control(17)=   "Image1"
      Tab(2).Control(18)=   "Shape3"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "�����"
      TabPicture(3)   =   "frmStart.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "������"
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
      Begin VB.CommandButton cmdUpdatePassword 
         Caption         =   "��������"
         Height          =   330
         Left            =   -71175
         TabIndex        =   131
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdUpdateGMC 
         Caption         =   "��������"
         Height          =   330
         Left            =   -71175
         TabIndex        =   129
         Top             =   3060
         Width           =   1140
      End
      Begin VB.CommandButton cmdUpdatePrice 
         Caption         =   "��������"
         Height          =   330
         Left            =   -71175
         TabIndex        =   128
         Top             =   2655
         Width           =   1140
      End
      Begin VB.CommandButton cmdUpdatePlot 
         Caption         =   "��������"
         Height          =   330
         Left            =   -71175
         TabIndex        =   127
         Top             =   2250
         Width           =   1140
      End
      Begin VB.CommandButton cmdUpdatePC 
         Caption         =   "��������"
         Height          =   330
         Left            =   -71175
         TabIndex        =   126
         Top             =   1845
         Width           =   1140
      End
      Begin VB.CommandButton cmdOpenStatForm 
         Caption         =   "����������"
         Height          =   690
         Left            =   -74775
         TabIndex        =   113
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
         Top             =   810
         Width           =   2200
      End
      Begin VB.CommandButton cmdUpdateStat 
         Caption         =   "�������� ������"
         Height          =   690
         Left            =   -67620
         TabIndex        =   104
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
         TabIndex        =   55
         Top             =   315
         Width           =   9795
         Begin Threed.SSCommand cmdDanger 
            Height          =   2310
            Left            =   3690
            TabIndex        =   85
            Top             =   4680
            Visible         =   0   'False
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   4075
            _StockProps     =   78
            Caption         =   "������"
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
            TabIndex        =   95
            Top             =   4590
            Width           =   3165
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "� ������"
               Height          =   195
               Left            =   180
               TabIndex        =   114
               Top             =   1125
               Width           =   660
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "����� ��������"
               Height          =   195
               Left            =   180
               TabIndex        =   103
               Top             =   1485
               Width           =   1260
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "������ � / �. ���."
               Height          =   195
               Left            =   180
               TabIndex        =   102
               Top             =   900
               Width           =   1380
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
               TabIndex        =   101
               Top             =   990
               Width           =   960
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "�������� ��������"
               Height          =   195
               Left            =   180
               TabIndex        =   100
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
               Top             =   270
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "�����������, ��."
               Height          =   195
               Left            =   180
               TabIndex        =   96
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "����� ��������"
            Height          =   2400
            Left            =   90
            TabIndex        =   86
            Top             =   4590
            Width           =   3570
            Begin Threed.SSCommand cmdKKM 
               Height          =   750
               Left            =   2655
               TabIndex        =   87
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
               Caption         =   "���� ���."
               Height          =   195
               Left            =   135
               TabIndex        =   94
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Price 
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
               TabIndex        =   93
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
               TabIndex        =   92
               Top             =   1485
               Width           =   2490
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "����� ���."
               Height          =   195
               Left            =   135
               TabIndex        =   91
               Top             =   1215
               Width           =   855
            End
            Begin VB.Label �������������� 
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
               TabIndex        =   90
               Top             =   540
               Width           =   1230
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "� / �"
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
               TabIndex        =   89
               Top             =   675
               Width           =   525
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "3"
               Height          =   195
               Left            =   3195
               TabIndex        =   88
               Top             =   585
               Width           =   90
            End
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   3450
            Left            =   45
            TabIndex        =   56
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
            Begin VB.Timer tmrDvsCompressorAnimation 
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
               Interval        =   60000
               Left            =   8790
               Top             =   2835
            End
            Begin Threed.SSPanel �����_��� 
               Height          =   1500
               Left            =   1800
               TabIndex        =   57
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
               Begin Threed.SSPanel ���������� 
                  Height          =   375
                  Left            =   135
                  TabIndex        =   58
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
                  Caption         =   "����."
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
                  TabIndex        =   59
                  Top             =   1170
                  Width           =   510
               End
               Begin VB.Image �����������_��� 
                  Height          =   480
                  Left            =   360
                  Picture         =   "frmStart.frx":00B8
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   300
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   0
                  Left            =   225
                  Picture         =   "frmStart.frx":02BA
                  Top             =   540
                  Width           =   600
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   1
                  Left            =   225
                  Picture         =   "frmStart.frx":065C
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   2
                  Left            =   225
                  Picture         =   "frmStart.frx":09FE
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   3
                  Left            =   225
                  Picture         =   "frmStart.frx":0DA0
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   4
                  Left            =   225
                  Picture         =   "frmStart.frx":1142
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Image ��� 
                  Height          =   600
                  Index           =   5
                  Left            =   225
                  Picture         =   "frmStart.frx":14E4
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   600
               End
            End
            Begin Threed.SSPanel �����_����� 
               Height          =   1500
               Left            =   3105
               TabIndex        =   60
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
               Begin Threed.SSPanel �_�����_����� 
                  Height          =   375
                  Left            =   135
                  TabIndex        =   61
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
                  Caption         =   "�����."
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
                  TabIndex        =   62
                  Top             =   1170
                  Width           =   660
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   0
                  Left            =   180
                  Picture         =   "frmStart.frx":1886
                  Top             =   540
                  Width           =   900
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   1
                  Left            =   180
                  Picture         =   "frmStart.frx":1E08
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   2
                  Left            =   180
                  Picture         =   "frmStart.frx":238A
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   3
                  Left            =   180
                  Picture         =   "frmStart.frx":290C
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   4
                  Left            =   180
                  Picture         =   "frmStart.frx":2E8E
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Image ���������� 
                  Height          =   600
                  Index           =   5
                  Left            =   180
                  Picture         =   "frmStart.frx":3410
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   900
               End
            End
            Begin Threed.SSPanel ������_���� 
               Height          =   1890
               Left            =   7230
               TabIndex        =   63
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
               Begin Threed.SSPanel �_���������� 
                  Height          =   375
                  Left            =   630
                  TabIndex        =   64
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
               Begin Threed.SSPanel ���������� 
                  Height          =   1725
                  Left            =   90
                  TabIndex        =   65
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
                  TabIndex        =   79
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
               TabIndex        =   66
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
               Begin Threed.SSPanel �_����������� 
                  Height          =   375
                  Left            =   585
                  TabIndex        =   67
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
               Begin Threed.SSPanel ����������� 
                  Height          =   1725
                  Left            =   90
                  TabIndex        =   68
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
            Begin VB.Label ���������_��� 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Caption         =   "��������� ��� 999999 �."
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   1305
               TabIndex        =   118
               Top             =   540
               Width           =   2130
            End
            Begin VB.Image ��������� 
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
            Begin VB.Label �_����_����� 
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
               TabIndex        =   69
               Top             =   2970
               Width           =   495
            End
            Begin VB.Image ��2 
               Height          =   480
               Index           =   0
               Left            =   1170
               Picture         =   "frmStart.frx":77FA
               Top             =   1935
               Width           =   480
            End
            Begin VB.Image ��2 
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
            Begin VB.Image ��1 
               Height          =   480
               Index           =   0
               Left            =   765
               Picture         =   "frmStart.frx":7E0E
               Top             =   2565
               Width           =   480
            End
            Begin VB.Image ��1 
               Height          =   480
               Index           =   1
               Left            =   720
               Picture         =   "frmStart.frx":8118
               Top             =   2565
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ��5 
               Height          =   480
               Index           =   0
               Left            =   7230
               Picture         =   "frmStart.frx":8422
               Top             =   400
               Width           =   480
            End
            Begin VB.Image ��6 
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
            Begin VB.Image ��3 
               Height          =   480
               Index           =   0
               Left            =   4950
               Picture         =   "frmStart.frx":8A36
               Top             =   1485
               Width           =   480
            End
            Begin VB.Image ��4 
               Height          =   480
               Index           =   0
               Left            =   270
               Picture         =   "frmStart.frx":8D40
               Top             =   1935
               Width           =   480
            End
            Begin VB.Image ��7 
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
               Caption         =   "��3"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   4635
               TabIndex        =   78
               Top             =   1620
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��4"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   77
               Top             =   1710
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��7"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   76
               Top             =   540
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��5"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Index           =   5
               Left            =   7740
               TabIndex        =   75
               Top             =   540
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��6"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   6
               Left            =   6075
               TabIndex        =   74
               Top             =   540
               Width           =   300
            End
            Begin VB.Label �_�����_��������� 
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
               TabIndex        =   73
               Top             =   2970
               Width           =   585
            End
            Begin VB.Image ��������� 
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
               Caption         =   "��2"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   990
               TabIndex        =   71
               Top             =   1710
               Width           =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��1"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   855
               TabIndex        =   70
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
            Begin VB.Shape ����� 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   2700
               Top             =   1575
               Width           =   675
            End
            Begin VB.Image ��4 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "frmStart.frx":9456
               Top             =   1935
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ��7 
               Height          =   480
               Index           =   1
               Left            =   270
               Picture         =   "frmStart.frx":9760
               Top             =   765
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ��3 
               Height          =   480
               Index           =   1
               Left            =   4950
               Picture         =   "frmStart.frx":9A6A
               Top             =   1485
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ��6 
               Height          =   480
               Index           =   1
               Left            =   5520
               Picture         =   "frmStart.frx":9D74
               Top             =   400
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ��5 
               Height          =   480
               Index           =   1
               Left            =   7230
               Picture         =   "frmStart.frx":A07E
               Top             =   400
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image ����� 
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
            Begin VB.Label �_���_��_����� 
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
               TabIndex        =   72
               Top             =   2970
               Width           =   585
            End
            Begin VB.Image ����� 
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
            TabIndex        =   80
            Top             =   6075
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   1614
            _StockProps     =   78
            Caption         =   "���� �����"
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
            TabIndex        =   81
            Top             =   4680
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   2090
            _StockProps     =   78
            Caption         =   "�����"
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
            TabIndex        =   82
            Top             =   4680
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   2090
            _StockProps     =   78
            Caption         =   "���"
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
            TabIndex        =   83
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
            Begin VB.Label ������������� 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "�������� ���������..."
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
               TabIndex        =   84
               Top             =   90
               Width           =   9450
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6975
         Index           =   1
         Left            =   -74910
         TabIndex        =   22
         Top             =   360
         Width           =   9615
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   54
            Text            =   "Text2"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   53
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
            TabIndex        =   52
            Text            =   "Text2"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   51
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   50
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
            TabIndex        =   49
            Text            =   "Text2"
            Top             =   1920
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   48
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
            TabIndex        =   47
            Text            =   "Text2"
            Top             =   2280
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   46
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
            TabIndex        =   45
            Text            =   "Text2"
            Top             =   2640
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   44
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
            TabIndex        =   43
            Text            =   "Text2"
            Top             =   3000
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   2400
            TabIndex        =   42
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
            TabIndex        =   41
            Text            =   "Text2"
            Top             =   3360
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   7320
            TabIndex        =   40
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
            TabIndex        =   39
            Text            =   "Text2"
            Top             =   840
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   7320
            TabIndex        =   38
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
            TabIndex        =   37
            Text            =   "Text2"
            Top             =   1200
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   7320
            TabIndex        =   36
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
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   1560
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   7320
            TabIndex        =   34
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
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   1920
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   7320
            TabIndex        =   32
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
            TabIndex        =   31
            Text            =   "Text2"
            Top             =   2280
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   7320
            TabIndex        =   30
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
            TabIndex        =   29
            Text            =   "Text2"
            Top             =   2640
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   7320
            TabIndex        =   28
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
            TabIndex        =   27
            Text            =   "Text2"
            Top             =   3000
            Width           =   2050
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   7320
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
            Text            =   "Text2"
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6975
         Index           =   0
         Left            =   -74865
         TabIndex        =   1
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame4 
            Height          =   2850
            Left            =   5310
            TabIndex        =   133
            Top             =   180
            Width           =   3885
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   16
               Left            =   1980
               TabIndex        =   161
               Top             =   630
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   16
               Left            =   2340
               TabIndex        =   160
               Top             =   675
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   17
               Left            =   1980
               TabIndex        =   159
               Top             =   990
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   17
               Left            =   2340
               TabIndex        =   158
               Top             =   1035
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   18
               Left            =   1980
               TabIndex        =   157
               Top             =   1350
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   18
               Left            =   2340
               TabIndex        =   156
               Top             =   1395
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   19
               Left            =   1980
               TabIndex        =   155
               Top             =   1710
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   19
               Left            =   2340
               TabIndex        =   154
               Top             =   1755
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   20
               Left            =   1980
               TabIndex        =   153
               Top             =   2070
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   20
               Left            =   2340
               TabIndex        =   152
               Top             =   2115
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   21
               Left            =   1980
               TabIndex        =   151
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   21
               Left            =   2340
               TabIndex        =   150
               Top             =   315
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   1
               Left            =   135
               TabIndex        =   149
               Top             =   2430
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   1
               Left            =   495
               TabIndex        =   148
               Top             =   2475
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   23
               Left            =   1980
               TabIndex        =   147
               Top             =   2430
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   23
               Left            =   2340
               TabIndex        =   146
               Top             =   2475
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   26
               Left            =   135
               TabIndex        =   145
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   26
               Left            =   495
               TabIndex        =   144
               Top             =   315
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   27
               Left            =   135
               TabIndex        =   143
               Top             =   630
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   27
               Left            =   495
               TabIndex        =   142
               Top             =   675
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   28
               Left            =   135
               TabIndex        =   141
               Top             =   990
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   28
               Left            =   495
               TabIndex        =   140
               Top             =   1035
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   29
               Left            =   135
               TabIndex        =   139
               Top             =   1350
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   29
               Left            =   495
               TabIndex        =   138
               Top             =   1395
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   30
               Left            =   135
               TabIndex        =   137
               Top             =   1710
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   30
               Left            =   495
               TabIndex        =   136
               Top             =   1755
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   31
               Left            =   135
               TabIndex        =   135
               Top             =   2070
               Width           =   255
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "� �� ����� �����"
               Height          =   195
               Index           =   31
               Left            =   495
               TabIndex        =   134
               Top             =   2115
               Width           =   1380
            End
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   209
            Top             =   810
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   208
            Top             =   855
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   8
            Left            =   135
            TabIndex        =   207
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   495
            TabIndex        =   206
            Top             =   1575
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   9
            Left            =   135
            TabIndex        =   205
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   9
            Left            =   495
            TabIndex        =   204
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   24
            Left            =   135
            TabIndex        =   203
            Top             =   1170
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   24
            Left            =   495
            TabIndex        =   202
            Top             =   1215
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   25
            Left            =   2700
            TabIndex        =   201
            Top             =   450
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   25
            Left            =   3060
            TabIndex        =   200
            Top             =   495
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   32
            Left            =   135
            TabIndex        =   199
            Top             =   1890
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   32
            Left            =   495
            TabIndex        =   198
            Top             =   1935
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   33
            Left            =   2700
            TabIndex        =   197
            Top             =   1890
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   33
            Left            =   3060
            TabIndex        =   196
            Top             =   1935
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   34
            Left            =   2700
            TabIndex        =   195
            Top             =   2250
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   34
            Left            =   3060
            TabIndex        =   194
            Top             =   2295
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   35
            Left            =   2700
            TabIndex        =   193
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   35
            Left            =   3060
            TabIndex        =   192
            Top             =   1575
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   36
            Left            =   2700
            TabIndex        =   191
            Top             =   810
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   36
            Left            =   3060
            TabIndex        =   190
            Top             =   855
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   37
            Left            =   2700
            TabIndex        =   189
            Top             =   1170
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   37
            Left            =   3060
            TabIndex        =   188
            Top             =   1215
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   15
            Left            =   135
            TabIndex        =   187
            Top             =   450
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   15
            Left            =   495
            TabIndex        =   186
            Top             =   495
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   5445
            TabIndex        =   185
            Top             =   4275
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   2
            Left            =   5760
            TabIndex        =   184
            Top             =   4320
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   5445
            TabIndex        =   183
            Top             =   4635
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   3
            Left            =   5760
            TabIndex        =   182
            Top             =   4680
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   5445
            TabIndex        =   181
            Top             =   4995
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   4
            Left            =   5760
            TabIndex        =   180
            Top             =   5040
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   5445
            TabIndex        =   179
            Top             =   5355
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   5
            Left            =   5760
            TabIndex        =   178
            Top             =   5400
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   6
            Left            =   5445
            TabIndex        =   177
            Top             =   5715
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   6
            Left            =   5760
            TabIndex        =   176
            Top             =   5760
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   7
            Left            =   5445
            TabIndex        =   175
            Top             =   6075
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   7
            Left            =   5760
            TabIndex        =   174
            Top             =   6120
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   10
            Left            =   7290
            TabIndex        =   173
            Top             =   4275
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   10
            Left            =   7650
            TabIndex        =   172
            Top             =   4320
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   11
            Left            =   7290
            TabIndex        =   171
            Top             =   4635
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   11
            Left            =   7650
            TabIndex        =   170
            Top             =   4680
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   12
            Left            =   7290
            TabIndex        =   169
            Top             =   4995
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   12
            Left            =   7650
            TabIndex        =   168
            Top             =   5400
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   13
            Left            =   7290
            TabIndex        =   167
            Top             =   5355
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   13
            Left            =   7650
            TabIndex        =   166
            Top             =   5040
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   14
            Left            =   7290
            TabIndex        =   165
            Top             =   5715
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   14
            Left            =   7650
            TabIndex        =   164
            Top             =   5760
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   22
            Left            =   7290
            TabIndex        =   163
            Top             =   6075
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   22
            Left            =   7650
            TabIndex        =   162
            Top             =   6120
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   47
            Left            =   495
            TabIndex        =   21
            Top             =   6165
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   47
            Left            =   135
            TabIndex        =   20
            Top             =   6120
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   46
            Left            =   495
            TabIndex        =   19
            Top             =   5805
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   46
            Left            =   135
            TabIndex        =   18
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   45
            Left            =   495
            TabIndex        =   17
            Top             =   5445
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   45
            Left            =   135
            TabIndex        =   16
            Top             =   5400
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   44
            Left            =   495
            TabIndex        =   15
            Top             =   5085
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   44
            Left            =   135
            TabIndex        =   14
            Top             =   5040
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   43
            Left            =   495
            TabIndex        =   13
            Top             =   4725
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   43
            Left            =   135
            TabIndex        =   12
            Top             =   4680
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   42
            Left            =   495
            TabIndex        =   11
            Top             =   4365
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   42
            Left            =   135
            TabIndex        =   10
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   41
            Left            =   495
            TabIndex        =   9
            Top             =   4005
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   41
            Left            =   135
            TabIndex        =   8
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   40
            Left            =   495
            TabIndex        =   7
            Top             =   3645
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   40
            Left            =   135
            TabIndex        =   6
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   39
            Left            =   495
            TabIndex        =   5
            Top             =   3285
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   39
            Left            =   135
            TabIndex        =   4
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "� �� ����� �����"
            Height          =   195
            Index           =   38
            Left            =   495
            TabIndex        =   3
            Top             =   2925
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   38
            Left            =   135
            TabIndex        =   2
            Top             =   2880
            Width           =   255
         End
      End
      Begin Threed.SSCommand SSExit 
         Height          =   1320
         Left            =   -74820
         TabIndex        =   115
         Top             =   5895
         Width           =   9390
         _Version        =   65536
         _ExtentX        =   16563
         _ExtentY        =   2328
         _StockProps     =   78
         Caption         =   "�����"
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
      Begin VB.Label lblAppVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "appVersion"
         Height          =   240
         Left            =   -74145
         TabIndex        =   132
         Top             =   1125
         Width           =   1860
      End
      Begin VB.Label Label20 
         Caption         =   "������"
         Height          =   285
         Left            =   -74415
         TabIndex        =   130
         Top             =   3555
         Width           =   1545
      End
      Begin VB.Label lbl_gnPlot 
         Caption         =   "0.000"
         Height          =   240
         Left            =   -71985
         TabIndex        =   125
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label Label19 
         Caption         =   "���� ����"
         Height          =   285
         Left            =   -74415
         TabIndex        =   124
         Top             =   2745
         Width           =   1545
      End
      Begin VB.Label Label18 
         Caption         =   "��������� ���"
         Height          =   285
         Left            =   -74415
         TabIndex        =   123
         Top             =   3150
         Width           =   1950
      End
      Begin VB.Label Label17 
         Caption         =   "��������� ����:"
         Height          =   285
         Left            =   -74415
         TabIndex        =   122
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label txtTimeDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "05 �������� 2022    00:00:00"
         Height          =   285
         Left            =   -68610
         TabIndex        =   116
         Top             =   870
         Width           =   3210
      End
      Begin VB.Label Label4 
         Caption         =   "����������� �����������:"
         Height          =   285
         Left            =   -74415
         TabIndex        =   121
         Top             =   1935
         Width           =   2310
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "������ ����������� ������� ���������� ������������ ������������� ������������ ���������� ""���������������"""
         Height          =   465
         Left            =   -74145
         TabIndex        =   120
         Top             =   675
         Width           =   5325
      End
      Begin VB.Label lblPC 
         Caption         =   "0.000"
         Height          =   240
         Left            =   -71985
         TabIndex        =   119
         Top             =   1935
         Width           =   780
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "��������� �����"
         Height          =   285
         Left            =   -67575
         TabIndex        =   117
         Top             =   675
         Width           =   2190
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   -74730
         Picture         =   "frmStart.frx":B25C
         Stretch         =   -1  'True
         Top             =   450
         Width           =   510
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "�� �������"
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
         TabIndex        =   112
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "�� �����"
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
         TabIndex        =   111
         Top             =   450
         Width           =   2205
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "�� ���"
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
         TabIndex        =   110
         Top             =   450
         Width           =   2205
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "�� �����"
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
         TabIndex        =   109
         Top             =   450
         Width           =   2250
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   1140
         Left            =   -74955
         Top             =   360
         Width           =   9690
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
   frmStart.SSTab1.Tab = 3
   �������������.BackColor = &HE0E0E0
   �������������.ForeColor = &HFF0000
   �������������.Caption = "�������� ���������..."
   lblAppVersion.Caption = "������ " & App.Major & "." & App.Minor & "." & App.Revision
   
   Show
   DoEvents
   
   InitAGNKS
   If isDebug Then
      frmDebug.Show vbModeless
      'frmDbDebug.Show
   End If
End Sub

Private Sub SSCmdStart_Click()
    If gbCmdStart = True Then
        gbCmdStart = False
        giStage = 1  '������� �� ���� ��������()
        giStage2 = 0
        giStage1 = 0
        ROn A1, 4 '������� �1
    Else
        '���� ���� �������� �������������
        If gbAkkum = True Then
            frm������.Show vbModeless
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

      ROff A1, 0    '������� � 1-6, ���� ���� 2
      ROn A1, 2 ' ���� ���
      toStage_0
      If gbDontStat = True Then
         StatRS_Insert
         gbDontStat = False    '����� �������� � ������
      End If

      Select Case Index
         Case 1 ' ������ ���� ���
            cmdSTOP(1).Enabled = False
         Case 0 ' ������ ���� �����
            cmdSTOP(0).Enabled = False
            'frmStart.Timer2.Enabled = False
            cmdDanger.Visible = True
            �������������.Caption = ������������()
      End Select
End Sub

Private Sub cmdStopCarRefueling_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s1       As String
    ROff A1, 191 '������� �5 (��������)
    gbDontStat = False         '����� �������� � ������
    StopOutput (2)
    StatRS_Insert
    If gbOnlyAkk = True Then
        ROff A1, 127 '������� �6 (���)
        frmStart.SSCmdStart.Enabled = True
        gbAkkum = True
        giStage = 1  '������� �� ���� ��������()
        giStage1 = 0
        giStage2 = 0
    Else
        '���������� ������������
        ROn A1, 128     '������� ��6
    End If

    '��������� ��������� �������� ���������� �� ����� �������� �������������
    frmStart.SSCmdStart.Enabled = True
    gbAkkum = True
End Sub

Private Sub cmdDanger_Click()
    frmStart.cmdDanger.Visible = False
    ROff A1, 1  '������� � 1-6
    '���� ���, ������� ���4 'TODO ��������� �����������
    toStage_0
    gbStopAGNKS = False
End Sub

Private Sub cmdKKM_Click()
    StatusKKM
    frmKKM.txtKKM.Text = frmStart.Label_Summa.Caption
    frmKKM.lblErrorKKM.Caption = gsErrorKKM    ' = Drvfr.ResultCodeDescription
    frmKKM.lblStatusKKM.Caption = gs��������    '= Drvfr.ECRModeDescription
    frmKKM.Show vbModal
End Sub

Private Sub cmdOpenStatForm_Click()
    frmSt.Show vbModeless
End Sub

Private Sub cmdUpdateStat_Click()
   frmStart.MousePointer = vbHourglass
   load_statistic_from_DB
   frmStart.MousePointer = vbArrow
End Sub

Private Sub cmdUpdatePC_Click()
    updatePC
End Sub

Private Sub cmdUpdatePlot_Click()
    updatePlot
End Sub

Private Sub cmdUpdatePrice_Click()
    updatePrice
End Sub

Private Sub cmdUpdateGMC_Click()
    updateGMC
End Sub

Private Sub cmdUpdatePassword_Click()
    updatePWD
End Sub
Private Sub Label1_Click(Index As Integer)
    'Dim Maska As Integer
    'Dim rez As Long
    'Dim i As Integer
    'Dim Temp As Integer
    ' ������� ����
    '   Maska = 1
    ' ��� ����� A0
    '  If (Index >= 0 And Index < 8) Then
    '   For i = 1 To Index
    '     Maska = Maska * 2
    '   Next i
    '     Temp = gn48DIO(0) '��������� ��������� ����� A0
    '     Temp = Temp Xor Maska

    '!!!!��� ���������
    'rez = W_48DIO_DO(A0, Temp)
    '     If gn������(Index).Data = 0 Then
    '       ROn A0, Maska
    '     Else
    '       ROff A0, Maska Xor 255
    '     End If

    '     gn48DIO(0) = Temp


    ' ��� ����� A1
    '   ElseIf (Index > 23 And Index < 32) Then
    '     For i = 1 To Index - 24
    '     Maska = Maska * 2
    '   Next i
    '     Temp = gn48DIO(3) '��������� ��������� ����� A1
    '     Temp = Temp Xor Maska

    '!!!��� ���������
    'rez = W_48DIO_DO(A1, Temp)
    '     If gn������(Index).Data = 0 Then
    '       ROn A1, Maska
    '     Else
    '       ROff A1, Maska Xor 255
    '     End If

    '     gn48DIO(3) = Temp

    '   End If


End Sub

Private Sub SSExit_Click()
    ' TODO ������ �� ������������� ������
    ' TODO �������� ��� ���������� ��������� Car ��� giStage
    If gbDontStat = True Then
        StatRS_Insert
        Debug.Print "��������� ��������� ��������"
    Else
       ' ����� ������� ��������� ��������
        saveGMC_in_DB
        Debug.Print "�������� ���������"
    End If
   ' FIXME ��������� ������� ����
   '   StatRS.Close
   '   StatDB.Close
   '   StatWS.Close
   
    DIO_DriverClose
    ISO813_DriverClose
   'TODO ������� ����������� � ���
    ExitWindowsEx 1, 0
    End
End Sub



' 500 ����
Private Sub Timer1_Timer()
    Dim i           As Integer
    Dim Dv, Akk, t  As Integer
    Dim Temp        As Double
    Dim s1          As String
    Dim bPSensorOsOk      As Boolean
    Dim v           As Double ' ����� ������������� ����
    Dim s           As String
    nTimer1Counter = nTimer1Counter + 1
    ���������
    
    For i = 0 To 47
      If gn������(i).Note <> "������" Then
         If gn������(i).Data = 0 Then
               Label1(i).BackColor = &HFF00&
         Else
               Label1(i).BackColor = &HFF
         End If
      Else
         Label1(i).BackColor = &HC0C0C0
      End If
    Next i

    ' ���� ��� ������������ ���������� ��������
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

    If glCounter >= glAver Then    '���� ������� �����, �� ���������
        For i = 2 To 16
            If sum(i) = -1 Then
                sum(i) = 0
                Text2(i - 1).ForeColor = &HFF
                Text1(i - 1).Text = "�� ��������"
            Else
                sum(i) = sum(i) / glCounter
                Text2(i - 1).ForeColor = &H80000012
                Text1(i - 1).Text = Format(sum(i), "##0.000")
            End If

            '��� ��������� ������
            Select Case i
                Case 2: �_����_�����.Caption = Format(sum(i) / 0.0981, "##0.0")
                Case 6: �_�����_�����.Caption = Format(sum(i) / 0.0981, "##0.0")
                Case 7
                  �_�����������.Caption = Format(sum(i) / 0.0981, "##0.0")
                  �����������.FloodPercent = getP_As_Percent(sum(i))
                Case 8: �_�����_���������.Caption = Format(sum(i), "#0.0")
                Case 9: �_���_��_�����.Caption = Format(sum(i), "#0.0")
                Case 4
                  �_����������.Caption = Format(sum(i) / 0.0981, "##0.0")
                  ����������.FloodPercent = getP_As_Percent(sum(i))
                Case 14: ����������.Caption = Format((sum(i) \ 100) * 100, "###0")
            End Select
            sum(i) = 0
        Next i
        glCounter = 0
    End If

    mnemonicScheme_Tic '���������� ������������

    ���������_��� = "��������� ��� " & Format((GMC + tmrMotorCounter) / 60, "00") & " �."
    

    '������� ������ �� �������� ����� ������
    txtKg.Caption = Format(gd��2, "0.00")

    v = Round(gd��2 / agnks�onfig.plot, 1) ' ���������� �� �������
    ��������������.Caption = Format(v, "0.0")
    Label_Summa.Caption = Format(v * agnks�onfig.Price, "##0.00")
    txtTime.Caption = formatSecToHHMMSS(GetTimeCounter_2) ' ����� ��������

    Label_Avg_Speed_Car = Format((GetMassExpense_2 * 60) / agnks�onfig.plot, "0.00")
    Label_Avg_Left_Time_Car = formatSecToHHMMSS(getLeftRefuelingTime)

    bPSensorOsOk = isAll_PSecnsor_OK
    '���� ������ ���������� ��� ���������� �������
    If (isHandControl) Or (bPSensorOsOk = False) Then
        '���� ������� �� ������ ����������
        �������������.BackColor = &HFF
        �������������.ForeColor = &HFFFF&
        �������������.Caption = "������ ���������� !!! - ��������� �� ��������� ���������� !"
        If bPSensorOsOk = False Then
            �������������.Caption = "���������� ������� !!! - ��������� �� ��������� ���������� !"
        End If
    Else
        �������������.BackColor = &HE0E0E0
        �������������.ForeColor = &HFF0000

        Select Case giStage
            Case 0:
                '<<<��������>>> 1 ����
                �������������.Caption = �������
                DoEvents
            Case 1:
                '<<<��������>>> 2 ����
                �������������.Caption = ��������
                DoEvents
            Case 2:
                '<<<��������>>> 3 ����
                �������������.Caption = ��������
                DoEvents
            Case 3:
                '��������� ���������
                Danger
                DoEvents
        End Select
    End If

    '�������� ��������� ��������
    s1 = ""
    s1 = Verify_Damage
    If s1 <> "" Then
        �������������.BackColor = &HFF
        �������������.ForeColor = &HFFFF&
        �������������.Caption = �������������.Caption + " " + s1
    Else
        �������������.BackColor = &HE0E0E0
        �������������.ForeColor = &HFF0000
    End If

   If (gbCmdStart) Then
      frmStart.SSCmdStart.Caption = "���� �����"
   Else
      frmStart.SSCmdStart.Caption = "��������"
   End If
End Sub

Private Sub Timer2_Timer()
    txtTimeDate = Format(Now, "dd.mmmm.yyyy    hh:nn:ss")
End Sub



' Interval 75 ms
Private Sub tmrDvsCompressorAnimation_Timer()
    Dim i           As Integer
    '����������� ������ ���, �����������, ���������
    If getDVS_RPM > 100 Then
        tmrMotor.Enabled = True    '������� ����������
        For i = 0 To 5
            If ���(i).Visible Then
                ���(i).Visible = False
                If isClutchOn Then
                    ����������(i).Visible = False
                End If
                If i < 5 Then
                    ���(i + 1).Visible = True
                    If isClutchOn Then
                        ����������(i + 1).Visible = True
                    End If
                    Exit For
                Else
                    ���(0).Visible = True
                    If isClutchOn Then
                        ����������(0).Visible = True
                    End If
                End If
            End If
        Next i
    Else
        tmrMotor.Enabled = False    '��������� ������� ����������
    End If
End Sub

' Interval 60 000 ms
Private Sub tmrMotor_Timer()
    tmrMotorCounter = tmrMotorCounter + 1
End Sub

Sub mnemonicScheme_Tic()
    With frmStart
        '������ ������
        .��1(0).Visible = Not (k1_isOpen)
        .��1(1).Visible = k1_isOpen

        .��2(0).Visible = Not (k2_isOpen)
        .��2(1).Visible = k2_isOpen
        .�����(0).Visible = k2_isOpen

        .��3(0).Visible = Not (k3_isOpen)
        .��3(1).Visible = k3_isOpen

        .��4(0).Visible = Not (k4_isOpen)
        .��4(1).Visible = k4_isOpen

        .��5(0).Visible = Not (k5_isOpen)
        .��5(1).Visible = k5_isOpen
        .������_����.Visible = k5_isOpen

        .��6(0).Visible = Not (k6_isOpen)
        .��6(1).Visible = k6_isOpen
 
        .��7(0).Visible = Not (k7_isOpen)
        .��7(1).Visible = k7_isOpen
        .�����(1).Visible = k7_isOpen

        ' TODO ���������� �� Visible
        If isClutchOn Then
            .�����.BackColor = &HFF&
        Else
            .�����.BackColor = &HC0C0C0
        End If

        '��������� ����������� ���������� ���
        'If gn������(33).Data = 1 Then
        'Else
        'End If

        '����� � ������ ���
        'If gn������(45).Data = 1 Then
        'Else
        'End If

        '����� � ��������������� ������
        'If isFireTech Then
        'Else
        'End If

        '��� � ������ ��� 10%
        'If (gn������(41).Data = 1) Then
        'Else
        'End If
        '��� � ������ ��� 20%
        'If (gn������(42).Data = 1) Then
        'Else
        'End If

        '��� � ��������������� ������ 10%
        'If (gn������(43).Data = 1) Then
        'Else
        'End If

        '��� � ��������������� ������ 20%
        'If (gn������(44).Data = 1) Then
        'Else
        'End If
    End With
End Sub
