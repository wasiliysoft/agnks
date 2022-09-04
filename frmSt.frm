VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmSt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Статистика"
   ClientHeight    =   5205
   ClientLeft      =   315
   ClientTop       =   585
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   6720
   Begin VB.CommandButton cmdClose 
      Caption         =   "Закрыть"
      Height          =   600
      Left            =   3375
      TabIndex        =   4
      Top             =   4455
      Width           =   3210
   End
   Begin VB.CommandButton smdShow 
      Caption         =   "Показать"
      Height          =   600
      Left            =   4815
      TabIndex        =   3
      Top             =   2475
      Width           =   1725
   End
   Begin VB.CommandButton cmdSetCalendarNow 
      Caption         =   "Сегодня"
      Height          =   600
      Left            =   3330
      TabIndex        =   2
      Top             =   2475
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      _Version        =   524288
      _ExtentX        =   5953
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   10
      Day             =   1
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
   cmdSetCalendarNow_Click
End Sub



Private Sub cmdSetCalendarNow_Click()
    Calendar1.Value = Now
    smdShow_Click
End Sub


Private Sub cmdClose_Click()
   frmSt.Hide
End Sub


Private Sub smdShow_Click()
    Const separatorStr = "=============================="
    Dim s           As String
    Dim s1          As String
    Dim d           As Date
    Dim sum         As Double
    Dim v           As Double
    
    sum = 0    
    d = frmSt.Calendar1.Value

    frmSt.List1.Clear
    frmSt.List1.AddItem ("Журнал за ") & Format(d, "dd.mmmm.yyyy")
    frmSt.List1.AddItem separatorStr
    s = Format(d, "\#mm\/dd\/yyyy 00:00:00\#")
    s1 = Format(d, "\#mm\/dd\/yyyy 23:59:59\#")
    Set SelectRS = StatDB.OpenRecordset("select * from stat where DATA between " & s & " AND " & s1)
    If SelectRS.RecordCount >= 1 Then
        SelectRS.MoveLast
        SelectRS.MoveFirst
        Do While Not SelectRS.EOF
            v = SelectRS("GAZ_CAR")
            sum = sum + v
            frmSt.List1.AddItem Format(SelectRS("Data"), "  hh:mm:ss") + "                   " + Format(v, "###0.00")
            SelectRS.MoveNext
        Loop            
    End If
    frmSt.List1.AddItem separatorStr
    frmSt.List1.AddItem ("Всего:                        " & Format(sum, "###0.00"))
End Sub
