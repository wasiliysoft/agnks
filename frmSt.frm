VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
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
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin Threed.SSCommand ssClose 
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   4320
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Закрыть"
      BevelWidth      =   3
      Font3D          =   1
   End
   Begin Threed.SSCommand ssShow 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Показать"
      BevelWidth      =   3
      Font3D          =   1
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   3120
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

Private Sub ssClose_Click()
  frmSt.Hide
End Sub


Private Sub ssShow_Click()
Dim s As String
Dim s1 As String
Dim d As Date
Dim sum As Double

frmSt.List1.Clear
sum = 0
d = frmSt.Calendar1.Value
s = Format(d, "mm/dd/yyyy")
  s = Convert_Date(s)
  s1 = Format(d + 1, "mm/dd/yyyy")
  s1 = Convert_Date(s1)
  'frmShow.MousePointer = vbHourglass
Set SelectRS = StatDB.OpenRecordset("select * from stat where DATA between " & s & " AND " & s1)
If SelectRS.RecordCount >= 1 Then
    SelectRS.MoveLast
    SelectRS.MoveFirst
    
 For i = 0 To SelectRS.RecordCount - 1
   s = ""
   s = Format(CStr(SelectRS("Data")), "hh:mm:ss") + "        " + Format(CStr(SelectRS("GAZ_CAR")), "###0.00")
   sum = sum + SelectRS("GAZ_CAR")
   frmSt.List1.AddItem (s)
   SelectRS.MoveNext
 Next i
 frmSt.List1.AddItem ("====================")
 frmSt.List1.AddItem ("Всего:    " & Format(CStr(sum), "###0.00"))
End If

End Sub


