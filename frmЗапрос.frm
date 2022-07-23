VERSION 4.00
Begin VB.Form frmЗапрос 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Вопрос оператору"
   ClientHeight    =   2505
   ClientLeft      =   765
   ClientTop       =   1605
   ClientWidth     =   7500
   Height          =   2910
   Left            =   705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Top             =   1260
   Width           =   7620
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Нет"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   4215
      TabIndex        =   2
      Top             =   1065
      Width           =   3045
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Да"
      Default         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   270
      TabIndex        =   1
      Top             =   1095
      Width           =   3300
   End
   Begin VB.Label lblВопрос 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   750
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7065
   End
End
Attribute VB_Name = "frmЗапрос"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
'Если выбрана заправка во время наполнения аккумуляторов
If gbAkkum = True Then
  giStage2 = 7
  Car = 1
Else
  giStage2 = giStage2 + 1
End If

If gbOnlyAkk = True Then
  giStage = 2
  giStage2 = 8
End If

giTrigger = 1
gbFrmShow = False
frmЗапрос.Hide
End Sub


Private Sub Command2_Click()
If gbAkkum = False Then
  giStage2 = giStage2 + 1
Else
  frmStart.SSCmdStart.Enabled = True
End If

'Если пистолет не вставлен во время заправки только от аккумуляторов, то на ПредПуск
If gbOnlyAkk = True Then
  giStage = 1
  giStage1 = 1
 frmStart.SSCmdStart.Enabled = True
  'gbAkkum = True
End If

giTrigger = 0
gbFrmShow = False
frmЗапрос.Hide
End Sub


Private Sub Form_Activate()
  lblВопрос.Caption = gsMsg
End Sub

Private Sub Form_Load()
Left = 10
Top = 10
   ' Включение атрибута TopMost.
 SetWindowPos hwnd, conHwndTopmost, 10, 10, 520, 200, conSwpNoActivate Or conSwpShowWindow
End Sub
