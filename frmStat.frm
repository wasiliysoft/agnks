VERSION 4.00
Begin VB.Form frmStat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Статистика"
   ClientHeight    =   1965
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3330
   Height          =   2370
   Left            =   1080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   3450
   Begin VB.TextBox txtStat 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   1725
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   495
      Width           =   1440
   End
   Begin VB.TextBox txtStat 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   0
      Left            =   1710
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   135
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   855
      TabIndex        =   2
      Top             =   1425
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Отпущено газа :"
      Height          =   270
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Пришло газа : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1260
   End
End
Attribute VB_Name = "frmStat"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Command1_Click()
  frmStat.Hide
End Sub


