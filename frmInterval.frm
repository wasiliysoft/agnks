VERSION 4.00
Begin VB.Form frmInterval 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Interval"
   ClientHeight    =   1590
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2595
   Height          =   1995
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   2715
   Begin VB.CommandButton Command1 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "500"
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmInterval"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   gnInterval = Val(Text1.Text)
   frmStart.Timer1.Interval = gnInterval
   frmStart.Text3(0).Text = Text1.Text
   frmInterval.Hide
End Sub


