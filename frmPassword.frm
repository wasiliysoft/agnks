VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4005
      TabIndex        =   2
      Top             =   585
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4005
      TabIndex        =   1
      Top             =   135
      Width           =   915
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   90
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1035
      Width           =   4830
   End
   Begin VB.Label lblDescription 
      Height          =   780
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   3660
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOk_Click()
    If Len (Trim(txtPassword))= 0 Then
      MsgBox "Пустой ввод",vbExclamation
    Else
      Hide
    End If
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

