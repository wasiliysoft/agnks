VERSION 5.00
Begin VB.Form frm������ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������ ���������"
   ClientHeight    =   2505
   ClientLeft      =   765
   ClientTop       =   1605
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2505
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "���"
      BeginProperty Font 
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
      Caption         =   "��"
      Default         =   -1  'True
      BeginProperty Font 
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
   Begin VB.Label lbl������ 
      Alignment       =   2  'Center
      Caption         =   "�������� �������� ?"
      BeginProperty Font 
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
Attribute VB_Name = "frm������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ������������ �� MSGBOX ������, ����� ��������� ��� �� �� ����� ����������� �������� �����

' �������� ��������
Private Sub Command1_Click()
    '���� ������� �������� �� ����� ���������� �������������
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
    frm������.Hide
End Sub

' �������� �� ��������
Private Sub Command2_Click()
    If gbAkkum = False Then
        giStage2 = giStage2 + 1
    Else
        frmStart.SSCmdStart.Enabled = True
    End If

    '���� �������� �� �������� �� ����� �������� ������ �� �������������, �� �� ��������
    If gbOnlyAkk = True Then
        giStage = 1
        giStage1 = 1
        frmStart.SSCmdStart.Enabled = True
        'gbAkkum = True
    End If

    giTrigger = 0
    gbFrmShow = False
    frm������.Hide
End Sub




Private Sub Form_Load()
    Left = 10
    Top = 10
    ' ��������� �������� TopMost.
    SetWindowPos hwnd, conHwndTopmost, 10, 10, 520, 200, conSwpNoActivate Or conSwpShowWindow
End Sub

