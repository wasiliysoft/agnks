VERSION 5.00
Begin VB.Form frmDbDebug 
   Caption         =   "frmDbDebug"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "frmDbDebug"
   ScaleHeight     =   4440
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
   End
End
Attribute VB_Name = "frmDbDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long
    Set SelectRS = StatDB.OpenRecordset("select * from stat")
    If SelectRS.RecordCount >= 1 Then
        SelectRS.MoveLast
        SelectRS.MoveFirst
        For i = 0 To SelectRS.RecordCount - 1
            List1.AddItem Format(SelectRS("Data"), "dd.nn.yyyy hh:mm:ss") + "          " + Format(SelectRS("GAZ_CAR"), "###0.00") + "         " + Format(SelectRS("gaz_ir1"), "###0.00") + "         " + Format(SelectRS("moto"), "###0.00")
            SelectRS.MoveNext
        Next i
    End If
    SelectRS.Close
End Sub
