Attribute VB_Name = "DIO"
Option Explicit
Global Const DIO_NoError = 0
Global Const DIO_DriverOpenError = 1
Global Const DIO_DriverNoOpen = 2
Global Const DIO_GetDriverVersionError = 3
Global Const DIO_InstallIrqError = 4
Global Const DIO_ClearIntCountError = 5
Global Const DIO_GetIntCountError = 6
Global Const DIO_ResetError = 7
Global Const DIO_RemoveIrqError = 8

Global Const DIO_GetTotalBoardError = 9
Global Const DIO_CardNotFound = 10
Global Const DIO_GetConfigError = 11
Global Const DIO_ExceedBoardNumber = 12
 

' The Driver functions
Declare Function DIO_DriverInit Lib "DIO.DLL" _
        (wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()

' The DIO functions
Declare Sub DIO_OutputByte Lib "DIO.DLL" _
        (ByVal address As Integer, ByVal dataout As Byte)
Declare Function DIO_InputByte Lib "DIO.DLL" _
        (ByVal address As Integer) As Integer
