Attribute VB_Name = "DIO"

Global Const DIO_NoError            = 0
Global Const DIO_DriverOpenError    = 1
Global Const DIO_DriverNoOpen       = 2
Global Const DIO_GetDriverVersionError = 3
Global Const DIO_InstallIrqError    = 4
Global Const DIO_ClearIntCountError = 5
Global Const DIO_GetIntCountError   = 6
Global Const DIO_ResetError         = 7   
Global Const DIO_RemoveIrqError     = 8

Global Const DIO_GetTotalBoardError = 9
Global Const DIO_CardNotFound       = 10
Global Const DIO_GetConfigError     = 11
Global Const DIO_ExceedBoardNumber  = 12
 

' The test functions
Declare Function DIO_ShortSub2 Lib "DIO.DLL" _
	(ByVal a As Integer, ByVal b As Integer) As Integer
Declare Function DIO_FloatSub2 Lib "DIO.DLL" _
	(ByVal a As Single, ByVal b As Single) As Single
Declare Function DIO_GetDllVersion Lib "DIO.DLL" () As Integer
Declare Function DIO_GetDriverVersion Lib "DIO.DLL" _
	(wDriverVersion As Integer) As Integer

' The Driver functions
Declare Function DIO_DriverInit Lib "DIO.DLL" _
	(wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()
Declare function DIO_GetConfigAddress Lib "DIO.DLL" _
	(wAddrBase As Integer, wCurrentBoard As Integer) As Integer
Declare function DIO_ActiveBoard Lib "DIO.DLL" _
	(ByVal wBoardNo As Integer ) As Integer

' The DIO functions
Declare Sub DIO_OutputByte Lib "DIO.DLL" _
	(ByVal address As Integer, ByVal dataout As Byte)
Declare Sub DIO_OutputWord Lib "DIO.DLL" _
	(ByVal address As Integer, ByVal dataout As Integer)
Declare Function DIO_InputByte Lib "DIO.DLL" _
	(ByVal address As Integer) As Integer
Declare Function DIO_InputWord Lib "DIO.DLL" _
	(ByVal address As Integer) As Integer

' The Interrupt functions
Declare Function DIO_IntInstall Lib "DIO.DLL" _
        (hEvent As Long) As Integer
Declare Function DIO_IntGetCount Lib "DIO.DLL" _
        (wIntCount As Long) As Integer
Declare Function DIO_IntRemove Lib "DIO.DLL" () As Integer

' DIO-48 Counter functions
Declare Sub DIO48_SetCounter Lib "DIO.DLL" _
        (ByVal wBase As Integer, ByVal wCounterNo As Integer, _
         ByVal bCounterMode As Integer, ByVal wCounterValue As Long )
Declare function  DIO48_ReadCounter Lib "DIO.DLL" _
        (ByVal wBase As Integer, ByVal wCounterNo As Integer, _
         ByVal bCounterMode As Integer) As Long
Declare Sub DIO48_SetCounterA Lib "DIO.DLL" _
        (ByVal wCounterNo As Integer, ByVal bCounterMode As Integer, _
         ByVal wCounterValue As Long)
Declare function  DIO48_ReadCounterA Lib "DIO.DLL" _
        (ByVal wCounterNo As Integer, ByVal bCounterMode As Integer ) _
         As Long

' DIO-64 Counter functions
Declare Sub DIO64_SetCounter Lib "DIO.DLL" _
        (ByVal wBase As Integer, ByVal wCounterNo As Integer, _
         ByVal bCounterMode As Integer, ByVal wCounterValue As Long)
Declare function  DIO64_ReadCounter Lib "DIO.DLL" _
        (ByVal wBase As Integer, ByVal wCounterNo As Integer, _
         ByVal bCounterMode As Integer) As Long
Declare Sub DIO64_SetCounterA Lib "DIO.DLL" _
        (ByVal wCounterNo As Integer, ByVal bCounterMode As Integer, _
         ByVal wCounterValue As Long)
Declare function  DIO64_ReadCounterA Lib "DIO.DLL" _
        (ByVal wCounterNo As Integer, ByVal bCounterMode As Integer) _
         As Long
