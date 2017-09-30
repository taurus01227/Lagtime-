Attribute VB_Name = "PISODIO"

Global Const PISODIO_NoError = 0
Global Const PISODIO_DriverOpenError = 1
Global Const PISODIO_DriverNoOpen = 2
Global Const PISODIO_GetDriverVersionError = 3
Global Const PISODIO_InstallIrqError = 4
Global Const PISODIO_ClearIntCountError = 5
Global Const PISODIO_GetIntCountError = 6
Global Const PISODIO_RegisterApcError = 7
Global Const PISODIO_RemoveIrqError = 8
Global Const PISODIO_FindBoardError = 9
Global Const PISODIO_ExceedBoardNumber = 10
Global Const PISODIO_ResetError = 11

' to trigger a interrupt when high -> low
Global Const PISODIO_ActiveLow = 0
' to trigger a interrupt when low -> high
Global Const PISODIO_ActiveHigh = 1


' ID
Global Const PISO_P16R16U = &H18000000                       ' for PISO-P16R16U

Global Const PISO_C64 = &H800800                           ' for PISO-C64
Global Const PISO_P64 = &H800810                           ' for PISO-P64
Global Const PISO_A64 = &H800850                           ' for PISO-A64
Global Const PISO_P32C32 = &H800820                        ' for PISO-P32C32
Global Const PISO_P32A32 = &H800870                        ' for PISO-P32A32
Global Const PISO_P8R8 = &H800830                          ' for PISO-P8R8
Global Const PISO_P8SSR8AC = &H800830                      ' for PISO-P8SSR8AC
Global Const PISO_P8SSR8DC = &H800830                      ' for PISO-P8SSR8DC
Global Const PISO_730 = &H800840                           ' for PISO-730
Global Const PISO_730A = &H800880                          ' for PISO-730A



' The Test functions
Declare Function PISODIO_ShortSub Lib "PISODIO.dll" (ByVal a As Integer, ByVal b As Integer) As Integer
Declare Function PISODIO_FloatSub Lib "PISODIO.dll" (ByVal a As Single, ByVal b As Single) As Single
Declare Function PISODIO_GetDllVersion Lib "PISODIO.dll" () As Integer

' The Driver functions
Declare Function PISODIO_DriverInit Lib "PISODIO.dll" () As Integer
Declare Sub PISODIO_DriverClose Lib "PISODIO.dll" ()
Declare Function PISODIO_SearchCard Lib "PISODIO.dll" (wBoards As Integer, ByVal dwPIOPISOCardID As Long) As Integer
Declare Function PISODIO_GetDriverVersion Lib "PISODIO.dll" (wDriverVersion As Integer) As Integer
Declare Function PISODIO_GetConfigAddressSpace Lib "PISODIO.dll" ( _
                           ByVal wBoardNo As Integer, wAddrBase As Long, wIrqNo As Integer, _
                           wSubVendor As Integer, wSubDevice As Integer, wSubAux As Integer, _
                           wSlotBus As Integer, wSlotDevice As Integer) As Integer

Declare Function PISODIO_ActiveBoard Lib "PISODIO.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function PISODIO_WhichBoardActive Lib "PISODIO.dll" () As Integer


' DIO functions
Declare Sub PISODIO_OutputByte Lib "PISODIO.dll" (ByVal address As Long, ByVal dataout As Integer)
Declare Sub PISODIO_OutputWord Lib "PISODIO.dll" (ByVal address As Long, ByVal dataout As Long)
Declare Function PISODIO_InputByte Lib "PISODIO.dll" (ByVal address As Long) As Integer
Declare Function PISODIO_InputWord Lib "PISODIO.dll" (ByVal address As Long) As Long

' Interrupt functions
Declare Function PISODIO_IntInstall Lib "PISODIO.dll" (ByVal wBoard As Integer, hEvent As Long, _
        ByVal wInterruptSource As Integer, ByVal wActiveMode As Integer) As Integer
Declare Function PISODIO_IntRemove Lib "PISODIO.dll" () As Integer
Declare Function PISODIO_IntGetCount Lib "PISODIO.dll" (dwIntCount As Long) As Integer
Declare Function PISODIO_IntResetCount Lib "PISODIO.dll" () As Integer


