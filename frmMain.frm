VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "IPS Lag Time Controller"
   ClientHeight    =   4065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Troubleshooter"
      Height          =   5775
      Left            =   7320
      TabIndex        =   24
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdStartPos 
         Caption         =   "Feed to starting position"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdReadIO 
         Caption         =   "Read IO"
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtIOCmd 
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Text            =   "@01"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdReqIOCfg 
         Caption         =   "Request IO Config"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdReqStatus 
         Caption         =   "Request Printer Status"
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdInitPrn 
         Caption         =   "Initialize Printer"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdTestPrint 
         Caption         =   "Test Print"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCHex 
         Caption         =   "Convert"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtHex 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtDecimal 
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Hex"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Number"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Input/Output"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "TestPLC"
         Height          =   375
         Left            =   2640
         TabIndex        =   53
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCount 
         Height          =   375
         Left            =   6120
         TabIndex        =   51
         Top             =   960
         Width           =   375
      End
      Begin VB.Timer TimerStart 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   5400
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4680
         Top             =   960
      End
      Begin VB.CheckBox chkInvert 
         Caption         =   "Invert Input Value"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbPIOMsg 
         Caption         =   "Ready"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   6375
      End
      Begin VB.Label lbGateLTR 
         Caption         =   "Input"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5880
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "LTR Gate"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "P20"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label32 
         Caption         =   "P7"
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label31 
         Caption         =   "P4"
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "P1"
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "P33"
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "P32"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "P31"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "P30"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "NO3"
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spNO3 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "NO2"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spNO2 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3240
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "NO1"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spNO1 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2760
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "NO0"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spNO0 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2280
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "DIA3"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spDI3 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "DIA2"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spDI2 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1080
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "DIA1"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spDI1 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   600
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "DIA0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape spDI0 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lag Time Reader"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtCOMBAUD 
         Height          =   285
         Left            =   5280
         TabIndex        =   52
         Top             =   1080
         Width           =   615
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   6000
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.TextBox txtLastRead 
         Height          =   285
         Left            =   4680
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Timer TimerGL 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6120
         Top             =   720
      End
      Begin VB.TextBox txtMSComm1 
         Height          =   285
         Left            =   4680
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtExitTime 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtPrintTime 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtTicketRead 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "Last Read"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "COM Port"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Duration"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape spCondi 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   3480
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "OK"
         Height          =   255
         Index           =   24
         Left            =   3720
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape spCondi 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   3480
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "Expired"
         Height          =   255
         Index           =   26
         Left            =   3720
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Shape spCondi 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   4560
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "Anti-Passback"
         Height          =   255
         Index           =   27
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape spCondi 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   4560
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "Invalid Ticket"
         Height          =   255
         Index           =   29
         Left            =   4800
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Exit Time"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Print Time"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Ticket Read"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection
Dim db_name, db_server, db_port, db_user, db_pass, constr As String

Dim blSend As Boolean
Dim iSerial As Integer
Dim iLTRGrace As Integer
Dim blWait As Boolean
Dim blStarted As Boolean

Dim wBaseAddr         As Long
Dim wIrq              As Integer
Dim wSubVendor        As Integer
Dim wSubDevice        As Integer
Dim wSubAux           As Integer
Dim wSlotBus          As Integer
Dim wSlotDevice       As Integer
Dim wTotalBoards      As Integer
Dim wInitialCode      As Integer
Dim Ra                As Integer
Dim Rb                As Integer
Dim iRBits            As Integer

Private Sub cmdCHex_Click()
  If Trim(txtDecimal.Text) = "" Then Exit Sub
  
  txtHex.Text = GetHEX(CInt(Trim(txtDecimal.Text)))
  
End Sub

Private Sub GetSettings()
  Dim i As Integer
  Dim rtn As Long
  Dim wRetVal As Integer
  Dim sBoard As String
  
  'txtMSComm1.Text = Trim(GetIni("COMPORT", "LTR", "1"))
  'txtCOMBAUD.Text = Trim(GetIni("COMPORT", "LTRBAUD", "9600"))
  
  'On Error Resume Next
  
'
'  iLTRGrace = GetIni("RATE", "LTRGRACE", "15")
'  If Trim(GetIni("SYSTEM", "INVERTINPUT", "NO")) = "YES" Then
'    chkInvert.Value = 1
'  End If
'
'  lbGateLTR.Caption = Trim(GetIni("RELAY", "LTRGATE", ""))
'
'
'  iRBits = &H0
'  '********************************************************************
'  '* NOTICE: call PISODIO_DriverInit() to initialize the driver.        *
'  '* Initial the device driver, and return the board number in the PC *
'  '********************************************************************
'  wInitialCode = PISODIO_DriverInit()
'
'
'  If wInitialCode <> PISODIO_NoError Then
'      rtn = MsgBox("Driver initialize error!!!", , "PISODIO Card Error")
'
'      Exit Sub
'  End If
'
'  sBoard = Trim(GetIni("SYSTEM", "IOCARD", ""))
'
'  Select Case sBoard
'  Case "PISO_P8R8"
'    If PISODIO_SearchCard(wTotalBoards, PISO_P8R8) <> PISODIO_NoError Then
'        rtn = MsgBox("Search Card Error!!", , "PISODIO PISO_P8R8 Card Error")
'    End If
'  Case "PISO_730"
'    If PISODIO_SearchCard(wTotalBoards, PISO_730) <> PISODIO_NoError Then
'        rtn = MsgBox("Search Card Error!!", , "PISODIO PISO_730 Card Error")
'    End If
'  Case Else
'    MsgBox "I/O Card not specified!"
'
'  End Select
'
'  'Get board's Configuration Space
'  wRetVal = PISODIO_GetConfigAddressSpace(0, wBaseAddr, wIrq, _
'                    wSubVendor, wSubDevice, wSubAux, _
'                    wSlotBus, wSlotDevice)
'  If wRetVal <> PISODIO_NoError Then
'      lbPIOMsg.Caption = "Get Config-Address-Space Error !!"
'      Exit Sub
'  End If
'
'  PISODIO_OutputByte (wBaseAddr + &HC0), iRBits 'Reset Output
  
End Sub




Private Sub Command1_Click()

    Dim Instring As String
    ' Use COM2.
    MSComm1.CommPort = 1
    ' 9600 baud, odd parity, 8 data, and 1 stop bit.
    MSComm1.Settings = "9600,O,8,1"
    ' Tell the control to read entire buffer when Input
    ' is used.
    MSComm1.InputLen = 0
    ' Open the port.
    MSComm1.PortOpen = True
    ' Send the attention command to the modem.
    MSComm1.Output = "%01#WCSR00021**" + Chr$(13)
    ' Wait for data to come back to the serial port.
    Do
    DoEvents
    Loop Until MSComm1.InBufferCount >= 2
    ' Read the "OK" response data in the serial port.
    Instring = MSComm1.Input
    txtLastRead.Text = Instring
    ' Close the serial port.
    MSComm1.PortOpen = False
    
    End Sub

Private Sub Form_Load()
  Dim sACode As String
  Dim sSCode As String
  Dim sCCode As String
  Dim sMCode As String
  
  Dim iTotal As Long
  Dim iACode As Long
  Dim i As Integer
  
  blSend = False
  blWait = False
  blStarted = False
  
  
  'MsgBox "Activation key ok!", vbExclamation
  TimerStart.Enabled = True
  
'  db_name = "ips_pcp"
'  db_server = "localhost"
'  db_port = "3306"    'default port is 3306
'  db_user = "pcp_admin"
'  db_pass = "master2x"

End Sub

Private Sub Form_Unload(Cancel As Integer)
 
  On Error Resume Next
  
  'PISODIO_DriverClose
   
  Timer1.Enabled = False
  
  conn.Close
  Set conn = Nothing
End Sub

Private Sub OpenServer() 'Connect MySQL Server Without ODBC setup
  Dim lgCount As Long
  Dim sServerIP As String
  Dim sConstr As String
  Dim sDataPath As String
  
  On Error Resume Next
  
  sServerIP = Trim(GetIni("SYSTEM", "SERVERIP", "192.168.2.5"))
  sDataPath = GetIni("SYSTEM", "DBPATH", "E:\DB\IPSAUTOPAY.GDB")
  
  sConstr = "DRIVER=Firebird/InterBase(r) driver;DBNAME=" & sServerIP & ":" & sDataPath & ";UID=MMS;PWD=MMS"
    
  lgCount = 0
  'constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 5.1 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
  Set conn = New ADODB.Connection
AGAIN:
  conn.Open sConstr
  lgCount = lgCount + 1
  txtCount.Text = lgCount
  DoEvents
  If conn.State = 0 Then
    conn.Close
    If lgCount > 30000 Then
        MsgBox "Unable to connect to database.  Check database settings."
        Unload Me
    Else:
        GoTo AGAIN
    End If
  End If
End Sub


Private Sub Timer1_Timer()
'  Dim wRetVal As Integer
'  Dim i       As Integer
'  Dim sInput  As String
'  Dim sInput2 As String
'  Dim t       As Long
'  Dim blCarPresent As Boolean
'  Dim blPushButton As Boolean
'  Dim blTestButton As Boolean
'
'  'Get board's Configuration Space
'  wRetVal = PISODIO_GetConfigAddressSpace(0, wBaseAddr, wIrq, _
'                    wSubVendor, wSubDevice, wSubAux, _
'                    wSlotBus, wSlotDevice)
'  If wRetVal <> PISODIO_NoError Then
'      lbPIOMsg.Caption = "Get Config-Address-Space Error !!"
'      Exit Sub
'  End If
'  PISODIO_OutputByte wBaseAddr, 1  ' enable DI/DO
'  Ra = PISODIO_InputByte(wBaseAddr + &HC0)   '// Input channel   0- 7
''  Rb = PISODIO_InputByte(wBaseAddr + &HC4)   '// Isolated Input channel   8-15
'
''  'Test
''  sInput = "11111111"
''  sInput2 = "11111111"
'
'  If chkInvert.Value = 1 Then
'    sInput = GetInvBinary(Ra)
''    sInput2 = GetInvBinary(Rb)
'  Else:
'    sInput = GetBinary(Ra)
''    sInput2 = GetBinary(Rb)
'  End If
'
'
'  If sInput = "" Then
'    lbPIOMsg.Caption = "Board input return null"
'    Exit Sub
'  End If
'
'  lbPIOMsg.Caption = sInput
'
'  If Mid$(sInput, 5, 1) = "1" Then
'    spDI3.FillColor = vbGreen
'  Else:
'    spDI3.FillColor = vbBlack
'  End If
'  If Mid$(sInput, 6, 1) = "1" Then
'    spDI2.FillColor = vbGreen
'  Else:
'    spDI2.FillColor = vbBlack
'  End If
'  If Mid$(sInput, 7, 1) = "1" Then
'    spDI1.FillColor = vbGreen
'  Else:
'    spDI1.FillColor = vbBlack
'  End If
'  If Mid$(sInput, 8, 1) = "1" Then
'    spDI0.FillColor = vbGreen
'  Else:
'    spDI0.FillColor = vbBlack
'  End If
'
'
'  If Ra = 255 Then
'    Exit Sub
'  End If
'
'  Select Case lbCarPresent.Caption
'  Case "DI0"
'    If Mid$(sInput, 8, 1) = "1" Then
'      spCP.FillColor = vbGreen
'      blCarPresent = True
'    Else:
'      spCP.FillColor = vbBlack
'      blCarPresent = False
'    End If
'  Case "DI1"
'    If Mid$(sInput, 7, 1) = "1" Then
'      spCP.FillColor = vbGreen
'      blCarPresent = True
'    Else:
'      spCP.FillColor = vbBlack
'      blCarPresent = False
'    End If
'  Case "DI2"
'    If Mid$(sInput, 6, 1) = "1" Then
'      spCP.FillColor = vbGreen
'      blCarPresent = True
'    Else:
'      spCP.FillColor = vbBlack
'      blCarPresent = False
'    End If
'  Case "DI3"
'    If Mid$(sInput, 5, 1) = "1" Then
'      spCP.FillColor = vbGreen
'      blCarPresent = True
'    Else:
'      spCP.FillColor = vbBlack
'      blCarPresent = False
'    End If
'  End Select
'
'  Select Case lbPushButton.Caption
'  Case "DI0"
'    If Mid$(sInput, 8, 1) = "1" Then
'      spPB.FillColor = vbGreen
'      blPushButton = True
'    Else:
'      spPB.FillColor = vbBlack
'      blPushButton = False
'    End If
'  Case "DI1"
'    If Mid$(sInput, 7, 1) = "1" Then
'      spPB.FillColor = vbGreen
'      blPushButton = True
'    Else:
'      spPB.FillColor = vbBlack
'      blPushButton = False
'    End If
'  Case "DI2"
'    If Mid$(sInput, 6, 1) = "1" Then
'      spPB.FillColor = vbGreen
'      blPushButton = True
'    Else:
'      spPB.FillColor = vbBlack
'      blPushButton = False
'    End If
'  Case "DI3"
'    If Mid$(sInput, 5, 1) = "1" Then
'      spPB.FillColor = vbGreen
'      blPushButton = True
'    Else:
'      spPB.FillColor = vbBlack
'      blPushButton = False
'    End If
'  End Select
'
'  If blCarPresent And blPushButton Then
'    PrintCOM False
'    Timer1.Enabled = False
'  End If
'
'  blTestButton = False
'  Select Case lbTestButton.Caption
'  Case "DI0"
'    If Mid$(sInput, 8, 1) = "1" Then
'      blTestButton = True
'    End If
'  Case "DI1"
'    If Mid$(sInput, 7, 1) = "1" Then
'      blTestButton = True
'    End If
'  Case "DI2"
'    If Mid$(sInput, 6, 1) = "1" Then
'      blTestButton = True
'    End If
'  Case "DI3"
'    If Mid$(sInput, 5, 1) = "1" Then
'      blTestButton = True
'    End If
'  End Select
'
'  If blTestButton Then
'    PrintCOM True
'  End If
'
End Sub





Private Sub TimerGL_Timer()
  TimerGL.Enabled = False
  RelayOutput lbGateLTR.Caption, 0
End Sub

Private Sub TimerStart_Timer()
  Dim sCheck As String

  If blStarted = False Then
    GetSettings
    OpenServer
    blStarted = True
    TimerStart.Enabled = False
    txtTicketRead.SetFocus
  Else:
    sCheck = SQLSearch("ticketno", "exitjournal", "ticketno = '0'")
  End If
End Sub

Private Sub ProcessLTR(ByVal sCode As String)
  Dim sMachineNo As String
  Dim sSerialNo As String
  Dim i As Integer
  Dim sCheck As String
  Dim s As String
  Dim sDay As String
  Dim sMth As String
  Dim sYear As String
  Dim sHour As String
  Dim sMin As String
  Dim iDay As Integer
  Dim iMth As Integer
  Dim lgHour As Long
  Dim lgMin As Long
  Dim lgLag As Long
  
  Dim dtTicket As Date
  Dim dtCurrent As Date
  Dim lgIndex As Long
  
  Dim iMachineNo As Integer
  Dim iMnoChk As Integer
  
  Dim sMethod As String
  Dim sCutOff As String
  Dim dtCutOffYes As String
  Dim dtCutOff As Date
  Dim dtPaid As Date
  
  Dim rs As ADODB.Recordset
  
  Dim blError As Boolean
  
  'Machine No
  sMachineNo = Left$(sCode, 1)
  iMnoChk = CInt(Trim(GetIni("RATE", "ACCEPTMNO", "5")))
  
  'Day and Month
  sDay = Mid$(sCode, 6, 2)
  sMth = Mid$(sCode, 8, 1)
  iDay = CInt(sDay)
  iMth = CInt(sMth)
  
  'Hour and Minutes
  sHour = Mid$(sCode, 9, 2)
  sMin = Mid$(sCode, 11, 2)
  lgHour = CLng(sHour)
  lgMin = CLng(sMin)
  
  'Checking barcode validity
  blError = False
  iMachineNo = CInt(sMachineNo)
  If iMnoChk <> iMachineNo Then
    blError = True
  End If
  
  If iDay = 0 Then
    blError = True
  End If
  If iDay > 31 And iDay < 41 Then
    blError = True
  End If
  If iDay > 71 Then
    blError = True
  End If
  
  If lgHour > 23 Then
    blError = True
  End If
  If lgMin > 59 Then
    blError = True
  End If
  
  If blError Then
    For i = 0 To 3
      spCondi(i).FillColor = vbBlack
    Next
    spCondi(3).FillColor = vbGreen
    Exit Sub
  End If
   
  sCheck = SQLSearch("ticketno", "EXITJOURNAL", "ticketno = '" & sCode & "'")
  If sCheck <> "-1" Then
    For i = 0 To 3
      spCondi(i).FillColor = vbBlack
    Next
    spCondi(2).FillColor = vbGreen
    'lgIndex = GetId("cp_passback") + 1
    'conn.Execute "insert into cp_passback values(" & lgIndex & ",'" & Format(Now, "YYYY-MM-DD HH:MM") & "','" & sCode & "')"
    Exit Sub
  End If
  
  'Serial No
  sSerialNo = Mid$(sCode, 2, 4)
  
  'Day and Month
  If iDay > 40 Then
    iDay = iDay - 40
    iMth = iMth + 10
    sDay = CStr(iDay)
    sMth = CStr(iMth)
  End If
 
  dtTicket = Format(sDay & "-" & sMth & "-" & Year(Date) & " " & sHour & ":" & sMin, "DD-MM-YY HH:MM")
  
  dtCurrent = Now
  
  lgLag = DateDiff("n", dtTicket, dtCurrent)
  
  iLTRGrace = GetIni("RATE", "LTRGRACE", "15")
  If lgLag <= iLTRGrace Then
    For i = 0 To 3
      spCondi(i).FillColor = vbBlack
    Next
    spCondi(0).FillColor = vbGreen 'OK
    
    'lgIndex = GetId("cp_granted") + 1
    'conn.Execute "insert into cp_granted values(" & lgIndex & ",'" & Format(Now, "YYYY-MM-DD HH:MM") & "','" & sCode & "')"
    conn.Execute "insert into EXITJOURNAL values('" & sCode & "','" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "HH:MM") & "','OK')"
    
    'Gate Up
    'RelayOutput lbGateLTR.Caption, 1
    'TimerGL.Enabled = True
    'Change to cash drawer signal
    cashDrawer
  
  Else:
    sMethod = Trim(GetIni("RATE", "METHOD", "MINUTE"))
    If sMethod = "DAILY" Then
      sCutOff = Trim(GetIni("RATE", "CUTOFF", "02:00:00"))
      dtCutOff = Format(Date, "DD-MMM-YYYY " & sCutOff)
      If Now() >= dtCutOff Then
        dtCutOff = DateAdd("d", 1, dtCutOff)
      End If
      dtCutOffYes = DateAdd("d", -1, dtCutOff)
      
      Set rs = New ADODB.Recordset
      rs.CursorLocation = adUseClient
      rs.Open "select OUTDATE,OUTTIME from JOURNAL where TICKETNO = '" & sCode & "'", conn, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
        dtTicket = Format(MediumDate(rs!OUTDATE) & " " & Format(rs!OUTTIME, "HH:MM"), "DD-MMM-YYYY HH:MM")
        
        If dtTicket < dtCutOffYes Then
            For i = 0 To 3
                spCondi(i).FillColor = vbBlack
            Next
            spCondi(1).FillColor = vbGreen 'Expired
        Else
            If dtTicket < dtCutOff Then
                conn.Execute "insert into EXITJOURNAL values('" & sCode & "','" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "HH:MM") & "','OK')"
                cashDrawer
                For i = 0 To 3
                    spCondi(i).FillColor = vbBlack
                Next
                spCondi(0).FillColor = vbGreen
            Else:
              For i = 0 To 3
                spCondi(i).FillColor = vbBlack
              Next
              spCondi(1).FillColor = vbGreen 'Expired
            End If
        End If
      Else:
        For i = 0 To 3
          spCondi(i).FillColor = vbBlack
        Next
        spCondi(3).FillColor = vbGreen
      End If
      rs.Close
      Set rs = Nothing
      
    Else:
      Set rs = New ADODB.Recordset
      rs.CursorLocation = adUseClient
      rs.Open "select OUTDATE,OUTTIME from JOURNAL where TICKETNO = '" & sCode & "'", conn, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
        dtTicket = Format(rs!OUTTIME, "DD-MM-YY HH:MM")
        lgLag = DateDiff("n", dtTicket, dtCurrent)
        If lgLag <= iLTRGrace Then
          conn.Execute "insert into EXITJOURNAL values('" & sCode & "','" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "HH:MM") & "','OK')"
          cashDrawer
          For i = 0 To 3
            spCondi(i).FillColor = vbBlack
          Next
          spCondi(0).FillColor = vbGreen
        Else:
          For i = 0 To 3
            spCondi(i).FillColor = vbBlack
          Next
          spCondi(1).FillColor = vbGreen 'Expired
        End If
      Else:
        For i = 0 To 3
          spCondi(i).FillColor = vbBlack
        Next
        spCondi(3).FillColor = vbGreen 'Invalid Ticket
      End If
      rs.Close
      Set rs = Nothing
    End If
    
    
    'lgIndex = GetId("cp_expired") + 1
    
    
  End If
  
End Sub

Private Sub cashDrawer()
  Dim sGateUp As String
  
  'sGateUp = Trim(GetIni("COMPORT", "GATEUPCMD", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"))
  'MSComm1.CommPort = CInt(txtMSComm1.Text)
  'MSComm1.PortOpen = True ' Open the serial port
  'MSComm1.Settings = txtCOMBAUD.Text & ",n, 8,1" ' Define the communication parameters per printer specifications
  'MSComm1.Output = sGateUp
  
  'MSComm1.PortOpen = False
  
  
    Dim Instring As String
    ' Use COM2.
    MSComm1.CommPort = 1
    ' 9600 baud, odd parity, 8 data, and 1 stop bit.
    MSComm1.Settings = "9600,O,8,1"
    ' Tell the control to read entire buffer when Input
    ' is used.
    MSComm1.InputLen = 0
    ' Open the port.
    MSComm1.PortOpen = True
    ' Send the attention command to the modem.
    MSComm1.Output = "%01#WCSR00021**" + Chr$(13)
    ' Wait for data to come back to the serial port.
    Do
    DoEvents
    Loop Until MSComm1.InBufferCount >= 2
    ' Read the "OK" response data in the serial port.
    Instring = MSComm1.Input
    txtLastRead.Text = Instring
    ' Close the serial port.
    MSComm1.PortOpen = False
  
  
  
End Sub

Private Function GetId(ByVal sTable As String) As Long
  Dim rs As New ADODB.Recordset
  
  rs.CursorLocation = adUseClient
  rs.Open "select max(id) as maxid from " & sTable, conn, adOpenKeyset, adLockReadOnly, adCmdText
  If IsNull(rs!maxid) Then
    GetId = 0
  Else:
    GetId = rs!maxid
  End If
  rs.Close
  Set rs = Nothing
End Function

Private Function SQLSearch(ByVal sField As String, sTable As String, sCondition As String)
  
  Dim adoRS As New ADODB.Recordset
  Dim sQuery As String
  
  If sCondition = "0" Then
    sQuery = "SELECT " & sField & " FROM " & sTable & " ORDER  BY " & sField & " ASC"
  Else:
    sQuery = "SELECT " & sField & " FROM " & sTable & " WHERE " & sCondition & " ORDER  BY " & sField & " ASC"
  End If
    
  'Debug.Print sQuery
    
  adoRS.Open sQuery, conn, , , adCmdText
  If adoRS.EOF = False Then
    
    SQLSearch = Trim(adoRS(0).Value)
    
  Else:
    SQLSearch = "-1"
  End If
  Set adoRS = Nothing
  
End Function

Private Sub RelayOutput(ByVal sRelayNo As String, iBit As Integer)
  Dim iValue As Integer
    
  Select Case sRelayNo
  Case "NO0"
    If iBit = 1 Then
      spNO0.FillColor = vbGreen
    Else:
      spNO0.FillColor = vbBlack
    End If
  Case "NO1"
    If iBit = 1 Then
      spNO1.FillColor = vbGreen
    Else:
      spNO1.FillColor = vbBlack
    End If
  Case "NO2"
    If iBit = 1 Then
      spNO2.FillColor = vbGreen
    Else:
      spNO2.FillColor = vbBlack
    End If
  Case "NO3"
    If iBit = 1 Then
      spNO3.FillColor = vbGreen
    Else:
      spNO3.FillColor = vbBlack
    End If
  End Select
  
  DoEvents
    
  iValue = 0
  
  If spNO0.FillColor = vbGreen Then
    iValue = iValue + 1
  End If
  If spNO1.FillColor = vbGreen Then
    iValue = iValue + 2
  End If
  If spNO2.FillColor = vbGreen Then
    iValue = iValue + 4
  End If
  If spNO3.FillColor = vbGreen Then
    iValue = iValue + 8
  End If
  
'  If cmdBin(4).Caption = "1" Then
'    iValue = iValue + 16
'  End If
'  If cmdBin(5).Caption = "1" Then
'    iValue = iValue + 32
'  End If
'  If cmdBin(6).Caption = "1" Then
'    iValue = iValue + 64
'  End If
'  If cmdBin(7).Caption = "1" Then
'    iValue = iValue + 128
'  End If
  
  iRBits = iValue
  
  PISODIO_OutputByte (wBaseAddr + &HC0), iValue

End Sub





Private Sub txtTicketRead_KeyPress(KeyAscii As Integer)
  Dim s As String
    
  If KeyAscii = vbKeyReturn Then
      If Len(Trim(txtTicketRead.Text)) >= 13 Then
        s = Left(Trim(txtTicketRead.Text), 13)
        If s <> "" Then
          ProcessLTR s
          txtTicketRead.Text = ""
          txtLastRead.Text = s
        End If
      Else:
        txtTicketRead.Text = ""
        txtTicketRead.SetFocus
      End If
    
  End If
End Sub
