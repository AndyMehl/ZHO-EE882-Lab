VERSION 5.00
Begin VB.Form frmDDE 
   Caption         =   "DDE to RSLinx"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPLCDDEInput 
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   41
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtPLCDDEInput 
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   40
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtPLCDDEInput 
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   39
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   34
      Top             =   5040
      Width           =   4215
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   3735
      End
      Begin VB.CommandButton cmdWriteOne 
         Caption         =   "Write to:"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboWritePLC 
         Height          =   315
         ItemData        =   "frmDDE.frx":0000
         Left            =   2040
         List            =   "frmDDE.frx":0028
         TabIndex        =   36
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdReadAll 
         Caption         =   "Read All"
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame fraInputs 
      Caption         =   "DDE Inputs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   25
      Top             =   240
      Width           =   4215
      Begin VB.TextBox txtPLCDDEInput 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEInput 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEInput 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   27
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEInput 
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   26
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pallet Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   33
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PLC Home Complete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Stop:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   31
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Scan:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame fraOutputs 
      Caption         =   "DDE Outputs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   23
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   8
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   7
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   6
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   5
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   8
         Left            =   2160
         TabIndex        =   4
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   9
         Left            =   2160
         TabIndex        =   3
         Top             =   5880
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   10
         Left            =   2160
         TabIndex        =   2
         Top             =   6480
         Width           =   1935
      End
      Begin VB.TextBox txtPLCDDEOutput 
         Height          =   375
         Index           =   11
         Left            =   2160
         TabIndex        =   1
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "BOM Setup Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "WatchDog Disable:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scanner Init:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   21
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Results Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Calc Complete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Scan Ack:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Station Fault:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   17
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part ID 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   16
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part ID 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   15
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part ID 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   14
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part ID 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   13
         Top             =   6600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part ID 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   12
         Top             =   7200
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmDDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const RSLINXDDETOPIC = "RSLinx|PROGSCAN"    'The topic name setup in RSLinx

'*** PLC Memory Location Constants ***
Private Const STARTSCANADDR = "B19/64"              'Memory Location for Start Scan
Private Const STARTSCANACKADDR = "B19/65"           'Memory Location for Start Scan Acknowledge
Private Const CALCCOMPLETEADDR = "B19/66"           'Memory Location for Calc Complete
Private Const TOOLINGSAFEADDR = "B19/67"            'Memory Location for Tooling Safe State
Private Const SCANNERINITADDR = "B19/68"            'Memory Location for Scanner Initialized and Ready
Private Const WATCHDOGDISABLEADDR = "B19/69"        'Memory Location for Graphics Mode
Private Const ESTOPADDR = "B19/70"                  'Memory Location for EStop
Private Const STATIONFAULTADDR = "B19/71"           'Memory Location for Station Fault

Private Const RESULTSADDR = "B218:12"               'Memory Location for Results
Private Const PARTIDONEADDR = "B218:15"             'Memory Location for Part ID #1
Private Const PARTIDTWOADDR = "B218:16"             'Memory Location for Part ID #2
Private Const PARTIDTHREEADDR = "B218:17"           'Memory Location for Part ID #3
Private Const PARTIDFOURADDR = "B218:18"            'Memory Location for Part ID #4
Private Const PARTIDFIVEADDR = "B218:19"            'Memory Location for Part ID #5

Private Const PALLETNUMADDR = "B218:29"             'Memory Location for Pallet Number
Private Const BOMSETUPCODEADDR = "B102:3"           'Memory Location for BOM Setup Code

'Enumerated types to represent the item sent to the PLC
Enum DDEInput
    StartScan = 0
    EStop = 1
    ToolingSafeState = 2
    PalletNum = 3
    SerialNum = 4
    DateCode = 5
    PalletLoad = 6
End Enum

Enum DDEOutput
    StartScanAck = 0
    CalcComplete = 1
    ResultsCode = 2
    ScannerInit = 3
    WatchdogDisable = 4
    BOMSetupCode = 5
    StationFault = 6
    PartIDWord1 = 7
    PartIDWord2 = 8
    PartIDWord3 = 9
    PartIDWord4 = 10
    PartIDWord5 = 11
    ReClamp = 12          '3.1ANM
    FirstScanFail = 13    '3.1ANM
End Enum

Public Sub PLCDDESetup()
'
'   PURPOSE: To setup the DDE communication properties
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Setup LinkTopic and LinkItem for each DDE Input
txtPLCDDEInput(DDEInput.StartScan).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(DDEInput.StartScan).LinkItem = STARTSCANADDR
txtPLCDDEInput(DDEInput.StartScan).LinkMode = vbLinkManual
txtPLCDDEInput(DDEInput.EStop).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(DDEInput.EStop).LinkItem = ESTOPADDR
txtPLCDDEInput(DDEInput.EStop).LinkMode = vbLinkManual
txtPLCDDEInput(DDEInput.ToolingSafeState).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(DDEInput.ToolingSafeState).LinkItem = TOOLINGSAFEADDR
txtPLCDDEInput(DDEInput.ToolingSafeState).LinkMode = vbLinkManual
txtPLCDDEInput(DDEInput.PalletNum).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(DDEInput.PalletNum).LinkItem = PALLETNUMADDR
txtPLCDDEInput(DDEInput.PalletNum).LinkMode = vbLinkManual
txtPLCDDEInput(4).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(4).LinkItem = PALLETNUMADDR
txtPLCDDEInput(4).LinkMode = vbLinkManual
txtPLCDDEInput(5).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(5).LinkItem = PALLETNUMADDR
txtPLCDDEInput(5).LinkMode = vbLinkManual
txtPLCDDEInput(6).LinkTopic = RSLINXDDETOPIC
txtPLCDDEInput(6).LinkItem = PALLETNUMADDR
txtPLCDDEInput(6).LinkMode = vbLinkManual

'Setup LinkTopic and LinkItem for each DDE Output
txtPLCDDEOutput(DDEOutput.StartScanAck).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.StartScanAck).LinkItem = STARTSCANACKADDR
txtPLCDDEOutput(DDEOutput.StartScanAck).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.CalcComplete).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.CalcComplete).LinkItem = CALCCOMPLETEADDR
txtPLCDDEOutput(DDEOutput.CalcComplete).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.ResultsCode).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.ResultsCode).LinkItem = RESULTSADDR
txtPLCDDEOutput(DDEOutput.ResultsCode).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.ScannerInit).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.ScannerInit).LinkItem = SCANNERINITADDR
txtPLCDDEOutput(DDEOutput.ScannerInit).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.WatchdogDisable).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.WatchdogDisable).LinkItem = WATCHDOGDISABLEADDR
txtPLCDDEOutput(DDEOutput.WatchdogDisable).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.BOMSetupCode).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.BOMSetupCode).LinkItem = BOMSETUPCODEADDR
txtPLCDDEOutput(DDEOutput.BOMSetupCode).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.StationFault).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.StationFault).LinkItem = STATIONFAULTADDR
txtPLCDDEOutput(DDEOutput.StationFault).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.PartIDWord1).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.PartIDWord1).LinkItem = PARTIDONEADDR
txtPLCDDEOutput(DDEOutput.PartIDWord1).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.PartIDWord2).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.PartIDWord2).LinkItem = PARTIDTWOADDR
txtPLCDDEOutput(DDEOutput.PartIDWord2).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.PartIDWord3).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.PartIDWord3).LinkItem = PARTIDTHREEADDR
txtPLCDDEOutput(DDEOutput.PartIDWord3).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.PartIDWord4).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.PartIDWord4).LinkItem = PARTIDFOURADDR
txtPLCDDEOutput(DDEOutput.PartIDWord4).LinkMode = vbLinkManual
txtPLCDDEOutput(DDEOutput.PartIDWord5).LinkTopic = RSLINXDDETOPIC
txtPLCDDEOutput(DDEOutput.PartIDWord5).LinkItem = PARTIDFIVEADDR
txtPLCDDEOutput(DDEOutput.PartIDWord5).LinkMode = vbLinkManual

End Sub

Public Function ReadDDEInput(PLCcommand As DDEInput) As Integer
'
'   PURPOSE: Reads a DDE input from the PLC
'
'  INPUT(S): None
'
' OUTPUT(S): None

On Error GoTo PLCCommError

'Read from PLC
If gudtMachine.PLCCommType Then
    txtPLCDDEInput(PLCcommand).LinkRequest
Else
    txtPLCDDEInput(PLCcommand).Text = "0"
End If

'Return value read from textbox
ReadDDEInput = CInt(txtPLCDDEInput(PLCcommand).Text)

Exit Function
PLCCommError:

    gintAnomaly = 30            'Identify anomaly as PLC Comm Error
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in ReadDDEInput: " & Err.Description, True, False)
End Function

Public Function ReadDDEOutput(PLCcommand As DDEOutput) As Integer
'
'   PURPOSE: Reads a DDE output from the PLC
'
'  INPUT(S): None
'
' OUTPUT(S): None

On Error GoTo PLCCommError

'Read from PLC
txtPLCDDEOutput(PLCcommand).LinkRequest

'Return value read from textbox
ReadDDEOutput = CInt(txtPLCDDEOutput(PLCcommand).Text)

Exit Function
PLCCommError:

    gintAnomaly = 31            'Identify anomaly as PLC Comm Error
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in ReadDDEOutput: " & Err.Description, True, False)
End Function

Public Sub WriteDDEOutput(PLCcommand As DDEOutput, WriteData As Integer)
'
'   PURPOSE: Performs a DDE Write to the PLC
'
'  INPUT(S): None
'
' OUTPUT(S): None

On Error GoTo PLCCommError

'Place item to write in text box
txtPLCDDEOutput(PLCcommand).Text = CStr(WriteData)

'Write to PLC
txtPLCDDEOutput(PLCcommand).LinkPoke

Exit Sub
PLCCommError:

    gintAnomaly = 32            'Identify anomaly as PLC Comm Error
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in WriteDDEInput: " & Err.Description, True, False)
End Sub

Private Sub cmdExit_Click()
'
'   PURPOSE: Makes the form invisible
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Set the window to invisible
frmDDE.Visible = False

End Sub

Private Sub cmdReadAll_Click()
'
'   PURPOSE: Reads all the DDE Inputs & Outputs from the PLC
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Setup DDE Link Topics and Items
Call PLCDDESetup

If InStr(command$, "NOHARDWARE") = 0 Then         'Nohardware check
    'Read Input Values
    Call ReadDDEInput(StartScan)
    Call ReadDDEInput(EStop)
    Call ReadDDEInput(ToolingSafeState)
    Call ReadDDEInput(PalletNum)

    'Read Output Values
    Call ReadDDEOutput(StartScanAck)
    Call ReadDDEOutput(CalcComplete)
    Call ReadDDEOutput(ResultsCode)
    Call ReadDDEOutput(ScannerInit)
    Call ReadDDEOutput(WatchdogDisable)
    Call ReadDDEOutput(BOMSetupCode)
    Call ReadDDEOutput(StationFault)
    Call ReadDDEOutput(PartIDWord1)
    Call ReadDDEOutput(PartIDWord2)
    Call ReadDDEOutput(PartIDWord3)
    Call ReadDDEOutput(PartIDWord4)
    Call ReadDDEOutput(PartIDWord5)

End If

End Sub

Private Sub cmdWriteOne_Click()
'
'   PURPOSE: Writes to one of the DDE Output memory locations in the PLC
'
'  INPUT(S): None
'
' OUTPUT(S): None

If InStr(command$, "NOHARDWARE") = 0 Then         'Nohardware check
    'Exit sub if no location selected
    If cboWritePLC.Text = "" Then Exit Sub
    
    'Otherwise, write to PLC
    Call WriteDDEOutput(cboWritePLC.ListIndex, txtPLCDDEOutput(cboWritePLC.ListIndex).Text)
End If

End Sub

