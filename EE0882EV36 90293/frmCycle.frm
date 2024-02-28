VERSION 5.00
Begin VB.Form frmCycle 
   Caption         =   "Part Cycle"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optFreq 
      Caption         =   "1 Hz"
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.9 Hz"
      Height          =   375
      Index           =   7
      Left            =   3480
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.8 Hz"
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.7 Hz"
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.6 Hz"
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.5 Hz"
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.4 Hz"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.3 Hz"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optFreq 
      Caption         =   "0.2 Hz"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtNumCycles 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "10"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Frequency:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCC 
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Current Cycle:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Cycles (2 - 1000):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.6ANM new form

Option Explicit
Private mblnStop As Boolean
Private msngVel As Single

Private Sub cmdStart_Click()

Dim lintNum, x As Integer

If gudtMachine.scanStart = 0 Then Call Pedal.FindStartScan

'Verify number of cycles
lintNum = CInt(txtNumCycles.Text)
If (lintNum < 2) Or (lintNum > 1000) Then
    MsgBox "Number of cycles must be 2 - 1000!", vbOKOnly, "Incorrect Number of Cycles"
    Exit Sub
End If

'Set scan arm Velocity
Call VIX500IE.SetVelocity(msngVel)
'Set scan arm Acceleration
Call VIX500IE.SetAcceleration(1)
'Set scan arm Deceleration
Call VIX500IE.SetDeceleration(1)

cmdStart.Enabled = False
cmdStop.Enabled = True

'Cycle thru number @ fequency
For x = 1 To lintNum
    If mblnStop = True Then Exit For
    Call Pedal.MoveToPosition(gudtMachine.scanEnd, 10)
    Call Pedal.MoveToPosition(gudtMachine.scanStart, 10)
    lblCC.Caption = CStr(x)
Next x

'Send back to load location
Call Pedal.MoveToLoadLocation
mblnStop = False
cmdStart.Enabled = True
cmdStop.Enabled = False

End Sub

Private Sub cmdStop_Click()
mblnStop = True
End Sub

Private Sub Form_Load()
cmdStop.Enabled = False
cmdStart.Enabled = True
mblnStop = False
lblCC.Caption = ""
msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.01
optFreq.Item(8).Value = True

End Sub

Private Sub optFreq_Click(Index As Integer)
Select Case Index
    Case 0    '0.2 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.00117
    Case 1    '0.3 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.0018
    Case 2    '0.4 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.00245
    Case 3    '0.5 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.00315
    Case 4    '0.6 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.004
    Case 5    '0.7 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.005
    Case 6    '0.8 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.006
    Case 7    '0.9 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.0075
    Case 8    '1.0 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.01
    Case Else 'Default to 1 Hz
        msngVel = (gudtMachine.scanEnd - gudtMachine.scanStart) * 0.01
End Select
    
    
End Sub
