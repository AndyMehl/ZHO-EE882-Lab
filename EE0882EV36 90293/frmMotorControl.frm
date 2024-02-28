VERSION 5.00
Begin VB.Form frmMotorControl 
   Caption         =   "Motor Control"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLoc 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Location:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindFace 
      Caption         =   "Find Pedal Face"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblFaceLoc 
      Caption         =   "0.0"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblFace 
      Caption         =   "Face Position:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblPos 
      Caption         =   "0.0"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLoc 
      Caption         =   "Current Position:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmMotorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvntFirstPosition As Variant

Private Sub cmdFindFace_Click()

Dim lsngStart As Single
Dim lsngStop As Single

gintAnomaly = 0
lsngStart = gudtMachine.preScanStart
lsngStop = gudtMachine.preScanStop
gudtMachine.preScanStart = gudtMachine.loadLocation + 1
gudtMachine.preScanStop = gudtMachine.preScanStart + 20

If gudtMachine.preScanStop > 85 Then gudtMachine.preScanStop = 85

Call Sensotec.ActivateTare(1)

If FindPedalFace Then
    lblFaceLoc.Caption = Format(gudtReading(CHAN0).pedalFaceLoc, "###.00")
Else
    'Display the error message
    gudtReading(CHAN0).pedalFaceLoc = 0
    lblFaceLoc.Caption = Format(gudtReading(CHAN0).pedalFaceLoc, "###.00")
    Call MsgBox("Unable to Find Pedal Face." & vbCrLf & "Check Force Sensing Equipment.", vbOKOnly, "Error!")
End If
gudtMachine.preScanStart = lsngStart
gudtMachine.preScanStop = lsngStop

Call Sensotec.DeActivateTare(1)

Call Pedal.MoveToLoadLocation

mvntFirstPosition = Position
lblPos.Caption = Format(mvntFirstPosition, "###.00")
gintAnomaly = 0

End Sub

Private Sub cmdMove_Click()

'Set distance
If txtLoc.Text = "" Then txtLoc.Text = "0"
Call VIX500IE.DefineMovement(gudtReading(CHAN0).pedalFaceLoc + CSng(txtLoc.Text))

'Start the motor
Call VIX500IE.StartMotor

'Verify motor has stopped
Do
    mvntFirstPosition = Position
    Call frmDAQIO.KillTime(50)
Loop Until mvntFirstPosition = Position

lblPos.Caption = Format(mvntFirstPosition, "###.00")

End Sub

Private Sub Form_Load()

'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)

Call Pedal.MoveToLoadLocation

gudtReading(CHAN0).pedalFaceLoc = 0
mvntFirstPosition = Position
lblPos.Caption = Format(mvntFirstPosition, "###.00")
lblFaceLoc.Caption = Format(gudtReading(CHAN0).pedalFaceLoc, "###.00")

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call Pedal.MoveToLoadLocation

End Sub

Private Sub txtLoc_KeyPress(KeyAscii As Integer)
'Accept only numbers
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 43             '+
        'Accept the character
    Case 45, 46         '- and .
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 13
        Call cmdMove_Click
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select
End Sub
