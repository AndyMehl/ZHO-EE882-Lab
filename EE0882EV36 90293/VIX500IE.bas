Attribute VB_Name = "VIX500IE"
Option Explicit

Private Const DEGPERREV = 360

'Enumerated type for direction
Enum DirectionType

    dtClockwise = 0
    dtCounterClockwise = 1

End Enum

'Enumerated type for servo mode
Enum ModeType

    mtModeContinuous = 0
    mtModeAbsolute = 1
    mtModeIncremental = 2

End Enum

Private msngGearRatio As Single                 'Ratio of any gearbox attached to the motor
Private msngStepsPerRev As Single               'Steps per motor revolution
Private mstrControllerNum As String             'Current controller number
Private msngMotorVelocity As Single             'Velocity in Rev/Sec
Private msngMotorAcceleration As Single         'Acceleration in Rev/Sec/Sec
Private msngMotorDeceleration As Single         'Deceleration in Rev/Sec/Sec
Private mmtServoMode As ModeType                'Servo Mode
Private mdtDirection As DirectionType           'Servo direction
Private mblnGoodLink As Boolean                 'Boolean to indicate if communication is OK

Public Sub AbsoluteMoveTo(PositionInDegrees As Single)
'
'   PURPOSE: To make an absolute move to a specified location.
'            NOTE: This routine puts the controller in Absolute Position
'                  Mode temporarily.  This happens each time it is called.
'                  The routine is inefficient if multiple absolute movements
'                  will take place consecutively, and should be replaced by
'                  a subset of the commands that are used in this routine.
'
'  INPUT(S): PositionInDegrees = The absolute position to move to, in degrees
'
' OUTPUT(S): none

Dim lmtPreviousMode As ModeType

'Save the previous Servo Mode
lmtPreviousMode = GetServoMode

'Set the controller in Absolute Position mode
Call SetServoMode(mtModeAbsolute)
'Define the position to move to
Call DefineMovement(PositionInDegrees)
'Start the movement
Call StartMotor

'Return the motor to the previous mode if it wasn't Absolute Position Mode
If lmtPreviousMode <> mtModeAbsolute Then
    Call SetServoMode(lmtPreviousMode)
End If

End Sub

Public Sub DeEnergizeMotor()
'
'   PURPOSE: To de-energize the servo motor.
'
'  INPUT(S): None
'
' OUTPUT(S): None

'De-energize the motor
Call SendData("OFF")

End Sub

Public Sub DefineMovement(Degrees As Single)
'
'   PURPOSE: To define the distance to move.
'            NOTE:  If in Mode Absolute (MA), the distance will cause a move
'                   to an absolute position defined as "Degrees" degrees from zero.
'                   If in Mode Incremental (MI), the distance will cause a move
'                   to a distance from the current location.
'  INPUT(S): Degrees = The location/distance to move to/move.
'
' OUTPUT(S): none

Dim lsngDistance As Single
Dim llngDistance As Long

'Calculate the distance to define
lsngDistance = (Degrees / DEGPERREV) * msngStepsPerRev * msngGearRatio

'Convert to a long to drop the decimal
llngDistance = CLng(lsngDistance)

'Send the command to set the distance/location
Call SendData("D" & CStr(llngDistance))

End Sub

Public Sub DisableLimits()
'
'   PURPOSE: To disable the CW/CCW limits
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Disable the clockwise/counterclockwise limits
Call SendData("LIMITS(3,0,0)")

End Sub

Public Sub EnableLimits()
'
'   PURPOSE: To enable the CW/CCW limits
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Enable the clockwise/counterclockwise limits
Call SendData("LIMITS(0,0,0)")

End Sub

Public Sub EnergizeMotor()
'
'   PURPOSE: To energize the servo motor.
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Energize the motor
Call SendData("ON")

End Sub

Public Function GetDirection() As DirectionType
'
'   PURPOSE:   Returns the value of the modular Direction variable
'
'  INPUT(S):   None
'
' OUTPUT(S):   None

GetDirection = mdtDirection

End Function

Public Function GetGearRatio() As Single
'
'   PURPOSE:   Returns the value of the modular GearRatio variable
'
'  INPUT(S):   None
'
' OUTPUT(S):   None

GetGearRatio = msngGearRatio

End Function

Public Function GetLinkStatus() As Boolean
'
'   PURPOSE:   Returns the value of the modular boolean communication good variable.
'
'  INPUT(S):   None
'
' OUTPUT(S):   None

GetLinkStatus = mblnGoodLink

End Function

Public Function GetServoMode() As ModeType
'
'   PURPOSE:   Returns the value of the modular ServoMode variable
'
'  INPUT(S):   None
'
' OUTPUT(S):   None

GetServoMode = mmtServoMode

End Function

Public Function GetStepsPerRev() As Single
'
'   PURPOSE:   Returns the value of the modular StepsPerRev variable
'
'  INPUT(S):   None
'
' OUTPUT(S):   None

GetStepsPerRev = msngStepsPerRev

End Function

Public Sub InitializeCommunication(ComPortNum As Integer)
'
'   PURPOSE: To establish RS232 Communication to the VIX500IE Motor Controller
'
'  INPUT(S): ComPortNum = the COM port to communicate with the
'            VIX500IE Motor Controller
'
' OUTPUT(S):

Dim lstrReadData As String

On Error GoTo NotAbleToEstablishLink

'Make sure the port is closed
If frmVIX500IE.SerialPort.PortOpen = True Then frmVIX500IE.SerialPort.PortOpen = False

frmVIX500IE.SerialPort.CommPort = ComPortNum        'Set the serial port #
frmVIX500IE.SerialPort.Settings = "9600,N,8,1"     'Set baud rate, parity, etc.
frmVIX500IE.SerialPort.Handshaking = comNone        'Set flow control
frmVIX500IE.SerialPort.PortOpen = True              'Open COM Port

'Initialize to good RS232 Link;
'Sending the reset will reset it to false if it doesn't work
Call SetLinkStatus(True)

'Turn off echoing
Call SetEcho(False)

'Delay 100msec
Call frmVIX500IE.KillTime(100)

'Request a reset
Call SendDataGetResponse("Z", lstrReadData)

'Delay 500 milliseconds for response from reset command
Call frmVIX500IE.KillTime(500)

'Turn off echoing
Call SetEcho(False)

'Delay 100msec
Call frmVIX500IE.KillTime(100)

'If the controller echoes the command, the communication link is good
If GetLinkStatus Then
    'Incremental is default servo mode
    mmtServoMode = mtModeIncremental
    'Clockwise is default servo direction
    mdtDirection = dtClockwise
    '4000 is default steps per rev
    msngStepsPerRev = 4000
    '1 rev/sec is default velocity
    msngMotorVelocity = 1
    '100 rev/sec/sec is default acceleration
    msngMotorAcceleration = 10
    '100 rev/sec/sec is default deceleration
    msngMotorDeceleration = 10
    'The link is good now
    mblnGoodLink = True
End If

Exit Sub
'Error Trap
NotAbleToEstablishLink:
    mblnGoodLink = False
    
End Sub

Public Function ReadPosition(PositionInDegrees As Single) As Boolean
'
'   PURPOSE: To read the motor's current position in degrees.
'
'  INPUT(S): None
'
' OUTPUT(S): PositionInDegrees = Current Position of the motor in °.

Dim lstrReadData As String

On Error GoTo ReadError

ReadPosition = False

'If we are currently setup to communicate to all controllers, we can't get
'the position of any particular one
If mstrControllerNum <> "" Then

    'Send the command to read the position
    Call VIX500IE.SendDataGetResponse("R(PA)", lstrReadData)

    'Convert to a single and divide by StepsPerRev * GearRatio, then multiply by DEGPERREV
    PositionInDegrees = (CSng(lstrReadData) / (msngStepsPerRev * msngGearRatio)) * DEGPERREV

    'Function completed successfully
    ReadPosition = True

End If

Exit Function
ReadError:
    ReadPosition = False
End Function

Public Function ReadDriveFault() As Integer
'
'   PURPOSE: To read the motor's last reported DF.
'
'  INPUT(S): None
'
' OUTPUT(S): None

Dim lstrReadData As String
Dim lintBitNum As Integer

On Error GoTo CommError:

'Send the command to read the servo errors
Call SendDataGetResponse("R(DF)", lstrReadData)

If lstrReadData <> "" Then
    'Remove the prefix AND _'S
    lstrReadData = Mid(lstrReadData, 2, 4) & Mid(lstrReadData, 7, 4) & Mid(lstrReadData, 12, 4) & Mid(lstrReadData, 17, 4) & Mid(lstrReadData, 2, 4) & Mid(lstrReadData, 22, 4) & Mid(lstrReadData, 27, 4) & Mid(lstrReadData, 32, 4)

    If lstrReadData <> "" Then
        'Return bit of error if found, else zero
        For lintBitNum = 1 To 32
            If Mid(lstrReadData, lintBitNum, 1) = "1" Then
                ReadDriveFault = lintBitNum
                Exit Function
            End If
        Next lintBitNum
        ReadDriveFault = 0
    Else
        mblnGoodLink = False
        'Motor not responding
        ReadDriveFault = 99
    End If
Else
    mblnGoodLink = False
    'Motor not responding
    ReadDriveFault = 99
End If

Exit Function
CommError:
    mblnGoodLink = False
    'Motor not responding
    ReadDriveFault = 99

End Function

Public Sub RelativeMove(DistanceInDegrees As Single)
'
'   PURPOSE: To make a relative move for a specified distance.
'            NOTE: This routine puts the controller in Incremental Position
'                  Mode temporarily.  This happens each time it is called.
'                  The routine is inefficient if many incremental movements
'                  will take place consecutively, and should be replaced by
'                  a subset of the commands that are used in this routine.
'
'  INPUT(S): DistanceInDegrees = The distance to move, in degrees
'
' OUTPUT(S): none

Dim lmtPreviousMode As ModeType

'Save the previous Servo Mode
lmtPreviousMode = GetServoMode

'Set the controller in Incremental Position Mode
Call SetServoMode(mtModeIncremental)
'Define the distance to move
Call DefineMovement(DistanceInDegrees)
'Start the movement
Call StartMotor
'Return the motor to the previous mode if it wasn't Incremental Position Mode
If lmtPreviousMode <> mtModeIncremental Then
    Call SetServoMode(lmtPreviousMode)
End If

End Sub

Public Sub RestoreFactorySettings()
'
'   PURPOSE: To restore the factory PID settings on the motor controller.
'
'  INPUT(S): none
'
' OUTPUT(S): none

Dim lstrReadData As String

Call SendData("RFS")

End Sub

Public Sub SendData(WriteData As String)
'
'     PURPOSE:  To write data to the serial port.
'
'    INPUT(S):  writeData = string variable to be written to the serial
'               port.
'
'   OUTPUT(S):  None.
'

Dim lstrWriteData As String

On Error GoTo CommError

'Clear the input buffer
frmVIX500IE.SerialPort.InBufferCount = 0

'Add the controller number and a carriage return to the data
lstrWriteData = mstrControllerNum & WriteData & vbCr

'Write the data to the motor controller
frmVIX500IE.SerialPort.Output = lstrWriteData

'Loop until the message is sent out
Do Until frmVIX500IE.SerialPort.OutBufferCount = 0
    DoEvents
Loop

Exit Sub
CommError:
    mblnGoodLink = False

End Sub

Public Sub SendDataGetResponse(WriteData As String, ReadData As String)
'
'     PURPOSE:  To write data to the serial port and read back a response
'               from the motor controller.
'
'    INPUT(S):  WriteData = string variable to be written to the serial
'               port.
'
'   OUTPUT(S):  ReadData = string read back from serial port.
'
Dim lsngStartTimer As Single
Dim lsngEndTimer As Single
Dim lblnTimeOut As Boolean
Dim lblnStringFound As Boolean
Dim lstrReadData As String
Dim lstrWriteData As String

On Error GoTo CommError

'Initialize the variables
lblnTimeOut = False

'Clear the input buffer
frmVIX500IE.SerialPort.InBufferCount = 0

'Add the controller number and a carriage return to the data
lstrWriteData = mstrControllerNum & WriteData & vbCr

'Write the data to the motor controller
frmVIX500IE.SerialPort.Output = lstrWriteData

lsngStartTimer = Timer
Do
    lstrReadData = lstrReadData & frmVIX500IE.SerialPort.Input

    'If we the end of the read data is a carriage return line feed...
    If right(lstrReadData, 2) = vbCrLf Then
        'The communication is finished
        lblnStringFound = True
    End If

    '10-second timeout
    lsngEndTimer = Timer
    If (lsngEndTimer - lsngStartTimer > 10) Then lblnTimeOut = True
Loop Until lblnStringFound Or lblnTimeOut

If Not lblnTimeOut Then
    'Assign string after eliminating the Carriage Return-Line Feed and the * prefix
    ReadData = Mid(lstrReadData, 2, Len(lstrReadData) - 3)
Else
    mblnGoodLink = False
End If

Exit Sub
CommError:
    mblnGoodLink = False

End Sub

Public Sub SetControllerNumber(ControllerNum As Integer)
'
'   PURPOSE: To set controller number in a daisy chain of controllers.
'
'  INPUT(S): ControllerNum = The controller to set the module to communicate
'                            with.  NOTE: 0 = global, or all controllers
' OUTPUT(S): none

'Entering zero into the routine will result in setting the modular variable to an
'empty string, which will not add any prefix to commands, thus making them global.
If ControllerNum = 0 Then
    mstrControllerNum = "0"
Else
    mstrControllerNum = CStr(ControllerNum)
End If

End Sub

Public Sub SetDirection(Direction As DirectionType)
'
'   PURPOSE: To set the direction of the servo controller
'
'  INPUT(S): Direction = the direction to set the servo controller in.
'
' OUTPUT(S): none

Dim lstrDirection As String

'Set the modular variable
mdtDirection = Direction

If Direction = dtClockwise Then
    lstrDirection = "H+"
ElseIf Direction = dtCounterClockwise Then
    lstrDirection = "H-"
End If

'Send the command to set the direction
Call SendData(lstrDirection)

End Sub

Public Sub SetEcho(EchoOn As Boolean)
'
'   PURPOSE: To set whether or not the controller returns echoes.
'
'  INPUT(S): EchoOn = Whether or not the controller ret
'
' OUTPUT(S): none

Dim lstrWriteData As String

If EchoOn Then
    lstrWriteData = "W(EQ,0)"
Else
    lstrWriteData = "W(EQ,2)"
End If

'Send the command to set the direction
Call SendData(lstrWriteData)

End Sub

Public Function SetGearRatio(Ratio As Single) As Single
'
'   PURPOSE:   Sets the value of the modular GearRatio variable.  This
'              value represents the ratio of the number of motor revolutions
'              per revolution of the final drive arm of the system.  The value
'              is used in calculation of distances, velocities, and
'              accelerations.  This allows the function calls to request
'              these values in terms of final drive distance/speed/
'              acceleration, rather than motor distance/speed/acceleration.
'              Note that the gear ratio of a motor with no external gearing
'              is one (one motor revolution per drive arm revolution).
'
'  INPUT(S):   Ratio = System Gear Ratio
'
' OUTPUT(S):   None

msngGearRatio = Ratio

End Function

Public Sub SetLinkStatus(GoodOrBad As Boolean)
'
'   PURPOSE:   Sets the value of the modular boolean communication good variable.
'
'  INPUT(S):   GoodOrBad = The value to set the variable to.
'
' OUTPUT(S):   None

mblnGoodLink = GoodOrBad

End Sub

Public Sub SetPIDParameters(FeedForwardGain As Integer, ProportionalGain As Integer, IntegralGain As Integer, DerivativeGain As Integer, IntegralLimit As Integer, PositionError As Integer, PositionErrorWindow As Integer, InPositionTime As Integer, FilterTime As Integer)
'
'   PURPOSE: To set the parameters for the Motor Controller.
'
'  INPUT(S): FeedForwardGain          = FeedForward gain term of PID digital
'                                       filter loop. 0 to 1023.
'            ProportionalGain         = Proportional gain term of the PID
'                                       digital filter loop.  0 to 1023.
'            IntegralGain             = Integral gain term of the PID digital
'                                       filter loop.  0 to 1023.
'            DerivativeGain           = Derivative gain term of the PID
'                                       digital filter loop  0 to 1023.
'            IntegralLimit            = Clamps the level of influence the
'                                       integral gain term has on the PID
'                                       digital filter loop.  0 to 65535.
'            PositionError            = Maximum allowable position error.
'                                       +/- 2147483648.
'            PositionErrorWindow      = Used to configure In-Position window.
'                                       0 to 65535.
'            InPositionTime           = Used to configure In-Position window.
'                                       0 to 500ms.
'            FilterTime               = Filter Time constant of the PID
'                                       digital filter.  0 to 255.
' OUTPUT(S): none
If Not mblnGoodLink Then Exit Sub
Call SendData("W(GF," & CStr(FeedForwardGain) & ")")          'FeedForward Gain
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(GP," & CStr(ProportionalGain) & ")")         'Proportional Gain
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(GI," & CStr(IntegralGain) & ")")             'Integral Gain
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(GV," & CStr(DerivativeGain) & ")")           'Derivative Gain
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(IW," & CStr(IntegralLimit) & ")")            'Integral Limit
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(PE," & CStr(PositionError) & ")")            'Position Error
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(EW," & CStr(PositionErrorWindow) & ")")      'In-Position Error Window
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(IT," + CStr(InPositionTime) & ")")           'In-Position Time
Call frmVIX500IE.KillTime(50)
If Not mblnGoodLink Then Exit Sub
Call SendData("W(FT," + CStr(FilterTime) & ")")               'FilterTime
Call frmVIX500IE.KillTime(50)
End Sub

Public Sub SetServoMode(ByVal Mode As ModeType)
'
'   PURPOSE: To set the mode of the servo controller.
'
'  INPUT(S): Mode = the mode to set the servo controller to.
'
' OUTPUT(S): none

Dim lstrCommand As String

If mblnGoodLink Then
    'Set the modular variable
    mmtServoMode = Mode

    'Select the command
    Select Case Mode
        Case mtModeAbsolute
            lstrCommand = "MA"
        Case mtModeIncremental
            lstrCommand = "MI"
        Case mtModeContinuous
            lstrCommand = "MC"
    End Select

    If mblnGoodLink Then
        'Set the mode
        Call SendData(lstrCommand)
    End If
End If

End Sub

Public Sub SetStepsPerRev(StepsPerRev As Single)
'
'   PURPOSE:   Sets the value of the modular StepsPerRev variable
'
'  INPUT(S):   The value to set the variable to.
'
' OUTPUT(S):   None

msngStepsPerRev = StepsPerRev

Call SendData("W(EM," & CInt(msngStepsPerRev) & ")")

End Sub

Public Sub SetDeceleration(DecelInRevPerSecSquared As Single)
'
'   PURPOSE: To set the deceleration parameter for the VIX500IE Motor Controller.
'
'  INPUT(S): DecelInRevPerSecSquared    = Deceleration for the servo, in
'                                         Rev/sec^2.
' OUTPUT(S): none

'Set the modular variables
msngMotorDeceleration = DecelInRevPerSecSquared

'Send the deceleration
Call SendData("AD" & Format(msngMotorDeceleration * Abs(msngGearRatio), "#0.00"))

End Sub

Public Sub SetAcceleration(AccelInRevPerSecSquared As Single)
'
'   PURPOSE: To set the accleration parameter for the VIX500IE Motor Controller.
'
'  INPUT(S): AccelInRevPerSecSquared    = Acceleration for the servo, in
'                                         Rev/sec^2.
' OUTPUT(S): none

'Set the modular variables
msngMotorAcceleration = AccelInRevPerSecSquared

'Send the acceleration
Call SendData("AA" & Format(msngMotorAcceleration * Abs(msngGearRatio), "#0.00"))

End Sub

Public Sub SetVelocity(VelocityInRevPerSec As Single)
'
'   PURPOSE: To set the velocity parameter for the VIX500IE Motor Controller.
'
'  INPUT(S): VelocityInRevPerSec        = Maximum velocity for the servo, in
'                                         Rev/sec.
' OUTPUT(S): none

'Set the modular variables
msngMotorVelocity = VelocityInRevPerSec

'Send the velocity
Call SendData("V" & Format(msngMotorVelocity * Abs(msngGearRatio), "#0.00"))

End Sub

Public Sub StartMotor()
'
'   PURPOSE: To start the motor moving with the current commands/parameters.
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Start the motor
Call SendData("G")

End Sub

Public Sub StopMotor()
'
'   PURPOSE: To decelerate the motor to a stop using the current acceleration value.
'
'  INPUT(S): None
'
' OUTPUT(S): None

'Stop the motor
Call SendData("S")

End Sub

Public Sub ZeroPosition()
'
'   PURPOSE: To define the current motor position as zero.
'
'  INPUT(S): none
'
' OUTPUT(S): none

'Send the command to set the position to zero
Call SendData("W(PA,0)")

End Sub


