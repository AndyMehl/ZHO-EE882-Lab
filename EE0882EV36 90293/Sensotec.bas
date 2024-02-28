Attribute VB_Name = "Sensotec"
'**********  Sensotec Serial Communication Module  **********
'
'   Scott R Calkins
'   CTS Automotive
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (219) 295-3575
'
'   The following program was written as a test panel / plug-in module
'   to allow serial communication with Sensotec SC devices.  Several
'   of the functions found here are not supported by the SC1000, such as
'   the limit functions and the peak/valley functions.
'   Although Sensotec SC devices can be bussed together on a single
'   serial communication line, this code was written assuming that would
'   not happen.  Future modification would be needed to support the
'   addressing of the different SC instruments on a common bus.  See the
'   Sensotec "SC Instrumentation Communications Guide" for details.
'   Currently, no SC instruments with more than one channel are in use
'   or planned to be used at CTS, but the functionality to control
'   different channels was included for robustness.
'   Note that the form can be added to a project and used to communicate
'   with a Sensotec SC instrument.  Add the form, verify the proper
'   COM port number, and use the functions.
'
'Ver     Date     By   Purpose of modification
'1.0  09/09/2002  SRC  First release of Sensotec software module.
'1.1  02/03/2003  SRC  Eliminated automatic error messages.
'1.2  03/12/2003  SRC  Updated per EE861 code review.
'1.3  04/11/2003  SRC  Disable killtime timer in Form_Unload
'1.4  06/27/2003  SRC  Modified ReadSerialNumber to work w/ new products.
'1.5  08/04/2003  SRC  Updated SetFreqResponse to work at real speed.
'1.6  09/29/2004  SRC  Updated Message Boxes.

Option Explicit

'Enumerated declaration for different energizing option of limits
Enum LimitEnergize
    
    leSignalLessThanSetPoint = 0
    leSignalGreaterThanSetPoint = 16
    leSignalInside = 32
    leSignalOutside = 48

End Enum

Private Const SENSOTECADDRESS = "00"

Private mblnGoodLink As Boolean         'Boolean to indicate if comm is OK

Public Sub InitializeCommunication(SerialPortNum As Integer)
'
'   PURPOSE: To establish RS232 Communication to the Sensotec SC Instrumentation
'
'  INPUT(S): SerialPortNum = the COM port to communicate with the SC Instrument
'
' OUTPUT(S):

Dim lstrRead As String

On Error GoTo NotAbleToEstablishLink

mblnGoodLink = False        'Initialize to bad RS232

'Make sure the port is closed
If frmSensotec.SerialPort.PortOpen = True Then frmSensotec.SerialPort.PortOpen = False

frmSensotec.SerialPort.CommPort = SerialPortNum        'Set the serial port #
frmSensotec.SerialPort.Settings = "38400,N,8,1"        'Set baud rate, parity, etc.
frmSensotec.SerialPort.PortOpen = True                 'Open COM Port

'Initiate communication with Sensotec
Call SendData("FICOMMUNICATING")        'Display a message to the front panel

Call ReadData(lstrRead)

If lstrRead = "OK" Then
    mblnGoodLink = True                 'Established RS232 communication
Else

'Error Trap
NotAbleToEstablishLink:
    
    mblnGoodLink = False
End If

End Sub

Public Sub Reset()
'
'     PURPOSE:  To reset the Sensotec Instrument.
'
'    INPUT(S):  None
'
'   OUTPUT(S):  None
'

'Send the Reset command
Call SendData("FR")

End Sub

Public Sub SendData(WriteData)
'
'     PURPOSE:  To write data to the serial port in a format understood by
'               the Sensotec.
'
'    INPUT(S):  writeData = string variable to be written to the serial
'               port.
'
'   OUTPUT(S):  None.
'
'*** Format the string to send to the Sensotec ***
'The format for writing to the Sensotec is:
'
'   # ad ch CO 00 n Cr
'               where:
'   #   is the "attention" character, ASCII code decimal 35
'   ad  is the two-character address of the SC instrument
'   ch  is the two-numeric ASCII-character channel number to which the
'         commmand will apply, 00 for system (global) commands
'   CO  is the two-character command to be executed
'   00  is the optional two-character parameter number needed by some
'         commands
'   n   is the optional argument to be written to the instrument
'   Cr  is "end" character, ASCII code deciaml 13

'  This module assumes that ch, CO, 00, and n will all be included in the
'  input argument write data where necessary.

'Clear the input buffer for the next response
frmSensotec.SerialPort.InBufferCount = 0

'Write the data to the motor controller
frmSensotec.SerialPort.Output = "#" & SENSOTECADDRESS & WriteData & vbCr

'Loop until the message is sent out
Do Until frmSensotec.SerialPort.OutBufferCount = 0
    DoEvents
Loop

End Sub

Public Sub ReadData(ByRef Data As String)
'
'     PURPOSE:  To wait to recieve a response from the Sensotec Instrument
'
'    INPUT(S):  ReadData = string variable to be written to the serial
'               port.
'
'   OUTPUT(S):  None.
'

Dim lblnStringFound As Boolean
Dim lintdatalength As Integer
Dim lsngStartTimer  As Single
Dim lsngEndTimer As Single
Dim lblnTimeOut As Boolean

'Initialize the data variable
Data = ""

lsngStartTimer = Timer
Do
    Data = Data + frmSensotec.SerialPort.Input
    If (right(Data, 2) = vbLf + vbCr) Then lblnStringFound = True
    '1-second timeout
    lsngEndTimer = Timer
    If (lsngEndTimer - lsngStartTimer > 1) Then lblnTimeOut = True
Loop Until lblnStringFound Or lblnTimeOut


'Take the length of the data
lintdatalength = Len(Data)

If lintdatalength > 2 Then
    'Cut off the carriage return and line feed
    Data = left(Data, lintdatalength - 2)
End If

End Sub

Public Sub ActivateTare(channelNum As Integer)
'
'     PURPOSE:  To turn the tare function on, or adding an offset
'               to declare the current measured value a zero point.
'
'    INPUT(S):  channelNum = the channel to tare
'
'   OUTPUT(S):  None.
'

Dim lstrRead As String

'Write the command to the Sensotec
Call SendData("0" + CStr(channelNum) + "F1")

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub DeActivateTare(channelNum As Integer)
'
'     PURPOSE:  To turn the tare function off, re-applying absolute
'               measure.
'
'    INPUT(S):  channelNum = the channel to un-tare
'
'   OUTPUT(S):  None.
'

Dim lstrRead As String

'Write the command to the Sensotec
Call SendData("0" + CStr(channelNum) + "F2")

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub ForceDACOutput(channelNum As Integer, Percentage As Single, Force As Boolean)
'
'   PURPOSE: To force the DAC output to a certain percentage of its full-scale.
'
'  INPUT(S): ChannelNum = the channel to force the output on
'            Percentage = the percentage of fullscale to force to
'            Force      = boolean representing whether or not to force
'
' OUTPUT(S): none


Dim lstrRead As String
Dim lsngPercentage As Single

If Force Then

    'Nominalize
    lsngPercentage = Percentage / 100
    
    'Send the command
    Call SendData("0" + CStr(channelNum) + "FH" + CStr(lsngPercentage))

Else

    Call SendData("0" + CStr(channelNum) + "FHAUTO")

End If

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub ClearPeakAndValley(channelNum As Integer)
'
'   PURPOSE: To clear the peak & valley readings
'
'  INPUT(S): ChannelNum = the channel to clear the peaks & valleys on
'
' OUTPUT(S): none

Dim lstrRead As String

'Send the command
Call SendData("0" + CStr(channelNum) + "FB")

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub SetFreqResponse(channelNum As Integer, Frequency As Integer)
'
'   PURPOSE: To set the frequency response
'
'  INPUT(S): ChannelNum = the channel to set the frequency response on
'            Frequency = the frequency response to set
'
' OUTPUT(S): none

'Note: 002, 008, 016, 032, 050, 100, 250, 500, & 800 are the only
'      available frequency responses.

Dim lstrRead As String

'Initialize the variable
lstrRead = ""

Select Case Frequency
    'Only try to send the data if it is a valid frequency response
    Case 2, 8, 16, 32, 50, 100, 250, 500, 800
        Call SendData("0" + CStr(channelNum) + "WU" + Format(Frequency, "000"))

        'Delay for the frequency response to be set
        Call frmSensotec.KillTime(2500)

        'Make sure the response was OK
        Call ReadData(lstrRead)
        If lstrRead <> "OK" Then
            mblnGoodLink = False
        End If

    Case Else
        MsgBox "Invalid Frequency Response Request", vbOKOnly, "Sensotec Error"

End Select

End Sub

Public Sub SetDACFullScale(channelNum As Integer, FullScale As Single)
'
'   PURPOSE: To set the DAC FullScale value
'
'  INPUT(S): ChannelNum = the channel to set the full-scale value on
'            FullScale = the full scale value to set
'
' OUTPUT(S): none


Dim lstrRead As String

Call SendData("0" + CStr(channelNum) + "WO" + CStr(FullScale))

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub SetLimitPoints(LimitNum As Integer, SetPoint As Single, ReturnPoint As Single)
'
'   PURPOSE: To set limit set and return points
'
'  INPUT(S): LimitNum = the channel to set limits on
'            SetPoint = the "set" point for the limit, or where it turns on.
'            ReturnPoint = the "return" point, or where the limit turns off.
'
' OUTPUT(S): none

Dim lstrRead As String

'Set the limit set point
Call SendData("WA0" + CStr(LimitNum) + CStr(SetPoint))

'Make sure the response was OK
Call ReadData(lstrRead)

'Only set the return point if the set point was set properly
If lstrRead = "OK" Then
    'Set the limit return point
    Call SendData("WB0" + CStr(LimitNum) + CStr(ReturnPoint))
    
    'Make sure the response was OK
    Call ReadData(lstrRead)
End If

If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub SetLimitOperation(LimitNum As Integer, channelNum As Integer, Enabled As Boolean, Mode As LimitEnergize)
'
'   PURPOSE: To set the operation of the limit outputs
'
'  INPUT(S): LimitNum   = Limit number to set up
'            ChannelNum = Channel number of the limit to set up
'            Enabled    = Whether or not the limit will be enabled
'            Mode       = Mode the limit will be set in
'
' OUTPUT(S): none

'Note: Each limit can be either enabled or disabled.
'      When enabled, the limit can be set to energize below or above
'      the set point, or inside or outside of the set point/ return
'      point window.  For readability, this was made an enumerated type.


Dim lintCommand As Integer
Dim lstrRead As String

'Initialize the command to the channel (defined by channelNum * 256)
lintCommand = 256 * channelNum
    
'Only define the limit if it is enabled (Adding 1 here enables it)
If Enabled Then lintCommand = lintCommand + 1 + Mode

'Send the command
Call SendData("WC0" + CStr(LimitNum) + CStr(lintCommand))

'Make sure the response was OK
Call ReadData(lstrRead)
If lstrRead <> "OK" Then
    mblnGoodLink = False
End If

End Sub

Public Sub ReadSerialNumber(channelNum As Integer, SerialNumber As String)
'
'   PURPOSE: To get the Serial Number of the force cell being used
'
'  INPUT(S): ChannelNum = the channel to read the
'
' OUTPUT(S): SerialNumber = the Serial Number of the attached force cell

Dim lstrRead As String

On Error GoTo InvalidResponse1

Call SendData("0" + CStr(channelNum) + "FE")

Call ReadData(lstrRead)

SerialNumber = lstrRead

Exit Sub
'Error Trap
InvalidResponse1:
    mblnGoodLink = False
End Sub

Public Sub ReadTracking(channelNum As Integer, TrackingValue As Single)
'
'   PURPOSE: To get the tracking reading
'
'  INPUT(S): ChannelNum = the channel to read the tracking on
'
' OUTPUT(S): TrackingValue = the peak value read back

Dim lstrRead As String

On Error GoTo InvalidResponse1

Call SendData("0" + CStr(channelNum) + "F0")

Call ReadData(lstrRead)

TrackingValue = CSng(lstrRead)

Exit Sub
'Error Trap
InvalidResponse1:
    mblnGoodLink = False
End Sub

Public Sub ReadPeak(channelNum As Integer, PeakValue As Single)
'
'   PURPOSE: To get the peak reading
'
'  INPUT(S): ChannelNum = the channel to read the peak on
'
' OUTPUT(S): PeakValue = the peak value read back

Dim lstrRead As String

On Error GoTo InvalidResponse2

Call SendData("0" + CStr(channelNum) + "F9")

Call ReadData(lstrRead)

PeakValue = CSng(lstrRead)

Exit Sub
'Error Trap
InvalidResponse2:
    mblnGoodLink = False
End Sub

Public Sub ReadValley(channelNum As Integer, ValleyValue As Single)
'
'   PURPOSE: To get the valley reading
'
'  INPUT(S): ChannelNum = the channel to read the valley on
'
' OUTPUT(S): ValleyValue = the valley value read back

Dim lstrRead As String

On Error GoTo InvalidResponse3

Call SendData("0" + CStr(channelNum) + "FA")

Call ReadData(lstrRead)

ValleyValue = CSng(lstrRead)

Exit Sub
'Error Trap
InvalidResponse3:
    mblnGoodLink = False
End Sub

Public Sub ReadLimitStatus(Limits() As Boolean)
'
'   PURPOSE: To read the status of the limit channels
'
'  INPUT(S): none
'
' OUTPUT(S): Limits() = An array of booleans representing the status
'                       of the limits
'
'Note: response from the SC Instrument is a binary representation
'      of the limits:
'       Limit1 = 0001
'       Limit2 = 0010
'       Limit3 = 0100
'       Limit4 = 1000
'      The limits are additive, so if Limit1 and Limit2 are both
'      on, the response will be 0011, or ASCII character "3"

Dim lstrRead As String
Dim lintLimits As Integer

On Error GoTo InvalidData

'Request the limit status
Call SendData("F6")

'Read back the reply
Call ReadData(lstrRead)

lintLimits = CInt(lstrRead)

'Make sure the code returned is valid
If lintLimits >= 0 And lintLimits < 16 Then
    
    'Decode the limits
    Limits(1) = CBool((lintLimits And &H1) = &H1)
    Limits(2) = CBool((lintLimits And &H2) = &H2)
    Limits(3) = CBool((lintLimits And &H4) = &H4)
    Limits(4) = CBool((lintLimits And &H8) = &H8)

Else

InvalidData:
    mblnGoodLink = False
End If

End Sub

Public Function GetLinkStatus() As Boolean
'
'   PURPOSE: To return the status of the moduler Good Link variable
'
'  INPUT(S): none.
'
' OUTPUT(S): GetLinkStatus

GetLinkStatus = mblnGoodLink

End Function

Public Sub SetLinkStatus(status As Boolean)
'
'   PURPOSE: To set the status of the moduler Good Link variable
'
'  INPUT(S): Status = Boolean to set the link status to.
'
' OUTPUT(S):

mblnGoodLink = status

End Sub

Private Sub ErrorMessage()
'
'   PURPOSE: To display the error message stating the communication is no
'            longer active
'
'  INPUT(S):
'
' OUTPUT(S): none

    MsgBox "There has been an error in the communication with the" _
       & vbCrLf & "Sensotec Instrumentation.  Please cycle power to the " _
       & vbCrLf & "instrument and reset the communication." _
       , vbOKOnly + vbCritical, "Sensotec Communication Error"

End Sub
