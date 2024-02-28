Attribute VB_Name = "MLX90293"
'************  MLX90293 Communication Module  ************
'
'   Andrew N Mehl
'   CTS Automotive
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (574) 295-3575
'
'   The following program was written as a test panel / plug-in module
'   to allow serial communication with the MLX90293 chip.
'
'Ver    Date       By   Purpose of modification
'1.0.0 01/14/2019  ANM  First release of MLX90293 software module.
'

Option Explicit
Public glngMLXID1 As Long
Public glngMLXID2 As Long
Public glngMLXID3 As Long
Public glngCUSTID1 As Long
Public glngCUSTID2 As Long
Public glngMLXLock As Long
Public glngSN As Double
Public glngLot As Double
Public gblnSolverFail As Boolean

Type PTC04A
    CommPortNum As Integer                          'Comm Port Number for the current programmer
    'Status Variables
    VendorID As Integer                             'PTC-04 Information
    ProductID As Integer                            'PTC-04 Information
    HardwareRevision As String                      'PTC-04 Information
    SerialCode As String                            'PTC-04 Information
    'Timing Variables
    Tpuls As Integer                                'Time Pulse time setting in uSec
    Tpor As Integer                                 'Power-up time setting in uSec
    Tprog As Integer                                'EEPROM charge time setting in uSec
    Thold As Integer                                'Hold time for return to regular mode before IC power-down in uSec
    Tsynchro As Integer
    TpulsMin As Integer
    TpulsMax As Integer
    SynchroDelay As Integer
    ByteData As Integer
    TSentTick As Integer
    Baudrate As Integer
    BaudrateSyncID As Integer
    VDDLow As Single
    VDDNorm As Single
    VDDComm As Single
    VBatLow As Single
    VBatNorm As Single
End Type

Public gudtPTC04(1 To 2) As PTC04A               'Instantiate the PTC-04 variables   (two - for two programmers)

Private mvntFirstPosition As Variant
Private Const MELEXISREADTIMEOUT = 10           'Melexis read timeout constant

Public MyDev(2) As PSF090293AAMLXDevice
Public lintDev1 As Integer
Public lintDev2 As Integer

Public MyDev1 As PSF090293AAMLXDevice
Public MyAdv1 As PSF090293AAMLXAdvanced
Public MySol1 As PSF090293AAMLXSolver

Public MyDev2 As PSF090293AAMLXDevice
Public MyAdv2 As PSF090293AAMLXAdvanced
Public MySol2 As PSF090293AAMLXSolver

Public MyDev3 As PSF090293AAMLXDevice
Public MyAdv3 As PSF090293AAMLXAdvanced
Public MySol3 As PSF090293AAMLXSolver

Public Const StrNotConnected As String = "Software is not connected. Press Connect button to try to connect"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function EncodePartID() As String
'
'   PURPOSE: To build the Part ID number based on the MLX ID information
'            in the IC associated with output #1.
'
'  INPUT(S): None
' OUTPUT(S): returns the PartID (String)

Dim ldblPartID As Double

'Y-Location on Wafer (Bit0 - Bit7)
ldblPartID = CDbl(gudtMLX90277(1).Read.X)
'X-Location on Wafer (Bit8 - Bit15)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.Y) * BIT8)
'Wafer Number(Bit16 - Bit20)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.Wafer) * BIT16)
'Lot Number (Bit21 - Bit38)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.Lot) * BIT21)

'Format as an eleven digit number and return
EncodePartID = Format(ldblPartID, "000000000000")

End Function

Public Sub EstablishCommunication()
'
'   PURPOSE: To establish Serial communications between the PC and
'            the Micronas programmer.
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim ComPortNum As Integer

On Error GoTo NotAbleToEstablishLink

Set MyDev(0) = CreateObject("MPT.PSF090293AAMLXDevice")
Call MyDev(0).ConnectChannel(CVar(CLng(PTC04PORT1)), dtSerial)
Call MyDev(0).CheckSetup(False)
Call MyDev(0).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
MyDev(0).Advanced.ChipVersion = 2

Set MyDev(1) = CreateObject("MPT.PSF090293AAMLXDevice")
Call MyDev(1).ConnectChannel(CVar(CLng(PTC04PORT2)), dtSerial)
Call MyDev(1).CheckSetup(False)
Call MyDev(1).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
MyDev(1).Advanced.ChipVersion = 2
        
gblnGoodPTC04Link = True

'gblnGoodPTC04Link = False
'
'ComPortNum = gudtPTC04(1).CommPortNum
''Make sure the port is closed
'If frmMLX90293.SerialPort(1).PortOpen = True Then frmMLX90293.SerialPort(1).PortOpen = False
'frmMLX90293.SerialPort(1).CommPort = ComPortNum     'Set the serial port #
'frmMLX90293.SerialPort(1).Settings = "115200,N,8,1" 'Set baud rate, parity, etc.
'frmMLX90293.SerialPort(1).Handshaking = comNone     'Set flow control
'frmMLX90293.SerialPort(1).InputLen = 0              'Set to read all characters
'frmMLX90293.SerialPort(1).PortOpen = True           'Open COM Port
'
'ComPortNum = gudtPTC04(2).CommPortNum
''Make sure the port is closed
'If frmMLX90293.SerialPort(2).PortOpen = True Then frmMLX90293.SerialPort(2).PortOpen = False
'frmMLX90293.SerialPort(2).CommPort = ComPortNum     'Set the serial port #
'frmMLX90293.SerialPort(2).Settings = "115200,N,8,1" 'Set baud rate, parity, etc.
'frmMLX90293.SerialPort(2).Handshaking = comNone     'Set flow control
'frmMLX90293.SerialPort(2).InputLen = 0              'Set to read all characters
'frmMLX90293.SerialPort(2).PortOpen = True           'Open COM Port
'
''Set the link variable to ok if we are able to Initialize the Programmer
'gblnGoodPTC04Link = MLX90293.InitializeProgrammer
'
'If gblnGoodPTC04Link Then Exit Sub

Exit Sub

NotAbleToEstablishLink:

MsgBox "There was an error trying to establish communications" _
        & vbCrLf & "with the 90293 programmer.", _
        vbOKOnly + vbCritical, "90293 Communication Error"

End Sub

Public Function InitializeProgrammer() As Boolean
'
'   PURPOSE: To initialize the instruction set, voltage levels, and
'            setup times for Melexis programmer.
'
'  INPUT(S): None.
'
' OUTPUT(S): Returns whether or not the programmer was initialized successfully

Dim lintProgrammerNum As Integer
Dim lsngVoltageLevel(1 To 2) As Single
Dim lintTimeInMicroSeconds(1 To 2) As Integer
Dim lsngCurrentLimit(1 To 2) As Single
Dim lintNumber(1 To 2) As Integer

'Initialize the value of the function to False
InitializeProgrammer = False

'If the Programmer is reset properly, assume communication is established
gblnGoodPTC04Link = ResetProgrammer

If gblnGoodPTC04Link Then

    'If any of the following routines does not complete successfully,
    'the code will exit the sub returning false as initialized above

'    '*** Setup voltage levels ***
'    'Set Channel 0 (Vdd) Ground Voltage to 0V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 0
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddGnd, lsngVoltageLevel()) Then Exit Function
'    'Set Channel 0 (Vdd) Nominal Voltage to 5V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 5
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddNom, lsngVoltageLevel()) Then Exit Function
'    'Set Channel 0 (Vdd) Programming Voltage to 9V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 9
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddProg, lsngVoltageLevel()) Then Exit Function
'    'Set Channel 1 (Vout) Low Level Voltage to 0V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 0
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutLow, lsngVoltageLevel()) Then Exit Function
'    'Set Channel 1 (Vout) Mid Level Voltage to 2.5V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 2.5
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutMid, lsngVoltageLevel()) Then Exit Function
'    'Set Channel 1 (Vout) High Level Voltage to 5V
'    For lintProgrammerNum = 1 To 2
'        lsngVoltageLevel(lintProgrammerNum) = 5
'    Next lintProgrammerNum
'    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutHigh, lsngVoltageLevel()) Then Exit Function
'
'    '*** Setup Timings ***
'    'Set Tpor
'    For lintProgrammerNum = 1 To 2
'        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tpor
'    Next lintProgrammerNum
'    If Not SetupTiming(ptcTpor, lintTimeInMicroSeconds()) Then Exit Function
'    'Set Thold
'    For lintProgrammerNum = 1 To 2
'        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Thold
'    Next lintProgrammerNum
'    If Not SetupTiming(ptcThold, lintTimeInMicroSeconds()) Then Exit Function
'    'Set Tprog
'    For lintProgrammerNum = 1 To 2
'        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tprog
'    Next lintProgrammerNum
'    If Not SetupTiming(ptcTprog, lintTimeInMicroSeconds()) Then Exit Function
'    'Set Tpuls
'    For lintProgrammerNum = 1 To 2
'        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tpuls
'    Next lintProgrammerNum
'    If Not SetupTiming(ptcTpuls, lintTimeInMicroSeconds()) Then Exit Function
'
'    '*** Set Measurement Delays & Filtering ***
'    'Set Measurement Delay to 5000 uS
'    For lintProgrammerNum = 1 To 2
'        lintNumber(lintProgrammerNum) = 5000
'    Next lintProgrammerNum
'    If Not SetupDelayOrFilter(ptcSetMeasureDelay, lintNumber()) Then Exit Function
'    'Set Sample Delay to 1 uS
'    For lintProgrammerNum = 1 To 2
'        lintNumber(lintProgrammerNum) = 1
'    Next lintProgrammerNum
'    If Not SetupDelayOrFilter(ptcSetSampleDelay, lintNumber()) Then Exit Function
'    'Set Measurement Filter to 1 measurement
'    For lintProgrammerNum = 1 To 2
'        lintNumber(lintProgrammerNum) = 1
'    Next lintProgrammerNum
'    If Not SetupDelayOrFilter(ptcSetMeasureFilter, lintNumber()) Then Exit Function
'
'    '*** Set the Current Limit ***
'    'Set Current Limit for Channel 4 (Vdd) to 200 mA
'    For lintProgrammerNum = 1 To 2
'        lsngCurrentLimit(lintProgrammerNum) = 200
'    Next lintProgrammerNum
'    If Not SetupCurrentLimit(ptcPPS_Vdd_I_Limit, lsngCurrentLimit()) Then Exit Function
'    'Set Current Limit for Channel 5 (Output) to 200 mA
'    For lintProgrammerNum = 1 To 2
'        lsngCurrentLimit(lintProgrammerNum) = 200
'    Next lintProgrammerNum
'    If Not SetupCurrentLimit(ptcPPS_Out_I_Limit, lsngCurrentLimit()) Then Exit Function
End If

'If everything went ok and we haven't left the Function yet, Initialization was successful
InitializeProgrammer = gblnGoodPTC04Link

End Function

Private Function ResetProgrammer() As Boolean
'
'   PURPOSE: To request a software reset from the programmers, then request
'            the programmers exit the bootloader portion of the firmware.
'            Finally, request hardware identification information from the
'            programmers.
'
'  INPUT(S): None
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

On Error GoTo ResetError

'Initialize the routine to return false
ResetProgrammer = False

'Build the command to reset the programmers
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcResetHardware)
Next lintProgrammerNum
'If the Reset command does not work, then exit the function
If Not SendCommandGetResponse(lstrWrite(), lstrResponse()) Then Exit Function

'Build the command to exit the bootloader program for both programmers
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcExit_BootLoader)
Next lintProgrammerNum
'Exit the bootloader firmware program
If Not SendCommandGetResponse(lstrWrite(), lstrResponse()) Then Exit Function

'Build the command to request hardware information from both programmers
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcGetHardwareID_Main)
Next lintProgrammerNum
'Request the hardware information from the Programmers
If Not SendCommandGetResponse(lstrWrite(), lstrResponse()) Then Exit Function

'The response from each programmer should be 35 characters long if it is correct
If Len(lstrResponse(1)) <> 35 Or Len(lstrResponse(2)) <> 35 Then
    'The GetHardwareID_Main command did not return the proper response
    Exit Function
End If

'The response format from the GetHardwareID command is as follows:

' B | BB | BB | B | BBBBBBB | B | B | BBBB | B | BBB | BBBBBBBBB | B
' 0 | 00 | 00 | 0 | 0000111 | 1 | 1 | 1111 | 1 | 222 | 222222233 | 3
' 0 | 12 | 34 | 5 | 6789012 | 3 | 4 | 5678 | 9 | 012 | 345678901 | 2

'Where:
'B00 = echo of command (GetHardwareID_Main = 01)
'B01 = Lower byte of Vendor ID
'B02 = Upper byte of Vendor ID
'B03 = Lower byte of Product ID
'B04 = Upper byte of Product ID
'B05 = Seperator byte: Hexadecimal 01
'B06 = M
'B07 = e
'B08 = l
'B09 = e
'B10 = x
'B11 = i
'B12 = s
'B13 = Seperator byte: Hexadecimal 01
'B14 = P
'B15 = T
'B16 = C
'B17 = 0
'B18 = 4
'B19 = Seperator byte: Hexadecimal 01
'B20 = Version #
'B21 = Version #
'B22 = Version #
'B23 = Version #
'B24 = Version #
'B25 = Seperator byte: Hexadecimal 01
'B26 = Serial Code Character #1
'B27 = Serial Code Character #2
'B28 = Serial Code Character #3
'B29 = Serial Code Character #4
'B30 = Serial Code Character #5
'B31 = Serial Code Character #6
'B32 = Serial Code Character #7
'B33 = Serial Code Character #8
'B34 = Seperator byte: Hexadecimal 01

For lintProgrammerNum = 1 To 2
    gudtPTC04(lintProgrammerNum).VendorID = Asc(Mid(lstrResponse(lintProgrammerNum), 2, 1)) + Asc(Mid(lstrResponse(lintProgrammerNum), 3, 1)) * 8
    gudtPTC04(lintProgrammerNum).ProductID = Asc(Mid(lstrResponse(lintProgrammerNum), 4, 1)) + Asc(Mid(lstrResponse(lintProgrammerNum), 5, 1)) * 8
    gudtPTC04(lintProgrammerNum).HardwareRevision = Mid(lstrResponse(lintProgrammerNum), 21, 5)
    gudtPTC04(lintProgrammerNum).SerialCode = Mid(lstrResponse(lintProgrammerNum), 27, 8)
Next lintProgrammerNum

ResetProgrammer = True      'Routine completed successfully

Exit Function
ResetError:
    ResetProgrammer = False
End Function

Public Function SendCommandGetResponse(command() As String, ByRef response() As String) As Boolean
'
'   PURPOSE: To add the necessary prefix and suffix bytes to the command
'            string, then send the string.  Then, receive the reponse from the
'            programmers.  Note that this routine handles communication with both
'            programmars simultaneously.
'
'  INPUT(S): command(1)  = string to be sent to programmer #1
'            command(2)  = string to be sent to programmer #2
' OUTPUT(S): response(1) = string read from programmer #1
'            response(2) = string read from programmer #2

Dim lintProgrammerNum As Integer
Dim lblnPerformCommunication(1 To 2) As Boolean
Dim lstrWriteData(1 To 2) As String
Dim lstrReadData(1 To 2) As String
Dim lblnGoodCRC As Boolean
Dim lsngStartTimer As Single
Dim lblnTimeOut As Boolean
Dim lintLengthOfResponse(1 To 2) As Integer
Dim lblnResponseReceived(1 To 2) As Boolean
Dim lintResponseCRC(1 To 2) As Integer
Dim lintCalculatedResponseCRC(1 To 2) As Integer

On Error GoTo ComError

'Initialize the routine to unsuccessful completion
SendCommandGetResponse = False

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    If command(lintProgrammerNum) <> "" Then
        'Add the number of bytes in the command as a prefix, and the CRC as a suffix
        lstrWriteData(lintProgrammerNum) = Chr$(Len(command(lintProgrammerNum))) & command(lintProgrammerNum) & Chr$(CalculateCommunicationCRC(command(lintProgrammerNum), lblnGoodCRC))
        'Verify the calculation of the CRC was successful
        If Not lblnGoodCRC Then GoTo ComError
        'If command() is an empty string, we skip that programmer number
        lblnPerformCommunication(lintProgrammerNum) = True
    End If
Next lintProgrammerNum

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Only send data if there is a valid string to send
    If lblnPerformCommunication(lintProgrammerNum) Then
        'Clear the input buffer for the next response
        frmMLX90293.SerialPort(lintProgrammerNum).InBufferCount = 0
        'Write the data to the programmer
        frmMLX90293.SerialPort(lintProgrammerNum).Output = lstrWriteData(lintProgrammerNum)
        'Loop until the message is sent out
        Do Until frmMLX90293.SerialPort(lintProgrammerNum).OutBufferCount = 0
            DoEvents
        Loop
        'Initialize the response to ""
        response(lintProgrammerNum) = ""
        'Initialize the ReadData to ""
        lstrReadData(lintProgrammerNum) = ""
        'Initialize the length of the response to 0
        lintLengthOfResponse(lintProgrammerNum) = 0
    End If
Next lintProgrammerNum

'Initialize the Response Timer
lsngStartTimer = Timer
Do
    'Loop through both programmers
    For lintProgrammerNum = 1 To 2
        'Only look for a response if we're performing the communication
        If lblnPerformCommunication(lintProgrammerNum) Then
            'Read from the serial buffer while looping, appending to the data variable
            lstrReadData(lintProgrammerNum) = lstrReadData(lintProgrammerNum) & frmMLX90293.SerialPort(lintProgrammerNum).Input
            'If we haven't found the length of the response yet, look for it
            If lintLengthOfResponse(lintProgrammerNum) = 0 Then
                'Look for the first character
                If Len(lstrReadData(lintProgrammerNum)) > 0 Then
                    'The length of the response is the ASCII code of the first character
                    lintLengthOfResponse(lintProgrammerNum) = Asc(left(lstrReadData(lintProgrammerNum), 1))
                End If
            'If we have found the length of the response...
            Else
                'Check if we've received all characters (Length of communication + prefix & suffix bytes)
                If Len(lstrReadData(lintProgrammerNum)) = lintLengthOfResponse(lintProgrammerNum) + 2 Then
                    lblnResponseReceived(lintProgrammerNum) = True
                End If
            End If
        Else
            'If there was no communication to send, pretend the response was received
            lblnResponseReceived(lintProgrammerNum) = True
        End If
        'Check for timeout
        If (Timer - lsngStartTimer > MELEXISREADTIMEOUT) Then lblnTimeOut = True
    Next lintProgrammerNum
    'If we've received both responses, exit the loop
    If lblnResponseReceived(1) And lblnResponseReceived(2) Then Exit Do
Loop Until lblnTimeOut

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'If we intended to perform a communication and we got the entire response...
    If lblnPerformCommunication(lintProgrammerNum) And lblnResponseReceived(lintProgrammerNum) Then
        'Get the CRC value (ASCII code number of the last byte)
        lintResponseCRC(lintProgrammerNum) = Asc(right(lstrReadData(lintProgrammerNum), 1))
        'Remove the prefix (length byte) and the suffix (CRC byte) to get the response
        response(lintProgrammerNum) = left(right(lstrReadData(lintProgrammerNum), lintLengthOfResponse(lintProgrammerNum) + 1), lintLengthOfResponse(lintProgrammerNum))
        'Calculate what the CRC byte should be
        lintCalculatedResponseCRC(lintProgrammerNum) = CalculateCommunicationCRC(response(lintProgrammerNum), lblnGoodCRC)
        'Verify the calculation of the CRC was successful
        If Not lblnGoodCRC Then GoTo ComError
        'Verify that the CRC byte receive from the programmer matches with what was calculated
        If lintResponseCRC(lintProgrammerNum) <> lintCalculatedResponseCRC(lintProgrammerNum) Then GoTo ComError
    ElseIf Not lblnResponseReceived(lintProgrammerNum) Then
        'Return whatever we did find (for ease of troubleshooting at higher levels)
        response(lintProgrammerNum) = lstrReadData(lintProgrammerNum)
        'Branch to show that the function did not complete successfully
        GoTo ComError
    End If
Next lintProgrammerNum

'The routine completed successfully, without error
SendCommandGetResponse = True

Exit Function
ComError:
    gblnGoodPTC04Link = False
End Function

Public Function EncodeCustomerID90293(DateCode As String) As Long
'
'   PURPOSE: To translate the date code into a Customer ID number
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lintYear As Integer
Dim lintJulianDate As Integer
Dim lstrShiftLetter As String
Dim lstShift As ShiftType
Dim llngCustomerID As Long
Dim lintStation As Integer

'********** Date Code Format **********
'  XX              XX              XXX         XX
' Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B18|B19   B17|B16|B15|B14|B13|B12|B11  B10|B09|B08|B07|B06|B05|B04|B03|B02  B01|B00
'Station        Year Beyond 2000                    Julian Date               Shift

'Initialize the Customer ID
llngCustomerID = 0

'Get the Year from the Date Code
lintYear = CInt(left(DateCode, 2))

'Get the Julian Date from the Date Code
lintJulianDate = CInt(Mid(DateCode, 3, 3))

'Get the Shift Letter from the Date Code
lstrShiftLetter = Mid(DateCode, 6, 1)

'Get the Station number from the Date Code
lintStation = CInt(Mid(DateCode, 7, 1))

'Encode the Shift Letter into a number
Select Case lstrShiftLetter
    Case "A"
        lstShift = stShiftA
    Case "B"
        lstShift = stShiftB
    Case "C"
        lstShift = stShiftC
End Select

'Add the Station (2-bit number bitshifted by 18 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintStation And &H3) * BIT18

'Add the Year (7-bit number bitshifted by 11 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintYear And &H7F) * BIT11

'Add the Julian Date (9-bit number bitshifted by 2 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintJulianDate And &H1FF) * BIT2

'Add the Shift (2-bit number) to the coded Date Code
llngCustomerID = llngCustomerID + (lstShift And &H3)

'Values for 90293 chip
glngCUSTID1 = CLng(lintJulianDate And &H1FF) + CLng(lintYear And &H7F) * BIT9
glngCUSTID2 = CLng(lstShift And &H3) + CLng(lintStation And &H3) * BIT2

'Return the Customer ID
EncodeCustomerID90293 = llngCustomerID

End Function

Public Function DecodeCustomerID90293() As String
'
'   PURPOSE: To translate the Customer ID number into a date code
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lstrYear As String
Dim lstrJulianDate As String
Dim lstrShiftLetter As String
Dim lstShift As ShiftType
Dim lintStation As Integer
Dim lstrDateCode As String

'********** Date Code Format **********
'   XX              XX              XXX         XX
' Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B18|B19   B17|B16|B15|B14|B13|B12|B11  B10|B09|B08|B07|B06|B05|B04|B03|B02  B01|B00
'Station        Year Beyond 2000                    Julian Date               Shift

'Initialize the Date Code
lstrDateCode = ""

'Get the Station (2-bit number bitshifted by 2 bits) from the Customer ID
lintStation = Format(((glngCUSTID2 \ BIT2) And &H3), "00")

'Get the Shift (2-bit number) from the Customer ID
lstShift = glngCUSTID2 And &H3

'Get the Year (7-bit number bitshifted by 9 bits) from the Customer ID
lstrYear = Format(((glngCUSTID1 \ BIT9) And &H7F), "00")

'Get the Julian Date (9-bit number) from the Customer ID
lstrJulianDate = Format((glngCUSTID1 And &H1FF), "000")

'Decode the Shift number into a letter
Select Case lstShift
    Case stShiftA
            lstrShiftLetter = "A"
    Case stShiftB
            lstrShiftLetter = "B"
    Case stShiftC
            lstrShiftLetter = "C"
    Case Else   'Anomalous Shift value
        lstrShiftLetter = "0"
End Select

'Build the date code
lstrDateCode = lstrYear & lstrJulianDate & lstrShiftLetter & CStr(lintStation)

'Return the Date Code
DecodeCustomerID90293 = lstrDateCode

End Function

Sub WriteCustID(ID As Integer)
'
'   PURPOSE: To write the Cust ID to the chip.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo WCIDErr

'Call MyDev1.SetEEParameterCode(CodeCUSTID1, glngCUSTID1)
'Call MyDev1.SetEEParameterCode(CodeCUSTID2, glngCUSTID2)

Exit Sub

WCIDErr:
    Call frmDDE.WriteDDEOutput(StationFault, 1)
    'gblnFaultSent = True
    MsgBox Err.Description, vbOKOnly, "Error in Write Cust ID"
    gintAnomaly = 52
    Err.Clear
End Sub

Sub ReadSN()
'
'   PURPOSE: To read the three MLX ID values and calc SN and lot.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lvntMLX1 As Variant
Dim lvntMLX2 As Variant
Dim lvntMLX3 As Variant
Dim lvntMLX4 As Variant
Dim lvntMLXL As Variant
Dim ldblMod As Double

On Error GoTo RSNErr

'lvntMLX1 = Hex$(CLng(MyDev.GetEEParameterCode(CodeMLXID1)) And &HFFFF)
'lvntMLX2 = Hex$(CLng(MyDev.GetEEParameterCode(CodeMLXID2)) And &HFFFF)
'lvntMLX3 = Hex$(CLng(MyDev.GetEEParameterCode(CodeMLXID3)) And &HFFFF)
'lvntMLX4 = Hex$(CLng(MyDev.GetEEParameterCode(CodeCUSTID1)) And &HFFFF)
'lvntMLXL = Hex$(CLng(MyDev.GetEEParameterCode(CodeLOCK)) And &HFFFF)
'
'glngMLXID1 = CLng("&H" & lvntMLX1)
'glngMLXID2 = CLng("&H" & lvntMLX2)
'glngMLXID3 = CLng("&H" & lvntMLX3)
'glngCUSTID1 = CLng("&H" & lvntMLX4)
'glngMLXLock = CLng("&H" & lvntMLXL) '1.2cANM
'
'If glngMLXID1 < 65000 Then
'    glngSN = glngMLXID2 + (glngMLXID1 * (2 ^ 16))
'    ldblMod = glngSN - (Int(glngSN / 1024) * 1024) '1.2bANM
'    glngLot = 8 & ldblMod + (Int(glngMLXID3 / 256) * 1024) '1.2bANM
'End If
'
'If (glngMLXID1 = 0) Or (glngMLXID2 = 0) Or (glngMLXID3 = 0) Or (glngMLXID1 = 65535) Or (glngMLXID2 = 65535) Or (glngMLXID3 = 65535) Or (glngCUSTID1 = 65535) Then
'    gintAnomaly = 52
'    Call ErrorLogFile(gintAnomaly)
'    MsgBox "Error reading 90333 Chip!", vbOKOnly, "MLX Error!"
'End If

Exit Sub

RSNErr:
    Call frmDDE.WriteDDEOutput(StationFault, 1)
    'gblnFaultSent = True
    MsgBox Err.Description, vbOKOnly, "Error in Read SN"
    gintAnomaly = 52
    Err.Clear
End Sub

Sub RunSolver()
'
'   PURPOSE: Programs the 90333 chip for HTT.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim Adv As PSF090333AAMLXAdvanced
'Dim Angle As Single
'Dim Level As Single
'Dim lintA As Integer
'
'On Error GoTo RSErr
'
'lintA = 0
'gblnSolverFail = False
'
''Set the Scan Velocity
'Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
''Set the Scan Acceleration
'Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
''Set the Scan Deceleration
'Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)
'
'If MyDev Is Nothing Then
'    Call frmDDE.WriteDDEOutput(StationFault, 1)
'    gblnFaultSent = True
'    MsgBox StrNotConnected
'    gintAnomaly = 150
'    Exit Sub
'Else
'    'MyDev.DeviceReplaced
'    'MyDev.SelectedDevice = 1
'    'MyDev.ReadFullDevice
'End If
'
''Clamps and solver defaults
'If MyDev Is Nothing Then
'    Call frmDDE.WriteDDEOutput(StationFault, 1)
'    gblnFaultSent = True
'    MsgBox StrNotConnected
'    gintAnomaly = 150
'    Exit Sub
'Else
'    Set Adv = MyDev.Advanced
'    'Call Adv.SetSolverSetting(SolverSettingClampingLow, 10)
'    'Call Adv.SetSolverSetting(SolverSettingClampingHigh, 90)
'    'Clamping levels set before programming, so gain and offset of D/A is correct
'
'    Call Adv.SetSolverSetting(SolverSettingApplicationSpeed, 0) '0 = low speed
'    Call Adv.SetSolverSetting(SolverSettingFilter, 5)           'filter = 5
'    Call Adv.SetSolverSetting(SolverSettingOutput1Mode, 2)      'out 1 = Analog mode2
'    Call Adv.SetSolverSetting(SolverSettingOutput2Mode, 0)      'out 2 = disabled
'    Call Adv.SetSolverSetting(SolverSettingAlphaDeadZone, 0)    ' Set Deadzone=0
'    Call Adv.SetSolverSetting(SolverSettingAlphaM180, 0)        ' Set M180 off
'    Call Adv.SetSolverSetting(SolverSettingBetaDeadZone, 0)     ' Set Deadzone=0
'    Call Adv.SetSolverSetting(SolverSettingBetaM180, 0)         ' Set M180 off
'    Call Adv.SetSolverSetting(SolverSettingAppMode, 1)          ' = 2D
'    Call Adv.SetSolverSetting(SolverSettingDP, 0)
'    Call Adv.SetSolverSetting(SolverSetting3Points, 1)          '3 point calibartion
'    Call Adv.SetSolverSetting(SolverSettingSPIMode, 0)          'Disable SPI
'
'    MyDev.Advanced.SolverSettingInUse(SolverSettingPWMFreq) = 0 'disable update of PWM
'
'    Call Adv.SetSolverSetting(SolverSettingMainMode, 3)         'Beta, derivate
'
'    MyDev.Solver.CopySolverSettingsToParameters
'
'    Call frmDAQIO.KillTime(100)
'
'    'Read/Disable all relays '1.2aANM \/\/
'    Dim lintIO As Integer
'    lintIO = MyDev.PTC04.GetDBIO
'    MyDev.PTC04.SetDBIO (0)
'    Call Adv.SetDBMux(0)
'
'    'DAC and MinGain
'    Call frmDAQIO.KillTime(100)
'    MyDev.Solver.CharacterizeOutputDAC
'    MyDev.Solver.SetGain
'
'    '1.2aANM \/\/ Check offset
'    Dim lsngOffset As Single
'    Dim lsngGain As Single
'    lsngOffset = MyDev.Advanced.GetSolverSetting(SolverSettingOutput1DACOffset)
'    lsngGain = MyDev.Advanced.GetSolverSetting(SolverSettingOutput1DACGain)
'    If (lsngOffset > 2) Or (lsngOffset < -2) Or (lsngGain > 2) Or (lsngGain < -2) Then
'        'Call frmDDE.WriteDDEOutput(StationFault, 1)
'        'gblnFaultSent = True
'        MsgBox "PTC-04 error # " & CStr(lintIO) & "! Reset device. Gain: " & CStr(lsngGain) & " Offset: " & CStr(lsngOffset)
'        gintAnomaly = 150
'        gblnFailure(ZERO) = True
'        Exit Sub
'    End If
'End If
'
''Move to idle (10%)
'Call VIX500IE.DefineMovement(gudtMachine.scanStart + gsngIndex1Loc + gintOffset)
'
''Start the motor
'Call VIX500IE.StartMotor
'
''Delay for motor move
'Call frmDAQIO.KillTime(150)
'
''Verify motor has stopped
'Do
'    mvntFirstPosition = Position
'    Call frmDAQIO.KillTime(50)
'Loop Until mvntFirstPosition = Position
'
''Def DP
'If MyDev Is Nothing Then
'    Call frmDDE.WriteDDEOutput(StationFault, 1)
'    gblnFaultSent = True
'    MsgBox StrNotConnected
'    gintAnomaly = 150
'    Exit Sub
'Else
'    Call MyDev.SetEEParameterValue(CodeALPHADP, 0) 'set dp to 0
'    Angle = MyDev.GetAlpha 'read first position
'
'    Select Case Angle
'        Case 0 To 90, 270 To 360  'in case first position = around 0 deg, then add DP = 180 deg
'            Call MyDev.SetEEParameterValue(CodeALPHADP, 180)
'        Case 91 To 270 'in case first position = around 180 deg, then do not change DP
'            Call MyDev.SetEEParameterValue(CodeALPHADP, 0)
'    End Select
'
'    'Set idle (Pt 1)
'    Level = gudtMachine.MLXProgPt1 '1.1aANM
'    MyDev.Solver.SetCoordinate0 (Level)
'
'    'Move to 30%
'    Call VIX500IE.DefineMovement(((gsngIndex2Loc - gsngIndex1Loc) / 4) + gintOffset + gudtMachine.scanStart + gsngIndex1Loc)
'
'    'Start the motor
'    Call VIX500IE.StartMotor
'
'    'Delay for motor move
'    Call frmDAQIO.KillTime(150)
'
'    'Verify motor has stopped
'    Do
'        mvntFirstPosition = Position
'        Call frmDAQIO.KillTime(50)
'    Loop Until mvntFirstPosition = Position
'
'    'Set Pt 2
'    Level = gudtMachine.MLXProgPt2 '1.1aANM
'    MyDev.Solver.SetCoordinateA (Level)
'
'    'Move to 50%
'    Call VIX500IE.DefineMovement(((gsngIndex2Loc - gsngIndex1Loc) / 2) + gintOffset + gudtMachine.scanStart + gsngIndex1Loc)
'
'    'Start the motor
'    Call VIX500IE.StartMotor
'
'    'Delay for motor move
'    Call frmDAQIO.KillTime(150)
'
'    'Verify motor has stopped
'    Do
'        mvntFirstPosition = Position
'        Call frmDAQIO.KillTime(50)
'    Loop Until mvntFirstPosition = Position
'
'    'Set Pt 3
'    Level = gudtMachine.MLXProgPt3 '1.1aANM
'    MyDev.Solver.SetCoordinateB (Level)
'
'    'Move to 70%
'    Call VIX500IE.DefineMovement((((gsngIndex2Loc - gsngIndex1Loc) * 3) / 4) + gintOffset + gudtMachine.scanStart + gsngIndex1Loc)
'
'    'Start the motor
'    Call VIX500IE.StartMotor
'
'    'Delay for motor move
'    Call frmDAQIO.KillTime(150)
'
'    'Verify motor has stopped
'    Do
'        mvntFirstPosition = Position
'        Call frmDAQIO.KillTime(50)
'    Loop Until mvntFirstPosition = Position
'
'    'Set Pt 4
'    Level = gudtMachine.MLXProgPt4 '1.1aANM
'    MyDev.Solver.SetCoordinateC (Level)
'
'    'Move to WOT (90%)
'    Call VIX500IE.DefineMovement((gsngIndex2Loc - gsngIndex1Loc) + gintOffset + gudtMachine.scanStart + gsngIndex1Loc)
'
'    'Start the motor
'    Call VIX500IE.StartMotor
'
'    'Delay for motor move
'    Call frmDAQIO.KillTime(150)
'
'    'Verify motor has stopped
'    Do
'        mvntFirstPosition = Position
'        Call frmDAQIO.KillTime(50)
'    Loop Until mvntFirstPosition = Position
'
'    'Set Pt 5
'    Level = gudtMachine.MLXProgPt5 '1.1aANM
'    MyDev.Solver.SetCoordinateD (Level)
'
'    'Program
'    Call MyDev.SetEEParameterValue(CodeCLAMPLOW, gudtMachine.MLXClampLo)
'    Call MyDev.SetEEParameterValue(CodeCLAMPHIGH, gudtMachine.MLXClampHi)
'    Call MyDev.SetEEParameterCode(CodeFHYST, 0)
'    Call MyDev.SetEEParameterCode(CodeDOUTINFAULT, 0)       ' Enable reset on fault
'    Call MyDev.SetEEParameterCode(CodeDRESONFAULT, 0)       ' Enable reset on fault
'
'ReDo:
'    lintA = lintA + 1
'    Dim lsngTimer As Single
'    Dim lsngTimeout As Single
'    lsngTimer = Timer '1.1.1ANM
'    MyDev.ProgramDevice
'    DoEvents
'    lsngTimeout = Timer - lsngTimer
'    If lsngTimeout > 5 Then
'        gblnSolverFail = True
'        Call frmDDE.WriteDDEOutput(StationFault, 1)
'        gblnFaultSent = True
'        MsgBox Err.Description, vbOKOnly, "Timeout in Run Solver"
'        gintAnomaly = 52
'    End If
'End If
'
''Return to the start location
'If gblnRevOnly Then '1.3eANM \/\/
'    'Move to End
'    Call VIX500IE.DefineMovement((gudtMachine.scanEnd + gintOffset) + overTravel)
'
'    'Start the motor
'    Call VIX500IE.StartMotor
'
'    'Delay for motor move
'    Call frmDAQIO.KillTime(150)
'
'    'Verify motor has stopped
'    Do
'        mvntFirstPosition = Position
'        Call frmDAQIO.KillTime(50)
'    Loop Until mvntFirstPosition = Position
'Else
'    Call MoveToLoadLocation
'End If
'
'MyDev.Advanced.DisconnectDevice
'
'Exit Sub
'
'RSErr:
'    Call frmDAQIO.KillTime(50)
'
'    If lintA < 3 Then
'        Resume ReDo
'    End If
'
'    gblnSolverFail = True
'
'    Call frmDDE.WriteDDEOutput(StationFault, 1)
'    'gblnFaultSent = True
'    MsgBox Err.Description, vbOKOnly, "Error in Run Solver"
'    gintAnomaly = 52
'    Err.Clear
End Sub

Sub RunSolver90293()
'
'   PURPOSE: Programs the 90293 chip.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim Adv1 As PSF090293AAMLXAdvanced
Dim Adv2 As PSF090293AAMLXAdvanced
Dim Angle As Single
Dim Level As Single
Dim lintA As Integer
Dim lsngIndexLoc(1 To 2) As Single
Dim fNTemp As String
Dim fNTemp2 As String
Dim llngAlg As Long
Dim llngFit As Long
Dim lsngErr As Single

llngAlg = 1

On Error GoTo RSErr

If MyDev(lintDev1) Is Nothing Then
    MsgBox StrNotConnected
Else
   Call MyDev(lintDev1).SetEEParameterCode(CodeVG, 19)
   'Call MyDev(lintDev1).ProgramDevice
   
   'Define settings
   Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPHIGH, gudtSolver(1).Clamp(2).IdealValue)
   Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPLOW, gudtSolver(1).Clamp(1).IdealValue)
   Call MyDev(lintDev1).SetEEParameterCode(CodeFILTER, 1)
   'Call mydev(lintDev1).Advanced.SetSetting(SettingDBLoad, 2)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingSolverType, SolverType2Points)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingOutput1Mode, 4)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingBoutMin, 5)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingBoutMax, 95)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDigOffset, 49152)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingBpivot, 128)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingFilter, 1)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingTreatSeq, 3)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS1, -128)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS2, -64)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS3, -64)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS5, 64)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS6, 64)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefOSYS7, 64)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS1, 0.824)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS2, 0.88)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS3, 0.936)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS5, 1.07)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS6, 1.07)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefSSYS7, 1.07)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefBTCmin, 10617)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingDefBTCmax, 54919)
   'mydev(lintDev1).Advanced.ChipVersion = 2
   'Call mydev(lintDev1).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
   Call MyDev(lintDev1).Solver.CopySolverSettingsToParameters
   
   Call MyDev(lintDev1).Solver.CharacterizeOutputDAC
   
   Call MyDev(lintDev1).Solver.SetFirstFitPoint(gudtSolver(1).Index(1).IdealValue)
End If

If MyDev(lintDev2) Is Nothing Then
    MsgBox StrNotConnected
Else
   Call MyDev(lintDev2).SetEEParameterCode(CodeVG, 19)
   'Call MyDev(lintDev2).ProgramDevice
   
   'Define settings
   Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPHIGH, gudtSolver(2).Clamp(2).IdealValue)
   Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPLOW, gudtSolver(2).Clamp(1).IdealValue)
   'Call mydev(lintDev2).SetEEParameterCode(CodeDISADCCLIP, 0)
   'Call mydev(lintDev2).Advanced.SetSetting(SettingDBLoad, 2)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingSolverType, SolverType2Points)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingOutput1Mode, 4)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingBoutMin, 5)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingBoutMax, 95)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDigOffset, 49152)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingBpivot, 128)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingFilter, 1)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingTreatSeq, 3)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS1, -64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS2, -64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS3, -64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS5, 64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS6, 64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefOSYS7, 64)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS1, 0.824)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS2, 0.88)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS3, 0.936)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS5, 1.07)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS6, 1.07)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefSSYS7, 1.07)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefBTCmin, 10617)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingDefBTCmax, 54919)
   'mydev(lintDev2).Advanced.ChipVersion = 2
   'Call mydev(lintDev2).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
   Call MyDev(lintDev2).Solver.CopySolverSettingsToParameters
   
   Call MyDev(lintDev2).Solver.CharacterizeOutputDAC
   
   Call MyDev(lintDev2).Solver.SetFirstFitPoint(gudtSolver(2).Index(1).IdealValue)
End If

'Find the appropriate position
'Activate the Sensotec SC2000's Tare function to remove any offset
If Sensotec.GetLinkStatus Then
    Call Sensotec.ActivateTare(1)
Else
    'Error Communicating with the Sensotec SC2000
    gintAnomaly = 50
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
End If
'Find the Second Programming Position
lsngIndexLoc(2) = FindSecondProgrammingPosition
'De-Activate the Sensotec SC2000's Tare Function
If Sensotec.GetLinkStatus Then
    Call Sensotec.DeActivateTare(1)
Else
    'Error Communicating with the Sensotec SC2000
    gintAnomaly = 50
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
End If
'Exit on system error
If gintAnomaly Then Exit Sub
'Set servo motor Velocity
Call VIX500IE.SetVelocity(gudtMachine.progVelocity)
'Set servo motor Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.progAcceleration)
'Set servo motor Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.progAcceleration)

'Move to Index 2
If Not (MoveToPosition(lsngIndexLoc(2), 1.5)) Then
    gintAnomaly = 163
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Motor Movement Failed While Solving.", True, True)
    Exit Sub
End If

'Get the current readings
gudtReading(0).mlxWCurrent = MyDev(lintDev1).GetIdd
'Lock the part if it is called for (and it almost always should be)
If gblnLockICs Then
    If (Not gblnLockRejects And (gintAnomaly <> 0)) Then gblnProgFailure = True
    If gblnProgFailure Then
        If gblnLockRejects Then
            'Set MemLock to True
            Call MyDev(lintDev1).SetEEParameterCode(CodeMEMLOCK, 3)
            Call MyDev(lintDev2).SetEEParameterCode(CodeMEMLOCK, 3)
            'Call MyDev(lintDev1).Solver.MEMLOCK
            'Call MyDev(lintDev2).Solver.MEMLOCK
            gblnLockedPart = True
        Else
            MsgBox "Part was not Locked!", vbOKOnly, "Melexis Status"
        End If
    Else
        'Set MemLock to True
        Call MyDev(lintDev1).SetEEParameterCode(CodeMEMLOCK, 3)
        Call MyDev(lintDev2).SetEEParameterCode(CodeMEMLOCK, 3)
        'Call MyDev(lintDev1).Solver.MEMLOCK
        'Call MyDev(lintDev2).Solver.MEMLOCK
        gblnLockedPart = True
    End If
Else
    MsgBox "Part was not Locked!", vbOKOnly, "Melexis Status"
End If


'Postion 2
If MyDev(lintDev1) Is Nothing Then
    MsgBox StrNotConnected
Else
  Call MyDev(lintDev1).Solver.SetNextFitPoint(gudtSolver(1).Index(2).IdealValue)
  Call MyDev(lintDev1).Solver.FitPoints(1, lsngErr)
  Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPHIGH, gudtSolver(1).Clamp(2).IdealValue)
  Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPLOW, gudtSolver(1).Clamp(1).IdealValue)
  Call MyDev(lintDev1).SetEEParameterCode(CodeFILTER, 1)
  Call MyDev(lintDev1).SetEEParameterCode(CodeENABLEHARDTHRESHOLD, 0)
  Call MyDev(lintDev1).SetEEParameterCode(CodeHARDTHRESHOLD, 0)
  Call MyDev(lintDev1).SetEEParameterCode(CodeDIAGSETTINGS, 264)
  Call MyDev(lintDev1).SetEEParameterCode(CodeOSMOD, 116)
  Call MyDev(lintDev1).ProgramDevice
End If

gudtReading(1).mlxWCurrent = MyDev(lintDev2).GetIdd

If MyDev(lintDev2) Is Nothing Then
    MsgBox StrNotConnected
Else
  Call MyDev(lintDev2).Solver.SetNextFitPoint(gudtSolver(2).Index(2).IdealValue)
  Call MyDev(lintDev2).Solver.FitPoints(1, lsngErr)
  Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPHIGH, gudtSolver(2).Clamp(2).IdealValue)
  Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPLOW, gudtSolver(2).Clamp(1).IdealValue)
  Call MyDev(lintDev2).SetEEParameterCode(CodeFILTER, 1)
  Call MyDev(lintDev2).SetEEParameterCode(CodeENABLEHARDTHRESHOLD, 0)
  Call MyDev(lintDev2).SetEEParameterCode(CodeHARDTHRESHOLD, 0)
  Call MyDev(lintDev2).SetEEParameterCode(CodeDIAGSETTINGS, 264)
  Call MyDev(lintDev2).SetEEParameterCode(CodeOSMOD, 116)
  Call MyDev(lintDev2).ProgramDevice
End If

'Move back
Call Pedal.MoveToLoadLocation

Exit Sub

RSErr:
    Call frmDAQIO.KillTime(50)
    
    gblnSolverFail = True
    
    Call frmDDE.WriteDDEOutput(StationFault, 1)
    'gblnFaultSent = True
    MsgBox Err.Description, vbOKOnly, "Error in Run Solver 90293"
    gintAnomaly = 52
    Err.Clear
    
End Sub

Sub SetLock()
'
'   PURPOSE: Locks the 90333 chip for HTT.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo SLErr

'Call MyDev.SetEEParameterCode(CodeMLXLOCK, 1)
'Call MyDev.SetEEParameterCode(CodeLOCK, 1)
'Call MyDev.ProgramDevice

Exit Sub

SLErr:
    Call frmDDE.WriteDDEOutput(StationFault, 1)
    'gblnFaultSent = True
    MsgBox Err.Description, vbOKOnly, "Error in Set Lock"
    gintAnomaly = 52
    Err.Clear
End Sub

Sub CreateDevice(bAutomatic As Boolean)
'    Dim PSFMan As PSF090293AAMLXManager
'    DevicesCol As ObjectCollection
'    i As Long
'
'    On Error GoTo lError
'
'    Call CleanUp
'
'    If bAutomatic Then
'        ' Automatic device scanning begins here
'        Set PSFMan = CreateObject("MPT.PSF090293AAMLXManager")
'        Set DevicesCol = PSFMan.ScanStandalone(dtAll)
'        If DevicesCol.Count <= 0 Then
'            MsgBox ("No 90293 were found!")
'            Exit Sub
'        End If
'
'        If DevicesCol.Count > 1 Then
'            For i = 1 To DevicesCol.Count - 1
'                'We are responsible to call Destroy(True) on device objects we do not need
'                Call DevicesCol(i).Destroy(True)
'            Next i
'        End If
'        Set MyDev = DevicesCol(0)
'        Set MyAdv = MyDev.Advanced
'        Set MySol = MyDev.Solver
'        'Call MyDev.Advanced.OpenProfile(ActiveWorkbook.Names("ProfileFileName").RefersToRange.Text)
'    Else
'        ' Manual connection begins here
'        Set MyDev1 = CreateObject("MPT.PSF090293AAMLXManager")
'        Call MyDev1.ConnectChannel(CVar(CLng(PTC04PORT1)), dtSerial)
'        ' Check if PTC04 programmer is connected to this channel
'        Call MyDev1.CheckSetup(False)
'
'        Set MyDev2 = CreateObject("MPT.PSF090293AAMLXManager")
'        Call MyDev2.ConnectChannel(CVar(CLng(PTC04PORT2)), dtSerial)
'        ' Check if PTC04 programmer is connected to this channel
'        Call MyDev2.CheckSetup(False)
'
'        Call Setup
'        gblnGoodPTC04Link = True
'    End If
'    Call MyDev.CheckSetup(False)
'
'    Exit Sub
'
'lError:
'    Set MyDev = Nothing
'    Call frmDDE.WriteDDEOutput(StationFault, 1)
'    'gblnFaultSent = True
'    MsgBox Err.Description, vbOKOnly, "Error in Create Device"
'    gintAnomaly = 52
'    gblnGoodPTC04Link = False
'    Err.Clear
End Sub

Sub Setup()
    'Settings
'    Call MyDev1.Advanced.SetSetting(SettingVDDLow, gudtPTC04(1).VDDLow)
'    Call MyDev1.Advanced.SetSetting(SettingVDDNorm, gudtPTC04(1).VDDNorm)
'    Call MyDev1.Advanced.SetSetting(SettingVDDComm, gudtPTC04(1).VDDComm)
'    Call MyDev1.Advanced.SetSetting(SettingVBatLow, gudtPTC04(1).VBatLow)
'    Call MyDev1.Advanced.SetSetting(SettingVBatNorm, gudtPTC04(1).VBatNorm)
'    Call MyDev1.Advanced.SetSetting(SettingTpor, gudtPTC04(1).Tpor)
'    Call MyDev1.Advanced.SetSetting(SettingTsynchro, gudtPTC04(1).Tsynchro)
'    Call MyDev1.Advanced.SetSetting(SettingTpulsMin, gudtPTC04(1).TpulsMin)
'    Call MyDev1.Advanced.SetSetting(SettingTpulsMax, gudtPTC04(1).TpulsMax)
'    Call MyDev1.Advanced.SetSetting(SettingSynchroDelay, gudtPTC04(1).SynchroDelay)
'    Call MyDev1.Advanced.SetSetting(SettingByteDelay, gudtPTC04(1).ByteData)
'    Call MyDev1.Advanced.SetSetting(SettingTSentTick, gudtPTC04(1).TSentTick)
'    Call MyDev1.Advanced.SetSetting(SettingBaudrate, gudtPTC04(1).Baudrate)
'    Call MyDev1.Advanced.SetSetting(SettingBaudrateSyncDiff, gudtPTC04(1).BaudrateSyncID)
'
'    Call MyDev2.Advanced.SetSetting(SettingVDDLow, gudtPTC04(2).VDDLow)
'    Call MyDev2.Advanced.SetSetting(SettingVDDNorm, gudtPTC04(2).VDDNorm)
'    Call MyDev2.Advanced.SetSetting(SettingVDDComm, gudtPTC04(2).VDDComm)
'    Call MyDev2.Advanced.SetSetting(SettingVBatLow, gudtPTC04(2).VBatLow)
'    Call MyDev2.Advanced.SetSetting(SettingVBatNorm, gudtPTC04(2).VBatNorm)
'    Call MyDev2.Advanced.SetSetting(SettingTpor, gudtPTC04(2).Tpor)
'    Call MyDev2.Advanced.SetSetting(SettingTsynchro, gudtPTC04(2).Tsynchro)
'    Call MyDev2.Advanced.SetSetting(SettingTpulsMin, gudtPTC04(2).TpulsMin)
'    Call MyDev2.Advanced.SetSetting(SettingTpulsMax, gudtPTC04(2).TpulsMax)
'    Call MyDev2.Advanced.SetSetting(SettingSynchroDelay, gudtPTC04(2).SynchroDelay)
'    Call MyDev2.Advanced.SetSetting(SettingByteDelay, gudtPTC04(2).ByteData)
'    Call MyDev2.Advanced.SetSetting(SettingTSentTick, gudtPTC04(2).TSentTick)
'    Call MyDev2.Advanced.SetSetting(SettingBaudrate, gudtPTC04(2).Baudrate)
'    Call MyDev2.Advanced.SetSetting(SettingBaudrateSyncDiff, gudtPTC04(2).BaudrateSyncID)
End Sub

Sub CleanUp()
'    Dim Man As CommManager
'
'    Set Man = CreateObject("MPT.CommManager")
'    If Not (MyDev Is Nothing) Then
'        ' Must call Destroy(True) to inform the object to prepare for shutdown
'        Call MyDev.Destroy(True)
'        Set MyDev = Nothing
'    End If
'    If Not (Man Is Nothing) Then
'        Man.Quit
'        Set Man = Nothing
'    End If
End Sub

Function hex2long(str As String) As Long
    hex2long = CLng("&H" & str)
End Function

Function long2hex(l As Long) As String
    long2hex = Hex$(l)
End Function
