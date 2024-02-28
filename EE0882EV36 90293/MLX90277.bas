Attribute VB_Name = "MLX90277"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) '2.5ANM

'*** Melexis 90277 Memory Address Map (Labels)***
'Melexis Area
Private Const MLXLOCK = 127:        Private Const DRIFT2 = 95
Private Const Y0 = 126:             Private Const TC2ND1 = 94
Private Const Y1 = 125:             Private Const CKANACH0 = 93
Private Const Y2 = 124:             Private Const TCW3TC04 = 92
Private Const Y3 = 123:             Private Const TCW2 = 91
Private Const Y4 = 122:             Private Const TCW3TC05 = 90
Private Const Y5 = 121:             Private Const FCKADJ3 = 89
Private Const Y6 = 120:             Private Const TC2ND5 = 88
Private Const X0 = 119:             Private Const TCW3TC06 = 87
Private Const X1 = 118:             Private Const FCKADJ0 = 86
Private Const X2 = 117:             Private Const CKANACH1 = 85
Private Const X3 = 116:             Private Const TCW3TC07 = 84
Private Const X4 = 115:             Private Const SLOW = 83
Private Const X5 = 114:             Private Const TC1 = 82
Private Const X6 = 113:             Private Const MLXPAR0 = 81
Private Const LOT0 = 112:           Private Const CKDACCH1 = 80
Private Const TCW3TC00 = 111:       Private Const CKDACCH0 = 79
Private Const TCW3TC01 = 110:       Private Const MLXPAR1 = 78
Private Const TC2ND2 = 109:         Private Const MLXPAR2 = 77
Private Const TCW3TC02 = 108:       Private Const AGND6B = 76
Private Const TCW0 = 107:           Private Const TC2 = 75
Private Const DRIFT1 = 106:         Private Const AGND3B = 74
Private Const FCKADJ1 = 105:        Private Const AGND5B = 73
Private Const TC2ND3 = 104:         Private Const AGND7B = 72
Private Const DRIFT3 = 103:         Private Const TC3 = 71
Private Const TC2ND0 = 102:         Private Const AGND8B = 70
Private Const TC4 = 101:            Private Const AGND1B = 69
Private Const TCW3TC03 = 100:       Private Const AGND2B = 68
Private Const TCW1 = 99:            Private Const TC0 = 67
Private Const DRIFT0 = 98:          Private Const AGND4B = 66
Private Const FCKADJ2 = 97:         Private Const AGND0B = 65
Private Const TC2ND4 = 96:          Private Const AGND9B = 64

'Customer Area
Private Const CLAMPHIGH9B = 63:     Private Const CLAMPHIGH4B = 31
Private Const OFFSET8B = 62:        Private Const CLAMPHIGH0B = 30
Private Const FILTER1 = 61:         Private Const CLAMPLOW4B = 29
Private Const FILTER0 = 60:         Private Const OFFSET4B = 28
Private Const TCW6TC00 = 59:        Private Const CLAMPHIGH2B = 27
Private Const FILTER3 = 58:         Private Const TCW6TC05 = 26
Private Const CLAMPLOW9B = 57:      Private Const MODE1 = 25
Private Const CLAMPLOW8B = 56:      Private Const FREE2 = 24
Private Const CLAMPHIGH7B = 55:     Private Const OFFSET0B = 23
Private Const OFFSET6B = 54:        Private Const CLAMPLOW0B = 22
Private Const CLAMPHIGH8B = 53:     Private Const RG2 = 21
Private Const FILTER2 = 52:         Private Const RG3 = 20
Private Const TCW6TC01 = 51:        Private Const INVERT = 19
Private Const TCW6TC02 = 50:        Private Const FG8 = 18
Private Const OFFSET9B = 49:        Private Const FG9 = 17
Private Const CLAMPLOW6B = 48:      Private Const FG5 = 16
Private Const CLAMPHIGH5B = 47:     Private Const FG3 = 15
Private Const CLAMPLOW5B = 46:      Private Const FG4 = 14
Private Const CLAMPHIGH6B = 45:     Private Const FG0 = 13
Private Const TCW6TC03 = 44:        Private Const FG6 = 12
Private Const TCW6TC04 = 43:        Private Const FG7 = 11
Private Const OFFSET7B = 42:        Private Const FG2 = 10
Private Const CLAMPLOW7B = 41:      Private Const FG1 = 9
Private Const OFFSET5B = 40:        Private Const RG0 = 8
Private Const CLAMPHIGH1B = 39:     Private Const RG1 = 7
Private Const CLAMPLOW1B = 38:      Private Const FAULTLEV = 6
Private Const OFFSET3B = 37:        Private Const MODE0 = 5
Private Const CLAMPLOW3B = 36:      Private Const STOP1 = 4
Private Const CLAMPHIGH3B = 35:     Private Const PARITY2 = 3
Private Const OFFSET1B = 34:        Private Const PARITY1 = 2
Private Const CLAMPLOW2B = 33:      Private Const PARITY0 = 1
Private Const OFFSET2B = 32:        Private Const MEMLOCK = 0

'*** Melexis 90277 Memory Address Map (IDs)***
'Melexis Area
Private Const FREE1 = 127:          Private Const TCW3TC08 = 95
Private Const WFR40 = 126:          Private Const TCW3TC09 = 94
Private Const WFR41 = 125:          Private Const TCW2TC00 = 93
Private Const WFR42 = 124:          Private Const TCW2TC01 = 92
Private Const WFR30 = 123:          Private Const TCW2TC02 = 91
Private Const WFR31 = 122:          Private Const TCW2TC03 = 90
Private Const WFR32 = 121:          Private Const TCW2TC04 = 89
Private Const WFR20 = 120:          Private Const TCW2TC05 = 88
Private Const WFR21 = 119:          Private Const TCW4TC150 = 87
Private Const WFR22 = 118:          Private Const TCW4TC151 = 86
Private Const WFR10 = 117:          Private Const TCW4TC152 = 85
Private Const WFR11 = 116:          Private Const TCW4TC153 = 84
Private Const WFR12 = 115:          Private Const TCW4TC154 = 83
Private Const WFR00 = 114:          Private Const TCW4TC155 = 82
Private Const WFR01 = 113:          Private Const TCW4TC310 = 81
Private Const WFR02 = 112:          Private Const TCW4TC311 = 80
Private Const LOT1 = 111:           Private Const TCW4TC312 = 79
Private Const LOT2 = 110:           Private Const TCW4TC313 = 78
Private Const LOT3 = 109:           Private Const TCW4TC314 = 77
Private Const LOT4 = 108:           Private Const TCW4TC315 = 76
Private Const LOT5 = 107:           Private Const TCW5TC310 = 75
Private Const LOT6 = 106:           Private Const TCW5TC311 = 74
Private Const LOT7 = 105:           Private Const TCW5TC312 = 73
Private Const LOT8 = 104:           Private Const TCW5TC313 = 72
Private Const LOT9 = 103:           Private Const TCW5TC314 = 71
Private Const LOT10 = 102:          Private Const TCW5TC315 = 70
Private Const LOT11 = 101:          Private Const MLXCRC0 = 69
Private Const LOT12 = 100:          Private Const MLXCRC1 = 68
Private Const LOT13 = 99:           Private Const MLXCRC2 = 67
Private Const LOT14 = 98:           Private Const MLXCRC3 = 66
Private Const LOT15 = 97:           Private Const MLXCRC4 = 65
Private Const LOT16 = 96:           Private Const MLXCRC5 = 64

'Customer Area
Private Const TCW5TC00 = 63:        Private Const HANDLING2 = 31
Private Const TCW5TC01 = 62:        Private Const CRC5 = 30
Private Const TCW5TC02 = 61:        Private Const CRC4 = 29
Private Const TCW5TC03 = 60:        Private Const CRC3 = 28
Private Const TCW5TC04 = 59:        Private Const CRC2 = 27
Private Const TCW5TC05 = 58:        Private Const CRC1 = 26
Private Const TCW3TC310 = 57:       Private Const CRC0 = 25
Private Const TCW3TC311 = 56:       Private Const CUSTID23 = 24
Private Const TCW3TC312 = 55:       Private Const CUSTID22 = 23
Private Const TCW3TC313 = 54:       Private Const CUSTID21 = 22
Private Const TCW3TC314 = 53:       Private Const CUSTID20 = 21
Private Const TCW3TC315 = 52:       Private Const CUSTID19 = 20
Private Const TCW2TC310 = 51:       Private Const CUSTID18 = 19
Private Const TCW2TC311 = 50:       Private Const CUSTID17 = 18
Private Const TCW2TC312 = 49:       Private Const CUSTID16 = 17
Private Const TCW2TC313 = 48:       Private Const CUSTID15 = 16
Private Const TCW2TC314 = 47:       Private Const CUSTID14 = 15
Private Const TCW2TC315 = 46:       Private Const CUSTID13 = 14
Private Const TCW1TC310 = 45:       Private Const CUSTID12 = 13
Private Const TCW1TC311 = 44:       Private Const CUSTID11 = 12
Private Const TCW1TC312 = 43:       Private Const CUSTID10 = 11
Private Const TCW1TC313 = 42:       Private Const CUSTID9 = 10
Private Const TCW1TC314 = 41:       Private Const CUSTID8 = 9
Private Const TCW1TC315 = 40:       Private Const CUSTID7 = 8
Private Const TCW4TC00 = 39:        Private Const CUSTID6 = 7
Private Const TCW4TC01 = 38:        Private Const CUSTID5 = 6
Private Const TCW4TC02 = 37:        Private Const CUSTID4 = 5
Private Const TCW4TC03 = 36:        Private Const CUSTID3 = 4
Private Const TCW4TC04 = 35:        Private Const CUSTID2 = 3
Private Const TCW4TC05 = 34:        Private Const CUSTID1 = 2
Private Const HANDLING0 = 33:       Private Const CUSTID0 = 1
Private Const HANDLING1 = 32:       Private Const FREE0 = 0

'*** Melexis 90277 Constants ***
Private Const MELEXISREADTIMEOUT = 10           'Melexis read timeout constant
Private Const NUMEEPROMLOCATIONS = 128          'Number of locations in the EEprom
Public Const CRCPOLYNOM = &H3                   'CRC Calculation Constant
Public Const CRCINIT = &H3F                     'CRC Calculation Constant
Public Const CRCXOR = &H3F                      'CRC Calculation Constant
Public Const CRCMASK = &H3F                     'CRC Calculation Constant
Public Const CRCHIGHBIT = &H20                  'CRC Calculation Constant

Type MelexisICContents
    MelexisLock     As Boolean                  'Melexis Lock
    MemoryLock      As Boolean                  'Memory Lock
    Y               As Integer                  'Y-Location of IC
    X               As Integer                  'X-Location of IC
    Wafer           As Integer                  'Wafer Number of IC
    Lot             As Double                   'Lot Number of IC
    Free            As Integer                  'Free address
    TCW1TC31        As Integer                  'TCW = 1, TC = 31 data  (6 bit signed integer)
    TCW2TC0         As Integer                  'TCW = 2, TC = 0 data   (6 bit signed integer)
    TCW2TC31        As Integer                  'TCW = 2, TC = 31 data  (6 bit signed integer)
    TCW3TC0         As Integer                  'TCW = 3, TC = 0 data   (10 bit signed integer)
    TCW3TC31        As Integer                  'TCW = 3, TC = 31 data  (6 bit signed integer)
    TCW4TC0         As Integer                  'TCW = 4, TC = 0 data   (6 bit signed integer)
    TCW4TC15        As Integer                  'TCW = 4, TC = 15 data  (6 bit signed integer)
    TCW4TC31        As Integer                  'TCW = 4, TC = 31 data  (6 bit signed integer)
    TCW5TC0         As Integer                  'TCW = 5, TC = 0 data   (6 bit signed integer)
    TCW5TC31        As Integer                  'TCW = 5, TC = 31 data  (6 bit signed integer)
    TCW6TC0         As Integer                  'TCW = 6, TC = 0 data   (6 bit signed integer)
    Handling        As Integer                  'Handling Information
    CustID          As Long                     'Customer Identification #
    MLXCRC          As Integer                  'Melexis CRC information
    CRC             As Integer                  'CRC Information
    TCWin           As Integer                  'Temperature Compensation Window
    TC              As Integer                  'First Order Temperature Compensation Coefficient
    TC2nd           As Integer                  'Second Order Temperature Compensation Coefficient
    Drift           As Integer                  'Offset Drift
    FCKADJ          As Integer                  'Oscillator Adjust
    CKANACH         As Integer                  'Clock Analog Select
    CKDACCH         As Integer                  'Clock Gen. Select
    SlowMode        As Boolean                  'Slow Mode Select
    Mode            As Integer                  'Mode Select
    InvertSlope     As Boolean                  'Invert Slope
    FaultLevel      As Boolean                  'EEPROM Fault
    RGain           As Integer                  'Rough Gain
    FGain           As Integer                  'Fine Gain
    offset          As Integer                  'Offset
    clampLow        As Integer                  'Low Output Clamp
    clampHigh       As Integer                  'High Output Clamp
    AGND            As Integer                  'Analog Ground
    Filter          As Integer                  'Filter Select
    Unlock          As Boolean                  'Stop/Unlock Bit
    MLXParity       As Integer                  'Parity count for MLX area of EEPROM
    TotalParity     As Integer                  'Total parity count for EEPROM
End Type

Type MLX90277
    'Represents Contents of entire EEPROM
    EEPROMContent(0 To (NUMEEPROMLOCATIONS - 1), 0 To 3)    As Boolean
    'Represents Contents of EEPROM LABEL Section
    LABELContent(0 To (NUMEEPROMLOCATIONS - 1))                    As Boolean
    'Represents Contents of EEPROM ID Section
    IDContent(0 To (NUMEEPROMLOCATIONS - 1))                       As Boolean
    'Reads & Writes
    Read As MelexisICContents                       'Reads from Melexis IC
    Write As MelexisICContents                      'Writes to Melexis IC
End Type

Type PTC04
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
End Type

Enum ModeFrameType             'Enumerated type to represent the different Melexis Mode Frames

    mftNormal = 0
    mftTemporary = 1
    mftBlockErase = 2
    mftBlockWrite = 3
    mftNormalRead = 4
    mftWordWrite = 5
    mftRowErase = 6

End Enum

Enum PTC04CommandType           'PTC-04 Enumerated Commands

    'Global Functions of the PTC-04
    ptcResetHardware = 0
    ptcGetHardwareID_Main = 1
    ptcGetHardwareID_Module = 2
    ptcGetSoftwareID = 3
    ptcGoto_BootLoader = 6
    ptcExit_BootLoader = 7
    ptcSendIntelHexFile = 8
    ptcGetProgramHexCode = 9
    'EEM128
    ptcGetContentsOfEEPROM = 10
    ptcSetContentsToEEPROM = 11
    ptcGetTextFromEEPROM = 12
    ptcSetTextToEEPROM = 13
    '12C
    ptcComm_12C = 24
    ptcGetContentsFrom_12C_EE = 25
    ptcSetContentsTo_12C_EE = 26
    'RAM
    ptcGetContentsFromCoreRAM = 30
    ptcGetSetContentsToCoreRam = 31
    'Other Commands
    ptcHello_World = 35
    'Timing
    ptcSetTiming = 60
    'Drivers
    ptcSetDAC = 70
    ptcSetFastDAC = 71
    ptcSetPPS = 72
    ptcSetLevel = 73
    ptcRunPattern = 74
    ptcSetRelay = 75
    ptcGetRelayStatus = 76
    'Measure
    ptcGetADC = 80
    ptcGetFilteredADC = 81
    ptcGetLevel = 82
    ptcSetMeasureDelay = 83
    ptcSetSampleDelay = 84
    ptcSetMeasureFilter = 85
    ptcGetSelectChannel = 86
    ptcGetCurrent = 87
    'Extension Support
    ptcSetDBIO = 90
    ptcWriteToDBExtension = 91
    ptcReadFromDBExtension = 92
    'MLX90251Cx/MLX90251FA Specific Functions   'V1.2 Updated, added #174
    ptc90251CxReadBack = 150
    ptc90251MeasureByRAM = 151
    ptc90251Program = 152
    ptc90251FastProgram = 153
    ptc90251FAReadBack = 174
    'MLX90244 Specific Functions                'V1.2
    ptc90244PWMByRAM = 170
    ptc90244ReadDump = 172
    ptc90244ReadBack = 174
End Enum

Enum PTC04SetupType         'PTC-04 Enumerated Channels & Timing

    'Channels
    ptcPPS_Vdd = 0          'Vdd Channel
    ptcPPS_Out = 1          'Output Channel

    ptcPPS_Vdd_I_Limit = 4  'Vdd Current Limit Channel
    ptcPPS_Out_I_Limit = 5  'Output Current Limit Channel

    'Levels
    ptcVddGnd = 0           'Vdd Ground Level
    ptcVddNom = 1           'Vdd Nominal Level
    ptcVddProg = 2          'Vdd Programming Level
    ptcOutLow = 0           'Vout Low Level
    ptcOutMid = 1           'Vout Mid Level
    ptcOutHigh = 2          'Vout High Level

    'Timing
    ptcTpor = 0             'Timing after going to a level
    ptcThold = 1            'Timing after Running Pulses before doing something else
    ptcTprog = 2            'Timing after Running Pulses before programming
    ptcTpuls = 3            'Timing between all edges of pulses on the output pin

End Enum

'Public gudtPTC04(1 To 2) As PTC04               'Instantiate the PTC-04 variables   (two - for two programmers)
Public gudtMLX90277(1 To 2) As MLX90277         'Instantiate the 90277 variables    (two - for two programmers)
Public gstrMLX90277Revision As String           'MLX90277 Revision Level
Public gblnGoodPTC04Link As Boolean             'Tracks whether or not communications are active with both programmers

Private Function BuildCommand(Mode As ModeFrameType, Data() As Boolean, Address As Integer) As String
'
'   PURPOSE: To build the command representing the sequence of 1's and 0's to
'            send to the programmer for communicating with the Melexis part.
'
'  INPUT(S): Mode       = Selects which Mode Frame to use
'            Data       = Data to be sent to the programmer
'            Address    = Address to be sent to the programmer
'
' OUTPUT(S): BuildCommand  = The function returns the command to send

Dim lintBitNum As Integer
Dim lstrSequence As String
Dim lintAddress As Integer
Dim lstrCommand As String
Dim lintByte As Integer
Dim lintBitPosition As Integer

'The "Mode Normal" sequence the programmer sends to the EEPROM contains:
'11 Mode Bits (The Mode Frame)
'1 Start Bit
'4 Data Bits (The Data Frame) (D0|D1|D2|D3 are voting bits in the EEPROM structure)
'7 Address Bits (The Address Frame)
'
'M0|M1|M2|M3|M4|M5|M6|M7|M8|M9|M10|SB|D0|D1|D2|D3|A0|A1|A2|A3|A4|A5|A6
'
'When using the "Temporary" Frame, the squence contains:
'11 Mode Bits (The Mode Frame)
'1 Start Bit
'123 Data Bits (The Temporary Frame) (B5|B6|B7... are the voted bits)
'M0|M1|M2|M3|M4|M5|M6|M7|M8|M9|M10|SB|B5|B6|B7|B8|...|B125|B126|B127
'

'Select the mode frame
Select Case Mode
    Case mftNormal
        lstrSequence = "01000000010"
    Case mftTemporary
        lstrSequence = "01100000010"
    Case mftBlockErase
        lstrSequence = "01000000001"
    Case mftBlockWrite
        lstrSequence = "01001000001"
    Case mftNormalRead
        lstrSequence = "01000100001"
    Case mftWordWrite
        lstrSequence = "01001100001"
    Case mftRowErase
        lstrSequence = "01011100001"
End Select

'Add Start Bit to the Sequence
lstrSequence = lstrSequence & "1"

'Add the Temporary or Data + Address Frame based on what mode we selected above
If Mode = mftTemporary Then
    '*** Temporary Frame ***
    'Loop through addresses 5 to 127
    For lintAddress = 5 To NUMEEPROMLOCATIONS - 1
        'Add Data (Bit D0)
        If Data(lintAddress, 0) Then
            lstrSequence = lstrSequence & "1"
        Else
            lstrSequence = lstrSequence & "0"
        End If
    Next lintAddress
Else
    '*** Data + Address Frame ***
    'Add Data (Bits D0 - D3)
    For lintBitNum = 0 To 3
    
        If Data(Address, lintBitNum) Then
            lstrSequence = lstrSequence & "1"
        Else
            lstrSequence = lstrSequence & "0"
        End If
    
    Next lintBitNum

    'Add Address (Bits A0 - A6)
    For lintBitNum = 0 To 6
    
        If (Address And (2 ^ lintBitNum)) > 0 Then
            lstrSequence = lstrSequence & "1"
        Else
            lstrSequence = lstrSequence & "0"
        End If
    
    Next lintBitNum

End If

'Now that we have defined the sequence of 1's and 0's, let's convert it
'to something the programmer will understand

'The first part of the command is the ASCII representation of the
'number of bits we are sending
lstrCommand = Chr$(32 + Len(lstrSequence))

'Now we iterate through the sequence, turning it into characters
'representing the sequence to the programmer:
'1)Iterate through each set of seven bits and determine the integer value
'  of those bits.
'2)Add 128 to force the MSB (eighth and somewhat imaginary bit) to be high.
'3)Determine the ASCII character for this integer, and add it to the
'  command string

'Initialize the variables
lintBitNum = 6: lintByte = 0

For lintBitPosition = 1 To Len(lstrSequence)
    If Mid$(lstrSequence, lintBitPosition, 1) = "1" Then lintByte = lintByte + 2 ^ lintBitNum

    'Decrement the bit number
    lintBitNum = lintBitNum - 1

    If lintBitNum < 0 Then
        'Add the proper character to the command string
        lstrCommand = lstrCommand & Chr$(lintByte + 128)
        'Reset the variables for next 7 bit word
        lintBitNum = 6: lintByte = 0
    End If

Next lintBitPosition

'If we didn't already add the last character, add it now
If lintBitNum <> 6 Then lstrCommand = lstrCommand & Chr$(lintByte + 128)


'Set the function equal to the final string
BuildCommand = lstrCommand

End Function

Public Function CalculateCommunicationCRC(command As String, calcOK As Boolean) As Integer
'
'   PURPOSE: To calculate CRC code for each command to the PTC-04 Programmer
'
'  INPUT(S): command = the command to be sent, interpreted as a series of bytes
'
' OUTPUT(S): calcOK  = returns whether or not the calculation successfully executed.
'            The function returns the calculated CRC code

Dim lintByteNum As Integer          'Index: which byte we're evaluating
Dim lintSum As Integer              'Sum of the instruction bytes
Dim lintCarries As Integer          'The sum of the carries for the additions
Dim lintCRC As Integer              'Calculated CRC
Dim lstrTest As String

On Error GoTo CRC_Error

calcOK = False                      'Initialize the routine to 'not finished correctly'
CalculateCommunicationCRC = 0       'Initialize the routine to return 0
lintSum = 0                         'Initialize the sum of instructions

'Loop through the bytes of the instruction, starting at byte 1
'(skip byte zero) and find the sum of the bytes of the instruction
For lintByteNum = 1 To Len(command)
    lintSum = lintSum + Asc(Mid(command, lintByteNum, 1))   'Keep a running Sum
    If lintSum > 255 Then lintSum = lintSum - 256 + 1       'If there is a carry, add the carry to the first byte
Next lintByteNum

'Complement the sum to get the CRC
lintCRC = 255 - lintSum

calcOK = True               'The function completed successfully
CalculateCommunicationCRC = lintCRC  'Return the value of the CRC

Exit Function
CRC_Error:
    calcOK = False      'The function did not complete successfully
End Function

Public Function CalculateDataCRC(CRCData() As Integer, Length As Integer, calcOK As Boolean) As Integer
'
'   PURPOSE: To calculate CRC codes (for 6-bit integers)
'            Based directly from Melexis C++ Source Code
'   * Melexis comment *
'   * fast bit by bit algorithm without augmented zero bytes.               *
'   * does not use lookup table, suited for polynom orders between 1...32.  *
'  INPUT(S): CRCData
'            length
'
' OUTPUT(S): The function returns the calculated CRC code

Dim lintCharacterNum As Integer     'Index: which character we're evaluating
Dim lintCRC As Integer              'CRC calculation variable
Dim lintBitMask As Integer          'BitMask For CRCData()
Dim lintBit As Integer              'Part of CRC Calculation

On Error GoTo CRC_Error

calcOK = False      'Initialize the routine to 'not finished correctly'
CalculateDataCRC = 0    'Initialize the routine to return 0

lintCRC = CRCINIT   'Initialize the CRC

'Loop through the six integers in the array passed in
For lintCharacterNum = 0 To Length - 1

    lintBitMask = BIT7  'Initialize the BitMask
    'Loop through the bits of the integer
    Do
        lintBit = lintCRC And CRCHIGHBIT
        'Discard the highest bits
        If lintCRC >= BIT14 Then
            lintCRC = lintCRC - BIT14
        End If
        lintCRC = lintCRC * 2
        If (CRCData(lintCharacterNum) And lintBitMask) Then
            lintBit = lintBit Xor CRCHIGHBIT
        End If
        If (lintBit) Then
            lintCRC = lintCRC Xor CRCPOLYNOM
        End If
        'Exit the loop after the the bitmask reaches 1
        If lintBitMask = 1 Then
            Exit Do
        Else
            lintBitMask = lintBitMask / 2
        End If
    Loop
  
Next lintCharacterNum
  
'Exclusively or with H3F, then And with H3F to invert the six bits
lintCRC = lintCRC Xor CRCXOR
lintCRC = lintCRC And CRCMASK

calcOK = True           'The function completed successfully
CalculateDataCRC = lintCRC  'Return the value of the CRC

Exit Function
CRC_Error:
    calcOK = False      'The function did not complete successfully
End Function

Private Sub ClearReadVariables(ProgrammerNum As Integer)
'
'   PURPOSE: To initialize all gudtMLX90277(programmerNum).Read.* variables to zero.
'
'  INPUT(S): None
' OUTPUT(S): None

gudtMLX90277(ProgrammerNum).Read.MelexisLock = False
gudtMLX90277(ProgrammerNum).Read.MemoryLock = False
gudtMLX90277(ProgrammerNum).Read.Y = 0
gudtMLX90277(ProgrammerNum).Read.X = 0
gudtMLX90277(ProgrammerNum).Read.Wafer = 0
gudtMLX90277(ProgrammerNum).Read.Lot = 0
gudtMLX90277(ProgrammerNum).Read.Free = 0
gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = 0
gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = 0
gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = 0
gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = 0
gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = 0
gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = 0
gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = 0
gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = 0
gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = 0
gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = 0
gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = 0
gudtMLX90277(ProgrammerNum).Read.Handling = 0
gudtMLX90277(ProgrammerNum).Read.CustID = 0
gudtMLX90277(ProgrammerNum).Read.MLXCRC = 0
gudtMLX90277(ProgrammerNum).Read.CRC = 0
gudtMLX90277(ProgrammerNum).Read.TCWin = 0
gudtMLX90277(ProgrammerNum).Read.TC = 0
gudtMLX90277(ProgrammerNum).Read.TC2nd = 0
gudtMLX90277(ProgrammerNum).Read.Drift = 0
gudtMLX90277(ProgrammerNum).Read.FCKADJ = 0
gudtMLX90277(ProgrammerNum).Read.CKANACH = 0
gudtMLX90277(ProgrammerNum).Read.CKDACCH = 0
gudtMLX90277(ProgrammerNum).Read.SlowMode = False
gudtMLX90277(ProgrammerNum).Read.Mode = 0
gudtMLX90277(ProgrammerNum).Read.InvertSlope = False
gudtMLX90277(ProgrammerNum).Read.FaultLevel = False
gudtMLX90277(ProgrammerNum).Read.RGain = 0
gudtMLX90277(ProgrammerNum).Read.FGain = 0
gudtMLX90277(ProgrammerNum).Read.offset = 0
gudtMLX90277(ProgrammerNum).Read.clampLow = 0
gudtMLX90277(ProgrammerNum).Read.clampHigh = 0
gudtMLX90277(ProgrammerNum).Read.AGND = 0
gudtMLX90277(ProgrammerNum).Read.Filter = 0
gudtMLX90277(ProgrammerNum).Read.Unlock = False
gudtMLX90277(ProgrammerNum).Read.MLXParity = 0
gudtMLX90277(ProgrammerNum).Read.TotalParity = 0

End Sub

Public Function CompareReadsAndWrites(ProgrammerNum As Integer) As Boolean
'
'   PURPOSE: To compare the read variables and write variables and make sure that
'            they are all matching.
'
'  INPUT(S): None.
' OUTPUT(S): Boolean representing whether or not the Reads and Writes are the same

'Initialize the routine to return False
CompareReadsAndWrites = False

'Exit the function if any of the reads don't match the writes
If gudtMLX90277(ProgrammerNum).Read.AGND <> gudtMLX90277(ProgrammerNum).Write.AGND Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.CKANACH <> gudtMLX90277(ProgrammerNum).Write.CKANACH Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.CKDACCH <> gudtMLX90277(ProgrammerNum).Write.CKDACCH Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.clampHigh <> gudtMLX90277(ProgrammerNum).Write.clampHigh Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.clampLow <> gudtMLX90277(ProgrammerNum).Write.clampLow Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.CRC <> gudtMLX90277(ProgrammerNum).Write.CRC Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.CustID <> gudtMLX90277(ProgrammerNum).Write.CustID Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Drift <> gudtMLX90277(ProgrammerNum).Write.Drift Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.FaultLevel <> gudtMLX90277(ProgrammerNum).Write.FaultLevel Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.FCKADJ <> gudtMLX90277(ProgrammerNum).Write.FCKADJ Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.FGain <> gudtMLX90277(ProgrammerNum).Write.FGain Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Filter <> gudtMLX90277(ProgrammerNum).Write.Filter Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Free <> gudtMLX90277(ProgrammerNum).Write.Free Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Handling <> gudtMLX90277(ProgrammerNum).Write.Handling Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.InvertSlope <> gudtMLX90277(ProgrammerNum).Write.InvertSlope Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Lot <> gudtMLX90277(ProgrammerNum).Write.Lot Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.MelexisLock <> gudtMLX90277(ProgrammerNum).Write.MelexisLock Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.MemoryLock <> gudtMLX90277(ProgrammerNum).Write.MemoryLock Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.MLXCRC <> gudtMLX90277(ProgrammerNum).Write.MLXCRC Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.MLXParity <> gudtMLX90277(ProgrammerNum).Write.MLXParity Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Mode <> gudtMLX90277(ProgrammerNum).Write.Mode Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.offset <> gudtMLX90277(ProgrammerNum).Write.offset Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.RGain <> gudtMLX90277(ProgrammerNum).Write.RGain Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.SlowMode <> gudtMLX90277(ProgrammerNum).Write.SlowMode Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TC <> gudtMLX90277(ProgrammerNum).Write.TC Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TC2nd <> gudtMLX90277(ProgrammerNum).Write.TC2nd Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW1TC31 <> gudtMLX90277(ProgrammerNum).Write.TCW1TC31 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW2TC0 <> gudtMLX90277(ProgrammerNum).Write.TCW2TC0 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW2TC31 <> gudtMLX90277(ProgrammerNum).Write.TCW2TC31 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW3TC0 <> gudtMLX90277(ProgrammerNum).Write.TCW3TC0 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW3TC31 <> gudtMLX90277(ProgrammerNum).Write.TCW3TC31 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW4TC0 <> gudtMLX90277(ProgrammerNum).Write.TCW4TC0 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW4TC15 <> gudtMLX90277(ProgrammerNum).Write.TCW4TC15 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW4TC31 <> gudtMLX90277(ProgrammerNum).Write.TCW4TC31 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW5TC0 <> gudtMLX90277(ProgrammerNum).Write.TCW5TC0 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW5TC31 <> gudtMLX90277(ProgrammerNum).Write.TCW5TC31 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCW6TC0 <> gudtMLX90277(ProgrammerNum).Write.TCW6TC0 Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TCWin <> gudtMLX90277(ProgrammerNum).Write.TCWin Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.TotalParity <> gudtMLX90277(ProgrammerNum).Write.TotalParity Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Unlock <> gudtMLX90277(ProgrammerNum).Write.Unlock Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Wafer <> gudtMLX90277(ProgrammerNum).Write.Wafer Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.X <> gudtMLX90277(ProgrammerNum).Write.X Then Exit Function
If gudtMLX90277(ProgrammerNum).Read.Y <> gudtMLX90277(ProgrammerNum).Write.Y Then Exit Function

'If we made it this far, everything compared ok
CompareReadsAndWrites = True

End Function

Public Sub CopyMLXReadsToMLXWrites(ProgrammerNum As Integer)
'
'   PURPOSE: To copy all MLX reads variables to the Write variables
'            NOTE: This subroutine is intended to be called after a Read, before
'            a Write.  This routine should be called BEFORE manipulating any
'            of the Write variables.  This insures that no memory locations
'            are left blank.
'
'  INPUT(S): None
' OUTPUT(S): None

gudtMLX90277(ProgrammerNum).Write.AGND = gudtMLX90277(ProgrammerNum).Read.AGND
gudtMLX90277(ProgrammerNum).Write.CKANACH = gudtMLX90277(ProgrammerNum).Read.CKANACH
gudtMLX90277(ProgrammerNum).Write.CKDACCH = gudtMLX90277(ProgrammerNum).Read.CKDACCH
gudtMLX90277(ProgrammerNum).Write.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh
gudtMLX90277(ProgrammerNum).Write.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow
gudtMLX90277(ProgrammerNum).Write.CRC = gudtMLX90277(ProgrammerNum).Read.CRC
gudtMLX90277(ProgrammerNum).Write.CustID = gudtMLX90277(ProgrammerNum).Read.CustID
gudtMLX90277(ProgrammerNum).Write.Drift = gudtMLX90277(ProgrammerNum).Read.Drift
gudtMLX90277(ProgrammerNum).Write.FaultLevel = gudtMLX90277(ProgrammerNum).Read.FaultLevel
gudtMLX90277(ProgrammerNum).Write.FCKADJ = gudtMLX90277(ProgrammerNum).Read.FCKADJ
gudtMLX90277(ProgrammerNum).Write.FGain = gudtMLX90277(ProgrammerNum).Read.FGain
gudtMLX90277(ProgrammerNum).Write.Filter = gudtMLX90277(ProgrammerNum).Read.Filter
gudtMLX90277(ProgrammerNum).Write.Free = gudtMLX90277(ProgrammerNum).Read.Free
gudtMLX90277(ProgrammerNum).Write.Handling = gudtMLX90277(ProgrammerNum).Read.Handling
gudtMLX90277(ProgrammerNum).Write.InvertSlope = gudtMLX90277(ProgrammerNum).Read.InvertSlope
gudtMLX90277(ProgrammerNum).Write.Lot = gudtMLX90277(ProgrammerNum).Read.Lot
gudtMLX90277(ProgrammerNum).Write.MelexisLock = gudtMLX90277(ProgrammerNum).Read.MelexisLock
gudtMLX90277(ProgrammerNum).Write.MemoryLock = gudtMLX90277(ProgrammerNum).Read.MemoryLock
gudtMLX90277(ProgrammerNum).Write.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC
gudtMLX90277(ProgrammerNum).Write.MLXParity = gudtMLX90277(ProgrammerNum).Read.MLXParity
gudtMLX90277(ProgrammerNum).Write.Mode = gudtMLX90277(ProgrammerNum).Read.Mode
gudtMLX90277(ProgrammerNum).Write.offset = gudtMLX90277(ProgrammerNum).Read.offset
gudtMLX90277(ProgrammerNum).Write.RGain = gudtMLX90277(ProgrammerNum).Read.RGain
gudtMLX90277(ProgrammerNum).Write.SlowMode = gudtMLX90277(ProgrammerNum).Read.SlowMode
gudtMLX90277(ProgrammerNum).Write.TC = gudtMLX90277(ProgrammerNum).Read.TC
gudtMLX90277(ProgrammerNum).Write.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd
gudtMLX90277(ProgrammerNum).Write.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31
gudtMLX90277(ProgrammerNum).Write.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0
gudtMLX90277(ProgrammerNum).Write.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31
gudtMLX90277(ProgrammerNum).Write.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0
gudtMLX90277(ProgrammerNum).Write.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31
gudtMLX90277(ProgrammerNum).Write.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0
gudtMLX90277(ProgrammerNum).Write.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15
gudtMLX90277(ProgrammerNum).Write.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31
gudtMLX90277(ProgrammerNum).Write.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0
gudtMLX90277(ProgrammerNum).Write.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31
gudtMLX90277(ProgrammerNum).Write.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0
gudtMLX90277(ProgrammerNum).Write.TCWin = gudtMLX90277(ProgrammerNum).Read.TCWin
gudtMLX90277(ProgrammerNum).Write.TotalParity = gudtMLX90277(ProgrammerNum).Read.TotalParity
gudtMLX90277(ProgrammerNum).Write.Unlock = gudtMLX90277(ProgrammerNum).Read.Unlock
gudtMLX90277(ProgrammerNum).Write.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer
gudtMLX90277(ProgrammerNum).Write.X = gudtMLX90277(ProgrammerNum).Read.X
gudtMLX90277(ProgrammerNum).Write.Y = gudtMLX90277(ProgrammerNum).Read.Y

End Sub

Public Function DecodeCustomerID(customerID As Long) As String
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
Dim lstrPalletLoad As String
Dim lstPallet As ShiftType

'********** Date Code Format **********
'    XX         XX              XX              XXX         XX
'PalletLoad | Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B20|B21     B18|B19   B17|B16|B15|B14|B13|B12|B11  B10|B09|B08|B07|B06|B05|B04|B03|B02  B01|B00
'PalletLoad  Station        Year Beyond 2000                    Julian Date               Shift

'Initialize the Date Code
lstrDateCode = ""

'Get the Pallet Load (2-bit number bitshifted by 20 bits) from the Customer ID
lstPallet = Format(((customerID \ BIT20) And &H3), "00")

'Get the Station (2-bit number bitshifted by 18 bits) from the Customer ID
lintStation = Format(((customerID \ BIT18) And &H3), "00")

'Get the Year (7-bit number bitshifted by 11 bits) from the Customer ID
lstrYear = Format(((customerID \ BIT11) And &H7F), "00")

'Get the Julian Date (9-bit number bitshifted by 2 bits) from the Customer ID
lstrJulianDate = Format(((customerID \ BIT2) And &H1FF), "000")

'Get the Shift (2-bit number) from the Customer ID
lstShift = customerID And &H3

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

'Decode the Pallet Load number into a letter
Select Case lstPallet
    Case stShiftA
            lstrPalletLoad = "A"
    Case stShiftB
            lstrPalletLoad = "B"
    Case Else   'Anomalous Pallet Load value
        lstrPalletLoad = "0"
End Select

'Build the date code
lstrDateCode = lstrYear & lstrJulianDate & lstrShiftLetter & CStr(lintStation) & lstrPalletLoad

'Return the Date Code
DecodeCustomerID = lstrDateCode

End Function

Public Function Decode705CustomerID(customerID As Long) As String
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
Dim lstrPalletLoad As String
Dim lstPallet As ShiftType
Dim lintStatus As Integer
Dim lintBOM As Integer

'********** Date Code Format **********
'XXXXX      XX         XX              XX              XXX         XX
' BOM | PalletLoad | Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B22-B26  B20|B21     B18|B19  B11-B17           B02-B10      B00|B01
' BOM     PalletLoad  Station  Year Beyond 2000  Julian Date  Shift

'Initialize the Date Code
lstrDateCode = ""

'Get the BOM (5-bit number bitshifted by 22 bits) from the Customer ID
lintBOM = Format(((customerID \ BIT22) And &H1F), "00")

'Get the Pallet Load (2-bit number bitshifted by 20 bits) from the Customer ID
lstPallet = Format(((customerID \ BIT20) And &H3), "00")

'Get the Station (2-bit number bitshifted by 18 bits) from the Customer ID
lintStation = Format(((customerID \ BIT18) And &H3), "00")

'Get the Year (7-bit number bitshifted by 11 bits) from the Customer ID
lstrYear = Format(((customerID \ BIT11) And &H7F), "00")

'Get the Julian Date (9-bit number bitshifted by 2 bits) from the Customer ID
lstrJulianDate = Format(((customerID \ BIT2) And &H1FF), "000")

'Get the Shift (2-bit number) from the Customer ID
lstShift = customerID And &H3

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

'Decode the Pallet Load number into a letter
Select Case lstPallet
    Case stShiftA
            lstrPalletLoad = "A"
    Case stShiftB
            lstrPalletLoad = "B"
    Case Else   'Anomalous Pallet Load value
        lstrPalletLoad = "0"
End Select

'Build the date code
lstrDateCode = lstrYear & lstrJulianDate & lstrShiftLetter & CStr(lintStation) & lstrPalletLoad

'Set variables
glngBOM = CLng(lintBOM)

'Return the Date Code
Decode705CustomerID = lstrDateCode

End Function

Public Sub DecodeDateCode(DateCode As String, Year As Integer, JulianDate As Integer, ShiftLetter As String, Station As Integer, PalletLoad As String)
'
'   PURPOSE: To decode the current Date Code into the date code information:
'            Julian Date, year, and shift
'
'  INPUT(S): DateCode   = The date code string
' OUTPUT(S): Year       = The year of the date code
'            JulianDate = The day of the year
'            Shift      = The shift letter
'            Station    = The station number
'            PalletLoad = The pallet load station

'********** Date Code Format **********
'       XX              XXX        X         X           X
'Year beyond 2000 | Julian Date | Shift | Station | PalletLoad

Year = CInt(Mid(gstrDateCode, 1, 2)) + 2000
JulianDate = CInt(Mid(gstrDateCode, 3, 3))
ShiftLetter = Mid(gstrDateCode, 6, 1)
Station = CInt(Mid(gstrDateCode, 7, 1))
PalletLoad = Mid(gstrDateCode, 8, 1)

End Sub

Public Sub DecodeEEpromRead(ProgrammerNum As Integer, EEPROMRead As String, VotingError As Boolean)
'
'   PURPOSE: To convert the bit data read from the EEprom into the
'            EEprom variable data.
'
'  INPUT(S): EEPROMRead = string read back from programmer
' OUTPUT(S): None.

Dim lintAddress As Integer
Dim lintNibble As Integer
Dim lintTemp As Integer
Dim lintBitLoc As Integer
Dim lintBitValue As Integer
Dim lblnWafer(0 To 4, 0 To 2) As Boolean
Dim lblnWaferVotes(0 To 4) As Boolean

'*** MELEXIS NOTES ***
'Note that each EEPROM location, 0 to 127, contains four bits (or a nibble),
'of information.  In the software, this information is stored in a
'two-dimensional boolean array:
'           EEPROMContent(EEPROM address, bit number), OR
'           EEPROMContent(0-127, 0-3)
'Each nibble represents two bits of information, an ID and a LABEL.
'The EEPROM can be thought of as two boolean arrays:
'           IDContent(EEPROM address), OR
'           IDContent(0-127)
'           LABELContent(EEPROM address), OR
'           LABELContent(0-127)
'Definitions for each EEPROM location are given in the constants section at the
'beginning of this module.
'  The nibble is defined as follows:
'   D3      D2      D1      D0
'    |       |       |       |
'    |       |       |       ---------------> LABEL section: voting bit 0
'    |       |       -----------------------> LABEL section: voting bit 1
'    |       -------------------------------> LABEL section: voting bit 2
'    ---------------------------------------> ID section: bit value

'Initialize the VotingError variable:
VotingError = False

'Decode the string read back from the programmer into the EEContent
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)

    'Convert the ASCII character from hex value to an integer
    lintNibble = Asc(Mid$(EEPROMRead, (lintAddress + 1), 1)) And &HF

    'Determine the individual bits in the EEPROM
    For lintBitLoc = 0 To 3
        If (lintNibble And (2 ^ lintBitLoc)) Then
            gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, lintBitLoc) = True
        Else
            gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, lintBitLoc) = False
        End If
    Next lintBitLoc

Next lintAddress

'Decode the values for the LABEL and ID sections of the EEPROM, and check for Voting Errors
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)

    'Set the contents of the LABEL array based on the voted bit
    gudtMLX90277(ProgrammerNum).LABELContent(lintAddress) = VoteBit(gudtMLX90277(ProgrammerNum).EEPROMContent(), lintAddress, VotingError)

    'The ID Section always = D3
    gudtMLX90277(ProgrammerNum).IDContent(lintAddress) = gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, 3)

Next lintAddress

'Convert LABEL and ID arrays into variable information
'Here, the address is looped through, from 0-127
'Two separate Case Statements assign the the variables with the
'data stored in the boolean arrays
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)

    '*** LABEL Section ***
    'Get the bit value from the LABEL Section
    'BitValue = 1 if True, 0 if False
    If gudtMLX90277(ProgrammerNum).LABELContent(lintAddress) Then
        lintBitValue = 1
    Else
        lintBitValue = 0
    End If

    'Select the current address and set the proper variables
    'according to the info at that address:
    Select Case lintAddress

        'Memlock
        Case MEMLOCK
            gudtMLX90277(ProgrammerNum).Read.MemoryLock = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

        'Parity
        Case PARITY0
            gudtMLX90277(ProgrammerNum).Read.TotalParity = gudtMLX90277(ProgrammerNum).Read.TotalParity + (lintBitValue * BIT0)
        Case PARITY1
            gudtMLX90277(ProgrammerNum).Read.TotalParity = gudtMLX90277(ProgrammerNum).Read.TotalParity + (lintBitValue * BIT1)
        Case PARITY2
            gudtMLX90277(ProgrammerNum).Read.TotalParity = gudtMLX90277(ProgrammerNum).Read.TotalParity + (lintBitValue * BIT2)

        'Stop1
        Case STOP1
            gudtMLX90277(ProgrammerNum).Read.Unlock = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

        'Mode
        Case MODE0
            gudtMLX90277(ProgrammerNum).Read.Mode = gudtMLX90277(ProgrammerNum).Read.Mode + (lintBitValue * BIT0)
        Case MODE1
            gudtMLX90277(ProgrammerNum).Read.Mode = gudtMLX90277(ProgrammerNum).Read.Mode + (lintBitValue * BIT1)

        'Fault Level
        Case FAULTLEV
            gudtMLX90277(ProgrammerNum).Read.FaultLevel = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

        'Rough Gain
        Case RG0
            gudtMLX90277(ProgrammerNum).Read.RGain = gudtMLX90277(ProgrammerNum).Read.RGain + (lintBitValue * BIT0)
        Case RG1
            gudtMLX90277(ProgrammerNum).Read.RGain = gudtMLX90277(ProgrammerNum).Read.RGain + (lintBitValue * BIT1)
        Case RG2
            gudtMLX90277(ProgrammerNum).Read.RGain = gudtMLX90277(ProgrammerNum).Read.RGain + (lintBitValue * BIT2)
        Case RG3
            gudtMLX90277(ProgrammerNum).Read.RGain = gudtMLX90277(ProgrammerNum).Read.RGain + (lintBitValue * BIT3)

        'Fine Gain
        Case FG0
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT0)
        Case FG1
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT1)
        Case FG2
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT2)
        Case FG3
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT3)
        Case FG4
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT4)
        Case FG5
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT5)
        Case FG6
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT6)
        Case FG7
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT7)
        Case FG8
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT8)
        Case FG9
            gudtMLX90277(ProgrammerNum).Read.FGain = gudtMLX90277(ProgrammerNum).Read.FGain + (lintBitValue * BIT9)

        'Invert Slope
        Case INVERT
            gudtMLX90277(ProgrammerNum).Read.InvertSlope = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

        'Offset
        Case OFFSET0B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT0)
        Case OFFSET1B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT1)
        Case OFFSET2B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT2)
        Case OFFSET3B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT3)
        Case OFFSET4B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT4)
        Case OFFSET5B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT5)
        Case OFFSET6B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT6)
        Case OFFSET7B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT7)
        Case OFFSET8B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT8)
        Case OFFSET9B
            gudtMLX90277(ProgrammerNum).Read.offset = gudtMLX90277(ProgrammerNum).Read.offset + (lintBitValue * BIT9)

        'Low Clamp
        Case CLAMPLOW0B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT0)
        Case CLAMPLOW1B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT1)
        Case CLAMPLOW2B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT2)
        Case CLAMPLOW3B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT3)
        Case CLAMPLOW4B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT4)
        Case CLAMPLOW5B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT5)
        Case CLAMPLOW6B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT6)
        Case CLAMPLOW7B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT7)
        Case CLAMPLOW8B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT8)
        Case CLAMPLOW9B
            gudtMLX90277(ProgrammerNum).Read.clampLow = gudtMLX90277(ProgrammerNum).Read.clampLow + (lintBitValue * BIT9)

        'High Clamp
        Case CLAMPHIGH0B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT0)
        Case CLAMPHIGH1B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT1)
        Case CLAMPHIGH2B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT2)
        Case CLAMPHIGH3B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT3)
        Case CLAMPHIGH4B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT4)
        Case CLAMPHIGH5B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT5)
        Case CLAMPHIGH6B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT6)
        Case CLAMPHIGH7B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT7)
        Case CLAMPHIGH8B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT8)
        Case CLAMPHIGH9B
            gudtMLX90277(ProgrammerNum).Read.clampHigh = gudtMLX90277(ProgrammerNum).Read.clampHigh + (lintBitValue * BIT9)

        'AGND
        Case AGND0B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT0)
        Case AGND1B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT1)
        Case AGND2B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT2)
        Case AGND3B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT3)
        Case AGND4B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT4)
        Case AGND5B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT5)
        Case AGND6B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT6)
        Case AGND7B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT7)
        Case AGND8B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT8)
        Case AGND9B
            gudtMLX90277(ProgrammerNum).Read.AGND = gudtMLX90277(ProgrammerNum).Read.AGND + (lintBitValue * BIT9)

        'MLX Parity
        Case MLXPAR0
            gudtMLX90277(ProgrammerNum).Read.MLXParity = gudtMLX90277(ProgrammerNum).Read.MLXParity + (lintBitValue * BIT0)
        Case MLXPAR1
            gudtMLX90277(ProgrammerNum).Read.MLXParity = gudtMLX90277(ProgrammerNum).Read.MLXParity + (lintBitValue * BIT1)
        Case MLXPAR2
            gudtMLX90277(ProgrammerNum).Read.MLXParity = gudtMLX90277(ProgrammerNum).Read.MLXParity + (lintBitValue * BIT2)

        'TC
        Case TC0
            gudtMLX90277(ProgrammerNum).Read.TC = gudtMLX90277(ProgrammerNum).Read.TC + (lintBitValue * BIT0)
        Case TC1
            gudtMLX90277(ProgrammerNum).Read.TC = gudtMLX90277(ProgrammerNum).Read.TC + (lintBitValue * BIT1)
        Case TC2
            gudtMLX90277(ProgrammerNum).Read.TC = gudtMLX90277(ProgrammerNum).Read.TC + (lintBitValue * BIT2)
        Case TC3
            gudtMLX90277(ProgrammerNum).Read.TC = gudtMLX90277(ProgrammerNum).Read.TC + (lintBitValue * BIT3)
        Case TC4
            gudtMLX90277(ProgrammerNum).Read.TC = gudtMLX90277(ProgrammerNum).Read.TC + (lintBitValue * BIT4)

        '2nd Order TC
        Case TC2ND0
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT0)
        Case TC2ND1
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT1)
        Case TC2ND2
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT2)
        Case TC2ND3
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT3)
        Case TC2ND4
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT4)
        Case TC2ND5
            gudtMLX90277(ProgrammerNum).Read.TC2nd = gudtMLX90277(ProgrammerNum).Read.TC2nd + (lintBitValue * BIT5)

        'Filter Select
        Case FILTER0
            gudtMLX90277(ProgrammerNum).Read.Filter = gudtMLX90277(ProgrammerNum).Read.Filter + (lintBitValue * BIT0)
        Case FILTER1
            gudtMLX90277(ProgrammerNum).Read.Filter = gudtMLX90277(ProgrammerNum).Read.Filter + (lintBitValue * BIT1)
        Case FILTER2
            gudtMLX90277(ProgrammerNum).Read.Filter = gudtMLX90277(ProgrammerNum).Read.Filter + (lintBitValue * BIT2)
        Case FILTER3
            gudtMLX90277(ProgrammerNum).Read.Filter = gudtMLX90277(ProgrammerNum).Read.Filter + (lintBitValue * BIT3)

        'Clock Gen. Select
        Case CKDACCH0
            gudtMLX90277(ProgrammerNum).Read.CKDACCH = gudtMLX90277(ProgrammerNum).Read.CKDACCH + (lintBitValue * BIT0)
        Case CKDACCH1
            gudtMLX90277(ProgrammerNum).Read.CKDACCH = gudtMLX90277(ProgrammerNum).Read.CKDACCH + (lintBitValue * BIT1)

        'Drift
        Case DRIFT0
            gudtMLX90277(ProgrammerNum).Read.Drift = gudtMLX90277(ProgrammerNum).Read.Drift + (lintBitValue * BIT0)
        Case DRIFT1
            gudtMLX90277(ProgrammerNum).Read.Drift = gudtMLX90277(ProgrammerNum).Read.Drift + (lintBitValue * BIT1)
        Case DRIFT2
            gudtMLX90277(ProgrammerNum).Read.Drift = gudtMLX90277(ProgrammerNum).Read.Drift + (lintBitValue * BIT2)
        Case DRIFT3
            gudtMLX90277(ProgrammerNum).Read.Drift = gudtMLX90277(ProgrammerNum).Read.Drift + (lintBitValue * BIT3)

        'TC Window
        Case TCW0
            gudtMLX90277(ProgrammerNum).Read.TCWin = gudtMLX90277(ProgrammerNum).Read.TCWin + (lintBitValue * BIT0)
        Case TCW1
            gudtMLX90277(ProgrammerNum).Read.TCWin = gudtMLX90277(ProgrammerNum).Read.TCWin + (lintBitValue * BIT1)
        Case TCW2
            gudtMLX90277(ProgrammerNum).Read.TCWin = gudtMLX90277(ProgrammerNum).Read.TCWin + (lintBitValue * BIT2)

        'Oscillator Adjust
        Case FCKADJ0
            gudtMLX90277(ProgrammerNum).Read.FCKADJ = gudtMLX90277(ProgrammerNum).Read.FCKADJ + (lintBitValue * BIT0)
        Case FCKADJ1
            gudtMLX90277(ProgrammerNum).Read.FCKADJ = gudtMLX90277(ProgrammerNum).Read.FCKADJ + (lintBitValue * BIT1)
        Case FCKADJ2
            gudtMLX90277(ProgrammerNum).Read.FCKADJ = gudtMLX90277(ProgrammerNum).Read.FCKADJ + (lintBitValue * BIT2)
        Case FCKADJ3
            gudtMLX90277(ProgrammerNum).Read.FCKADJ = gudtMLX90277(ProgrammerNum).Read.FCKADJ + (lintBitValue * BIT3)

        'Clock Analog Select
        Case CKANACH0
            gudtMLX90277(ProgrammerNum).Read.CKANACH = gudtMLX90277(ProgrammerNum).Read.CKANACH + (lintBitValue * BIT0)
        Case CKANACH1
            gudtMLX90277(ProgrammerNum).Read.CKANACH = gudtMLX90277(ProgrammerNum).Read.CKANACH + (lintBitValue * BIT1)

        'SlowMode Select
        Case SLOW
            gudtMLX90277(ProgrammerNum).Read.SlowMode = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

        'Free   (FREE0 & FREE1 are in the ID Section)
        Case FREE2
            gudtMLX90277(ProgrammerNum).Read.Free = gudtMLX90277(ProgrammerNum).Read.Free + (lintBitValue * BIT2)

        'TC Table   (TCW6TC0)
        Case TCW6TC00
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 + (lintBitValue * BIT0)
        Case TCW6TC01
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 + (lintBitValue * BIT1)
        Case TCW6TC02
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 + (lintBitValue * BIT2)
        Case TCW6TC03
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 + (lintBitValue * BIT3)
        Case TCW6TC04
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 + (lintBitValue * BIT4)
        Case TCW6TC05
            gudtMLX90277(ProgrammerNum).Read.TCW6TC0 = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 - (lintBitValue * BIT5)

        'TC Table   (TCW3TC0) (TCW3TC08 & TCW3TC09 are in the ID section)
        Case TCW3TC00
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT0)
        Case TCW3TC01
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT1)
        Case TCW3TC02
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT2)
        Case TCW3TC03
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT3)
        Case TCW3TC04
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT4)
        Case TCW3TC05
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT5)
        Case TCW3TC06
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT6)
        Case TCW3TC07
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT7)

        'Lot   (LOT1 - LOT16 are in the ID Section)
        Case LOT0
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT0)

        'X-location (of IC on wafer)
        Case X0
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT0)
        Case X1
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT1)
        Case X2
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT2)
        Case X3
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT3)
        Case X4
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT4)
        Case X5
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT5)
        Case X6
            gudtMLX90277(ProgrammerNum).Read.X = gudtMLX90277(ProgrammerNum).Read.X + (lintBitValue * BIT6)

        'Y-location (of IC on wafer)
        Case Y0
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT0)
        Case Y1
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT1)
        Case Y2
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT2)
        Case Y3
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT3)
        Case Y4
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT4)
        Case Y5
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT5)
        Case Y6
            gudtMLX90277(ProgrammerNum).Read.Y = gudtMLX90277(ProgrammerNum).Read.Y + (lintBitValue * BIT6)

        'MLX Lock
        Case MLXLOCK
            gudtMLX90277(ProgrammerNum).Read.MelexisLock = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)

    End Select

    '*** ID Section ***
    'Get the bit value from the ID Section
    'BitValue = 1 if True, 0 if False
    If gudtMLX90277(ProgrammerNum).IDContent(lintAddress) Then
        lintBitValue = 1
    Else
        lintBitValue = 0
    End If

    'Select the current address and set the proper variables
    'according to the info at that address:
    Select Case lintAddress

        'Free   (FREE2 is in the LABEL Section)
        Case FREE0
            gudtMLX90277(ProgrammerNum).Read.Free = gudtMLX90277(ProgrammerNum).Read.Free + (lintBitValue * BIT0)
        Case FREE1
            gudtMLX90277(ProgrammerNum).Read.Free = gudtMLX90277(ProgrammerNum).Read.Free + (lintBitValue * BIT1)

        'Cusomter ID
        Case CUSTID0
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT0)
        Case CUSTID1
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT1)
        Case CUSTID2
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT2)
        Case CUSTID3
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT3)
        Case CUSTID4
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT4)
        Case CUSTID5
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT5)
        Case CUSTID6
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT6)
        Case CUSTID7
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT7)
        Case CUSTID8
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT8)
        Case CUSTID9
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT9)
        Case CUSTID10
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT10)
        Case CUSTID11
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT11)
        Case CUSTID12
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT12)
        Case CUSTID13
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT13)
        Case CUSTID14
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT14)
        Case CUSTID15
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT15)
        Case CUSTID16
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT16)
        Case CUSTID17
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT17)
        Case CUSTID18
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT18)
        Case CUSTID19
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT19)
        Case CUSTID20
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT20)
        Case CUSTID21
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT21)
        Case CUSTID22
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT22)
        Case CUSTID23
            gudtMLX90277(ProgrammerNum).Read.CustID = gudtMLX90277(ProgrammerNum).Read.CustID + (lintBitValue * BIT23)

        'CRC
        Case CRC0
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT0)
        Case CRC1
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT1)
        Case CRC2
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT2)
        Case CRC3
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT3)
        Case CRC4
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT4)
        Case CRC5
            gudtMLX90277(ProgrammerNum).Read.CRC = gudtMLX90277(ProgrammerNum).Read.CRC + (lintBitValue * BIT5)

        'Handling
        Case HANDLING0
            gudtMLX90277(ProgrammerNum).Read.Handling = gudtMLX90277(ProgrammerNum).Read.Handling + (lintBitValue * BIT0)
        Case HANDLING1
            gudtMLX90277(ProgrammerNum).Read.Handling = gudtMLX90277(ProgrammerNum).Read.Handling + (lintBitValue * BIT1)
        Case HANDLING2
            gudtMLX90277(ProgrammerNum).Read.Handling = gudtMLX90277(ProgrammerNum).Read.Handling + (lintBitValue * BIT2)

        'TC Table   (TCW4TC0)
        Case TCW4TC00
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 + (lintBitValue * BIT0)
        Case TCW4TC01
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 + (lintBitValue * BIT1)
        Case TCW4TC02
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 + (lintBitValue * BIT2)
        Case TCW4TC03
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 + (lintBitValue * BIT3)
        Case TCW4TC04
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 + (lintBitValue * BIT4)
        Case TCW4TC05
            gudtMLX90277(ProgrammerNum).Read.TCW4TC0 = gudtMLX90277(ProgrammerNum).Read.TCW4TC0 - (lintBitValue * BIT5)

        'TC Table   (TCW1TC31)
        Case TCW1TC310
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 + (lintBitValue * BIT0)
        Case TCW1TC311
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 + (lintBitValue * BIT1)
        Case TCW1TC312
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 + (lintBitValue * BIT2)
        Case TCW1TC313
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 + (lintBitValue * BIT3)
        Case TCW1TC314
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 + (lintBitValue * BIT4)
        Case TCW1TC315
            gudtMLX90277(ProgrammerNum).Read.TCW1TC31 = gudtMLX90277(ProgrammerNum).Read.TCW1TC31 - (lintBitValue * BIT5)

        'TC Table   (TCW2TC31)
        Case TCW2TC310
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 + (lintBitValue * BIT0)
        Case TCW2TC311
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 + (lintBitValue * BIT1)
        Case TCW2TC312
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 + (lintBitValue * BIT2)
        Case TCW2TC313
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 + (lintBitValue * BIT3)
        Case TCW2TC314
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 + (lintBitValue * BIT4)
        Case TCW2TC315
            gudtMLX90277(ProgrammerNum).Read.TCW2TC31 = gudtMLX90277(ProgrammerNum).Read.TCW2TC31 - (lintBitValue * BIT5)

        'TC Table   (TCW3TC31)
        Case TCW3TC310
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 + (lintBitValue * BIT0)
        Case TCW3TC311
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 + (lintBitValue * BIT1)
        Case TCW3TC312
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 + (lintBitValue * BIT2)
        Case TCW3TC313
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 + (lintBitValue * BIT3)
        Case TCW3TC314
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 + (lintBitValue * BIT4)
        Case TCW3TC315
            gudtMLX90277(ProgrammerNum).Read.TCW3TC31 = gudtMLX90277(ProgrammerNum).Read.TCW3TC31 - (lintBitValue * BIT5)

        'TC Table   (TCW5TC0)
        Case TCW5TC00
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 + (lintBitValue * BIT0)
        Case TCW5TC01
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 + (lintBitValue * BIT1)
        Case TCW5TC02
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 + (lintBitValue * BIT2)
        Case TCW5TC03
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 + (lintBitValue * BIT3)
        Case TCW5TC04
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 + (lintBitValue * BIT4)
        Case TCW5TC05
            gudtMLX90277(ProgrammerNum).Read.TCW5TC0 = gudtMLX90277(ProgrammerNum).Read.TCW5TC0 - (lintBitValue * BIT5)

        'MLX CRC
        Case MLXCRC0
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT0)
        Case MLXCRC1
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT1)
        Case MLXCRC2
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT2)
        Case MLXCRC3
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT3)
        Case MLXCRC4
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT4)
        Case MLXCRC5
            gudtMLX90277(ProgrammerNum).Read.MLXCRC = gudtMLX90277(ProgrammerNum).Read.MLXCRC + (lintBitValue * BIT5)

        'TC Table   (TCW5TC31)
        Case TCW5TC310
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 + (lintBitValue * BIT0)
        Case TCW5TC311
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 + (lintBitValue * BIT1)
        Case TCW5TC312
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 + (lintBitValue * BIT2)
        Case TCW5TC313
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 + (lintBitValue * BIT3)
        Case TCW5TC314
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 + (lintBitValue * BIT4)
        Case TCW5TC315
            gudtMLX90277(ProgrammerNum).Read.TCW5TC31 = gudtMLX90277(ProgrammerNum).Read.TCW5TC31 - (lintBitValue * BIT5)

        'TC Table   (TCW4TC31)
        Case TCW4TC310
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 + (lintBitValue * BIT0)
        Case TCW4TC311
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 + (lintBitValue * BIT1)
        Case TCW4TC312
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 + (lintBitValue * BIT2)
        Case TCW4TC313
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 + (lintBitValue * BIT3)
        Case TCW4TC314
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 + (lintBitValue * BIT4)
        Case TCW4TC315
            gudtMLX90277(ProgrammerNum).Read.TCW4TC31 = gudtMLX90277(ProgrammerNum).Read.TCW4TC31 - (lintBitValue * BIT5)

        'TC Table   (TCW4TC15)
        Case TCW4TC150
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 + (lintBitValue * BIT0)
        Case TCW4TC151
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 + (lintBitValue * BIT1)
        Case TCW4TC152
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 + (lintBitValue * BIT2)
        Case TCW4TC153
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 + (lintBitValue * BIT3)
        Case TCW4TC154
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 + (lintBitValue * BIT4)
        Case TCW4TC155
            gudtMLX90277(ProgrammerNum).Read.TCW4TC15 = gudtMLX90277(ProgrammerNum).Read.TCW4TC15 - (lintBitValue * BIT5)

        'TC Table   (TCW2TC0)
        Case TCW2TC00
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 + (lintBitValue * BIT0)
        Case TCW2TC01
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 + (lintBitValue * BIT1)
        Case TCW2TC02
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 + (lintBitValue * BIT2)
        Case TCW2TC03
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 + (lintBitValue * BIT3)
        Case TCW2TC04
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 + (lintBitValue * BIT4)
        Case TCW2TC05
            gudtMLX90277(ProgrammerNum).Read.TCW2TC0 = gudtMLX90277(ProgrammerNum).Read.TCW2TC0 - (lintBitValue * BIT5)

        'TC Table   (TCW3TC0) (TCW3TC01 - TCW3TC07 are in the LABEL section)
        Case TCW3TC08
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 + (lintBitValue * BIT8)
        Case TCW3TC09
            gudtMLX90277(ProgrammerNum).Read.TCW3TC0 = gudtMLX90277(ProgrammerNum).Read.TCW3TC0 - (lintBitValue * BIT9)

        'Lot    (LOT0 is in the LABEL Section)
        Case LOT1
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT1)
        Case LOT2
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT2)
        Case LOT3
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT3)
        Case LOT4
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT4)
        Case LOT5
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT5)
        Case LOT6
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT6)
        Case LOT7
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT7)
        Case LOT8
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT8)
        Case LOT9
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT9)
        Case LOT10
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT10)
        Case LOT11
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT11)
        Case LOT12
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT12)
        Case LOT13
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT13)
        Case LOT14
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT14)
        Case LOT15
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT15)
        Case LOT16
            gudtMLX90277(ProgrammerNum).Read.Lot = gudtMLX90277(ProgrammerNum).Read.Lot + (lintBitValue * BIT16)

        'Wafer
        'NOTE: The WFR bits are voted like the LABEL section, but are "stacked"
        '      across three different ID addresses.
        Case WFR00
            lblnWafer(0, 0) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR01
            lblnWafer(0, 1) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR02
            lblnWafer(0, 2) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR10
            lblnWafer(1, 0) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR11
            lblnWafer(1, 1) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR12
            lblnWafer(1, 2) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR20
            lblnWafer(2, 0) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR21
            lblnWafer(2, 1) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR22
            lblnWafer(2, 2) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR30
            lblnWafer(3, 0) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR31
            lblnWafer(3, 1) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR32
            lblnWafer(3, 2) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR40
            lblnWafer(4, 0) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR41
            lblnWafer(4, 1) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
        Case WFR42
            lblnWafer(4, 2) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
            
    End Select

Next lintAddress

'Decode the Wafer locations (Voted Bits)
For lintAddress = 0 To 4  'Loop through the 5 Wafer spots

    'Vote the lblnWaferVotes bit
    lblnWaferVotes(lintAddress) = VoteBit(lblnWafer(), lintAddress, VotingError)

    'BitValue = 1 if True, 0 if False
    If lblnWaferVotes(lintAddress) Then
        lintBitValue = 1
    Else
        lintBitValue = 0
    End If

    'Determine the Wafer integer value
    Select Case lintAddress
        Case 0
            gudtMLX90277(ProgrammerNum).Read.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer + (lintBitValue * BIT0)
        Case 1
            gudtMLX90277(ProgrammerNum).Read.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer + (lintBitValue * BIT1)
        Case 2
            gudtMLX90277(ProgrammerNum).Read.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer + (lintBitValue * BIT2)
        Case 3
            gudtMLX90277(ProgrammerNum).Read.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer + (lintBitValue * BIT3)
        Case 4
            gudtMLX90277(ProgrammerNum).Read.Wafer = gudtMLX90277(ProgrammerNum).Read.Wafer + (lintBitValue * BIT4)

    End Select

Next lintAddress
    
End Sub

Public Function EncodeCustomerID(DateCode As String) As Long
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
Dim lstrPalletLoad As String
Dim lstPallet As ShiftType

'********** Date Code Format **********
'    XX         XX              XX              XXX         XX
'PalletLoad | Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B20|B21     B18|B19   B17|B16|B15|B14|B13|B12|B11  B10|B09|B08|B07|B06|B05|B04|B03|B02  B01|B00
'PalletLoad  Station        Year Beyond 2000                    Julian Date               Shift

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

'Get the pallet load letter from the Date Code
lstrPalletLoad = right(DateCode, 1)

'Encode the Shift Letter into a number
Select Case lstrShiftLetter
    Case "A"
        lstShift = stShiftA
    Case "B"
        lstShift = stShiftB
    Case "C"
        lstShift = stShiftC
End Select

'Encode the Pallet Load Letter into a number
Select Case lstrPalletLoad
    Case "A"
        lstPallet = stShiftA
    Case "B"
        lstPallet = stShiftB
End Select

'Add the Pallet Load (2-bit number bitshifted by 20 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lstPallet And &H3) * BIT20

'Add the Station (2-bit number bitshifted by 18 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintStation And &H3) * BIT18

'Add the Year (7-bit number bitshifted by 11 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintYear And &H7F) * BIT11

'Add the Julian Date (9-bit number bitshifted by 2 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintJulianDate And &H1FF) * BIT2

'Add the Shift (2-bit number) to the coded Date Code
llngCustomerID = llngCustomerID + (lstShift And &H3)

'Return the Customer ID
EncodeCustomerID = llngCustomerID

End Function

Public Function Encode705CustomerID(DateCode As String) As Long
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
Dim lstrPalletLoad As String
Dim lstPallet As ShiftType

'********** Date Code Format **********
'XXXXX      XX         XX              XX              XXX         XX
' BOM | PalletLoad | Station | Year beyond 2000 | Julian Date |  Shift

'********************* Customer ID Format (Date Code) *********************
'B22-B26  B20|B21     B18|B19   B17|B16|B15|B14|B13|B12|B11  B10|B09|B08|B07|B06|B05|B04|B03|B02  B01|B00
'  BOM    PalletLoad  Station        Year Beyond 2000                    Julian Date               Shift

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

'Get the pallet load letter from the Date Code
lstrPalletLoad = right(DateCode, 1)

'Encode the Shift Letter into a number
Select Case lstrShiftLetter
    Case "A"
        lstShift = stShiftA
    Case "B"
        lstShift = stShiftB
    Case "C"
        lstShift = stShiftC
End Select

'Encode the Pallet Load Letter into a number
Select Case lstrPalletLoad
    Case "A"
        lstPallet = stShiftA
    Case "B"
        lstPallet = stShiftB
End Select

'Add the BOM (5-bit number bitshifted by 22 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(gudtMachine.BOMNumber And &H1F) * BIT22

'Add the Pallet Load (2-bit number bitshifted by 20 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lstPallet And &H3) * BIT20

'Add the Station (2-bit number bitshifted by 18 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintStation And &H3) * BIT18

'Add the Year (7-bit number bitshifted by 11 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintYear And &H7F) * BIT11

'Add the Julian Date (9-bit number bitshifted by 2 bits) to the coded Date Code
llngCustomerID = llngCustomerID + CLng(lintJulianDate And &H1FF) * BIT2

'Add the Shift (2-bit number) to the coded Date Code
llngCustomerID = llngCustomerID + (lstShift And &H3)

'Return the Customer ID
Encode705CustomerID = llngCustomerID

End Function

Public Sub EncodeEEpromWrite(ProgrammerNum As Integer)
'
'   PURPOSE: To convert the EEprom variable data into bit data to be
'            written to the EEprom.  Also calculates parity.
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lintAddress As Integer
Dim lblnBitValue As Boolean
Dim lintTotalParityCnt As Integer
Dim lintMLXParityCnt As Integer

'Initialize the Parity Count
lintTotalParityCnt = 0
lintMLXParityCnt = 0

'Convert EEprom variables into LABEL data
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)

    'Default to false
    lblnBitValue = False

    Select Case lintAddress

        'Memlock
        Case MEMLOCK
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MemoryLock

        'Parity
        'Calculated below & set below

        'Stop1
        Case STOP1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Unlock

        'Mode
        Case MODE0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Mode And BIT0
        Case MODE1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Mode And BIT1

        'Fault Level
        Case FAULTLEV
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FaultLevel

        'Rough Gain
        Case RG0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.RGain And BIT0
        Case RG1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.RGain And BIT1
        Case RG2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.RGain And BIT2
        Case RG3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.RGain And BIT3

        'Fine Gain
        Case FG0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT0
        Case FG1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT1
        Case FG2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT2
        Case FG3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT3
        Case FG4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT4
        Case FG5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT5
        Case FG6
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT6
        Case FG7
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT7
        Case FG8
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT8
        Case FG9
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FGain And BIT9

        'Invert Slope
        Case INVERT
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.InvertSlope

        'Offset
        Case OFFSET0B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT0
        Case OFFSET1B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT1
        Case OFFSET2B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT2
        Case OFFSET3B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT3
        Case OFFSET4B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT4
        Case OFFSET5B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT5
        Case OFFSET6B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT6
        Case OFFSET7B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT7
        Case OFFSET8B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT8
        Case OFFSET9B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.offset And BIT9

        'Low Clamp
        Case CLAMPLOW0B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT0
        Case CLAMPLOW1B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT1
        Case CLAMPLOW2B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT2
        Case CLAMPLOW3B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT3
        Case CLAMPLOW4B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT4
        Case CLAMPLOW5B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT5
        Case CLAMPLOW6B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT6
        Case CLAMPLOW7B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT7
        Case CLAMPLOW8B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT8
        Case CLAMPLOW9B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampLow And BIT9

        'High Clamp
        Case CLAMPHIGH0B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT0
        Case CLAMPHIGH1B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT1
        Case CLAMPHIGH2B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT2
        Case CLAMPHIGH3B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT3
        Case CLAMPHIGH4B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT4
        Case CLAMPHIGH5B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT5
        Case CLAMPHIGH6B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT6
        Case CLAMPHIGH7B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT7
        Case CLAMPHIGH8B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT8
        Case CLAMPHIGH9B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.clampHigh And BIT9

        'AGND
        Case AGND0B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT0
        Case AGND1B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT1
        Case AGND2B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT2
        Case AGND3B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT3
        Case AGND4B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT4
        Case AGND5B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT5
        Case AGND6B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT6
        Case AGND7B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT7
        Case AGND8B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT8
        Case AGND9B
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.AGND And BIT9

        'MLX Parity
        Case MLXPAR0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXParity And BIT0
        Case MLXPAR1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXParity And BIT1
        Case MLXPAR2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXParity And BIT2

        'TC
        Case TC0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC And BIT0
        Case TC1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC And BIT1
        Case TC2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC And BIT2
        Case TC3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC And BIT3
        Case TC4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC And BIT4

        '2nd Order TC
        Case TC2ND0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT0
        Case TC2ND1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT1
        Case TC2ND2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT2
        Case TC2ND3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT3
        Case TC2ND4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT4
        Case TC2ND5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TC2nd And BIT5

        'Filter Select
        Case FILTER0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Filter And BIT0
        Case FILTER1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Filter And BIT1
        Case FILTER2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Filter And BIT2
        Case FILTER3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Filter And BIT3

        'Clock Gen. Select
        Case CKDACCH0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CKDACCH And BIT0
        Case CKDACCH1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CKDACCH And BIT1

        'Drift
        Case DRIFT0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Drift And BIT0
        Case DRIFT1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Drift And BIT1
        Case DRIFT2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Drift And BIT2
        Case DRIFT3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Drift And BIT3

        'TC Window
        Case TCW0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCWin And BIT0
        Case TCW1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCWin And BIT1
        Case TCW2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCWin And BIT2

        'Oscillator Adjust
        Case FCKADJ0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FCKADJ And BIT0
        Case FCKADJ1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FCKADJ And BIT1
        Case FCKADJ2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FCKADJ And BIT2
        Case FCKADJ3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.FCKADJ And BIT3

        'Clock Analog Select
        Case CKANACH0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CKANACH And BIT0
        Case CKANACH1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CKANACH And BIT1

        'SlowMode Select
        Case SLOW
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.SlowMode

        'Free   (FREE0 & FREE1 are in the ID Section)
        Case FREE2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Free And BIT2

        'TC Table   (TCW6TC0)
        Case TCW6TC00
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 And BIT0
        Case TCW6TC01
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 And BIT1
        Case TCW6TC02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 And BIT2
        Case TCW6TC03
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 And BIT3
        Case TCW6TC04
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 And BIT4
        Case TCW6TC05   'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW6TC0 < 0

        'TC Table   (TCW3TC0) (TCW3TC08 & TCW3TC09 are in the ID section)
        Case TCW3TC00
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT0
        Case TCW3TC01
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT1
        Case TCW3TC02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT2
        Case TCW3TC03
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT3
        Case TCW3TC04
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT4
        Case TCW3TC05
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT5
        Case TCW3TC06
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT6
        Case TCW3TC07
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT7

        'Lot   (LOT1 - LOT16 are in the ID Section)
        Case LOT0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT0

        'X-location (of IC on wafer)
        Case X0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT0
        Case X1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT1
        Case X2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT2
        Case X3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT3
        Case X4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT4
        Case X5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT5
        Case X6
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.X And BIT6

        'Y-location (of IC on wafer)
        Case Y0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT0
        Case Y1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT1
        Case Y2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT2
        Case Y3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT3
        Case Y4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT4
        Case Y5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT5
        Case Y6
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Y And BIT6

        'MLX Lock
        Case MLXLOCK
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MelexisLock

    End Select

    'Set the LABEL value at this address
    gudtMLX90277(ProgrammerNum).LABELContent(lintAddress) = lblnBitValue

    'Count for Parity, but don't count for the three Parity bits
    If ((lintAddress <> PARITY0) And (lintAddress <> PARITY1) And (lintAddress <> PARITY2)) Then
        If lblnBitValue Then
            lintTotalParityCnt = lintTotalParityCnt + 1
        End If
    End If

Next lintAddress

'Calculate TotalParity (0-127)
gudtMLX90277(ProgrammerNum).Write.TotalParity = (2 ^ (3 - (lintTotalParityCnt And &H3))) - 1

'Fill the Total Parity Bit locations
gudtMLX90277(ProgrammerNum).LABELContent(PARITY0) = gudtMLX90277(ProgrammerNum).Write.TotalParity And BIT0
gudtMLX90277(ProgrammerNum).LABELContent(PARITY1) = gudtMLX90277(ProgrammerNum).Write.TotalParity And BIT1
gudtMLX90277(ProgrammerNum).LABELContent(PARITY2) = gudtMLX90277(ProgrammerNum).Write.TotalParity And BIT2

'Convert EEprom variables into ID data
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)

    'Default to false
    lblnBitValue = False

    Select Case lintAddress
    
        'Free   (FREE2 is in the LABEL Section)
        Case FREE0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Free And BIT0
        Case FREE1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Free And BIT1

        'Cusomter ID
        Case CUSTID0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT0
        Case CUSTID1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT1
        Case CUSTID2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT2
        Case CUSTID3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT3
        Case CUSTID4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT4
        Case CUSTID5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT5
        Case CUSTID6
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT6
        Case CUSTID7
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT7
        Case CUSTID8
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT8
        Case CUSTID9
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT9
        Case CUSTID10
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT10
        Case CUSTID11
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT11
        Case CUSTID12
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT12
        Case CUSTID13
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT13
        Case CUSTID14
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT14
        Case CUSTID15
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT15
        Case CUSTID16
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT16
        Case CUSTID17
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT17
        Case CUSTID18
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT18
        Case CUSTID19
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT19
        Case CUSTID20
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT20
        Case CUSTID21
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT21
        Case CUSTID22
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT22
        Case CUSTID23
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CustID And BIT23

        'CRC
        Case CRC0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT0
        Case CRC1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT1
        Case CRC2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT2
        Case CRC3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT3
        Case CRC4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT4
        Case CRC5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.CRC And BIT5

        'Handling
        Case HANDLING0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Handling And BIT0
        Case HANDLING1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Handling And BIT1
        Case HANDLING2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Handling And BIT2

        'TC Table   (TCW4TC0)
        Case TCW4TC00
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 And BIT0
        Case TCW4TC01
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 And BIT1
        Case TCW4TC02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 And BIT2
        Case TCW4TC03
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 And BIT3
        Case TCW4TC04
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 And BIT4
        Case TCW4TC05   'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC0 < 0

        'TC Table   (TCW1TC31)
        Case TCW1TC310
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 And BIT0
        Case TCW1TC311
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 And BIT1
        Case TCW1TC312
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 And BIT2
        Case TCW1TC313
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 And BIT3
        Case TCW1TC314
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 And BIT4
        Case TCW1TC315  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW1TC31 < 0

        'TC Table   (TCW2TC31)
        Case TCW2TC310
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 And BIT0
        Case TCW2TC311
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 And BIT1
        Case TCW2TC312
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 And BIT2
        Case TCW2TC313
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 And BIT3
        Case TCW2TC314
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 And BIT4
        Case TCW2TC315  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC31 < 0

        'TC Table   (TCW3TC31)
        Case TCW3TC310
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 And BIT0
        Case TCW3TC311
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 And BIT1
        Case TCW3TC312
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 And BIT2
        Case TCW3TC313
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 And BIT3
        Case TCW3TC314
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 And BIT4
        Case TCW3TC315  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC31 < 0

        'TC Table   (TCW5TC0)
        Case TCW5TC00
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 And BIT0
        Case TCW5TC01
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 And BIT1
        Case TCW5TC02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 And BIT2
        Case TCW5TC03
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 And BIT3
        Case TCW5TC04
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 And BIT4
        Case TCW5TC05   'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC0 < 0

        'MLX CRC
        Case MLXCRC0
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT0
        Case MLXCRC1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT1
        Case MLXCRC2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT2
        Case MLXCRC3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT3
        Case MLXCRC4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT4
        Case MLXCRC5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.MLXCRC And BIT5

        'TC Table   (TCW5TC31)
        Case TCW5TC310
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 And BIT0
        Case TCW5TC311
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 And BIT1
        Case TCW5TC312
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 And BIT2
        Case TCW5TC313
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 And BIT3
        Case TCW5TC314
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 And BIT4
        Case TCW5TC315  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW5TC31 < 0

        'TC Table   (TCW4TC31)
        Case TCW4TC310
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 And BIT0
        Case TCW4TC311
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 And BIT1
        Case TCW4TC312
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 And BIT2
        Case TCW4TC313
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 And BIT3
        Case TCW4TC314
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 And BIT4
        Case TCW4TC315  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC31 < 0

        'TC Table   (TCW4TC15)
        Case TCW4TC150
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 And BIT0
        Case TCW4TC151
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 And BIT1
        Case TCW4TC152
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 And BIT2
        Case TCW4TC153
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 And BIT3
        Case TCW4TC154
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 And BIT4
        Case TCW4TC155  'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW4TC15 < 0

        'TC Table   (TCW2TC0)
        Case TCW2TC00
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 And BIT0
        Case TCW2TC01
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 And BIT1
        Case TCW2TC02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 And BIT2
        Case TCW2TC03
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 And BIT3
        Case TCW2TC04
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 And BIT4
        Case TCW2TC05   'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW2TC0 < 0

        'TC Table   (TCW3TC0) (TCW3TC01 - TCW3TC07 are in the LABEL section)
        Case TCW3TC08
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 And BIT8
        Case TCW3TC09   'Sign bit: Is the code less than 0?
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.TCW3TC0 < 0

        'Lot    (LOT0 is in the LABEL Section)
        Case LOT1
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT1
        Case LOT2
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT2
        Case LOT3
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT3
        Case LOT4
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT4
        Case LOT5
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT5
        Case LOT6
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT6
        Case LOT7
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT7
        Case LOT8
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT8
        Case LOT9
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT9
        Case LOT10
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT10
        Case LOT11
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT11
        Case LOT12
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT12
        Case LOT13
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT13
        Case LOT14
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT14
        Case LOT15
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT15
        Case LOT16
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Lot And BIT16

        'Wafer  (All three bits should be the same for the Wafer voting scheme)
        Case WFR00, WFR01, WFR02
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Wafer And BIT0
        Case WFR10, WFR11, WFR12
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Wafer And BIT1
        Case WFR20, WFR21, WFR22
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Wafer And BIT2
        Case WFR30, WFR31, WFR32
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Wafer And BIT3
        Case WFR40, WFR41, WFR42
            lblnBitValue = gudtMLX90277(ProgrammerNum).Write.Wafer And BIT4

    End Select

    'Set the ID value at this address
    gudtMLX90277(ProgrammerNum).IDContent(lintAddress) = lblnBitValue

Next lintAddress

'Fill The EEPROMContent array
For lintAddress = 0 To (NUMEEPROMLOCATIONS - 1)
    'If the LABEL content is true, then all 3 voting bits are set to one
    'If the LABEL content is false, then all 3 voting bits are set to zero
    gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, 0) = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)
    gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, 1) = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)
    gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, 2) = gudtMLX90277(ProgrammerNum).LABELContent(lintAddress)
    'Bit 3 = ID content
    gudtMLX90277(ProgrammerNum).EEPROMContent(lintAddress, 3) = gudtMLX90277(ProgrammerNum).IDContent(lintAddress)
Next lintAddress

End Sub

Private Function EncodeFloatString(ByVal number As Single) As String
'
'   PURPOSE: To create a 4-character string representing the floating point
'            passed in.
'  INPUT(S): number = floating point (single) number to be converted
' OUTPUT(S): None.

Dim lblnSign As Boolean
Dim lintExponent As Integer
Dim lsngMantissa As Single
Dim lstrSignString As String
Dim lstrExponentString As String
Dim lstrMantissaString As String
Dim lstr32BitString As String
Dim i As Integer
Dim j As Integer
Dim lstrByte(0 To 3) As String
Dim lintByte(0 To 3) As Integer

'IEEE Single Precision floating point numbers are 32-bit numbers stored as follows:
'33222222 22221111 11111100 00000000
'10987654 32109876 54321098 76543210
'SEEEEEEE EMMMMMMM MMMMMMMM MMMMMMMM
' Byte 3   Byte 2   Byte 1   Byte 0

'Where S represents the sign
'      E represents the exponent
'      M represents the mantissa

'Initialize variables
lintExponent = 0
lstrSignString = ""
lstrExponentString = ""
lstrMantissaString = ""

'Set the sign bit if the number is negative
If number < 0 Then
    lblnSign = True
    'Make sure we're dealing with positives from here on out
    number = Abs(number)
Else
    lblnSign = False
End If

'Decode the exponent
If number = 0 Then
    'If the number equals zero, offset the exponent by -127
    lintExponent = -127
ElseIf number > 2 Then
    'Numbers greater than 2 should be divided by two until between 1 and 2
    Do While number > 2
        'Increment the counter
        lintExponent = lintExponent + 1
        number = number / 2
    Loop
ElseIf number < 1 Then
    'Numbers less than 1 should be mulitplied by two until between 1 and 2
    Do While number < 1
        'Decrement the counter
        lintExponent = lintExponent - 1
        number = number * 2
    Loop
End If

'The mantissa is the remainder of the last operation
lsngMantissa = number
'The "1" is assumed unless the number  = 0
If number <> 0 Then lsngMantissa = lsngMantissa - 1
'The exponent is offset by -127
lintExponent = lintExponent + 127

'Generate the string associated with the sign
If lblnSign Then
    lstrSignString = "1"
Else
    lstrSignString = "0"
End If

'Generate the string associated with the exponent
For i = 0 To 7
    'Bitmask the exponent with the current power of two,
    'divide by the power of two to create a "1" or a "0",
    'then concatenate.  Note that string is created left to right
    '(bit 0 to bit 7).
    lstrExponentString = lstrExponentString & CStr((lintExponent And (2 ^ i)) / (2 ^ i)) '& lstrExponentString
Next i

'Generate the string associated with the mantissa
For i = 1 To 23
    'The mantissa represents the decimal portion of the stored number
    'Determine whether or not the mantissa is greater than the current
    'power of two.  If so, the mantissa contains that power of two, so
    'we add a "1" in that bit position and divide by that power of two.
    'If not, add a "0" in that bit position.  Note that the string is
    'created right to left (bit 0 to bit 23).
    If lsngMantissa >= 2 ^ (-i) Then
        lsngMantissa = lsngMantissa - 2 ^ (-i)
        lstrMantissaString = "1" & lstrMantissaString '& "1"
    Else
        lstrMantissaString = "0" & lstrMantissaString '& "0"
    End If
Next i

'Concatenate the sign, exponent, and mantissa
lstr32BitString = lstrMantissaString & lstrExponentString & lstrSignString

'Create the four bytes of the float from the 32-bit string
For i = 0 To 3
    'Initialize the variables
    lstrByte(i) = ""
    lintByte(i) = 0
    'Parse out the byte
    lstrByte(i) = Mid(lstr32BitString, (i * 8) + 1, 8)
    'Loop through each of the bits of the byte adding the power of two if called for
    For j = 0 To 7
        lintByte(i) = lintByte(i) + CInt(Mid(lstrByte(i), j + 1, 1) * 2 ^ j)
    Next j
Next i

'Return the four characters
EncodeFloatString = Chr$(lintByte(0)) & Chr$(lintByte(1)) & Chr$(lintByte(2)) & Chr$(lintByte(3))

End Function

Public Function EncodePartID() As String
'
'   PURPOSE: To build the Part ID number based on the MLX ID information
'            in the IC associated with output #1.
'
'  INPUT(S): None
' OUTPUT(S): returns the PartID (String)

Dim ldblPartID As Double

'Y-Location on Wafer (Bit0 - Bit6)
ldblPartID = CDbl(gudtMLX90277(1).Read.Y)
'X-Location on Wafer (Bit7 - Bit13)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.X) * BIT7)
'Wafer Number(Bit14 - Bit18)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.Wafer) * BIT14)
'Lot Number (Bit19 - Bit35)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(1).Read.Lot) * BIT19)

'Format as an eleven digit number and return
EncodePartID = Format(ldblPartID, "00000000000")

End Function

Public Function EncodePartID2() As String
'
'   PURPOSE: To build the Part ID number based on the MLX ID information
'            in the IC associated with output #1.
'
'  INPUT(S): None
' OUTPUT(S): returns the PartID (String)
'V1.3 new sub

Dim ldblPartID As Double

'Y-Location on Wafer (Bit0 - Bit6)
ldblPartID = CDbl(gudtMLX90277(2).Read.Y)
'X-Location on Wafer (Bit7 - Bit13)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(2).Read.X) * BIT7)
'Wafer Number(Bit14 - Bit18)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(2).Read.Wafer) * BIT14)
'Lot Number (Bit19 - Bit35)
ldblPartID = ldblPartID + (CDbl(gudtMLX90277(2).Read.Lot) * BIT19)

'Format as an eleven digit number and return
EncodePartID2 = Format(ldblPartID, "00000000000")

End Function

Private Function EncodeInteger(number As Integer) As String
'
'   PURPOSE: To calculate a two-byte string representing an integer
'
'  INPUT(S): number = integer to be converted
'
' OUTPUT(S): None
'

Dim lstrLowByte As String
Dim lstrHighByte As String

'Mask off the lower byte and convert to ASCII code
lstrLowByte = Chr$(number And &HFF)
'Mask off the upper byte, bitshift, and convert to ASCII code
lstrHighByte = Chr$((number And &HFF00) / &H100)

'Concatenate the Low Byte and the High Byte
EncodeInteger = lstrLowByte & lstrHighByte

End Function

Public Function EraseEEPROMRow(RowNum As Integer) As Boolean
'
'   PURPOSE: To send the data stream command to erase one row of the EEprom.
'            Performs action for both programmers.
'
'  INPUT(S): RowNum = identifies the row location in the EEprom.
' OUTPUT(S): None

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

On Error GoTo EraseEEPROMRowError

'Initialize the return value of the routine
EraseEEPROMRow = False

'Build the command to erase the EEPROM Rows
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptc90251Program) & BuildCommand(mftRowErase, gudtMLX90277(lintProgrammerNum).EEPROMContent(), (RowNum * 8))
Next lintProgrammerNum
EraseEEPROMRow = SendCommandGetResponse(lstrWrite(), lstrResponse())

Exit Function
EraseEEPROMRowError:
    gblnGoodPTC04Link = False

End Function

Public Sub EstablishCommunication()
'
'   PURPOSE: To establish Serial communications between the PC and
'            the Melexis PTC-04 programmers.  If communication cannot be
'            established, display communication error.
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer

On Error GoTo NotAbleToEstablishMelexisLink

gblnGoodPTC04Link = False    'Initialize the link variable to false (bad link)

For lintProgrammerNum = 1 To 2
    'Be sure that the COM port is closed
    If frmMLX90277.SerialPort(lintProgrammerNum).PortOpen = True Then
        frmMLX90277.SerialPort(lintProgrammerNum).PortOpen = False
    End If
    'Set COM Port #
    frmMLX90277.SerialPort(lintProgrammerNum).CommPort = gudtPTC04(lintProgrammerNum).CommPortNum
    'Set baud rate, parity, data bits, stop bits
    frmMLX90277.SerialPort(lintProgrammerNum).Settings = "115200,N,8,1"
    'Set the flow control
    frmMLX90277.SerialPort(lintProgrammerNum).Handshaking = comNone
    'Set to read all characters
    frmMLX90277.SerialPort(lintProgrammerNum).InputLen = 0
    'Open the COM Port
    frmMLX90277.SerialPort(lintProgrammerNum).PortOpen = True
Next lintProgrammerNum

'Set the link variable to ok if we are able to Initialize the Programmer
gblnGoodPTC04Link = InitializeProgrammer

If gblnGoodPTC04Link Then Exit Sub

NotAbleToEstablishMelexisLink:
    gblnGoodPTC04Link = False
    MsgBox "There was an error trying to establish communications" _
           & vbCrLf & "with a Melexis programmer", _
           vbOKOnly + vbCritical, "Melexis Communication Error: " & Err.Description
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

    '*** Setup voltage levels ***
    'Set Channel 0 (Vdd) Ground Voltage to 0V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 0
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddGnd, lsngVoltageLevel()) Then Exit Function
    'Set Channel 0 (Vdd) Nominal Voltage to 5V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 5
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddNom, lsngVoltageLevel()) Then Exit Function
    'Set Channel 0 (Vdd) Programming Voltage to 9V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 9
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Vdd, ptcVddProg, lsngVoltageLevel()) Then Exit Function
    'Set Channel 1 (Vout) Low Level Voltage to 0V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 0
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutLow, lsngVoltageLevel()) Then Exit Function
    'Set Channel 1 (Vout) Mid Level Voltage to 2.5V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 2.5
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutMid, lsngVoltageLevel()) Then Exit Function
    'Set Channel 1 (Vout) High Level Voltage to 5V
    For lintProgrammerNum = 1 To 2
        lsngVoltageLevel(lintProgrammerNum) = 5
    Next lintProgrammerNum
    If Not SetupVoltageLevel(ptcPPS_Out, ptcOutHigh, lsngVoltageLevel()) Then Exit Function

    '*** Setup Timings ***
    'Set Tpor
    For lintProgrammerNum = 1 To 2
        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tpor
    Next lintProgrammerNum
    If Not SetupTiming(ptcTpor, lintTimeInMicroSeconds()) Then Exit Function
    'Set Thold
    For lintProgrammerNum = 1 To 2
        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Thold
    Next lintProgrammerNum
    If Not SetupTiming(ptcThold, lintTimeInMicroSeconds()) Then Exit Function
    'Set Tprog
    For lintProgrammerNum = 1 To 2
        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tprog
    Next lintProgrammerNum
    If Not SetupTiming(ptcTprog, lintTimeInMicroSeconds()) Then Exit Function
    'Set Tpuls
    For lintProgrammerNum = 1 To 2
        lintTimeInMicroSeconds(lintProgrammerNum) = gudtPTC04(lintProgrammerNum).Tpuls
    Next lintProgrammerNum
    If Not SetupTiming(ptcTpuls, lintTimeInMicroSeconds()) Then Exit Function

    '*** Set Measurement Delays & Filtering ***
    'Set Measurement Delay to 5000 uS
    For lintProgrammerNum = 1 To 2
        lintNumber(lintProgrammerNum) = 5000
    Next lintProgrammerNum
    If Not SetupDelayOrFilter(ptcSetMeasureDelay, lintNumber()) Then Exit Function
    'Set Sample Delay to 1 uS
    For lintProgrammerNum = 1 To 2
        lintNumber(lintProgrammerNum) = 1
    Next lintProgrammerNum
    If Not SetupDelayOrFilter(ptcSetSampleDelay, lintNumber()) Then Exit Function
    'Set Measurement Filter to 1 measurement
    For lintProgrammerNum = 1 To 2
        lintNumber(lintProgrammerNum) = 1
    Next lintProgrammerNum
    If Not SetupDelayOrFilter(ptcSetMeasureFilter, lintNumber()) Then Exit Function

    '*** Set the Current Limit ***
    'Set Current Limit for Channel 4 (Vdd) to 200 mA
    For lintProgrammerNum = 1 To 2
        lsngCurrentLimit(lintProgrammerNum) = 200
    Next lintProgrammerNum
    If Not SetupCurrentLimit(ptcPPS_Vdd_I_Limit, lsngCurrentLimit()) Then Exit Function
    'Set Current Limit for Channel 5 (Output) to 200 mA
    For lintProgrammerNum = 1 To 2
        lsngCurrentLimit(lintProgrammerNum) = 200
    Next lintProgrammerNum
    If Not SetupCurrentLimit(ptcPPS_Out_I_Limit, lsngCurrentLimit()) Then Exit Function
End If

'If everything went ok and we haven't left the Function yet, Initialization was successful
InitializeProgrammer = gblnGoodPTC04Link

End Function

Public Function ReadEEPROM(RevLevel As String, VotingError As Boolean) As Boolean
'
'   PURPOSE: To send the Melexis command to read back all 128 bits of
'            each EEPROM.
'
'  INPUT(S): RevLevel = Revision Level of the Melexis 90277 IC  'V1.2
' OUTPUT(S): VotingError = Whether or not there was a voting error in one of the EEPROMs

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String
Dim lstrEEPROMRead(1 To 2) As String
Dim lblnVotingError(1 To 2) As Boolean

On Error GoTo ReadEEPROMError

'Initialize the return value of the routine
ReadEEPROM = False

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Zero the read variables
    Call MLX90277.ClearReadVariables(lintProgrammerNum)
    'Readback is different for the different revision levels of the Melexis IC
    If RevLevel = "Cx" Then     'V1.2
        'Request a readback from address 0 to 128 with a readback filter of 10
        'Command format:
        '|ReadBackCommand|StartAddress|StopAddress|RBFilterLowByte|RBFilterHighByte|
        lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptc90251CxReadBack) & Chr$(0) & Chr$(NUMEEPROMLOCATIONS) & EncodeInteger(10)
    ElseIf RevLevel = "FA" Then 'V1.2
        'Request a readback from address 0 to 128 with a readback filter of 10
        'Command format:
        '|ReadBackCommand|StartAddress|StopAddress|RBFilterLowByte|RBFilterHighByte|
        lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptc90251FAReadBack) & Chr$(0) & Chr$(NUMEEPROMLOCATIONS) & EncodeInteger(10)
    End If
Next lintProgrammerNum

If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then

    'Loop through both programmers
    For lintProgrammerNum = 1 To 2
        'Cut off the command from the response
        lstrEEPROMRead(lintProgrammerNum) = right(lstrResponse(lintProgrammerNum), Len(lstrResponse(lintProgrammerNum)) - 1)
        'Convert EEprom bits to EEprom variables
        Call DecodeEEpromRead(lintProgrammerNum, lstrEEPROMRead(lintProgrammerNum), lblnVotingError(lintProgrammerNum))
    Next lintProgrammerNum

    'The voting error returns whether or not there was a voting error on either chip
    VotingError = lblnVotingError(1) Or lblnVotingError(2)
    'If we made it this far, the routine executed successfully
    ReadEEPROM = True
End If

Exit Function
ReadEEPROMError:
    gblnGoodPTC04Link = False
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
        frmMLX90277.SerialPort(lintProgrammerNum).InBufferCount = 0
        'Write the data to the programmer
        frmMLX90277.SerialPort(lintProgrammerNum).Output = lstrWriteData(lintProgrammerNum)
        'Loop until the message is sent out
        Do Until frmMLX90277.SerialPort(lintProgrammerNum).OutBufferCount = 0
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
            lstrReadData(lintProgrammerNum) = lstrReadData(lintProgrammerNum) & frmMLX90277.SerialPort(lintProgrammerNum).Input
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

Private Function SetupCurrentLimit(ByVal channel As PTC04SetupType, currentInMilliAmps() As Single) As Boolean
'
'   PURPOSE: To send a setup command and verify that it was received properly.
'            by the PTC-04 Programmer.
'
'  INPUT(S): channel              = The PTC-04 channel to setup
'            currentInMilliAmps() = The current in mA to set the current limit to
'
' OUTPUT(S): Function returns whether or not the command was sent
'            a the proper response was received.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

SetupCurrentLimit = False

'Create the string to write to the programmer
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(channel) & EncodeFloatString(currentInMilliAmps(lintProgrammerNum))
Next lintProgrammerNum

'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    'The responses from the programmers should be the command
    SetupCurrentLimit = ((lstrResponse(1) = Chr$(PTC04CommandType.ptcSetPPS)) And (lstrResponse(2) = Chr$(PTC04CommandType.ptcSetPPS)))
End If

End Function

Public Sub GetCurrent()
'
'   PURPOSE: To get the current from the PTC-04 Programmers.
'
'  INPUT(S): None
' OUTPUT(S): None
'
'3.3ANM new sub
'3.6cANM updated for MLX VI

Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String
Dim MLXbyte(3) As Byte
Dim MLXbyte2(3) As Byte
Dim PPSStr As String
Dim PPSStr2 As String
Dim lsng5V As Single
Dim lsng5V2 As Single
Dim X As Integer
Dim lintF1 As Integer
Dim lintF2 As Integer
Dim lsngF As Single

'Set float then get bytes from float
If gblnMLXVI Then
    lsng5V = CSng(frmMLXVI.txtS1.Text)
    lsng5V2 = CSng(frmMLXVI.txtS2.Text)
Else
    lsng5V = 5
    lsng5V2 = 5
End If

'Set vals
lsngF = 1000
X = 0 'Mem issue
lintF2 = (lsngF \ (2 ^ 8)) And &HFF
X = 0 'Mem issue
lintF1 = (lsngF - (lintF2 * (2 ^ 8)))

'Convert bytes to string
PPSStr = Chr$(lintF1) & Chr$(lintF2)
PPSStr2 = Chr$(lintF1) & Chr$(lintF2)

'Create the string to write to the programmer
lstrWrite(1) = Chr$(PTC04CommandType.ptcSetMeasureFilter) & PPSStr
lstrWrite(2) = Chr$(PTC04CommandType.ptcSetMeasureFilter) & PPSStr2
    
'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    CopyMemory MLXbyte(0), lsng5V, 4
    CopyMemory MLXbyte2(0), lsng5V2, 4
    
    'Convert bytes to string
    PPSStr = Chr$(MLXbyte(0)) & Chr$(MLXbyte(1)) & Chr$(MLXbyte(2)) & Chr$(MLXbyte(3))
    PPSStr2 = Chr$(MLXbyte2(0)) & Chr$(MLXbyte2(1)) & Chr$(MLXbyte2(2)) & Chr$(MLXbyte2(3))
    
    'Create the string to write to the programmer
    lstrWrite(1) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(0) & PPSStr
    lstrWrite(2) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(0) & PPSStr2
    
    'Send the string
    If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
        'Create the string to write to the programmer
        lstrWrite(1) = Chr$(PTC04CommandType.ptcSetRelay) & Chr$(1) & Chr$(1)
        lstrWrite(2) = Chr$(PTC04CommandType.ptcSetRelay) & Chr$(1) & Chr$(1)
        
        'Send the string
        If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
            'Create the string to write to the programmer
            lstrWrite(1) = Chr$(PTC04CommandType.ptcGetCurrent) & Chr$(1)
            lstrWrite(2) = Chr$(PTC04CommandType.ptcGetCurrent) & Chr$(1)
            
            'Send the string
            If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
                For X = 1 To 2
                    'The responses from the programmers should be the current
                    MLXbyte(0) = Asc(Mid$(lstrResponse(X), 2, 1))
                    MLXbyte(1) = Asc(Mid$(lstrResponse(X), 3, 1))
                    MLXbyte(2) = Asc(Mid$(lstrResponse(X), 4, 1))
                    MLXbyte(3) = Asc(Mid$(lstrResponse(X), 5, 1))
                    
                    CopyMemory gudtReading(X - 1).mlxCurrent, MLXbyte(0), 4
                Next X
            End If
            
            If gblnMLXVI Then
                'Create the string to write to the programmer
                lstrWrite(1) = Chr$(PTC04CommandType.ptcGetLevel) & Chr$(13)
                lstrWrite(2) = Chr$(PTC04CommandType.ptcGetLevel) & Chr$(13)
                
                'Send the string
                If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
                    For X = 1 To 2
                        'The responses from the programmers should be the supply
                        MLXbyte(0) = Asc(Mid$(lstrResponse(X), 2, 1))
                        MLXbyte(1) = Asc(Mid$(lstrResponse(X), 3, 1))
                        MLXbyte(2) = Asc(Mid$(lstrResponse(X), 4, 1))
                        MLXbyte(3) = Asc(Mid$(lstrResponse(X), 5, 1))
                        
                        CopyMemory gudtReading(X - 1).mlxSupply, MLXbyte(0), 4
                    Next X
                End If
            End If
        End If
    End If
End If

End Sub

Public Sub GetCurrentW()
'
'   PURPOSE: To get the current from the PTC-04 Programmers.
'
'  INPUT(S): None
' OUTPUT(S): None
'
'3.6fANM new sub

Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String
Dim MLXbyte(3) As Byte
Dim MLXbyte2(3) As Byte
Dim PPSStr As String
Dim PPSStr2 As String
Dim lsng5V As Single
Dim lsng5V2 As Single
Dim X As Integer
Dim lintF1 As Integer
Dim lintF2 As Integer
Dim lsngF As Single

'Set float then get bytes from float
If gblnMLXVI Then
    lsng5V = CSng(frmMLXVI.txtS1.Text)
    lsng5V2 = CSng(frmMLXVI.txtS2.Text)
Else
    lsng5V = 5
    lsng5V2 = 5
End If

'Set vals
lsngF = 1000
X = 0 'Mem issue
lintF2 = (lsngF \ (2 ^ 8)) And &HFF
X = 0 'Mem issue
lintF1 = (lsngF - (lintF2 * (2 ^ 8)))

'Convert bytes to string
PPSStr = Chr$(lintF1) & Chr$(lintF2)
PPSStr2 = Chr$(lintF1) & Chr$(lintF2)

'Create the string to write to the programmer
lstrWrite(1) = Chr$(PTC04CommandType.ptcSetMeasureFilter) & PPSStr
lstrWrite(2) = Chr$(PTC04CommandType.ptcSetMeasureFilter) & PPSStr2
    
'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    CopyMemory MLXbyte(0), lsng5V, 4
    CopyMemory MLXbyte2(0), lsng5V2, 4
    
    'Convert bytes to string
    PPSStr = Chr$(MLXbyte(0)) & Chr$(MLXbyte(1)) & Chr$(MLXbyte(2)) & Chr$(MLXbyte(3))
    PPSStr2 = Chr$(MLXbyte2(0)) & Chr$(MLXbyte2(1)) & Chr$(MLXbyte2(2)) & Chr$(MLXbyte2(3))
    
    'Create the string to write to the programmer
    lstrWrite(1) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(0) & PPSStr
    lstrWrite(2) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(0) & PPSStr2
    
    'Send the string
    If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
        'Create the string to write to the programmer
        lstrWrite(1) = Chr$(PTC04CommandType.ptcSetRelay) & Chr$(1) & Chr$(1)
        lstrWrite(2) = Chr$(PTC04CommandType.ptcSetRelay) & Chr$(1) & Chr$(1)
        
        'Send the string
        If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
            'Create the string to write to the programmer
            lstrWrite(1) = Chr$(PTC04CommandType.ptcGetCurrent) & Chr$(1)
            lstrWrite(2) = Chr$(PTC04CommandType.ptcGetCurrent) & Chr$(1)
            
            'Send the string
            If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
                For X = 1 To 2
                    'The responses from the programmers should be the current
                    MLXbyte(0) = Asc(Mid$(lstrResponse(X), 2, 1))
                    MLXbyte(1) = Asc(Mid$(lstrResponse(X), 3, 1))
                    MLXbyte(2) = Asc(Mid$(lstrResponse(X), 4, 1))
                    MLXbyte(3) = Asc(Mid$(lstrResponse(X), 5, 1))
                    
                    CopyMemory gudtReading(X - 1).mlxWCurrent, MLXbyte(0), 4
                Next X
            End If
            
            If gblnMLXVI Then
                'Create the string to write to the programmer
                lstrWrite(1) = Chr$(PTC04CommandType.ptcGetLevel) & Chr$(13)
                lstrWrite(2) = Chr$(PTC04CommandType.ptcGetLevel) & Chr$(13)
                
                'Send the string
                If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
                    For X = 1 To 2
                        'The responses from the programmers should be the supply
                        MLXbyte(0) = Asc(Mid$(lstrResponse(X), 2, 1))
                        MLXbyte(1) = Asc(Mid$(lstrResponse(X), 3, 1))
                        MLXbyte(2) = Asc(Mid$(lstrResponse(X), 4, 1))
                        MLXbyte(3) = Asc(Mid$(lstrResponse(X), 5, 1))
                        
                        CopyMemory gudtReading(X - 1).mlxSupply, MLXbyte(0), 4
                    Next X
                End If
            End If
        End If
    End If
End If

End Sub

Private Function SetupDelayOrFilter(command As PTC04CommandType, number() As Integer) As Boolean
'
'   PURPOSE: To send a setup command and verify that it was received properly.
'            by the PTC-04 Programmers.
'
'  INPUT(S): command = The command to execute
'            number  = The number to be sent to the programmer
'
' OUTPUT(S): Function returns whether or not the command was sent
'            a the proper response was received.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

SetupDelayOrFilter = False

For lintProgrammerNum = 1 To 2
    'Create the string to write to the programmer
    lstrWrite(lintProgrammerNum) = Chr$(command) & EncodeInteger(number(lintProgrammerNum))
Next lintProgrammerNum

'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    'The response from the programmer should be the command
    SetupDelayOrFilter = ((lstrResponse(1) = Chr$(command)) And (lstrResponse(2) = Chr$(command)))
End If

End Function

Private Function SetupTiming(ByVal timingNumber As PTC04SetupType, timeInMicroSeconds() As Integer) As Boolean
'
'   PURPOSE: To send a setup command and verify that it was received properly.
'            by each PTC-04 Programmer.
'
'  INPUT(S): timingNumber         = The timing to set up
'            timeInMicroSeconds() = The number of microseconds to set the time to
'
' OUTPUT(S): Function returns whether or not the command was sent
'            and the proper response was received.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

SetupTiming = False

'Create the string to write to the programmer
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcSetTiming) & Chr$(timingNumber) & EncodeInteger(timeInMicroSeconds(lintProgrammerNum))
Next lintProgrammerNum

'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    'The responses from the programmers should be the command
    SetupTiming = ((lstrResponse(1) = Chr$(PTC04CommandType.ptcSetTiming)) And (lstrResponse(2) = Chr$(PTC04CommandType.ptcSetTiming)))
End If

End Function

Public Function SetupVoltageLevel(ByVal channel As PTC04SetupType, ByVal levelNumber As PTC04SetupType, voltageLevel() As Single) As Boolean
'
'   PURPOSE: To send a setup command and verify that it was received properly.
'            by the PTC-04 Programmer.
'
'  INPUT(S): channel        = The PTC-04 channel to setup
'            levelNumber    = the level to set up on the selected channel
'            voltageLevel() = The voltage level to set the selected level to
'
' OUTPUT(S): Function returns whether or not the command was sent
'            a the proper response was received.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

SetupVoltageLevel = False

'Create the string to write to the programmer
For lintProgrammerNum = 1 To 2
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptcSetLevel) & Chr$(channel) & Chr$(levelNumber) & EncodeFloatString(voltageLevel(lintProgrammerNum))
Next lintProgrammerNum

'Send the string
If SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
    'The responses from the programmers should be the command
    SetupVoltageLevel = ((lstrResponse(1) = Chr$(PTC04CommandType.ptcSetLevel)) And (lstrResponse(2) = Chr$(PTC04CommandType.ptcSetLevel)))
End If

End Function

Public Function VerifyCustomerCRC(ProgrammerNum As Integer) As Boolean
'
'   PURPOSE: To verify the (Customer) CRC value.  If this value is not correct,
'             the TC Table is not good, and the part should be rejected.
'
'  INPUT(S): programmerNum = which programmer (which IC)
' OUTPUT(S): Returns (Customer) CRC value based on IC contents

Dim lintDataCRC(0 To 9) As Integer
Dim lblnCalcOk As Boolean

VerifyCustomerCRC = False    'Initialize the value of the routine

'Get the first six bits of TCW6TC0
lintDataCRC(0) = gudtMLX90277(ProgrammerNum).Read.TCW6TC0 And &H3F
'Get the first six bits of TCW5TC0
lintDataCRC(1) = (gudtMLX90277(ProgrammerNum).Read.TCW5TC0 And &H3F)
'Get the last six bits of TCW3TC31
lintDataCRC(2) = (gudtMLX90277(ProgrammerNum).Read.TCW3TC31 And &H3F)
'Get the first six bits of TCW2TC31
lintDataCRC(3) = (gudtMLX90277(ProgrammerNum).Read.TCW2TC31 And &H3F)
'Get the last four bits of TCW1TC31
lintDataCRC(4) = (gudtMLX90277(ProgrammerNum).Read.TCW1TC31 And &H3F)
'Get the first six  bits of TCW4TC0
lintDataCRC(5) = (gudtMLX90277(ProgrammerNum).Read.TCW4TC0 And &H3F)
'Get the first six  bits of HANDLING
lintDataCRC(6) = (gudtMLX90277(ProgrammerNum).Read.Handling And &H7)
'Fill with zeros
lintDataCRC(7) = 0
'Fill with zeros
lintDataCRC(8) = 0
'Fill with zeros
lintDataCRC(9) = 0

If gudtMLX90277(ProgrammerNum).Read.CRC = CalculateDataCRC(lintDataCRC(), 8, lblnCalcOk) Then
    VerifyCustomerCRC = lblnCalcOk
Else
    VerifyCustomerCRC = False
End If

End Function

Public Function VerifyMLXCRC(ProgrammerNum As Integer) As Boolean
'
'   PURPOSE: To verify the MLXCRC value.  If this value is not correct, the
'            TC Table is not good, and the part should be rejected.
'
'  INPUT(S): programmerNum = which programmer (which IC)
' OUTPUT(S): Returns MLXCRC value based on IC contents
  
Dim lintDataCRC(0 To 9) As Integer
Dim lblnCalcOk As Boolean

VerifyMLXCRC = False    'Initialize the value of the routine

'Get the first six bits of the Lot
lintDataCRC(0) = gudtMLX90277(ProgrammerNum).Read.Lot And &H3F
'Get the second six bits of the Lot and bitshift 6 times
lintDataCRC(1) = (gudtMLX90277(ProgrammerNum).Read.Lot And &HFC0) / (2 ^ 6)
'Get the last six bits of the Lot and bitshift 12 times
lintDataCRC(2) = (gudtMLX90277(ProgrammerNum).Read.Lot And &H1F000) / (2 ^ 12)
'Get the first six bits of TCW3TC0
lintDataCRC(3) = (gudtMLX90277(ProgrammerNum).Read.TCW3TC0 And &H3F)
'Get the last four bits of TCW3TC0 and bitshift 6 times
lintDataCRC(4) = (gudtMLX90277(ProgrammerNum).Read.TCW3TC0 And &H3C0) / (2 ^ 6)
'Get the first six  bits of TCW2TC0
lintDataCRC(5) = (gudtMLX90277(ProgrammerNum).Read.TCW2TC0 And &H3F)
'Get the first six  bits of TCW4TC15
lintDataCRC(6) = (gudtMLX90277(ProgrammerNum).Read.TCW4TC15 And &H3F)
'Get the first six  bits of TCW4TC31
lintDataCRC(7) = (gudtMLX90277(ProgrammerNum).Read.TCW4TC31 And &H3F)
'Get the first six  bits of TCW5TC31
lintDataCRC(8) = (gudtMLX90277(ProgrammerNum).Read.TCW5TC31 And &H3F)
'Fill with zeros
lintDataCRC(9) = 0

If gudtMLX90277(ProgrammerNum).Read.MLXCRC = CalculateDataCRC(lintDataCRC(), 10, lblnCalcOk) Then
    VerifyMLXCRC = lblnCalcOk
Else
    VerifyMLXCRC = False
End If

End Function

Private Function VoteBit(BooleanArray() As Boolean, Address As Integer, VoteError As Boolean) As Boolean
'
'   PURPOSE: To Vote a logical bit based on a boolean array of bits and
'            the Melexis voting algorithm.
'
'  INPUT(S): BooleanArray = Array of Boolean data to be evaluated
'            Address      = Address within the array to evaluate
' OUTPUT(S): VoteError    = Whether or not a voting error occurred
'            Function returns the voted bit

'Voting Algorithm:
'The bits D2, D1, and D0 "vote" the value of the LABEL section bit
'The voting algorithm is:
'       If (D0 = D1)
'           Vote D0     (The bit takes on the value of D0)
'       Else
'           Vote D2     (The bit takes on the value of D2)
'       End If
'
'NOTE: The intention of the voting bits is for D2=D1=D0.  This provides
'      maximum data redundancy.  If any of these three does not match
'      the other two, a voting error has occured, meaning that an EEPROM
'      cell is likely bad.

'If D0 = D1, Vote D0
If BooleanArray(Address, 0) = BooleanArray(Address, 1) Then

    'Return D0
    VoteBit = BooleanArray(Address, 0)

    'Check to see if D1 = D2, if not, set the VotingError variable to True
    If BooleanArray(Address, 1) <> BooleanArray(Address, 2) Then VoteError = True

'Else, Vote D2
Else
    'Return D2
    VoteBit = BooleanArray(Address, 2)

    'Set the VotingError variable to True
    VoteError = True

End If

End Function

Public Function WriteEEPROMBlockByRows(startRow As Integer, stopRow As Integer) As Boolean
'
'   PURPOSE: First, send the command to erase the EEprom and then write
'            data to the EEPROM.
'
'  INPUT(S): startRow = Row to start writing to EEPROM
'            stopRow  = Row to stop writing to EEPROM
' OUTPUT(S): None.

Dim lintAddress As Integer
Dim lintRowNum As Integer

On Error GoTo WriteEEPROMBlockByRowsError

'Initialize the Function to return False, or unsuccessful completion
WriteEEPROMBlockByRows = False

'Make sure that we don't try to erase anything beyond row seven
If (startRow > stopRow) Or (stopRow > 7) Or (startRow < 0) Then
    Exit Function
End If

'Clear the contents of the EEPROM, exit if unsuccessful
For lintRowNum = startRow To stopRow
    If Not EraseEEPROMRow(lintRowNum) Then Exit Function
Next lintRowNum

'Loop though all the appropriate addresses, starting at the highest address
'and working down towards address 0 (MemLock)
For lintAddress = ((stopRow * 8) + 7) To (startRow * 8) Step -1
    'Exit the function (returning false or unsuccessful) if the Write does not work
    If Not WriteEEpromWord(lintAddress) Then
        Exit Function
    End If
Next lintAddress

'Everything completed successfully
WriteEEPROMBlockByRows = True

Exit Function
WriteEEPROMBlockByRowsError:
    gblnGoodPTC04Link = False
End Function

Public Function WriteEEpromWord(Address As Integer) As Boolean
'
'   PURPOSE: To send the data stream command to write data to an address
'            of the EEprom.
'
'  INPUT(S): Address = address of EEprom where data is written.
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String
Dim lblnWriteToEEPROM(1 To 2) As Boolean

On Error GoTo WriteEEpromWordError

'Initialize the return value of the routine
WriteEEpromWord = False

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Only write to the address if there are any desired "1"s at that address
    If gudtMLX90277(lintProgrammerNum).EEPROMContent(Address, 0) Or gudtMLX90277(lintProgrammerNum).EEPROMContent(Address, 1) Or gudtMLX90277(lintProgrammerNum).EEPROMContent(Address, 2) Or gudtMLX90277(lintProgrammerNum).EEPROMContent(Address, 3) Then
        'Build the command to Program the EEPROM Word
        lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptc90251Program) & BuildCommand(mftWordWrite, gudtMLX90277(lintProgrammerNum).EEPROMContent(), Address)
    Else
        'An empty string skips the write for that programmernumber
        lstrWrite(lintProgrammerNum) = ""
    End If
Next lintProgrammerNum

'Only write if there is data to write to at least one programmer
If lstrWrite(1) = "" And lstrWrite(2) = "" Then
    'Nothing to write, no problems
    WriteEEpromWord = True
Else
    'Return how the write went
    WriteEEpromWord = SendCommandGetResponse(lstrWrite(), lstrResponse())
End If

Exit Function
WriteEEpromWordError:
    gblnGoodPTC04Link = False
End Function

Public Function WriteTempRAM() As Boolean
'
'   PURPOSE: To send the data stream command to write data to Temp RAM
'
'  INPUT(S): None.
'
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer
Dim lstrWrite(1 To 2) As String
Dim lstrResponse(1 To 2) As String

On Error GoTo WriteTempRAMError

'Initialize the return value of the routine
WriteTempRAM = False

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Build the command to write to Temp RAM
    lstrWrite(lintProgrammerNum) = Chr$(PTC04CommandType.ptc90251MeasureByRAM) & Chr$(0) & BuildCommand(mftTemporary, gudtMLX90277(lintProgrammerNum).EEPROMContent(), 0)
Next lintProgrammerNum

WriteTempRAM = SendCommandGetResponse(lstrWrite(), lstrResponse())

Exit Function
WriteTempRAMError:
    gblnGoodPTC04Link = False
End Function
