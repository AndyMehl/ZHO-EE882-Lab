Attribute VB_Name = "Pedal"
'**********************************Pedal.BAS**********************************
'
'   Pedal Programming & Scanning Software, supplemented by SeriesXXX.Bas.
'   This module should handle all 700/702/703/705 series production prog/
'   scanners, test lab programmer/scanners, and database recall software.
'   The software is to be kept in the pedal software library, EE947.
'
'VER    DATE      BY   PURPOSE OF MODIFICATION                          TAG
'1.0  08/18/2005  SRC  First release per SCN 700F-001.                  1.0SRC
'1.1  11/09/2005  ANM  Updates per SCN# 700F-009 (3096) and             1.1ANM
'                      SCN# 702FP-030 (3098).
'1.2  12/16/2005  ANM  Updated to not find pedal face twice when scan   1.2ANM
'                      and prog, and added delay before reverse scan
'                      per SCN# 702FPSC-003 (3203)
'1.3  02/06/2006  ANM  Updated KD algorithm per SCN# 702FP-037 (3310).  1.3ANM
'                      Also added pallet load station
'1.4  03/01/2006  ANM  Update for new FPT per SCN# 700F-018 (3350).     1.4ANM
'1.5  03/13/2006  ANM  Update for rework per PR 14229-A and SCN#s       1.5ANM
'                      700F-016(3337) and 702FP-038 (3333).
'1.6  03/21/2006  ANM  Update for SCN# MISC-092 (3365).                 1.6ANM
'1.7  05/04/2006  ANM  Update for SCN# MISC-094 (3423).                 1.7ANM
'1.8  05/18/2006  ANM  Fix for good code on prog rejects.               1.8ANM
'1.9  08/29/2006  ANM  Updates SmartKD and per SCN# MISC-100 (3521).    1.9ANM
'2.0  12/04/2006  ANM  Updates per SCN# MISC-101 (3636).                2.0ANM
'2.1  03/08/2007  ANM  Updates for TR 8501-E (705 Prod.) and            2.1ANM
'                      SCN# MISC-102 (3702).
'2.2  03/15/2007  ANM  Updates per SCN# MISC-104 (3789).                2.2ANM
'2.3  05/04/2007  ANM  Updates per SCN# 705F-003 (3858).                2.3ANM
'2.4  05/17/2007  ANM  Updates per SCN# 705T-010 (3876).                2.4ANM
'2.5  06/12/2007  ANM  Updates per SCN# 702FP-042 (3886).               2.5ANM
'2.6  08/07/2007  ANM  Added ability to use new AMAD.                   2.6ANM
'2.7  09/24/2007  ANM  Updates for UDB.                                 2.7ANM
'2.8  02/18/2008  ANM  Updates for SCN#s 4066 & 4067. Moved shiftcalc.  2.8ANM
'2.9  03/31/2008  ANM  Update to add MLX current check per SCN# 4124.   2.9ANM
'3.0  05/02/2008  ANM  Add force only raw data save per SCN# 4139.      3.0ANM
'3.0a 11/10/2009  ANM  MLX Checks / lock fix / SS.                      3.0aANM
'3.2  06/21/2010  ANM  Update MLX Idds per SCN# 4585.                   3.2ANM
'3.2a 01/30/2012  ANM  Update filter 0 per SCN# 4933.                   3.2aANM
'

Option Explicit

'*****************
'*   Constants   *
'*****************

'Bit Masks
Public Const BIT0 = &H1                 'Bit mask (xxxx xxxx xxxx xxxx xxxx xxx1)
Public Const BIT1 = &H2                 'Bit mask (xxxx xxxx xxxx xxxx xxxx xx1x)
Public Const BIT2 = &H4                 'Bit mask (xxxx xxxx xxxx xxxx xxxx x1xx)
Public Const BIT3 = &H8                 'Bit mask (xxxx xxxx xxxx xxxx xxxx 1xxx)
Public Const BIT4 = &H10                'Bit mask (xxxx xxxx xxxx xxxx xxx1 xxxx)
Public Const BIT5 = &H20                'Bit mask (xxxx xxxx xxxx xxxx xx1x xxxx)
Public Const BIT6 = &H40                'Bit mask (xxxx xxxx xxxx xxxx x1xx xxxx)
Public Const BIT7 = &H80                'Bit mask (xxxx xxxx xxxx xxxx 1xxx xxxx)
Public Const BIT8 = &H100               'Bit mask (xxxx xxxx xxxx xxx1 xxxx xxxx)
Public Const BIT9 = &H200               'Bit mask (xxxx xxxx xxxx xx1x xxxx xxxx)
Public Const BIT10 = &H400              'Bit mask (xxxx xxxx xxxx x1xx xxxx xxxx)
Public Const BIT11 = &H800              'Bit mask (xxxx xxxx xxxx 1xxx xxxx xxxx)
Public Const BIT12 = &H1000             'Bit mask (xxxx xxxx xxx1 xxxx xxxx xxxx)
Public Const BIT13 = &H2000             'Bit mask (xxxx xxxx xx1x xxxx xxxx xxxx)
Public Const BIT14 = &H4000             'Bit mask (xxxx xxxx x1xx xxxx xxxx xxxx)
Public Const BIT15 = &H8000&            'Bit mask (xxxx xxxx 1xxx xxxx xxxx xxxx)
Public Const BIT16 = &H10000            'Bit mask (xxxx xxx1 xxxx xxxx xxxx xxxx)
Public Const BIT17 = &H20000            'Bit mask (xxxx xx1x xxxx xxxx xxxx xxxx)
Public Const BIT18 = &H40000            'Bit mask (xxxx x1xx xxxx xxxx xxxx xxxx)
Public Const BIT19 = &H80000            'Bit mask (xxxx 1xxx xxxx xxxx xxxx xxxx)
Public Const BIT20 = &H100000           'Bit mask (xxx1 xxxx xxxx xxxx xxxx xxxx)
Public Const BIT21 = &H200000           'Bit mask (xx1x xxxx xxxx xxxx xxxx xxxx)
Public Const BIT22 = &H400000           'Bit mask (x1xx xxxx xxxx xxxx xxxx xxxx)
Public Const BIT23 = &H800000           'Bit mask (1xxx xxxx xxxx xxxx xxxx xxxx)

'Ports
Public Const PORT0 = 0                          'DIO Port #0
Public Const PORT1 = 1                          'DIO Port #1
Public Const PORT2 = 2                          'DIO Port #2
Public Const PORT3 = 3                          'DIO Port #3
Public Const PORT4 = 4                          'DIO Port #4
Public Const PORT5 = 5                          'DIO Port #5
Public Const PORT6 = 6                          'DIO Port #6
Public Const PORT7 = 7                          'DIO Port #7
Public Const PORT8 = 8                          'DIO Port #8
Public Const PORT9 = 9                          'DIO Port #9
Public Const PORT10 = 10                        'DIO Port #10
Public Const PORT11 = 11                        'DIO Port #11

'Channels
Public Const CHAN0 = 0                          'Channel #0 output
Public Const CHAN1 = 1                          'Channel #1 output
Public Const CHAN2 = 2                          'Channel #2 output
Public Const CHAN3 = 3                          'Channel #3 output
Public Const CHAN4 = 4                          'Channel #4 output
Public Const CHAN5 = 5                          'Channel #5 output
Public Const CHAN6 = 6                          'Channel #6 output
Public Const CHAN7 = 7                          'Channel #7 output

'Hardware Constants
Public Const MAXVOLTS = 9.999694824             'Maximum voltage read by A/D card
Public Const SUPPLYIDEAL = 5#                   'Nominal supply voltage
Public Const MAXBITS = 32767                    'Maximum digital count (15 bits)
Public Const SUPPLYHIGH = 5.01                  'High tolerance for supply voltage (5.010 V)
Public Const SUPPLYLOW = 4.99                   'Low  tolerance for supply voltage (4.990 V)
Public Const SUPPLYMAX = 5.025                  'Max supply voltage (5.025 V)
Public Const SUPPLYMIN = 4.975                  'Min supply voltage (4.975 V)
Public Const MOTORSTEPSPERREV = 4000            'Number of steps per rev on the motor's encoder
Public Const ROVERALL = 5000                    'Resistance in series with OA resistor (CC224 board)
Public Const RSERIES = 500                      'Resistance in series with SR resistor (CC224 board)
Public Const PTBASEADDRESS = &H6400             'Base Address of the PT board (Base Address > 7)
Public Const MLXCLAMP1 = 768                    'Default MLX clamp value for output 1 '2.2ANM
Public Const MLXCLAMP2 = 256                    'Default MLX clamp value for output 2 '2.2ANM
Public Const MLXCLAMPC = 512                    'Default MLX clamp value for C parts  '2.6ANM

'NOTE:  Since the 3 least significant bits are reserved for internal
'       addressing, the base address must have a value greater than 7.
'       The value of 6400 Hex was randomly chosen as our standard.

'Watchdog Timers
Public Const SCANTIMEOUT = 10                   'Timeout for scanning
Public Const SCANTIMEOUTTL = 60                 'Timeout for scanning TL '3.0aANM
Public Const ZFINDTIMEOUT = 15                  'Timeout for Z-find after 15 seconds
Public Const MOVETOLOADLOCATIONTIMEOUT = 5      'Timeout for movement to load/unload location
Public Const SECONDSPERDAY = 86400              'Number of Seconds per day
Public Const GOODPALLET = 10                    'Good PLC Reject Code  '1.5ANM
Public Const BASEYEAR = 2000                    'Base Year for Date Code calculations

'Conversion factors
Public Const HUNDREDPERCENT = 100               '100%
Public Const DEGPERREV = 360                    'Number of degrees per revolutions
Public Const VOLTSPERLSB = MAXVOLTS / MAXBITS   'Convert digital counts to volts
Public Const NEWTONSPERLBF = 4.44822            'Number of Newtons/LBF

'Display Constants
Public Const NUMROWSPROGRESULTSDISPLAY = 30     'Maximum number of rows in the Programming Results display
Public Const NUMROWSPROGSTATSDISPLAY = 34       'Maximum number of rows in the Programming Stats display
Public Const SCANRESULTSGRID = 0                'Scan Results grid
Public Const PROGRESULTSGRID = 1                'Programming Results grid
Public Const SCANSTATSGRID = 2                  'Scan Stats grid
Public Const PROGSTATSGRID = 3                  'Programming Stats grid

'**************** Programming Failure Definitions ****************
'
'   Programming Failures are setup the same as Scan Failures
'   except that PROGFAULTCNT is used to signify the
'   last fault number instead of MACFAULTCNT.  The integer
'   array used is gintProgFailures(progNum,faultNum)
'

Public Const HIGHPROGINDEX1 = 1
Public Const LOWPROGINDEX1 = 2

Public Const HIGHPROGINDEX2 = 3
Public Const LOWPROGINDEX2 = 4

Public Const HIGHCLAMPLOW = 5
Public Const LOWCLAMPLOW = 6

Public Const HIGHCLAMPHIGH = 7
Public Const LOWCLAMPHIGH = 8

'2.1ANM removed offset drift
Public Const AGNDFAILURE = 9
Public Const FCKADJFAILURE = 10
Public Const CKANACHFAILURE = 11
Public Const CKDACCHFAILURE = 12
Public Const SLOWMODEFAILURE = 13

Public Const PROGFAULTCNT = 13

'******************************************************************
'*                         ENUMERATED TYPES                       *
'******************************************************************

Enum VoltageRefMode             'Enumerated type to represent the Voltage Reference Mode

    vrmReferenceIC = 0
    vrmSWControlled = 1

End Enum

Enum PLCCommType                'Enumerated type to represent the PLC Communication Type

    pctNoPLC = 0
    pctDDE = 1
    pctTTL = 2

End Enum

Enum ParameterResultsDisplay    'Enumerated Types for partResults

    prdGood = 1
    prdReject = 2
    prdNotChecked = 3
    prdEmpty = 4

End Enum

Enum PartResultsText            'Enumerated Type for Good/Bad/No Part

    prtGood = 0
    prtReject = 1
    prtNoPart = 2
    prtBlank = 3

End Enum

Enum SummaryTextBox             'Enumerated Type for Summary Text Boxes

    stbTotalUnits = 0
    stbGoodUnits = 1
    stbRejectedUnits = 2
    stbSevereUnits = 3
    stbSystemErrors = 4
    stbCurrentYield = 5
    stbLotYield = 6

End Enum

Enum ShiftType                  'Enumerated Type for Shift (A, B, or C)

    stShiftA = 1
    stShiftB = 2
    stShiftC = 3

End Enum

'******************************************************************
'*                         TYPE DEFINTIONS                        *
'******************************************************************

'*** Global Root Type Definitions ***
Type MachineParameters
    parameterName           As String               'Parameter file name
    parameterRev            As String               'Parameter file revision
    seriesID                As String               'Series number
    BOMNumber               As Integer              'BOM Setup Code
    stationCode             As Integer              'Programmer, Scanner, or Both
    loadLocation            As Single               'Location for Pedal Drive Arm to rest between scans in ° from Encoder 0°
    preScanStart            As Single               'Location to start looking for Pedal-At-Rest Location in ° from Encoder 0°
    preScanStop             As Single               'Location to stop looking for Pedal-At-Rest Location in ° from Encoder 0°
    EndScanStopForce        As Single               'Force value to search for to get end scan  '1.4ANM
    overTravel              As Single               'Distance to move beyond scan end in °
    offset4StartScan        As Single               'Distance before found Pedal-At-Rest Location to start acquiring Scan data in °
    scanLength              As Single               'Distance to scan in °
    scanStart               As Single               'Location to start scanning in ° from Encoder 0°
    scanEnd                 As Single               'Location to stop scanning in ° from Encoder 0°
    countsPerTrigger        As Integer              'How often to take readings in pulses
    encReso                 As Single               '# of encoder pulses per revolution
    gearRatio               As Single               'Total Gear Ratio of any gearing in the drive system; X:1
    FPTSlope                As Single               'Full-Pedal-Travel Transition Slope
    FPTWindow               As Single               'Full-Pedal-Travel Transition Window Length
    FPTPercentage           As Single               'Full-Pedal-Travel Transition Percentage
    FKSlope                 As Single               'Force Knee Transition Slope
    FKWindow                As Single               'Force Knee Transition Window Length
    FKPercentage            As Single               'Force Knee Transition Percentage
    slopeInterval           As Integer              'Span of slope checks, in # of data points
    slopeIncrement          As Integer              'Distance between slope checks, in # of data points
    kickdown                As Boolean              'Kickdown (Test/Don't Test)
    KDStartSlope            As Single               'Kickdown Start Transition Slope
    KDStartWindow           As Single               'Kickdown Start Transition Window Length
    KDStartPercentage       As Single               'Kickdown Start Transition Percentage
    preScanVelocity         As Single               'PreScan velocity in rps
    preScanAcceleration     As Single               'PreScan acceleration in rps²
    scanVelocity            As Single               'Scan velocity in rps
    scanAcceleration        As Single               'Scan acceleration in rps²
    scanVelocityB           As Single               'Scan velocity in rps         '3.0aANM
    scanAccelerationB       As Single               'Scan acceleration in rps²    '3.0aANM
    progVelocity            As Single               'Programming velocity in rps
    progAcceleration        As Single               'Programming acceleration in rps²
    graphZeroOffset         As Single               'Distance between Encoder 0° and Datum 0° on part; used to shift graphs
    currentPartCount        As Integer              '# of parts for hourly yield basis
    yieldGreen              As Single               'Percentage above which Yield is shown as green
    yieldYellow             As Single               'Percentage above which Yield is shown as yellow
    xAxisLow                As Single               'Graph x-axis minimum value
    xAxisHigh               As Single               'Graph x-axis maximum value
    blockOffset             As Single               'Location of Home Block in ° from Encoder 0°
    pedalAtRestLocForce     As Single               'Force at which pedal-at-rest location occurs
    filterLoc(0 To 3)       As Integer              'Filter Location (1-6, filter1-4 & Loads 1-2)
    VRefMode                As VoltageRefMode       'VRef Mode, IC controlled or SW controlled
    maxLBF                  As Single               'Maximum Force measured by the Sensotec SC2000, in LBF
    CustomerPartNum         As String               'Customer Part Number '1.1ANM
    PLCCommType             As PLCCommType          'What type of communication utilized between PC & PLC
    PCRfile                 As String               'PCR File for Laser Marker Setup
End Type
Public gudtMachine      As MachineParameters

Type SummaryCounts
    currentTotal            As Integer              'Number of total units (current)
    currentGood             As Integer              'Number of good units  (current)
    totalUnits              As Integer              'Number of total units
    totalGood               As Integer              'Number of good units
    totalReject             As Integer              'Number of rejected units
    totalSevere             As Integer              'Number of severe units
    totalNoTest             As Integer              'Number of scan errors
End Type
Public gudtScanSummary As SummaryCounts
Public gudtProgSummary As SummaryCounts

Type ProgrammingStats
    indexVal(1 To 2)        As Statistics           'Stat counts for programmed index
    indexLoc(1 To 2)        As Statistics           'Stat counts for programmed index location
    clampLow                As Statistics           'Stat counts for clamp low
    clampHigh               As Statistics           'Stat counts for clamp high
    offsetCode              As Statistics           'Stat counts for offset code
    roughGainCode           As Statistics           'Stat counts for rough gain code
    fineGainCode            As Statistics           'Stat counts for fine gain code
    clampLowCode            As Statistics           'Stat counts for clamp low code
    clampHighCode           As Statistics           'Stat counts for clamp high code
    OffsetDriftCode         As Statistics           'Stat counts for of offset drift code
    AGNDCode                As Statistics           'Stat counts for mode code
    OscillatorAdjCode       As Statistics           'Stat counts for oscillator adjust code
    CapFreqAdjCode          As Statistics           'Stat counts for capacitor frequency adjust code
    DACFreqAdjCode          As Statistics           'Stat counts for DAC frequency adjust code
    SlowModeCode            As Statistics           'Stat counts for slow code
    OffsetSeedCode          As Statistics           'Stat counts for offset seed code
    RoughGainSeedCode       As Statistics           'Stat counts for rough gain seed code
    FineGainSeedCode        As Statistics           'Stat counts for fine gain seed code
End Type
Public gudtProgStats(1 To 2)     As ProgrammingStats

'**********************************
'*  Global Variable Declarations  *
'**********************************

Public gintAnomaly As Integer

Public gintForward() As Integer, gintReverse() As Integer
Public gintForSupply() As Integer, gintRevSupply() As Integer
Public gintPreScanForce() As Integer
Public gintDatabaseStartNum As Integer    '1.1ANM
Public gintDatabaseStopNum As Integer     '1.1ANM
Public gintPLCReject As Integer           '1.5ANM
Public gdblDelay As Double                '2.1ANM

Public gvntGraph() As Variant
Public gsngMultipleGraphArray() As Single

Public gsngMonitorData(CHAN0 To CHAN7) As Single

Public gblnPLCStart As Boolean
Public gblnAdministrator As Boolean
Public gblnGraphEnable As Boolean
Public gblnLockICs As Boolean
Public gblnVRefStartupDone As Boolean
Public gblnAnalogDone As Boolean
Public gblnScanStart As Boolean
Public gblnProgramStart As Boolean
Public gblnEStop As Boolean
Public gblnScanFailure As Boolean
Public gblnSevere As Boolean
Public gblnProgFailure As Boolean
Public gblnSaveRawData As Boolean
Public gblnSaveScanResultsToFile As Boolean
Public gblnSaveProgResultsToFile As Boolean
Public gblnStartUpDone As Boolean
Public gblnParFileSelected As Boolean
Public gblnLotFileSelected As Boolean
Public gblnGoodSerialNumber As Boolean
Public gblnGoodDateCode As Boolean
Public gblnGoodOffsetAndGainCodes As Boolean
Public gblnForceOnly As Boolean                 '1.7ANM
Public gblnLockedPart As Boolean                '2.0ANM
Public gblnLockRejects As Boolean               '2.0ANM
Public gblnTLScanner As Boolean                 '2.2ANM
Public gblnMasterPara As Boolean                '2.5ANM
Public gblnUseNewAmad As Boolean                '2.6ANM
Public gblnBnmkTest As Boolean                  '2.6ANM
Public gblnLockSkip As Boolean                  '2.8ANM
Public gblnReClampEnable As Boolean             '2.8ANM
Public gblnReClamp As Boolean                   '2.8ANM
Public gblnReScanEnable As Boolean              '2.8ANM
Public gblnReScanRun As Boolean                 '2.8ANM
Public gblnMLXOk As Boolean                     '3.0aANM

Public gintNumRowsInResultsDisplay As Integer
Public gintNumRowsInStatsDisplay As Integer
Public gintSevere(MAXCHANNUM, MAXFAULTCNT) As Integer
Public gintFailure(MAXCHANNUM, MAXFAULTCNT) As Integer
Public gintProgFailure(1 To 2, PROGFAULTCNT) As Integer
Public gintPalletNumber As Integer          'Pallet Number
Public gintPointer As Integer               'Pointer to the current graph Number
Public gintMaxData As Integer               'Number of data points in scan arrays

Public gsngVRef As Single                   'Measured supply voltage (V)
Public gsngVRefGain As Single               'Measured VRef/DAC Gain
Public gsngVRefOffset As Single             'Measured VRef Offset (V)
Public gsngVRefSetPoint As Single           'Set Point for VRef (V)
Public gsngMeanSupplyVoltage As Single      'Mean Scan Supply Voltage (V)
Public gsngForceAmplifierOffset As Single   'Force Amplifier Offset
Public gsngNewtonsPerVolt As Single         'Force Amplifier output conversion
Public gsngResolution As Single             'Data Resolution, data points / degree
Public gsngCycleTimerStart As Single        'Cycle-Time Timer
Public gsngEndForce As Single               'Force found at end scan  '1.4ANM
Public gsngEndPos As Single                 'End scan force position  '1.4ANM
Public gsngForceOffset As Single            'Force Offset             '2.0ANM
Public gsngForceGain As Single              'Force Gain               '2.0ANM

Public gstrSystemName                       'Name of the Station
Public gstrSerialNumber As String           'Serial Number
Public gstrDateCode As String               'Date Code
Public gstrLotName As String                'Lot Name
Public gstrStatFilePath As String           'Stat File Path
Public gstrErrorFileName As String          'Error Log File Name
Public gstrDateCode2 As String              'Date Code for rework     '1.5ANM
Public gstrPalletLoad As String             'Pallet Load for rework   '1.5ANM
Public gstrSN1 As String                    'Temp SN 1                '2.5ANM
Public gstrSN2 As String                    'Temp SN 2                '2.5ANM
Public gstrSN3 As String                    'Temp SN 3                '2.5ANM
Public gstrSampleNum As String              'Sample Number            '3.0ANM

Public gfsoFileSystemObject As New FileSystemObject

Public Sub AdjustVRef()
'
'   PURPOSE: Measures the supply voltage and calculates the gain factor.
'            Also automatically adjusts voltage to nominal is using the
'            DAQ board's DAC output.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintAdjustCnt As Integer
Dim lsngVRefSetPointPast As Single
Dim lsngVRefAdjust As Single

On Error GoTo AdjustVRef_Err
    
If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    'If this is the first time adjusting VRef, initialize the DAC output
    If Not gblnVRefStartupDone Then
        Call CalculateVRefParameters(1)
        gblnVRefStartupDone = True
    End If

    'Initialize the set point
    gsngVRefSetPoint = ((SUPPLYIDEAL - gsngVRefOffset) / gsngVRefGain)

    'Try five adjustments, if necessary
    For lintAdjustCnt = 1 To 5
        'Write to the D/A
        frmDAQIO.cwaoVRef.SingleWrite (gsngVRefSetPoint)

        'Delay (25 msec) between write & read
        Call frmDAQIO.KillTime(25)

        'Read the Reference Voltage
        gsngVRef = frmDAQIO.ReadVRef

        'Check if VRef is within the tolerance band
        If (gsngVRef < SUPPLYHIGH) And (gsngVRef > SUPPLYLOW) Then
            'VRef in desired limits, exit For Loop
            Exit For
        Else
            lsngVRefSetPointPast = gsngVRefSetPoint                     'Store current voltage
            lsngVRefAdjust = (SUPPLYIDEAL - gsngVRef)                   'Calculate delta voltage
            'Calculate the new voltage to write
            gsngVRefSetPoint = ((lsngVRefAdjust) / gsngVRefGain) + lsngVRefSetPointPast
        End If
    Next lintAdjustCnt

End If

Exit Sub
AdjustVRef_Err:

    gintAnomaly = 6
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in Pedal.AdjustVRef: " & Err.Description, True, True)

End Sub

Public Function CalcDayOfYear(ByVal Month As Long, ByVal Day As Long, LeapYear As Boolean) As Integer
'
'   PURPOSE:    To calculate the numbered day of the year.
'
'  INPUT(S):    Month         = Current month
'               Day           = Current day of month
'               LeapYear      = Whether or not the current year is a leap year
' OUTPUT(S):    CalcDayOfYear = Number day of the year

CalcDayOfYear = Int((275 * Month) / 9) - IIf(LeapYear, 1, 2) * Int((Month + 9) / 12) + Day - 30

End Function

Public Sub CalcLBFPerVolt()
'
'   PURPOSE:    To calculate the coefficient for converting voltage from the force
'               amplifier to LBF, as well as the offset inherent with the DAC.
'
'     NOTES:    This is done to correct for any differences between the force
'               amplifier's output's reference and the DAQ board's reference, as
'               well as any gain associated with filtering and buffering.
'               The coefficient is a single precision number representing Newtons
'               per Volt, while the offset is a single precision number representing
'               Newtons to be added to the (Volts measured) * (Newtons per Volt)
'               result.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None
'

Dim lsngDataX(8) As Single
Dim lsngDataY(8) As Single
Dim lvntRead As Variant
Dim lsngSlope As Single
Dim lsngIntercept As Single
Dim i As Integer

'*** Calculate the voltage->force coefficient ***
'This is accomplished by forcing several different percentages of the full scale
'output, reading each, and determining a transfer function with a best-fit line
'Note: The SC2000 interprets compression force as a negative reading.
'      This is accounted for here, at essentially the lowest level.
'      The conversion factor is calculated to interpret negative voltages
'      (compression force) as positive force readings.
For i = 0 To 8
    'Force the force amplifier's DAC output to (10 * i)% of full-scale
    Call Sensotec.ForceDACOutput(1, (-10 * i), True)   'Negative because we're measuring compression
   '500 millisecond delay
    Call frmDAQIO.KillTime(500)
    'Read the force channel
    Call frmDAQIO.cwaiForce.SingleRead(lvntRead)
    'X-axis data: Determine the average of the read voltages
    lsngDataX(i) = frmDAQIO.cwStat1.Mean(lvntRead)
    'Y-axis data: pseudo force output of the amplifier in Newtons
    lsngDataY(i) = ((10 * i) / HUNDREDPERCENT) * (gudtMachine.maxLBF * NEWTONSPERLBF)
Next i

'*** Assign the calculated coefficient value ***
Call CalcLSQLine(lsngDataY(), lsngDataX(), 0, 8, 1, lsngSlope, lsngIntercept)

'Assign the coefficient (single precision, Newton/Volt)
gsngNewtonsPerVolt = lsngSlope
'Assign the offset (single precision, LBF)
gsngForceAmplifierOffset = lsngIntercept

'Reset the DACOutput to auto mode
Call Sensotec.ForceDACOutput(1, 0, False)

End Sub

Public Sub CalcLimitLineMandB(ByVal StartLoc As Single, ByVal startHigh As Single, ByVal startLow As Single, ByVal StopLoc As Single, ByVal stopHigh As Single, ByVal stopLow As Single, highLimitM As Single, lowLimitM As Single, highLimitB As Single, lowLimitB As Single)
'
'   PURPOSE:    To calculate the slopes and intercepts of high and low limit lines
'
'  INPUT(S):    startLoc    = start location of the region
'               startHigh   = high limit at the start of the region
'               startLow    = low limit at the start of the region
'               stopLoc     = stop location of the region
'               stopHigh    = high limit at the stop of the region
'               stopLow     = low limit at the stop of the region
'
' OUTPUT(S):    highLimitM  = Slope of the High Limit Line
'               highLimitB  = Intercept for the High Limit Line
'               lowLimitM   = Slope of the Low Limit Line
'               lowLimitB   = Intercept for the Low Limit Line

'Avoid division-by-zero
If StopLoc - StartLoc <> 0 Then
    'High Limit Slope
    highLimitM = (stopHigh - startHigh) / (StopLoc - StartLoc)
    'Low Limit Slope
    lowLimitM = (stopLow - startLow) / (StopLoc - StartLoc)
Else
    'High Limit Slope
    highLimitM = 0
    'Low Limit Slope
    lowLimitM = 0
End If

'High Limit Interecept
highLimitB = startHigh - (StartLoc * highLimitM)
'Low Limit Interecept
lowLimitB = startLow - (StartLoc * lowLimitM)

End Sub

Private Sub CalcLSQLine(yDataArray() As Single, xDataArray, ByVal evaluateStart As Single, ByVal evaluateStop As Single, resolution As Single, m As Single, b As Single)
'
'   PURPOSE:    To calculate the Least Squares Approximation (BEST FIT LINE)
'               of the X and Y data passed in
'
'     NOTES:    For the linear portion of the curve, the ideal value follows
'               the equation of a line:  y(x) = m * (x - n) + b, where:
'
'               y(x) =  ideal value @ point x
'                 m  =  ideal slope
'                 x  =  location of ideal value point
'                 n  =  location of index point
'                 b  =  output at index point
'
'               The linearity checks are typically performed on only the forward data.
'
'  INPUT(S):    ydataArray    : Y-axis data
'               xDataArray    : X-axis data
'               evaluateStart : Start point of evaluation
'               evaluateStop  : End point of evalutation
'               resolution    : Resolution of data
' OUTPUT(S):    m             : Least Squares Slope
'               B             : Least Squares Offset

Dim i As Integer                        'loop counter
Dim lintN As Integer                    'count of sums
Dim lsngSigmaX As Single                'sum of X's
Dim lsngSigmaY As Single                'sum of Y's
Dim lsngSigmaX2 As Single               'sum of X^2's
Dim lsngSigmaXY As Single               'sum of Y^2's
Dim lsngXBar As Single                  'average of X's
Dim lsngYBar As Single                  'average of Y's
Dim lsngNumerator As Single             'numerator for Least Squares slope calculation
Dim lsngDenominator As Single           'denominator for Least Squares slope calculation

'*** Initialize variables for arrays ***
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate summations required for Least Squares Approximation of Best Fit Line
For i = evaluateStart To evaluateStop
    lsngSigmaX = lsngSigmaX + xDataArray(i)
    lsngSigmaX2 = lsngSigmaX2 + (xDataArray(i) ^ 2)
    lsngSigmaY = lsngSigmaY + yDataArray(i)
    lsngSigmaXY = lsngSigmaXY + (xDataArray(i) * yDataArray(i))
    lintN = lintN + 1
Next

' LEAST SQUARES APPROXIMATION
'
' Given the equation of a line is:   y(x) = mx + b
'
' Calculate m & b as follows:
'
'                            SigmaXY - n * xBar * yBar
'                       m = ___________________________
'
'                             SigmaX^2 - n * xBar^2
'
'
'                       b =  yBar - m * xBar
'

'Calculate averages of X & Y
lsngXBar = lsngSigmaX / lintN
lsngYBar = lsngSigmaY / lintN

'Calculate LEAST SQUARES APPROXIMATION SLOPE:
lsngNumerator = lsngSigmaXY - (lintN * lsngXBar * lsngYBar)
lsngDenominator = lsngSigmaX2 - (lintN * lsngXBar ^ 2)
m = lsngNumerator / lsngDenominator

b = lsngYBar - (m * lsngXBar)

End Sub

Public Function CalculateKickdownOnLocation() As Single
'
'     PURPOSE:  To determine the ideal Kickdown On Location
'
'    INPUT(S):  None.
'   OUTPUT(S):  Returns Ideal Kickdown On Location

Dim lintProgrammerNum As Integer            'Programmer Number
Dim lsngKDOnLocation As Single              'Calculated Kickdown On Location
Dim lsngWOTLoc As Single                    'Calculated WOT location based on KDON   '1.3ANM \/\/
Dim lsngSlope As Single                     'Calculated 'ideal' slope
Dim lsngAdjustmentResolution As Single      'Calculated adjustment resolution
Dim lsngWOTToleranceBand As Single          'Calculated WOT tolerance band
Dim lsngWOTGuardBand As Single              'Calculated WOT guardband
Dim lsngKDONToleranceBand As Single         'Calculated KDON tolerance band
Dim lsngKDONGuardBand As Single             'Calculated KDON guardband
Dim lsngKDONLocHigh As Single               'Calculated KDON high limit
Dim lsngKDONLocLow As Single                'Calculated KDON low limit
Dim lsngWOTLocLow As Single                 'Calculated WOT location low limit
Dim lsngWOTLocHigh As Single                'Calculated WOT location high limit

Dim X As Integer                                                                     '1.3ANM /\/\

'For 700 series parts, Kickdown On Location is relative to Kickdown Start Location
If gudtMachine.seriesID = "700" Then
    'Return the ideal Kickdown On Location Relative to Pedal 0°
    '(Kickdown Start Location +  Ideal Kickdown On Span)
    lsngKDOnLocation = FindKickdownStartLocation + gudtTest(CHAN0).kickdownOnSpan.ideal
'For 702&703 series parts, Kickdown On Location is relative to Kickdown Peak Location
ElseIf (gudtMachine.seriesID = "702") Or (gudtMachine.seriesID = "703") Then
    'Return the ideal Kickdown On Location Relative to Pedal 0°
    '(Kickdown Peak Location +  Ideal Kickdown On Span)
    lsngKDOnLocation = FindKickdownPeakLocation + gudtTest(CHAN0).kickdownOnSpan.ideal
End If

'1.9ANM \/\/ Smart KD fix
Dim lintIndex As Integer                    'Index value
If gudtMachine.seriesID = "700" Then
    lintIndex = 3
Else
    lintIndex = 2
End If

'Calculate slope, WOTLoc, tolerance bands, adjustment resolution, and high/low
If gudtMachine.seriesID = "703" Then
    lsngSlope = ((gudtSolver(1).Index(2).IdealValue - gudtTest(CHAN0).Index(1).ideal) / (lsngKDOnLocation - gudtTest(CHAN0).riseTarget))
    lsngWOTLoc = lsngKDOnLocation - ((gudtSolver(1).Index(2).IdealValue - gudtTest(CHAN0).Index(lintIndex).ideal) / lsngSlope)
    lsngWOTToleranceBand = (gudtTest(CHAN0).Index(lintIndex).high - gudtTest(CHAN0).Index(lintIndex).low) / lsngSlope
    lsngWOTLocLow = gudtTest(CHAN0).Index(lintIndex).location - ((gudtTest(CHAN0).Index(lintIndex).ideal - gudtTest(CHAN0).Index(lintIndex).low) / lsngSlope)
    lsngWOTLocHigh = ((gudtTest(CHAN0).Index(lintIndex).high - gudtTest(CHAN0).Index(lintIndex).ideal) / lsngSlope) + gudtTest(CHAN0).Index(lintIndex).location
Else
    lsngSlope = ((gudtSolver(1).Index(2).IdealValue - gudtTest(CHAN0).Index(1).ideal) / (lsngKDOnLocation - gudtTest(CHAN0).riseTarget))
    lsngWOTLoc = lsngKDOnLocation - ((gudtSolver(1).Index(2).IdealValue - gudtTest(CHAN0).Index(lintIndex).location) / lsngSlope)
    lsngWOTToleranceBand = gudtTest(CHAN0).Index(lintIndex).high - gudtTest(CHAN0).Index(lintIndex).low
    lsngWOTLocLow = gudtTest(CHAN0).Index(lintIndex).low
    lsngWOTLocHigh = gudtTest(CHAN0).Index(lintIndex).high
End If

lsngKDONToleranceBand = gudtTest(CHAN0).kickdownOnSpan.high - gudtTest(CHAN0).kickdownOnSpan.low
lsngAdjustmentResolution = (lsngKDONToleranceBand / 10)
lsngKDONLocHigh = lsngKDOnLocation + (gudtTest(CHAN0).kickdownOnSpan.high - gudtTest(CHAN0).kickdownOnSpan.ideal)
lsngKDONLocLow = lsngKDOnLocation - (gudtTest(CHAN0).kickdownOnSpan.ideal - gudtTest(CHAN0).kickdownOnSpan.low)

'Guardband 10%
lsngWOTGuardBand = (0.1 * lsngWOTToleranceBand)
lsngKDONGuardBand = (0.1 * lsngKDONToleranceBand)

'Loop Until WOT is inside of the guardbanded tolerance region
Do While (lsngWOTLoc < (lsngWOTLocLow + lsngWOTGuardBand)) Or (lsngWOTLoc > (lsngWOTLocHigh - lsngWOTGuardBand))
    'Adjust KD On Loc
    If (lsngWOTLoc > (lsngWOTLocHigh - lsngWOTGuardBand)) Then
        lsngKDOnLocation = lsngKDOnLocation - lsngAdjustmentResolution
    Else
        lsngKDOnLocation = lsngKDOnLocation + lsngAdjustmentResolution
    End If
    
    'Exit the Loop if KDOnSpan is outside of it's tolerance
    If ((lsngKDOnLocation > (lsngKDONLocHigh - lsngKDONGuardBand)) Or (lsngKDOnLocation < (lsngKDONLocLow + lsngKDONGuardBand))) Then Exit Do
    
    'Calculate the new theoretical slope
    lsngSlope = ((gudtSolver(1).Index(lintIndex).IdealValue - gudtTest(CHAN0).Index(1).ideal) / (lsngKDOnLocation - gudtTest(CHAN0).riseTarget))
    'Calculate the new WOT Loc
    If gudtMachine.seriesID = "703" Then
        lsngWOTLoc = lsngKDOnLocation - ((gudtSolver(1).Index(lintIndex).IdealValue - gudtTest(CHAN0).Index(lintIndex).ideal) / lsngSlope)
    Else
        lsngWOTLoc = lsngKDOnLocation - ((gudtSolver(1).Index(lintIndex).IdealValue - gudtTest(CHAN0).Index(lintIndex).location) / lsngSlope)
    End If
Loop
'1.9ANM /\/\

'Index 2 should occur at the Ideal Kickdown On Location
'Note: This is for Kickdown parts ONLY!!! Index 2 should be a static location
'      for other parts
For lintProgrammerNum = 1 To 2
    gudtSolver(lintProgrammerNum).Index(2).IdealLocation = lsngKDOnLocation
Next lintProgrammerNum

'Return the Kickdown On Location
CalculateKickdownOnLocation = lsngKDOnLocation

End Function

Public Sub CalculateVRefParameters(gainEstimate As Single)
'
'     PURPOSE:  To determine the gain & offset for the VRef/DAC relationship
'
'    INPUT(S):  gainEstimate : An estimate of the gain for the CC259 Board
'   OUTPUT(S):  None.

Dim lintIteration As Integer            'Iteration number
Dim lsngVRef As Single                  'Measured VRef
Dim lsngWrite As Integer                'Voltage to Write
Dim lintN As Integer                    'Count of sums
Dim lsngSigmaX As Single                'Sum of X's
Dim lsngSigmaY As Single                'Sum of Y's
Dim lsngSigmaX2 As Single               'Sum of X^2's
Dim lsngSigmaXY As Single               'Sum of Y^2's
Dim lsngXBar As Single                  'Average of X's
Dim lsngYBar As Single                  'Average of Y's
Dim lsngNumerator As Single             'Numerator for Least Squares slope calculation
Dim lsngDenominator As Single           'Denominator for Least Squares slope calculation

'Write to D/A and calculate summations for best-fit line of VRef response
For lintIteration = 1 To 10
    lsngWrite = lintIteration / gainEstimate        'Calculate a value to write
    frmDAQIO.cwaoVRef.SingleWrite (lsngWrite)       'Write to the D/A
    Call frmDAQIO.KillTime(50)                      'Wait 50 msec for VRef to settle
    lsngVRef = frmDAQIO.ReadVRef                    'Read VRef
    'Ignore the data if it is within one millivolt of the maximum A/D reading
    If (lsngVRef <= MAXVOLTS - 0.001) Then
        'Keep track of writes and reads for best-fit line approximation
        lsngSigmaX = lsngSigmaX + lsngWrite
        lsngSigmaX2 = lsngSigmaX2 + (lsngWrite ^ 2)
        lsngSigmaY = lsngSigmaY + lsngVRef
        lsngSigmaXY = lsngSigmaXY + (lsngWrite * lsngVRef)
        lintN = lintN + 1
    End If
Next lintIteration

' LEAST SQUARES APPROXIMATION
'
' Given the equation of a line is:   y(x) = mx + b
'
' Calculate m & b as follows:
'
'                            SigmaXY - n * xBar * yBar
'                       m = ___________________________
'
'                             SigmaX^2 - n * xBar^2
'
'
'                       b =  yBar - m * xBar
'

'Calculate averages of X & Y
lsngXBar = lsngSigmaX / lintN
lsngYBar = lsngSigmaY / lintN

'Calculate LEAST SQUARES APPROXIMATION SLOPE:
lsngNumerator = lsngSigmaXY - lintN * lsngXBar * lsngYBar
lsngDenominator = lsngSigmaX2 - lintN * lsngXBar ^ 2
gsngVRefGain = lsngNumerator / lsngDenominator

'Calculate LEAST SQUARES APPROXIMATION Y-INTERCEPT
gsngVRefOffset = lsngYBar - (gsngVRefGain * lsngXBar)

End Sub

Public Sub CheckForProgrammingFaults()
'
'     PURPOSE:  To check for programming faults and set the pass/fail boolean
'
'    INPUT(S):  None.
'   OUTPUT(S):  None.

Dim lintProgrammerNum As Integer
Dim lintFaultNum As Integer
Dim lsngIdealSlope As Single

'Check the Solver outputs for pass/fail
For lintProgrammerNum = 1 To 2
    'NOTE: The Index checks are based on the actual position WOT was programmed at:
    'Calculate the ideal slope to use in calculating Index limits based on actual locations
    lsngIdealSlope = (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(1).IdealValue) / (gudtSolver(lintProgrammerNum).Index(2).IdealLocation - gudtSolver(lintProgrammerNum).Index(1).IdealLocation)
    'Check Index 1 (Idle)
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, LOWPROGINDEX1, HIGHPROGINDEX1, gintProgFailure())
    'Check Index 2 (WOT)
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, LOWPROGINDEX2, HIGHPROGINDEX2, gintProgFailure())
    'Check the Low Clamp
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue - gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue + gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, LOWCLAMPLOW, HIGHCLAMPLOW, gintProgFailure())
    'Check the High Clamp
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue - gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue + gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, LOWCLAMPHIGH, HIGHCLAMPHIGH, gintProgFailure())
    'Check Offset Drift Code
    '2.1ANM gintProgFailure(lintProgrammerNum, HIGHOFFSETDRIFT) = (gudtMLX90277(lintProgrammerNum).Read.Drift > gudtSolver(lintProgrammerNum).MaxOffsetDrift)
    'Check AGND Code
    Call Calc.CheckFault(lintProgrammerNum, gudtMLX90277(lintProgrammerNum).Read.AGND, gudtMLX90277(lintProgrammerNum).Read.AGND, gudtSolver(lintProgrammerNum).MinAGND, gudtSolver(lintProgrammerNum).MaxAGND, AGNDFAILURE, AGNDFAILURE, gintProgFailure()) '2.0ANM fixed AGND issue
    'Check Oscillator Adjust Code
    gintProgFailure(lintProgrammerNum, FCKADJFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.FCKADJ <> gudtSolver(lintProgrammerNum).FCKADJ)
    'Check Capacitor Frequency Adjust Code
    gintProgFailure(lintProgrammerNum, CKANACHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKANACH <> gudtSolver(lintProgrammerNum).CKANACH)
    'Check DAC Code Frequency Adjust Code
    gintProgFailure(lintProgrammerNum, CKDACCHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKDACCH <> gudtSolver(lintProgrammerNum).CKDACCH)
    'Check Slow Code
    gintProgFailure(lintProgrammerNum, SLOWMODEFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.SlowMode <> gudtSolver(lintProgrammerNum).SlowMode)
Next lintProgrammerNum

'Check each output
For lintProgrammerNum = 1 To 2
    'Check every fault on each output
    For lintFaultNum = 1 To PROGFAULTCNT
        If gintProgFailure(lintProgrammerNum, lintFaultNum) Then
            gblnProgFailure = True         'Failure occured
        End If
    Next lintFaultNum
Next lintProgrammerNum

'Set the part status control
If gblnProgFailure Then
    frmMain.ctrStatus1.StatusOnText(1) = "REJECT"
    frmMain.ctrStatus1.StatusOnColor(1) = vbRed
Else
    frmMain.ctrStatus1.StatusOnText(1) = "GOOD"
    frmMain.ctrStatus1.StatusOnColor(1) = vbGreen
End If

'Turn the status control on
frmMain.ctrStatus1.StatusValue(1) = True

End Sub

Public Function CheckStartPosition(ScanDirection As String) As Single
'
'   PURPOSE:    To verify that the start position is not within the scan region.
'
'  INPUT(S):    None
' OUTPUT(S):    None

Dim lsngCurrentPosition As Single

'Before the forward scan, verify that the start position is not in the scan region
lsngCurrentPosition = Position

If ScanDirection = "Pre-Scan" Then
    If gudtMachine.preScanStart < gudtMachine.preScanStop Then
        If lsngCurrentPosition > gudtMachine.preScanStart Then gintAnomaly = 104
    Else
        If lsngCurrentPosition < gudtMachine.preScanStart Then gintAnomaly = 104
    End If
ElseIf ScanDirection = "Forward" Then
    If gudtMachine.scanStart < gudtMachine.scanEnd Then
        If lsngCurrentPosition > gudtMachine.scanStart Then gintAnomaly = 105
    Else
        If lsngCurrentPosition < gudtMachine.scanStart Then gintAnomaly = 105
    End If
ElseIf ScanDirection = "Reverse" Then
    If gudtMachine.scanStart < gudtMachine.scanEnd Then
        If lsngCurrentPosition < gudtMachine.scanEnd Then gintAnomaly = 106
    Else
        If lsngCurrentPosition > gudtMachine.scanEnd Then gintAnomaly = 106
    End If
End If

'Display the anomaly message if needed
If gintAnomaly Then
    'Log the error to the error log and display the error message
    Call ErrorLogFile("The Current Position is in the Scan Region." & vbCrLf & _
                      "Please Home the motor using the Function Menu", True, True)
End If

End Function

Public Function CheckSupplyArray() As Single
'
'   PURPOSE:    To verify that no points in the VRef Array are outside the limits,
'               as well as calculate the mean supply voltage.
'
'  INPUT(S):    None
' OUTPUT(S):    None

Dim i As Integer
Dim lsngSupplyTotal As Single
Dim lsngForwardSupply As Single
Dim lsngReverseSupply As Single

For i = 0 To (gintMaxData - 1)
    'Translate counts to volts
    lsngForwardSupply = gintForSupply(i) * VOLTSPERLSB
    lsngReverseSupply = gintRevSupply(i) * VOLTSPERLSB

    'Check forward and reverse against Max Supply voltage
    If (lsngForwardSupply >= SUPPLYMAX) Or (lsngReverseSupply >= SUPPLYMAX) Then
        gintAnomaly = 2
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Reference Supply Too High During Scan", True, True)
        Exit For
    'Check forward and reverse against Min Supply voltage
    ElseIf (lsngForwardSupply <= SUPPLYMIN) Or (lsngReverseSupply <= SUPPLYMIN) Then
        gintAnomaly = 3
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Reference Supply Too Low During Scan", True, True)
        Exit For
    End If
    'Keep a running total of the forward and reverse supply voltage
    lsngSupplyTotal = lsngSupplyTotal + lsngForwardSupply + lsngReverseSupply
Next i

'Calculate the average supply voltage of the forward and reverse scans
gsngMeanSupplyVoltage = lsngSupplyTotal / (gintMaxData * 2)

End Function

Public Sub ClearCounter()
'
'     PURPOSE:  To clear the 24-bit counter from the Position Trigger Board
'
'    INPUT(S):  None.
'   OUTPUT(S):  None.

Dim Address As Long
Dim Data As Variant
Dim LSBAddr As Integer, MSBAddr As Integer
Dim LSBData As Variant, MIDData As Variant, MSBData As Variant

'*** Get Clear Trigger Count Address ***
'Note:  By multiplying by 8 Hex ("1000" binary), the base address is being
'       shifted to the left by 3 bits to accommodate the opcode for the
'       address.  The opcode is then added to the address. The opcode resides
'       in the 3 least significant bits.

Address = (PTBASEADDRESS) + &H3             'Address of clear trigger count register
LSBAddr = Address And &HFF                  'Get LSB address
MSBAddr = (Address \ BIT8) And &HFF         'Get MSB address
                                            
'*** Read Position Data ***
Call frmDAQIO.ReadPTBoardData(LSBAddr, MSBAddr, LSBData, MIDData, MSBData)

End Sub

Public Function DirectoryExists(DirPath As String) As Boolean
'
'   PURPOSE: To determine whether the input directory exists and if not
'            create it.
'
'  INPUT(S): DirPath = Path to the directory
'
' OUTPUT(S): None

On Error GoTo DirCheckErr

'Initialize variable(s)
DirectoryExists = False

'Create directory if it doesn't exist
If Not gfsoFileSystemObject.FolderExists(DirPath) Then
    gfsoFileSystemObject.CreateFolder (DirPath)
    DirectoryExists = True                          'Set to True, if no error detected
Else
    DirectoryExists = True                          'Set to True, if directory already exists
End If

Exit Function

DirCheckErr:

    MsgBox "Error encountered while trying to create the directory:  " & DirPath _
           & vbCrLf + vbCrLf & "                    Verify the security on this directory.", vbOKOnly, DirPath & " Directory Not Found"

End Function

Public Sub DisplayInitialization()
'
'   PURPOSE: To initialize the display of the results and statistics screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim llngRowNum As Long
Dim llngColNum As Long
Dim lvntAlignmentType As Variant

'Tab numbers:
'Tab 0 = Scan Results Tab
'Tab 1 = Programming Results Tab
'Tab 2 = Scan Graphs Tab
'Tab 3 = Scan Stats Tab
'Tab 4 = Programming Stats Tab

'Define the Tab names
frmMain.ctrResultsTabs1.TabName(0) = "Scan Results"
frmMain.ctrResultsTabs1.TabName(1) = "Prog Results"
frmMain.ctrResultsTabs1.TabName(2) = "Scan Graphs"
frmMain.ctrResultsTabs1.TabName(3) = "Scan Statistics"
frmMain.ctrResultsTabs1.TabName(4) = "Prog Statistics"

'Initialize number of rows in displays
frmMain.ctrResultsTabs1.NumberOfRows(SCANRESULTSGRID) = NUMROWSSCANRESULTSDISPLAY + 1
frmMain.ctrResultsTabs1.NumberOfRows(PROGRESULTSGRID) = NUMROWSPROGRESULTSDISPLAY + 1
frmMain.ctrResultsTabs1.NumberOfRows(SCANSTATSGRID) = NUMROWSSCANSTATSDISPLAY + 1
frmMain.ctrResultsTabs1.NumberOfRows(PROGSTATSGRID) = NUMROWSPROGSTATSDISPLAY + 1

'Define Cell properties for the Scan Results grid
frmMain.ctrResultsTabs1.TotalCellFontSize(SCANRESULTSGRID) = 12
frmMain.ctrResultsTabs1.BoldRow(SCANRESULTSGRID, 0) = True
frmMain.ctrResultsTabs1.RowAlignment(SCANRESULTSGRID, 0) = flexAlignCenterCenter
frmMain.ctrResultsTabs1.TotalWordWrap(SCANRESULTSGRID) = True
frmMain.ctrResultsTabs1.TotalCellFont(SCANRESULTSGRID) = "Arial"

'Define Cell properties for the Programming Results grid
frmMain.ctrResultsTabs1.TotalCellFontSize(PROGRESULTSGRID) = 12
frmMain.ctrResultsTabs1.BoldRow(PROGRESULTSGRID, 0) = True
frmMain.ctrResultsTabs1.RowAlignment(PROGRESULTSGRID, 0) = flexAlignCenterCenter
frmMain.ctrResultsTabs1.TotalWordWrap(PROGRESULTSGRID) = True
frmMain.ctrResultsTabs1.TotalCellFont(PROGRESULTSGRID) = "Arial"

'Define Cell properties for Scan Stats grid
frmMain.ctrResultsTabs1.TotalCellFontSize(SCANSTATSGRID) = 12
frmMain.ctrResultsTabs1.BoldRow(SCANSTATSGRID, 0) = True
frmMain.ctrResultsTabs1.RowAlignment(SCANSTATSGRID, 0) = flexAlignCenterCenter
frmMain.ctrResultsTabs1.TotalWordWrap(SCANSTATSGRID) = True
frmMain.ctrResultsTabs1.TotalCellFont(SCANSTATSGRID) = "Arial"

'Define Cell properties for Programming Stats grid
frmMain.ctrResultsTabs1.TotalCellFontSize(PROGSTATSGRID) = 12
frmMain.ctrResultsTabs1.BoldRow(PROGSTATSGRID, 0) = True
frmMain.ctrResultsTabs1.RowAlignment(PROGSTATSGRID, 0) = flexAlignCenterCenter
frmMain.ctrResultsTabs1.TotalWordWrap(PROGSTATSGRID) = True
frmMain.ctrResultsTabs1.TotalCellFont(PROGSTATSGRID) = "Arial"

'Define column widths
'Total available width is 13665, and a vertical scroll bar will use 300.
'Dimensions snap to multiples of 15

'Define Scan Results grid column spacing
frmMain.ctrResultsTabs1.NumberOfColumns(SCANRESULTSGRID) = 5
frmMain.ctrResultsTabs1.ColumnSpacing(SCANRESULTSGRID, 0) = 4590
frmMain.ctrResultsTabs1.ColumnSpacing(SCANRESULTSGRID, 1) = 3525
frmMain.ctrResultsTabs1.ColumnSpacing(SCANRESULTSGRID, 2) = 2840
frmMain.ctrResultsTabs1.ColumnSpacing(SCANRESULTSGRID, 3) = 1205
frmMain.ctrResultsTabs1.ColumnSpacing(SCANRESULTSGRID, 4) = 1205

'Define Programming Results grid column spacing
frmMain.ctrResultsTabs1.NumberOfColumns(PROGRESULTSGRID) = 5
frmMain.ctrResultsTabs1.ColumnSpacing(PROGRESULTSGRID, 0) = 4590
frmMain.ctrResultsTabs1.ColumnSpacing(PROGRESULTSGRID, 1) = 3525
frmMain.ctrResultsTabs1.ColumnSpacing(PROGRESULTSGRID, 2) = 2840
frmMain.ctrResultsTabs1.ColumnSpacing(PROGRESULTSGRID, 3) = 1205
frmMain.ctrResultsTabs1.ColumnSpacing(PROGRESULTSGRID, 4) = 1205

'Define Scan Stats grid column spacing
frmMain.ctrResultsTabs1.NumberOfColumns(SCANSTATSGRID) = 9
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 0) = 4590
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 1) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 2) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 3) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 4) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 5) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 6) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 7) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(SCANSTATSGRID, 8) = 1095

'Define Programming Stats grid column spacing
frmMain.ctrResultsTabs1.NumberOfColumns(PROGSTATSGRID) = 9
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 0) = 4590
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 1) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 2) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 3) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 4) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 5) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 6) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 7) = 1095
frmMain.ctrResultsTabs1.ColumnSpacing(PROGSTATSGRID, 8) = 1095

'Define Scan Results grid headings
frmMain.ctrResultsTabs1.Data(SCANRESULTSGRID, 0, 0) = "Parameter"
frmMain.ctrResultsTabs1.Data(SCANRESULTSGRID, 0, 1) = "Value"
frmMain.ctrResultsTabs1.Data(SCANRESULTSGRID, 0, 2) = "Status"
frmMain.ctrResultsTabs1.Data(SCANRESULTSGRID, 0, 3) = "High Count"
frmMain.ctrResultsTabs1.Data(SCANRESULTSGRID, 0, 4) = "Low Count"

'Define Programming Results grid headings
frmMain.ctrResultsTabs1.Data(PROGRESULTSGRID, 0, 0) = "Parameter"
frmMain.ctrResultsTabs1.Data(PROGRESULTSGRID, 0, 1) = "Value"
frmMain.ctrResultsTabs1.Data(PROGRESULTSGRID, 0, 2) = "Status"
frmMain.ctrResultsTabs1.Data(PROGRESULTSGRID, 0, 3) = "High Count"
frmMain.ctrResultsTabs1.Data(PROGRESULTSGRID, 0, 4) = "Low Count"

'Define Scan Stats grids headings
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 0) = "Parameter"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 1) = "AVG"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 2) = "STD"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 3) = "Cpk"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 4) = "Cp"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 5) = "Range High"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 6) = "Range Low"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 7) = "High Count"
frmMain.ctrResultsTabs1.Data(SCANSTATSGRID, 0, 8) = "Low Count"

'Define Programming Stats grids headings
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 0) = "Parameter"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 1) = "AVG"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 2) = "STD"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 3) = "Cpk"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 4) = "Cp"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 5) = "Range High"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 6) = "Range Low"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 7) = "High Count"
frmMain.ctrResultsTabs1.Data(PROGSTATSGRID, 0, 8) = "Low Count"

'Define the Row heights
'Total available row height is 7545, and rows will be snapped to multiples of 15.
'Total used height = NUMROWSxxxDISPLAY * TotalRowHeight + Row(0) Height

'Define the row height for all rows
frmMain.ctrResultsTabs1.TotalRowHeight(SCANRESULTSGRID) = 360
frmMain.ctrResultsTabs1.TotalRowHeight(PROGRESULTSGRID) = 360
frmMain.ctrResultsTabs1.TotalRowHeight(SCANSTATSGRID) = 360
frmMain.ctrResultsTabs1.TotalRowHeight(PROGSTATSGRID) = 360
'Re-define the row height for the top row (headings)
frmMain.ctrResultsTabs1.RowHeight(SCANRESULTSGRID, 0) = 870
frmMain.ctrResultsTabs1.RowHeight(PROGRESULTSGRID, 0) = 870
frmMain.ctrResultsTabs1.RowHeight(SCANSTATSGRID, 0) = 870
frmMain.ctrResultsTabs1.RowHeight(PROGSTATSGRID, 0) = 870

'Set the Cell Alignment properties for the Scan Results Tab
For llngRowNum = 0 To frmMain.ctrResultsTabs1.NumberOfRows(SCANRESULTSGRID) - 1
    For llngColNum = 0 To frmMain.ctrResultsTabs1.NumberOfColumns(SCANRESULTSGRID) - 1
        'Center everything in the first row
        If llngRowNum = 0 Then
            lvntAlignmentType = flexAlignCenterCenter
        Else
            'Select the alignment for the remaining rows
            Select Case llngColNum
            Case 0
                lvntAlignmentType = flexAlignLeftCenter
            Case 1
                lvntAlignmentType = flexAlignRightCenter
            Case 2, 3, 4
                lvntAlignmentType = flexAlignCenterCenter
            Case Else
                lvntAlignmentType = flexAlignCenterCenter
            End Select
        End If
        'Set the alignment
        frmMain.ctrResultsTabs1.TextAlignment(SCANRESULTSGRID, llngRowNum, llngColNum) = lvntAlignmentType
    Next llngColNum
Next llngRowNum

'Set the Cell Alignment properties for the Programming Results Tab
For llngRowNum = 0 To frmMain.ctrResultsTabs1.NumberOfRows(PROGRESULTSGRID) - 1
    For llngColNum = 0 To frmMain.ctrResultsTabs1.NumberOfColumns(PROGRESULTSGRID) - 1
        'Center everything in the first row
        If llngRowNum = 0 Then
            lvntAlignmentType = flexAlignCenterCenter
        Else
            'Select the alignment for the remaining rows
            Select Case llngColNum
            Case 0
                lvntAlignmentType = flexAlignLeftCenter
            Case 1
                lvntAlignmentType = flexAlignRightCenter
            Case Else
                lvntAlignmentType = flexAlignCenterCenter
            End Select
        End If
        'Set the alignment
        frmMain.ctrResultsTabs1.TextAlignment(PROGRESULTSGRID, llngRowNum, llngColNum) = lvntAlignmentType
    Next llngColNum
Next llngRowNum

'Set the Cell Alignment properties for the Stats Tab
For llngRowNum = 0 To frmMain.ctrResultsTabs1.NumberOfRows(SCANSTATSGRID) - 1
    For llngColNum = 0 To frmMain.ctrResultsTabs1.NumberOfColumns(SCANSTATSGRID) - 1
        'Center everything in the first row
        If llngRowNum = 0 Then
            lvntAlignmentType = flexAlignCenterCenter
        Else
            'Select the alignment for the remaining rows
            Select Case llngColNum
            Case 0
                lvntAlignmentType = flexAlignLeftCenter
            Case 1, 2, 3, 4, 5, 6, 7, 8
                lvntAlignmentType = flexAlignCenterCenter
            Case Else
                lvntAlignmentType = flexAlignCenterCenter
            End Select
        End If
        'Set the alignment
        frmMain.ctrResultsTabs1.TextAlignment(SCANSTATSGRID, llngRowNum, llngColNum) = lvntAlignmentType
    Next llngColNum
Next llngRowNum

'Set the Cell Alignment properties for the Stats Tab
For llngRowNum = 0 To frmMain.ctrResultsTabs1.NumberOfRows(PROGSTATSGRID) - 1
    For llngColNum = 0 To frmMain.ctrResultsTabs1.NumberOfColumns(PROGSTATSGRID) - 1
        'Center everything in the first row
        If llngRowNum = 0 Then
            lvntAlignmentType = flexAlignCenterCenter
        Else
            'Select the alignment for the remaining rows
            Select Case llngColNum
            Case 0
                lvntAlignmentType = flexAlignLeftCenter
            Case 1, 2, 3, 4, 5, 6
                lvntAlignmentType = flexAlignCenterCenter
            Case Else
                lvntAlignmentType = flexAlignCenterCenter
            End Select
        End If
        'Set the alignment
        frmMain.ctrResultsTabs1.TextAlignment(PROGSTATSGRID, llngRowNum, llngColNum) = lvntAlignmentType
    Next llngColNum
Next llngRowNum

'Display the parameter names
Call DisplayScanResultsNames
Call DisplayProgResultsNames
Call DisplayScanStatisticsNames
Call DisplayProgStatisticsNames

'Initialize part status display captions
frmMain.ctrStatus1.StatusCaption(1) = "Program"
frmMain.ctrStatus1.StatusCaption(2) = "Scan"
frmMain.ctrStatus1.StatusCaption(3) = ""
frmMain.ctrStatus1.StatusCaption(4) = ""

'Initialize part status display Off Color   'Default to gray
frmMain.ctrStatus1.StatusOffColor(1) = &HC0C0C0
frmMain.ctrStatus1.StatusOffColor(2) = &HC0C0C0
frmMain.ctrStatus1.StatusOffColor(3) = &HC0C0C0
frmMain.ctrStatus1.StatusOffColor(4) = &HC0C0C0

'Initialize part status display Off Text
frmMain.ctrStatus1.StatusOffText(1) = "No Part"
frmMain.ctrStatus1.StatusOffText(2) = "No Part"
frmMain.ctrStatus1.StatusOffText(3) = "No Part"
frmMain.ctrStatus1.StatusOffText(4) = "No Part"

'Initialize part status display On Color    'Default to gray
frmMain.ctrStatus1.StatusOnColor(1) = &HC0C0C0
frmMain.ctrStatus1.StatusOnColor(2) = &HC0C0C0
frmMain.ctrStatus1.StatusOnColor(3) = &HC0C0C0
frmMain.ctrStatus1.StatusOnColor(4) = &HC0C0C0

'Initialize part status display On Text     'Default to no text
frmMain.ctrStatus1.StatusOnText(1) = ""
frmMain.ctrStatus1.StatusOnText(2) = ""
frmMain.ctrStatus1.StatusOnText(3) = ""
frmMain.ctrStatus1.StatusOnText(4) = ""

'Initialize the Status Font Size
frmMain.ctrStatus1.StatusFont(1) = "Arial"
frmMain.ctrStatus1.StatusFont(2) = "Arial"
frmMain.ctrStatus1.StatusFont(3) = "Arial"
frmMain.ctrStatus1.StatusFont(4) = "Arial"

'Initialize the Status Font Size
frmMain.ctrStatus1.StatusFontSize(1) = 12
frmMain.ctrStatus1.StatusFontSize(2) = 12
frmMain.ctrStatus1.StatusFontSize(3) = 12
frmMain.ctrStatus1.StatusFontSize(4) = 12

'Initialize the Status Font to Bold
frmMain.ctrStatus1.StatusFontBold(1) = True
frmMain.ctrStatus1.StatusFontBold(2) = True
frmMain.ctrStatus1.StatusFontBold(3) = True
frmMain.ctrStatus1.StatusFontBold(4) = True

'Initialize part status display visiblity   'Show two
frmMain.ctrStatus1.StatusVisible(1) = True  'Program
frmMain.ctrStatus1.StatusVisible(2) = True  'Scan
frmMain.ctrStatus1.StatusVisible(3) = False
frmMain.ctrStatus1.StatusVisible(4) = False

'Initialize the Status Value                'All off
frmMain.ctrStatus1.StatusValue(1) = False
frmMain.ctrStatus1.StatusValue(2) = False
frmMain.ctrStatus1.StatusValue(3) = False
frmMain.ctrStatus1.StatusValue(4) = False

End Sub

Public Sub DisplayProgResultsCountsPrioritized()
'
'   PURPOSE: To display the failure counts to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

Dim lintProgrammerNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lvntHighCount(1 To NUMROWSPROGRESULTSDISPLAY) As Variant
Dim lvntLowCount(1 To NUMROWSPROGRESULTSDISPLAY) As Variant

'Row numbers:
'1 = Output #1 Label
'2 = Final Index 1 (Idle), output 1
'3 = Final Index 2 (WOT), output 1
'4 = Final Clamp Low Value, output 1
'5 = Final Clamp High Value, output 1
'6 = Offset Code, output 1
'7 = Rough Gain Code, output 1
'8 = Fine Gain Code, output 1
'9 = Clamp Low Code, output 1
'10 = Clamp High Code, output 1
'11 = AGND Code, output 1
'12 = Oscillator Adjust Code, output 1
'13 = Capacitor Frequency Adjust Code, output 1
'14 = DAC Frequency Adjust Code, output 1
'15 = Slow Mode Code, output 1
'16 = Output #2 Label
'17 = Final Index 1 (Idle), output 2
'18 = Final Index 2 (WOT), output 2
'19 = Final Clamp Value, output 2
'20 = Final Clamp Value, output 2
'21 = Offset Code, output 2
'22 = Rough Gain Code, output 2
'23 = Fine Gain Code, output 2
'24 = Clamp Low Code, output 2
'25 = Clamp High Code, output 2
'26 = AGND Code, output 2
'27 = Oscillator Adjust Code, output 2
'28 = Capacitor Frequency Adjust Code, output 2
'29 = DAC Frequency Adjust Code, output 2
'30 = Slow Mode Code, output 2

lintRow = 1   'Initialize the Row Number

'Loop through the two programmers
For lintProgrammerNum = 1 To 2

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (Idle)
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low

    lintRow = lintRow + 1

    'Index 2 (WOT) Output
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low

    lintRow = lintRow + 1

    'Low Clamp
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).clampLow.failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).clampLow.failCount.low

    lintRow = lintRow + 1

    'High Clamp
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).clampHigh.failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).clampHigh.failCount.low

    lintRow = lintRow + 1

    'Offset Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Rough Gain Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Fine Gain Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Clamp Low Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Clamp High Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'AGND Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Oscillator Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Capacitor Frequency Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'DAC Frequency Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Slow Mode Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

'2.0ANM
'    'Offset Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1
'
'    'Rough Gain Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1
'
'    'Fine Gain Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1

Next lintProgrammerNum

'Back up one row
lintRow = lintRow - 1

'Send the results to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateResultsCounts(PROGRESULTSGRID, llngRowNum, lvntHighCount(llngRowNum), lvntLowCount(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayProgResultsData()

'
'   PURPOSE: To display the programming results data to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

Dim lintProgrammerNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lstrValueAndLocation(1 To NUMROWSPROGRESULTSDISPLAY) As String
Dim lprdParameterResults(1 To NUMROWSPROGRESULTSDISPLAY) As ParameterResultsDisplay

'Default all parameters to REJECT
For llngRowNum = 1 To NUMROWSPROGRESULTSDISPLAY
    lprdParameterResults(llngRowNum) = prdReject
Next llngRowNum

'Row numbers:
'1 = Output #1 Label
'2 = Final Index 1 (Idle), output 1
'3 = Final Index 2 (WOT), output 1
'4 = Final Clamp Low Value, output 1
'5 = Final Clamp High Value, output 1
'6 = Offset Code, output 1
'7 = Rough Gain Code, output 1
'8 = Fine Gain Code, output 1
'9 = Clamp Low Code, output 1
'10 = Clamp High Code, output 1
'11 = AGND Code, output 1
'12 = Oscillator Adjust Code, output 1
'13 = Capacitor Frequency Adjust Code, output 1
'14 = DAC Frequency Adjust Code, output 1
'15 = Slow Mode Code, output 1
'16 = Output #2 Label
'17 = Final Index 1 (Idle), output 2
'18 = Final Index 2 (WOT), output 2
'19 = Final Clamp Value, output 2
'20 = Final Clamp Value, output 2
'21 = Offset Code, output 2
'22 = Rough Gain Code, output 2
'23 = Fine Gain Code, output 2
'24 = Clamp Low Code, output 2
'25 = Clamp High Code, output 2
'26 = AGND Code, output 2
'27 = Oscillator Adjust Code, output 2
'28 = Capacitor Frequency Adjust Code, output 2
'29 = DAC Frequency Adjust Code, output 2
'30 = Slow Mode Code, output 2

lintRow = 1   'Initialize the Row Number

For lintProgrammerNum = 1 To 2

    'Output # Label
    lprdParameterResults(lintRow) = prdEmpty

    lintRow = lintRow + 1

    'Final Index 1 (Idle)
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalIndexVal(1), "##0.00") & "% at " & Format(gudtSolver(lintProgrammerNum).FinalIndexLoc(1), "##0.00") & "° "
    If Not (gintProgFailure(lintProgrammerNum, HIGHPROGINDEX1) Or (gintProgFailure(lintProgrammerNum, LOWPROGINDEX1))) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Final Index 2 (WOT)
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalIndexVal(2), "##0.00") & "% at " & Format(gudtSolver(lintProgrammerNum).FinalIndexLoc(2), "##0.00") & "° "
    If Not (gintProgFailure(lintProgrammerNum, HIGHPROGINDEX2) Or (gintProgFailure(lintProgrammerNum, LOWPROGINDEX2))) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Final Clamp Low Value
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalClampLowVal, "##0.00") & "% "
    If Not (gintProgFailure(lintProgrammerNum, HIGHCLAMPLOW) Or (gintProgFailure(lintProgrammerNum, LOWCLAMPLOW))) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Final Clamp High Value
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalClampHighVal, "##0.00") & "% "
    If Not (gintProgFailure(lintProgrammerNum, HIGHCLAMPHIGH) Or (gintProgFailure(lintProgrammerNum, LOWCLAMPHIGH))) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Final Offset Code
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalOffsetCode, "##0") & " "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'Final Rough Gain Code
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalRGCode, "##0") & " "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'Final Fine Gain Code
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalFGCode, "##0") & " "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'Clamp Low Code
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalClampLowCode, "##0") & " "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'Clamp High Code
    lstrValueAndLocation(lintRow) = Format(gudtSolver(lintProgrammerNum).FinalClampHighCode, "##0") & " "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'AGND Code
    lstrValueAndLocation(lintRow) = Format(gudtMLX90277(lintProgrammerNum).Read.AGND, "##0") & " "
    If Not gintProgFailure(lintProgrammerNum, AGNDFAILURE) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Oscillator Adjust Code
    lstrValueAndLocation(lintRow) = Format(gudtMLX90277(lintProgrammerNum).Read.FCKADJ, "##0") & " "
    If Not gintProgFailure(lintProgrammerNum, FCKADJFAILURE) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Capacitor Frequency Adjust Code
    lstrValueAndLocation(lintRow) = Format(gudtMLX90277(lintProgrammerNum).Read.CKANACH, "##0") & " "
    If Not gintProgFailure(lintProgrammerNum, CKANACHFAILURE) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'DAC Frequency Adjust Code
    lstrValueAndLocation(lintRow) = Format(gudtMLX90277(lintProgrammerNum).Read.CKDACCH, "##0") & " "
    If Not gintProgFailure(lintProgrammerNum, CKDACCHFAILURE) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Slow Mode Code
    If gudtMLX90277(lintProgrammerNum).Read.SlowMode Then
        lstrValueAndLocation(lintRow) = "1 "     'Translate Boolean "True" to 1
    Else
        lstrValueAndLocation(lintRow) = "0 "     'Translate Boolean "False" to 0
    End If
    If Not gintProgFailure(lintProgrammerNum, SLOWMODEFAILURE) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

'2.0ANM \/\/
'    'Offset Seed Code
'    lstrValueAndLocation(lintRow) = "N/A" '1.9ANM Format(gudtSolver(lintProgrammerNum).OffsetSeedCode, "##0") & " "
'    lprdParameterResults(lintRow) = prdNotChecked
'
'    lintRow = lintRow + 1
'
'    'Rough Gain Seed Code
'    lstrValueAndLocation(lintRow) = "N/A" '1.9ANM Format(gudtSolver(lintProgrammerNum).RoughGainSeedCode, "##0") & " "
'    lprdParameterResults(lintRow) = prdNotChecked
'
'    lintRow = lintRow + 1
'
'    'Fine Gain Seed Code
'    lstrValueAndLocation(lintRow) = "N/A" '1.9ANM Format(gudtSolver(lintProgrammerNum).FineGainSeedCode, "##0") & " "
'    lprdParameterResults(lintRow) = prdNotChecked
'
'    lintRow = lintRow + 1

Next lintProgrammerNum

'Back up one row
lintRow = lintRow - 1

'Send the results to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateResultsData(PROGRESULTSGRID, llngRowNum, lstrValueAndLocation(llngRowNum), lprdParameterResults(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayProgResultsNames()
'
'   PURPOSE: To display the results parameter names to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

'Output #1
Call UpdateName(PROGRESULTSGRID, 1, "Output #1", True, flexAlignCenterCenter)
Call UpdateName(PROGRESULTSGRID, 2, "Final Index 1 (Idle) Output", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 3, "Final Index 2 (WOT) Output", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 4, "Final Clamp Low Value", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 5, "Final Clamp High Value", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 6, "Offset Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 7, "Rough Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 8, "Fine Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 9, "Clamp Low Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 10, "Clamp High Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 11, "AGND Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 12, "Oscillator Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 13, "Capacitor Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 14, "DAC Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 15, "Slow Code", False, flexAlignLeftCenter)
'Output #2
Call UpdateName(PROGRESULTSGRID, 16, "Output #2", True, flexAlignCenterCenter)
Call UpdateName(PROGRESULTSGRID, 17, "Final Index 1 (Idle) Output", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 18, "Final Index 2 (WOT) Output", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 19, "Final Clamp Low Value", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 20, "Final Clamp High Value", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 21, "Offset Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 22, "Rough Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 23, "Fine Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 24, "Clamp Low Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 25, "Clamp High Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 26, "AGND Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 27, "Oscillator Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 28, "Capacitor Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 29, "DAC Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGRESULTSGRID, 30, "Slow Code", False, flexAlignLeftCenter)

End Sub

Public Sub DisplayProgStatisticsCountsPrioritized()
'
'   PURPOSE: To display the programming failure counts to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

Dim lintProgrammerNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lvntHighCount(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntLowCount(1 To NUMROWSPROGSTATSDISPLAY) As Variant

'1 = Output #1 Label
'2 = Final Index 1 (Idle) Value, output 1
'3 = Final Index 1 (Idle) Location, output 1
'4 = Final Index 2 (WOT) Value, output 1
'5 = Final Index 2 (WOT) Location, output 1
'6 = Final Clamp Low Value, output 1
'7 = Final Clamp High Value, output 1
'8 = Offset Code, output 1
'9 = Rough Gain Code, output 1
'10 = Fine Gain Code, output 1
'11 = Clamp Low Code, output 1
'12 = Clamp High Code, output 1
'13 = AGND Code, output 1
'14 = Oscillator Adjust Code, output 1
'15 = Capacitor Frequency Adjust Code, output 1
'16 = DAC Frequency Adjust Code, output 1
'17 = Slow Mode Code, output 1
'18 = Output #2 Label
'19 = Final Index 1 (Idle) Value, output 2
'20 = Final Index 1 (Idle) Location, output 2
'21 = Final Index 2 (WOT) Value, output 2
'22 = Final Index 2 (WOT) Location, output 2
'23 = Final Clamp Value, output 2
'24 = Final Clamp Value, output 2
'25 = Offset Code, output 2
'26 = Rough Gain Code, output 2
'27 = Fine Gain Code, output 2
'28 = Clamp Low Code, output 2
'29 = Clamp High Code, output 2
'30 = AGND Code, output 2
'31 = Oscillator Adjust Code, output 2
'32 = Capacitor Frequency Adjust Code, output 2
'33 = DAC Frequency Adjust Code, output 2
'34 = Slow Mode Code, output 2

lintRow = 1   'Initialize the Row Number

For lintProgrammerNum = 1 To 2

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (Idle) Value
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low

    lintRow = lintRow + 1

    'Index 1 (Idle) Location
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Index 2 (WOT) Value
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low

    lintRow = lintRow + 1

    'Index 2 (WOT) Location
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Low Clamp
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).clampLow.failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).clampLow.failCount.low

    lintRow = lintRow + 1

    'High Clamp
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).clampHigh.failCount.high
    lvntLowCount(lintRow) = gudtProgStats(lintProgrammerNum).clampHigh.failCount.low

    lintRow = lintRow + 1

    'Offset Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Rough Gain Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Fine Gain Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Clamp Low Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Clamp High Code
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'AGND Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Oscillator Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Capacitor Frequency Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'DAC Frequency Adjust Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Slow Mode Code
    lvntHighCount(lintRow) = gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

'2.0ANM
'    'Offset Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1
'
'    'Rough Gain Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1
'
'    'Fine Gain Seed Code
'    lvntHighCount(lintRow) = "N/A"
'    lvntLowCount(lintRow) = "N/A"
'
'    lintRow = lintRow + 1

Next lintProgrammerNum

'Back up one row
lintRow = lintRow - 1

'Send the stats to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateStatisticsCounts(PROGSTATSGRID, llngRowNum, lvntHighCount(llngRowNum), lvntLowCount(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayProgStatisticsData()
'
'   PURPOSE: To display the programming statistics to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

Dim lintProgrammerNum As Integer
Dim lintRow As Integer
Dim llngRowNum As Long
Dim lvntAvg(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntStdDev(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntCpk(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntCp(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntRangehigh(1 To NUMROWSPROGSTATSDISPLAY) As Variant
Dim lvntRangeLow(1 To NUMROWSPROGSTATSDISPLAY) As Variant

'1 = Output #1 Label
'2 = Final Index 1 (Idle) Value, output 1
'3 = Final Index 1 (Idle) Location, output 1
'4 = Final Index 2 (WOT) Value, output 1
'5 = Final Index 2 (WOT) Location, output 1
'6 = Final Clamp Low Value, output 1
'7 = Final Clamp High Value, output 1
'8 = Offset Code, output 1
'9 = Rough Gain Code, output 1
'10 = Fine Gain Code, output 1
'11 = Clamp Low Code, output 1
'12 = Clamp High Code, output 1
'13 = AGND Code, output 1
'14 = Oscillator Adjust Code, output 1
'15 = Capacitor Frequency Adjust Code, output 1
'16 = DAC Frequency Adjust Code, output 1
'17 = Slow Mode Code, output 1
'18 = Output #2 Label
'19 = Final Index 1 (Idle) Value, output 2
'20 = Final Index 1 (Idle) Location, output 2
'21 = Final Index 2 (WOT) Value, output 2
'22 = Final Index 2 (WOT) Location, output 2
'23 = Final Clamp Value, output 2
'24 = Final Clamp Value, output 2
'25 = Offset Code, output 2
'26 = Rough Gain Code, output 2
'27 = Fine Gain Code, output 2
'28 = Clamp Low Code, output 2
'29 = Clamp High Code, output 2
'30 = AGND Code, output 2
'31 = Oscillator Adjust Code, output 2
'32 = Capacitor Frequency Adjust Code, output 2
'33 = DAC Frequency Adjust Code, output 2
'34 = Slow Mode Code, output 2

'Start at row one
lintRow = 1

For lintProgrammerNum = 1 To 2

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (Idle) Value
    If gudtProgStats(lintProgrammerNum).indexVal(1).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(1).sigma / gudtProgStats(lintProgrammerNum).indexVal(1).n, 3), "##0.000")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).indexVal(1).sigma2 - gudtProgStats(lintProgrammerNum).indexVal(1).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexVal(1).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).indexVal(1).sigma2 - gudtProgStats(lintProgrammerNum).indexVal(1).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexVal(1).n) / (gudtProgStats(lintProgrammerNum).indexVal(1).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If ((gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance) - lvntAvg(lintRow)) < (lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance)) Then
                lvntCpk(lintRow) = Format(((gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance) - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance)) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((2 * gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(1).max, 3), "##0.000")
        'Range Low
        lvntRangeLow(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(1).min, 3), "##0.000")
    End If

    lintRow = lintRow + 1

    'Index 1 (Idle) Location
    If gudtProgStats(lintProgrammerNum).indexLoc(1).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(1).sigma / gudtProgStats(lintProgrammerNum).indexLoc(1).n, 2), "##0.00")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2 - gudtProgStats(lintProgrammerNum).indexLoc(1).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexLoc(1).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2 - gudtProgStats(lintProgrammerNum).indexLoc(1).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexLoc(1).n) / (gudtProgStats(lintProgrammerNum).indexLoc(1).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(1).max, 2), "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(1).min, 2), "##0.00")
    End If

    lintRow = lintRow + 1

    'Index 2 (WOT) Value
    If gudtProgStats(lintProgrammerNum).indexVal(2).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(2).sigma / gudtProgStats(lintProgrammerNum).indexVal(2).n, 3), "##0.000")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).indexVal(2).sigma2 - gudtProgStats(lintProgrammerNum).indexVal(2).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexVal(2).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).indexVal(2).sigma2 - gudtProgStats(lintProgrammerNum).indexVal(2).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexVal(2).n) / (gudtProgStats(lintProgrammerNum).indexVal(2).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If ((gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance) - lvntAvg(lintRow)) < (lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance)) Then
                lvntCpk(lintRow) = Format(((gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance) - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance)) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((2 * gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(2).max, 3), "##0.000")
        'Range Low
        lvntRangeLow(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexVal(2).min, 3), "##0.000")
    End If

    lintRow = lintRow + 1

    'Index 2 (WOT) Location
    If gudtProgStats(lintProgrammerNum).indexLoc(2).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(2).sigma / gudtProgStats(lintProgrammerNum).indexLoc(2).n, 2), "##0.00")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2 - gudtProgStats(lintProgrammerNum).indexLoc(2).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexLoc(2).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2 - gudtProgStats(lintProgrammerNum).indexLoc(2).sigma ^ 2 / gudtProgStats(lintProgrammerNum).indexLoc(2).n) / (gudtProgStats(lintProgrammerNum).indexLoc(2).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(2).max, 2), "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(Round(gudtProgStats(lintProgrammerNum).indexLoc(2).min, 2), "##0.00")
    End If

    lintRow = lintRow + 1

    'Clamp Low
    If gudtProgStats(lintProgrammerNum).clampLow.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLow.sigma / gudtProgStats(lintProgrammerNum).clampLow.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).clampLow.sigma2 - gudtProgStats(lintProgrammerNum).clampLow.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampLow.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).clampLow.sigma2 - gudtProgStats(lintProgrammerNum).clampLow.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampLow.n) / (gudtProgStats(lintProgrammerNum).clampLow.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If ((gudtSolver(lintProgrammerNum).Clamp(1).IdealValue + gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance) - lvntAvg(lintRow)) < (lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Clamp(1).IdealValue - gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance)) Then
                lvntCpk(lintRow) = Format(((gudtSolver(lintProgrammerNum).Clamp(1).IdealValue + gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance) - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Clamp(1).IdealValue - gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance)) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((2 * gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLow.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLow.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Clamp High
    If gudtProgStats(lintProgrammerNum).clampHigh.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHigh.sigma / gudtProgStats(lintProgrammerNum).clampHigh.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).clampHigh.sigma2 - gudtProgStats(lintProgrammerNum).clampHigh.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampHigh.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).clampHigh.sigma2 - gudtProgStats(lintProgrammerNum).clampHigh.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampHigh.n) / (gudtProgStats(lintProgrammerNum).clampHigh.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If ((gudtSolver(lintProgrammerNum).Clamp(2).IdealValue + gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance) - lvntAvg(lintRow)) < (lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Clamp(2).IdealValue - gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance)) Then
                lvntCpk(lintRow) = Format(((gudtSolver(lintProgrammerNum).Clamp(2).IdealValue + gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance) - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - (gudtSolver(lintProgrammerNum).Clamp(2).IdealValue - gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance)) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((2 * gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHigh.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHigh.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Offset Code
    If gudtProgStats(lintProgrammerNum).offsetCode.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).offsetCode.sigma / gudtProgStats(lintProgrammerNum).offsetCode.n, "###0")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).offsetCode.sigma2 - gudtProgStats(lintProgrammerNum).offsetCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).offsetCode.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).offsetCode.sigma2 - gudtProgStats(lintProgrammerNum).offsetCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).offsetCode.n) / (gudtProgStats(lintProgrammerNum).offsetCode.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).offsetCode.max, "###0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).offsetCode.min, "###0")
    End If

    lintRow = lintRow + 1

    'Rough Gain Code
    If gudtProgStats(lintProgrammerNum).roughGainCode.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).roughGainCode.sigma / gudtProgStats(lintProgrammerNum).roughGainCode.n, "#0")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).roughGainCode.sigma2 - gudtProgStats(lintProgrammerNum).roughGainCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).roughGainCode.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).roughGainCode.sigma2 - gudtProgStats(lintProgrammerNum).roughGainCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).roughGainCode.n) / (gudtProgStats(lintProgrammerNum).roughGainCode.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).roughGainCode.max, "#0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).roughGainCode.min, "#0")
    End If

    lintRow = lintRow + 1

    'Fine Gain Code
    If gudtProgStats(lintProgrammerNum).fineGainCode.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).fineGainCode.sigma / gudtProgStats(lintProgrammerNum).fineGainCode.n, "###0")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).fineGainCode.sigma2 - gudtProgStats(lintProgrammerNum).fineGainCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).fineGainCode.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).fineGainCode.sigma2 - gudtProgStats(lintProgrammerNum).fineGainCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).fineGainCode.n) / (gudtProgStats(lintProgrammerNum).fineGainCode.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).fineGainCode.max, "###0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).fineGainCode.min, "###0")
    End If

    lintRow = lintRow + 1

    'Clamp Low Code
    If gudtProgStats(lintProgrammerNum).clampLowCode.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLowCode.sigma / gudtProgStats(lintProgrammerNum).clampLowCode.n, "###0")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).clampLowCode.sigma2 - gudtProgStats(lintProgrammerNum).clampLowCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampLowCode.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).clampLowCode.sigma2 - gudtProgStats(lintProgrammerNum).clampLowCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampLowCode.n) / (gudtProgStats(lintProgrammerNum).clampLowCode.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLowCode.max, "###0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampLowCode.min, "###0")
    End If

    lintRow = lintRow + 1

    'Clamp High Code
    If gudtProgStats(lintProgrammerNum).clampHighCode.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHighCode.sigma / gudtProgStats(lintProgrammerNum).clampHighCode.n, "###0")
        'Calculate Standard Deviation
        If (gudtProgStats(lintProgrammerNum).clampHighCode.sigma2 - gudtProgStats(lintProgrammerNum).clampHighCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampHighCode.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).clampHighCode.sigma2 - gudtProgStats(lintProgrammerNum).clampHighCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).clampHighCode.n) / (gudtProgStats(lintProgrammerNum).clampHighCode.n - 1)), "##0.00")
        Else
            lvntStdDev(lintRow) = "0.00"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHighCode.max, "###0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).clampHighCode.min, "###0")
    End If

    lintRow = lintRow + 1

    'AGND Code
    'No Average
    lvntAvg(lintRow) = "N/A"
    'No Standard Deviation
    lvntStdDev(lintRow) = "N/A"
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'No Range High
    lvntRangehigh(lintRow) = "N/A"
    'No Range Low
    lvntRangeLow(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Oscillator Adjust Code
    'No Average
    lvntAvg(lintRow) = "N/A"
    'No Standard Deviation
    lvntStdDev(lintRow) = "N/A"
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'No Range High
    lvntRangehigh(lintRow) = "N/A"
    'No Range Low
    lvntRangeLow(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Capacitor Frequency Adjust Code
    'No Average
    lvntAvg(lintRow) = "N/A"
    'No Standard Deviation
    lvntStdDev(lintRow) = "N/A"
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'No Range High
    lvntRangehigh(lintRow) = "N/A"
    'No Range Low
    lvntRangeLow(lintRow) = "N/A"

    lintRow = lintRow + 1

    'DAC Frequency Adjust Code
    'No Average
    lvntAvg(lintRow) = "N/A"
    'No Standard Deviation
    lvntStdDev(lintRow) = "N/A"
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'No Range High
    lvntRangehigh(lintRow) = "N/A"
    'No Range Low
    lvntRangeLow(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Slow Mode Code
    'No Average
    lvntAvg(lintRow) = "N/A"
    'No Standard Deviation
    lvntStdDev(lintRow) = "N/A"
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'No Range High
    lvntRangehigh(lintRow) = "N/A"
    'No Range Low
    lvntRangeLow(lintRow) = "N/A"

    lintRow = lintRow + 1

'2.0ANM
'    'Offset Seed Code
'    If gudtProgStats(lintProgrammerNum).OffsetSeedCode.n > 1 Then
'        'Calculate Average
'        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma / gudtProgStats(lintProgrammerNum).OffsetSeedCode.n, "###0")
'        'Calculate Standard Deviation
'        If (gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).OffsetSeedCode.n) > 0 Then
'            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).OffsetSeedCode.n) / (gudtProgStats(lintProgrammerNum).OffsetSeedCode.n - 1)), "##0.00")
'        Else
'            lvntStdDev(lintRow) = "0.00"
'        End If
'        'Cpk & CP are N/A
'        lvntCpk(lintRow) = "N/A"
'        lvntCp(lintRow) = "N/A"
'        'Range High
'        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, "###0")
'        'Range Low
'        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, "###0")
'    End If
'
'    lintRow = lintRow + 1
'
'    'Rough Gain Seed Code
'    If gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n > 1 Then
'        'Calculate Average
'        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma / gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n, "#0")
'        'Calculate Standard Deviation
'        If (gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n) > 0 Then
'            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n) / (gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n - 1)), "##0.00")
'        Else
'            lvntStdDev(lintRow) = "0.00"
'        End If
'        'Cpk & CP are N/A
'        lvntCpk(lintRow) = "N/A"
'        lvntCp(lintRow) = "N/A"
'        'Range High
'        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, "#0")
'        'Range Low
'        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, "#0")
'    End If
'
'    lintRow = lintRow + 1
'
'    'Fine Gain Seed Code
'    If gudtProgStats(lintProgrammerNum).FineGainSeedCode.n > 1 Then
'        'Calculate Average
'        lvntAvg(lintRow) = Format(gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma / gudtProgStats(lintProgrammerNum).FineGainSeedCode.n, "###0")
'        'Calculate Standard Deviation
'        If (gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).FineGainSeedCode.n) > 0 Then
'            lvntStdDev(lintRow) = Format(Sqr((gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2 - gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma ^ 2 / gudtProgStats(lintProgrammerNum).FineGainSeedCode.n) / (gudtProgStats(lintProgrammerNum).FineGainSeedCode.n - 1)), "##0.00")
'        Else
'            lvntStdDev(lintRow) = "0.00"
'        End If
'        'Cpk & CP are N/A
'        lvntCpk(lintRow) = "N/A"
'        lvntCp(lintRow) = "N/A"
'        'Range High
'        lvntRangehigh(lintRow) = Format(gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, "###0")
'        'Range Low
'        lvntRangeLow(lintRow) = Format(gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, "###0")
'    End If
'
'    lintRow = lintRow + 1

Next lintProgrammerNum

'Back up one row
lintRow = lintRow - 1

'Send the stats to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateStatisticsData(PROGSTATSGRID, llngRowNum, lvntAvg(llngRowNum), lvntStdDev(llngRowNum), lvntCpk(llngRowNum), lvntCp(llngRowNum), lvntRangehigh(llngRowNum), lvntRangeLow(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayProgStatisticsNames()
'
'   PURPOSE: To display the results parameter names to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM removed offset drift

'Output #1
Call UpdateName(PROGSTATSGRID, 1, "Output #1", True, flexAlignCenterCenter)
Call UpdateName(PROGSTATSGRID, 2, "Final Index 1 (Idle) Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 3, "Final Index 1 (Idle) Location", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 4, "Final Index 2 (WOT) Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 5, "Final Index 2 (WOT) Location", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 6, "Final Clamp Low Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 7, "Final Clamp High Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 8, "Offset Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 9, "Rough Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 10, "Fine Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 11, "Clamp Low Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 12, "Clamp High Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 13, "AGND Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 14, "Oscillator Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 15, "Capacitor Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 16, "DAC Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 17, "Slow Code", False, flexAlignLeftCenter)
'Output #2
Call UpdateName(PROGSTATSGRID, 18, "Output #2", True, flexAlignCenterCenter)
Call UpdateName(PROGSTATSGRID, 19, "Final Index 1 (Idle) Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 20, "Final Index 1 (Idle) Location", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 21, "Final Index 2 (WOT) Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 22, "Final Index 2 (WOT) Location", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 23, "Final Clamp Low Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 24, "Final Clamp High Value", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 25, "Offset Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 26, "Rough Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 27, "Fine Gain Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 28, "Clamp Low Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 29, "Clamp High Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 30, "AGND Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 31, "Oscillator Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 32, "Capacitor Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 33, "DAC Frequency Adjust Code", False, flexAlignLeftCenter)
Call UpdateName(PROGSTATSGRID, 34, "Slow Code", False, flexAlignLeftCenter)

End Sub

Public Sub DisplayProgSummary()
'
'   PURPOSE: To display the Programming Lot Summary data to a screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lsngCurrentYield As Single
Dim lsngLotYield As Single

'Calculate the Current xxx part Yield
If gudtProgSummary.currentTotal <> 0 Then
    lsngCurrentYield = gudtProgSummary.currentGood / gudtProgSummary.currentTotal
Else
    lsngCurrentYield = 0
End If

'Calculate the Lot Yield
If gudtProgSummary.totalUnits <> 0 Then
    lsngLotYield = gudtProgSummary.totalGood / gudtProgSummary.totalUnits
Else
    lsngLotYield = 0
End If
    
'Display the number of Current parts in the Current Yield Label
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbCurrentYield) = Format(gudtProgSummary.currentTotal, "0") & " Yield"
 
'Display the values in the Scan Summary Boxes
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbTotalUnits) = gudtProgSummary.totalUnits
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbGoodUnits) = gudtProgSummary.totalGood
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbRejectedUnits) = gudtProgSummary.totalUnits - gudtProgSummary.totalGood
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbSevereUnits) = gudtProgSummary.totalSevere
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbSystemErrors) = gudtProgSummary.totalNoTest
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbCurrentYield) = Format(lsngCurrentYield, "0.00%")
frmMain.ctrProgSummary.TextBoxText(SummaryTextBox.stbLotYield) = Format(lsngLotYield, "0.00%")

'Set the appropriate Background color for Current xxx Part Yield
If gudtProgSummary.currentTotal = 0 Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbWhite
ElseIf lsngCurrentYield * HUNDREDPERCENT >= gudtMachine.yieldGreen Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbGreen
ElseIf lsngCurrentYield * HUNDREDPERCENT >= gudtMachine.yieldYellow Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbYellow
Else
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbRed
End If

'Set the appropriate Background color for Lot Yield
If gudtProgSummary.totalUnits = 0 Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbWhite
ElseIf lsngLotYield * HUNDREDPERCENT >= gudtMachine.yieldGreen Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbGreen
ElseIf lsngLotYield * HUNDREDPERCENT >= gudtMachine.yieldYellow Then
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbYellow
Else
    frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbRed
End If

End Sub

Public Sub EncodeSerialNumberAndDateCode(ByVal Lot As Double, ByVal Wafer As Integer, ByVal XPos As Integer, ByVal YPos As Integer, ByVal DateCode As String, EncodedInformation() As Integer)
'
'   PURPOSE: To encode the serial number and date code for transmission to
'            the PLC.
'
'  INPUT(S): Lot          = 17-bit number for Lot ID of Serial Number
'            Wafer        = 5-bit number for Wafer of Serial Number
'            XPos         = 7-bit number for X-Position of Serial Number
'            YPos         = 7-bit number for Y-Position of Serial Number
'            DateCode     = String representing the Julian Date Code
'
' OUTPUT(S): EncodedInformation() = 5-position integer array of
'                                   encoded serial number & date code
'                                   information for transmission to PLC
'1.1ANM added station number
'1.3ANM added pallet load number

Dim llngWordArray(1 To 5) As Long
Dim lintWordNum As Integer
Dim lintJulianDate As Integer
Dim lintYear As Integer
Dim lstrShiftLetter As String
Dim lintShift As ShiftType
Dim lintStation As Integer
Dim lstrPalletLoad As String
Dim lintPallet As ShiftType

'Decode the Date Code string into it's pieces
Call MLX90277.DecodeDateCode(DateCode, lintYear, lintJulianDate, lstrShiftLetter, lintStation, lstrPalletLoad)

'Subtract the Base Year from the year variable
lintYear = lintYear - BASEYEAR

'Decode the Shift number into a letter
Select Case lstrShiftLetter
    Case "A"
        lintShift = stShiftA
    Case "B"
        lintShift = stShiftB
    Case "C"
        lintShift = stShiftC
    Case Else   'Anomalous Shift value
        lintShift = 0
End Select

'Decode the Pallet Load number into a letter
Select Case lstrPalletLoad
    Case "A"
        lintPallet = stShiftA
    Case "B"
        lintPallet = stShiftB
    Case Else   'Anomalous Shift value
        lintPallet = 0
End Select

'Generate first and second words
If Lot >= 2 ^ 17 Then               'If bit 17 is set in Lot variable, remove it in bit 6 of wafer variable
    Lot = Lot - (2 ^ 17)
    Wafer = Wafer + (2 ^ 6)
    If Lot >= 2 ^ 16 Then           'If bit 16 is set in Lot variable, remove it in bit 5 of wafer variable
        llngWordArray(1) = Lot - (2 ^ 16)
        llngWordArray(2) = Wafer + (2 ^ 5)
    Else                            'Bit 16 not set
        llngWordArray(1) = Lot
        llngWordArray(2) = Wafer
    End If
Else                                'If bit 17 isn't set, check bit 16
    If Lot >= 2 ^ 16 Then           'If bit 16 is set in Lot variable, remove it in bit 5 of wafer variable
        llngWordArray(1) = Lot - (2 ^ 16)
        llngWordArray(2) = Wafer + (2 ^ 5)
    Else                            'Bit 16 not set
        llngWordArray(1) = Lot
        llngWordArray(2) = Wafer
    End If
End If

'Generate third word
llngWordArray(3) = XPos + (YPos * (2 ^ 7))          'Place Ypos in high bits and Xpos in low bits

'Generate fourth word
llngWordArray(4) = lintJulianDate                   'Julian date is alone

'Generate fifth word by placing shift/station/pallet in high bits and year in low bits
llngWordArray(5) = lintYear + (lintShift * (2 ^ 7)) + (lintStation * (2 ^ 9)) + (lintPallet * (2 ^ 11))

'Loop through the 5 words to fill returned array (with integers)
For lintWordNum = 1 To 5
    EncodedInformation(lintWordNum) = CInt(llngWordArray(lintWordNum))
Next lintWordNum

End Sub

Public Sub ErrorLogFile(MessageText As String, DisplayMessage As Boolean, ThrowStationFault As Boolean)
'
'   PURPOSE:   To output error and message to log file for tracking purposes
'              and display an error message to the screen if requested.
'
'  INPUT(S):   MessageText       = Special text to display along with the error message
'              DisplayMessage    = Whether or not to display an error message
'              ThrowStationFault = Whether or not to try to set the PLC station fault line
'
' OUTPUT(S):
'
'****************************************************************************
'*                          SYSTEM FAULT DEFINITIONS                        *
'****************************************************************************
'
'NOTE(s):   System faults are defined via the variable, gintAnomaly.
'           To simplify identification of the various system faults
'           that can occur, a range of error codes have been identified
'           for different categories of faults.  The definitions of
'           those categories are as follows:
'
'               gintAnomaly         System Fault Category
'
'                   0               No system errors
'                 1 - 69            Misc. errors (includes supply errors)
'                70 - 99            Database Errors
'               100 - 159           DAQ errors
'               160 - 199           Programming errors
'               200 - 299           Motor errors
'               300 - 399           PT Board errors
'
'           gintAnomaly         System Fault Description
'               1               Software Run-Time Error
'               2               Supply Too High During Scan
'               3               Supply Too Low During Scan
'               4               Supply Too High Before Scan
'               5               Supply Too Low Before Scan
'               6               Error Adjusting VRef
'
'               20              Error generating Date Code
'
'               30              Error Reading PLC DDE Input
'               31              Error Reading PLC DDE Output
'               32              Error Writing PLC DDE Output
'               33              Error Verifying Program Results Code Received by PLC
'               34              Error Verifying Scan Results Code Received by PLC
'               35              Error Verifying Serial Number & Date Code Received by PLC
'
'               50              Error Communicating with Sensotec SC2000
'
'               53              Error Communicatiing with Laser Marker
'
'               70              Error Initializing Database Connection
'               71              Error Switching Database Connection
'               72              Error Saving Raw Data
'               73              Error Retrieving Raw Data
'
'               80              Error Saving Force Calibration Data
'               81              Error Saving Lot Name
'               82              Error Saving Machine Parameters
'               83              Error Saving Programming Parameters
'               84              Error Saving Scanning Parameters
'               85              Error Saving Serial Number
'               86              Error Saving Unserialized Programming Data
'               97              Error Saving Unserialized Scan Data
'               88              Error Saving Programming Results
'               89              Error Saving Scanning Results
'
'               90              Error Reading Force Calibration Data
'               91              Error Reading Lot Name
'               92              Error Reading Machine Parameters
'               93              Error Reading Programming Parameters
'               94              Error Reading Scanning Parameters
'               95              Error Reading Serial Number
'               96              Error Reading Unserialized Programming Data
'               97              Error Reading Unserialized Scan Data
'               98              Error Reading Programming Results
'               99              Error Reading Scanning Results
'
'              101              Scan watchdog timeout: Pre Scan
'              102              Scan watchdog timeout: Forward Scan
'              103              Scan watchdog timeout: Reverse Scan
'              104              Scan started in Scan Region: Pre Scan
'              105              Scan started in Scan Region: Forward Scan
'              106              Scan started in Scan Region: Reverse Scan
'              107              Unexpected Trigger Count: Pre Scan
'              108              Unexpected Trigger Count: Forward Scan
'              109              Unexpected Trigger Count: Reverse Scan
'
'              120              DAQ Error in FoutDAQSetup
'              121              DAQ Error in ReadFout
'              122              DAQ Error in OffPort1
'              123              DAQ Error in OffPort2
'              124              DAQ Error in OnPort1
'              125              DAQ Error in OnPort2
'              126              DAQ Error in ScanDAQSetup
'              127              DAQ Error in DIO1_Setup
'              128              DAQ Error in DIO2_Setup
'              129              DAQ Error in ForceSetup
'              130              DAQ Error in MonitorDAQRead
'              131              DAQ Error in MonitorDAQSetup
'              132              DAQ Error in PeakForceDAQSetup
'              133              DAQ Error in VRefDAQSetup
'              134              DAQ Error in VoutDAQSetup
'              135              DAQ Error in ReadDIOLine1
'              136              DAQ Error in ReadDIOLine2
'              137              DAQ Error in ReadPTBoardData
'              138              DAQ Error in ReadVout
'              139              DAQ Error in WritePTBoardData
'
'              151              Unable to Find Force Knee during Pre-Scan
'              152              Unable to Find Force Knee during Forward Scan
'              156              Unable to Find Kickdown Start Location
'              157              Unable to Find Force Knee
'              158              Unable to Find Full-Pedal-Travel
'              159              Kickdown Start Location at Full-Pedal-Travel
'
'              160              EvaluateTests calculations failed
'              161              MakeSolverMeasurements Error
'              162              Failure to find Pedal Face while solving
'              163              Motor movement failure while solving
'              164              Error Reading EEPROM
'              165              Error Writing to EEPROM
'              166              Reads do not match Writes after EEPROM loading
'              167              EEPROM Voting Error
'              168              Programmer Not Initialized
'              169              No Serial Number
'              170              No Date Code
'              171              EEPROM Locked; Cannot Be Re-Programmed
'              172              Kickdown Peak Force Out Of Range: Cannot Find Kickdown On Location
'              173              Part Saturated '2.0ANM
'              174              MLX Clamp Error '2.2ANM
'
'              181              MLX Lock or TC       '3.0aANM
'              182              MLX Lots don't match '3.0aANM
'
'              301              PT Board lost home signal

' Example output to file
'    DATE    |   TIME      |    SOURCE   | ERROR# |  DESCRIPTION    | USER '
' 04/20/2005 , 11:34:52 AM ,    MISC     ,  3  ,  SUPPLY TOO HIGH   ,  TLM

Dim lintFileNum As Integer          'Free file number
Dim lstrType As String              'Error Type
Dim lstrErrorDescription As String  'Error Description

If gblnUseNewAmad Then '2.6ANM
    lstrErrorDescription = Pedal.NewErrorLogFile(MessageText)
Else
    'Determine Error Source/Description
    If gintAnomaly > 0 And gintAnomaly <= 79 Then
        lstrType = "Miscellaneous"
        Select Case gintAnomaly
            Case 1
                lstrErrorDescription = "Software Run-Time Error"
            Case 2
                lstrErrorDescription = "Voltage Reference Too High During Scan"
            Case 3
                lstrErrorDescription = "Voltage Reference Too Low During Scan"
            Case 4
                lstrErrorDescription = "Voltage Reference Too High Before Scan"
            Case 5
                lstrErrorDescription = "Voltage Reference Too Low Before Scan"
            Case 6
                lstrErrorDescription = "Software error while Adjusting Voltage Reference"
            Case 20
                lstrErrorDescription = "Error Generating Date Code"
            Case 30
                lstrErrorDescription = "Error Reading PLC DDE Input"
            Case 31
                lstrErrorDescription = "Error Reading PLC DDE Output"
            Case 32
                lstrErrorDescription = "Error Writing PLC DDE Output"
            Case 33
                lstrErrorDescription = "Error Verifying Program Results Code Received by PLC"
            Case 34
                lstrErrorDescription = "Error Verifying Scan Results Code Received by PLC"
            Case 35
                lstrErrorDescription = "Error Verifying Serial Number & Date Code Received by PLC"
            Case 50
                lstrErrorDescription = "Error Communicating with SC2000"
            Case 53
                lstrErrorDescription = "Error Communicating with Laser Marker"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    ElseIf gintAnomaly > 79 And gintAnomaly <= 99 Then
        lstrType = "Database Error"
        Select Case gintAnomaly
            Case 70
                lstrErrorDescription = "Error Initializing Database Connection"
            Case 71
                lstrErrorDescription = "Error Switching Database Connection"
            Case 72
                lstrErrorDescription = "Error Saving Raw Data"
            Case 73
                lstrErrorDescription = "Error Retrieving Raw Data"
            Case 74
                lstrErrorDescription = "Error Finding Serial Number ID"
            Case 80
                lstrErrorDescription = "Error Saving Force Calibration Data"
            Case 81
                lstrErrorDescription = "Error Saving Lot Name"
            Case 82
                lstrErrorDescription = "Error Saving Machine Parameters"
            Case 83
                lstrErrorDescription = "Error Saving Programming Parameters"
            Case 84
                lstrErrorDescription = "Error Saving Scanning Parameters"
            Case 85
                lstrErrorDescription = "Error Saving Serial Number"
            Case 86
                lstrErrorDescription = "Error Saving Unserialized Programming Data"
            Case 87
                lstrErrorDescription = "Error Saving Unserialized Scan Data"
            Case 88
                lstrErrorDescription = "Error Saving Programming Results"
            Case 89
                lstrErrorDescription = "Error Saving Scanning Results"
            Case 90
                lstrErrorDescription = "Error Reading Force Calibration Data"
            Case 91
                lstrErrorDescription = "Error Reading Lot Name"
            Case 92
                lstrErrorDescription = "Error Reading Machine Parameters"
            Case 93
                lstrErrorDescription = "Error Reading Programming Parameters"
            Case 94
                lstrErrorDescription = "Error Reading Scanning Parameters"
            Case 95
                lstrErrorDescription = "Error Reading Serial Number"
            Case 96
                lstrErrorDescription = "Error Reading Unserialized Programming Data"
            Case 97
                lstrErrorDescription = "Error Reading Unserialized Scan Data"
            Case 98
                lstrErrorDescription = "Error Reading Programming Results"
            Case 99
                lstrErrorDescription = "Error Reading Scanning Results"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    ElseIf gintAnomaly > 99 And gintAnomaly <= 159 Then
        lstrType = "Data Acquisition"
        Select Case gintAnomaly
            Case 101
                lstrErrorDescription = "Scan Watchdog Timeout: Pre-Scan"
            Case 102
                lstrErrorDescription = "Scan Watchdog Timeout: Forward Scan"
            Case 103
                lstrErrorDescription = "Scan Watchdog Timeout: Reverse Scan"
            Case 104
                lstrErrorDescription = "Scan Started in Scan Region on Pre-Scan"
            Case 105
                lstrErrorDescription = "Scan Started in Scan Region on Forward Scan"
            Case 106
                lstrErrorDescription = "Scan Started in Scan Region on Reverse Scan"
            Case 107
                lstrErrorDescription = "Unexpected Trigger Count on Pre-Scan"
            Case 108
                lstrErrorDescription = "Unexpected Trigger Count on Forward Scan"
            Case 109
                lstrErrorDescription = "Unexpected Trigger Count on Reverse Scan"
            Case 120 '2.3ANM
                lstrErrorDescription = MessageText
            Case 121 '2.3ANM
                lstrErrorDescription = MessageText
            Case 122 '2.3ANM
                lstrErrorDescription = MessageText
            Case 123 '2.3ANM
                lstrErrorDescription = MessageText
            Case 124 '2.3ANM
                lstrErrorDescription = MessageText
            Case 125 '2.3ANM
                lstrErrorDescription = MessageText
            Case 126 '2.3ANM
                lstrErrorDescription = MessageText
            Case 127 '2.3ANM
                lstrErrorDescription = MessageText
            Case 128 '2.3ANM
                lstrErrorDescription = MessageText
            Case 129 '2.3ANM
                lstrErrorDescription = MessageText
            Case 130 '2.3ANM
                lstrErrorDescription = MessageText
            Case 131 '2.3ANM
                lstrErrorDescription = MessageText
            Case 132 '2.3ANM
                lstrErrorDescription = MessageText
            Case 133 '2.3ANM
                lstrErrorDescription = MessageText
            Case 134 '2.3ANM
                lstrErrorDescription = MessageText
            Case 135 '2.3ANM
                lstrErrorDescription = MessageText
            Case 136 '2.3ANM
                lstrErrorDescription = MessageText
            Case 137 '2.3ANM
                lstrErrorDescription = MessageText
            Case 138 '2.3ANM
                lstrErrorDescription = MessageText
            Case 139 '2.3ANM
                lstrErrorDescription = MessageText
            Case 151
                lstrErrorDescription = "Unable to Find Pedal Face during Pre-Scan"
            Case 152
                lstrErrorDescription = "Unable to Find Pedal Face during Forward Scan"
            Case 156
                lstrErrorDescription = "Unable to Find Kickdown Start Location"
            Case 157
                lstrErrorDescription = "Unable to Find Force Knee Location"
            Case 158
                lstrErrorDescription = "Unable to Find Full-Pedal-Travel Location"
            Case 159
                lstrErrorDescription = "Kickdown Start Location is at Find Full-Pedal-Travel Location"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    ElseIf gintAnomaly > 159 And gintAnomaly <= 199 Then
        lstrType = "Programming Error"
        Select Case gintAnomaly
            Case 160
                lstrErrorDescription = "EvaluateTests Calculations Failed"
            Case 161
                lstrErrorDescription = "MakeSolverMeasurements Failed"
            Case 162
                lstrErrorDescription = "Unable to Find Pedal Face while Solving"
            Case 163
                lstrErrorDescription = "Motor Movement Failure while Solving"
            Case 164
                lstrErrorDescription = "Error Reading EEPROM"
            Case 165
                lstrErrorDescription = "Error Writing to EEPROM"
            Case 166
                lstrErrorDescription = "Reads do not Match Writes after EEPROM Loading"
            Case 167
                lstrErrorDescription = "EEPROM Voting Error"
            Case 168
                lstrErrorDescription = "Programmer Not Initialized"
            Case 169
                lstrErrorDescription = "No Serial Number Found"
            Case 170
                lstrErrorDescription = "No Date Code Found"
            Case 171
                lstrErrorDescription = "EEPROM Locked; Cannot Be Re-Programmed"
            Case 172
                lstrErrorDescription = "Kickdown Peak Force Out Of Range: Cannot Find Kickdown On Location"
            Case 173 '2.0ANM
                lstrErrorDescription = MessageText
            Case 174 '2.2ANM
                lstrErrorDescription = MessageText
            Case 181 '3.0aANM
                lstrErrorDescription = "MLX Checks Failed!"
            Case 182 '3.0aANM
                lstrErrorDescription = "MLX Checks Failed!"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    ElseIf gintAnomaly > 199 And gintAnomaly <= 299 Then
        lstrType = "Motor Controller"
        Select Case gintAnomaly
            Case 299
                lstrErrorDescription = "Motor Drive Not Responding"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    ElseIf gintAnomaly > 299 And gintAnomaly <= 399 Then
        lstrType = "Position-Trigger Board"
        Select Case gintAnomaly
            Case 301
                lstrErrorDescription = "Position Trigger Board lost Home Signal"
            Case Else
                lstrErrorDescription = "Error or Unknown Type"
        End Select
    End If
End If

'Get free file number
lintFileNum = FreeFile

'Open the error log file for append
Open ERRORPATH & ERRORLOG For Append As #lintFileNum

'Print the details of the error to the file
Print #lintFileNum, Format$(Now, "mmmm d yyyy"); ","; Format$(Now, "h:mm:ss AM/PM"); ","; lstrType; ","; gintAnomaly; ","; MessageText; ","; frmMain.ctrSetupInfo1.Operator

'Close the error log file
Close #lintFileNum

'If a Station Fault was requested...
If ThrowStationFault Then
    If gudtMachine.PLCCommType = pctDDE Then
        'Throw a Station Fault
        Call frmDDE.WriteDDEOutput(StationFault, 1)
    ElseIf gudtMachine.PLCCommType = pctTTL Then
        'Disable the Watchdog Timer
        Call frmDAQIO.OffPort2(PORT8, BIT7)
    End If
End If

'Display the Error Message if called for
If DisplayMessage Then
    MsgBox "System Error: " & lstrType & " #" & Format(gintAnomaly, "###") _
            & vbCrLf & vbCrLf & MessageText, vbOKOnly + vbCritical, lstrErrorDescription
End If

'If a Station Fault was requested...
If ThrowStationFault Then
    If gudtMachine.PLCCommType = pctDDE Then
        'Clear the Station Fault
        Call frmDDE.WriteDDEOutput(StationFault, 0)
    ElseIf gudtMachine.PLCCommType = pctTTL Then
        'Enable the Watchdog Timer
        Call frmDAQIO.OnPort2(PORT8, BIT7)
    End If
End If

End Sub

Public Function FindKickdownPeakLocation() As Single
'
'     PURPOSE:  To determine the Kickdown Perak Location
'
'    INPUT(S):  None.
'   OUTPUT(S):  Returns Measured Kickdown Peak Location

Dim lsngFwdForce() As Single                'Calculated Forward Force data array
Dim lsngNotUsed1 As Single                  'Dummy variable
Dim lsngNotUsed2 As Single                  'Dummy variable
Dim lsngKickdownPeakLocation As Single      'Kickdown Peak Location
Dim lsngKickdownPeakForce As Single         'Kickdown Peak Force
Dim lsngHighForceLimit As Single            'Highest Allowable Kickdown Peak Force
Dim lsngLowForceLimit As Single             'Lowest Allowable Kickdown Peak Force

ReDim lsngFwdForce(gintMaxData)
    
'Create Forward Force gradient
Call Calc.CalcScaledDataArray(CHAN2, gintForward(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, VOLTSPERLSB * gsngNewtonsPerVolt, gsngForceAmplifierOffset, gsngResolution, lsngFwdForce())

'Calculate Kickdown Peak Location and Force
Call Calc.CalcMinMax(lsngFwdForce(), gudtTest(CHAN0).kickdownForceSpan.start, gudtTest(CHAN0).kickdownForceSpan.stop, gsngResolution, lsngNotUsed1, lsngNotUsed2, lsngKickdownPeakForce, lsngKickdownPeakLocation)

'Highest allowable kickdown peak force is the high force limit at ForcePt(2) + the high Kickdown Peak force limit
lsngHighForceLimit = gudtTest(CHAN0).fwdForcePt(2).high + gudtTest(CHAN0).kickdownForceSpan.high
'Lowest allowable kickdown peak force is the low force limit at ForcePt(3) + the low Kickdown Peak force limit
lsngLowForceLimit = gudtTest(CHAN0).fwdForcePt(2).low + gudtTest(CHAN0).kickdownForceSpan.low

If (lsngKickdownPeakForce > lsngHighForceLimit) Or (lsngKickdownPeakForce < lsngLowForceLimit) Then
    gintAnomaly = 172
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Programming Error in Pedal.FindKickdownPeakLocation:" & vbCrLf & _
                      "Measured Kickdown Peak Force Out of Range." & vbCrLf & _
                      "Verify Kickdown Module Installed Correctly" & vbCrLf & _
                      "and Check Force Sensing Equipment.", True, True)
    'Exit the Function
    Exit Function
End If

'Return the Kickdown Peak Location
FindKickdownPeakLocation = lsngKickdownPeakLocation

End Function

Public Function FindKickdownStartLocation() As Single
'
'     PURPOSE:  To determine the Kickdown Start Location
'
'    INPUT(S):  None.
'   OUTPUT(S):  Returns Kickdown Start Location

Dim lsngFwdForce() As Single                'Calculated Forward Force data array
Dim lsngNotUsed As Single                   'Dummy variable
Dim lsngForceKneeLocation As Single         'Force Knee Location
Dim lsngFullPedalTravelLocation As Single   'Full-Pedal-Travel Location
Dim lsngKickdownStartLocation As Single     'Kickdown Start Location
Dim lblnForceKneeFound As Boolean           'Whether or not the Force Knee was found
Dim lblnFullPedalTravelFound As Boolean     'Whether or not Full-Pedal-Travel was found
Dim lblnKDStartFound As Boolean             'Whether or not Kickdown Start was found

ReDim lsngFwdForce(gintMaxData)

'Create Forward Force gradient
Call Calc.CalcScaledDataArray(CHAN2, gintForward(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, VOLTSPERLSB * gsngNewtonsPerVolt, gsngForceAmplifierOffset, gsngResolution, lsngFwdForce())
'Calculate Force Knee Location & Force
Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.FKSlope, False, gudtMachine.FKPercentage, gudtMachine.FKWindow, gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, gsngResolution, lsngForceKneeLocation, lsngNotUsed, lblnForceKneeFound)
If gintAnomaly Then Exit Function           'Exit on system error
If Not lblnForceKneeFound Then
    gintAnomaly = 157
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Force Knee Not Found: Verify that force cell and amplifier" & vbCrLf & _
                      "                      are functioning properly and the pedal" & vbCrLf & _
                      "                      is properly clamped in the fixture.", True, True)
    'Exit the Function
    Exit Function
End If

'1.6ANM 'Calculate Full-Pedal-Travel Location & Force
'1.6ANM 'NOTE: Search from FPT low limit to evaluate stop avoid finding kickdown or causing errors
'1.6ANM Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.FPTSlope, True, gudtMachine.FPTPercentage, gudtMachine.FPTWindow, gudtTest(CHAN0).fullPedalTravel.low, gudtTest(CHAN0).evaluate.stop, gsngResolution, lsngFullPedalTravelLocation, lsngNotUsed, lblnFullPedalTravelFound)
'1.6ANM If gintAnomaly Then Exit Function           'Exit on system error
'1.6ANM If Not lblnFullPedalTravelFound Then
'1.6ANM     gintAnomaly = 158
'1.6ANM     'Log the error to the error log and display the error message
'1.6ANM     Call ErrorLogFile("Full-Pedal-Travel Not Found: Verify that force cell and amplifier" & vbCrLf & _
'1.6ANM                       "                             are functioning properly and the pedal" & vbCrLf & _
'1.6ANM                       "                             is properly clamped in the fixture.", True, True)
'1.6ANM     'Exit the Function
'1.6ANM     Exit Function
'1.6ANM End If

'Calculate and Check the Kickdown Start Location
'NOTE: Search from Measured Force Knee Location to Measured FPT Location
If gudtMachine.seriesID = "703" Then  '2.0ANM make one pedal.bas
    Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.KDStartSlope, True, gudtMachine.KDStartPercentage, gudtMachine.KDStartWindow, lsngForceKneeLocation, (gintMaxData / gsngResolution), gsngResolution, lsngKickdownStartLocation, lsngNotUsed, lblnKDStartFound)
Else
    Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.KDStartSlope, True, gudtMachine.KDStartPercentage, gudtMachine.KDStartWindow, lsngForceKneeLocation, gudtReading(CHAN0).fullPedalTravel.location, gsngResolution, lsngKickdownStartLocation, lsngNotUsed, lblnKDStartFound) '1.6ANM
End If
If gintAnomaly Then Exit Function           'Exit on system error
If Not lblnKDStartFound Then
    gintAnomaly = 156
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Kickdown Start Location Not Found." & vbCrLf & _
                      "Verify Kickdown Module Installed Correctly" & vbCrLf & _
                      "and Check Force Sensing Equipment.", True, True)
    'Exit the Function
    Exit Function
End If

'If KD Start = Full-Pedal-Travel, something is WRONG.  Throw a system error.
If (lsngKickdownStartLocation = lsngFullPedalTravelLocation) Then
    gintAnomaly = 159
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Kickdown Start Location is at Full-Pedal-Travel." & vbCrLf & _
                      "Verify Kickdown Module Installed Correctly" & vbCrLf & _
                      "and Check Force Sensing Equipment.", True, True)
    'Exit the Function
    Exit Function
End If

'Return the Kickdown Start Location
FindKickdownStartLocation = lsngKickdownStartLocation

End Function

Public Function FindPedalFace() As Single
'
'   PURPOSE: To find the pedal face relative to the encoder zero
'
'  INPUT(S): none
' OUTPUT(S): returns whether or not the pedal face was found
'

Dim lintDataPoint As Integer
Dim lsngForce As Single
Dim lblnPedalFaceFound As Boolean

'Initialize the routine to return false
FindPedalFace = False

'Prescan pedal and look for a force change
Call ScanForPedalAtRestLocation

If gintAnomaly Then Exit Function

'Initialize the boolean
lblnPedalFaceFound = False

'Look for gudtMachine.pedalAtRestLocForce Newtons
For lintDataPoint = 1 To gintMaxData - 1
    lsngForce = (CSng(gintPreScanForce(lintDataPoint)) * gsngNewtonsPerVolt * VOLTSPERLSB) + gsngForceAmplifierOffset
    If lsngForce > gudtMachine.pedalAtRestLocForce Then
        'We found the position of the Pedal Face
        gudtReading(CHAN0).pedalFaceLoc = gudtMachine.preScanStart + ((lintDataPoint - 1) / gsngResolution)
        'Report that the Pedal Face was found
        FindPedalFace = True
        Exit For
    End If
Next lintDataPoint

End Function

Public Function FindSecondProgrammingPosition() As Single
'
'   PURPOSE: To find the second programming position:
'            - Ideal Index 2 Location for Non-Kickdown Parts
'            - Ideal Kickdown On Location for Kickdown Parts
'
'  INPUT(S): None.
' OUTPUT(S): Returns the location found to program at
'1.6ANM adjusted for proper calls

If Not gudtMachine.kickdown Then
    'Search for the Pedal Face
    If FindPedalFace Then
        Call FindStartScan
        
        'Return the WOT location relative to encoder zero
        FindSecondProgrammingPosition = gudtReading(CHAN0).pedalFaceLoc + gudtSolver(1).Index(2).IdealLocation
    Else
        gintAnomaly = 162
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Unable to Find Pedal Face on Solver Pre-Scan." & vbCrLf & _
                          "Check Force Sensing Equipment.", True, True)
        'Exit on System Error
        Exit Function
    End If
Else
    'Search for the Pedal Face
    If FindPedalFace Then
        Call FindStartScan
        
        'De-Activate the Sensotec SC2000's Tare Function
        If Sensotec.GetLinkStatus Then
            Call Sensotec.DeActivateTare(1)
        Else
            'Error Communicating with the Sensotec SC2000
            gintAnomaly = 50
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
        End If
    
        'Continue testing if no errors
        If gintAnomaly = 0 Then
    
            'Move to the found Start Scan Position - OverTravel
            Call MoveToPosition((gudtMachine.scanStart - gudtMachine.overTravel), 2)
    
            'Activate the Sensotec SC2000's Tare function to remove any offset
            If Sensotec.GetLinkStatus Then
                Call Sensotec.ActivateTare(1)
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If
    
            'Scan the DUT to find the Peak Force Location
            If (gintAnomaly = 0) Then Call ScanForwardOnly
    
            'De-Activate the Sensotec SC2000's Tare Function
            If Sensotec.GetLinkStatus Then
                Call Sensotec.DeActivateTare(1)
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If
    
            'Find the Pedal At Rest Location and Manage the Data Arrays Accordingly
            Call FindPedalZeroAndTruncateData(True)
    
            'Return the Kickdown On Location relative to encoder zero
            FindSecondProgrammingPosition = CalculateKickdownOnLocation + gudtReading(CHAN0).pedalFaceLoc
        End If
    Else
        gintAnomaly = 162
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Unable to Find Pedal Face on Solver Pre-Scan." & vbCrLf & _
                          "Check Force Sensing Equipment.", True, True)
        'Exit on System Error
        Exit Function
    End If
End If

End Function

Public Sub FindStartScan()
'
'   PURPOSE: To find Start Scan
'
'  INPUT(S): none
' OUTPUT(S): none

'If we already programmed the part, calculate start scan   '1.2ANM \/\/
If gblnProgramStart And (Not gblnLockSkip) Then            '2.8ANM
    gudtReading(CHAN0).pedalAtRestLoc = gudtMachine.blockOffset - gudtReading(CHAN0).pedalFaceLoc
    'Calculate ScanStart and ScanEnd based on direction of pre-scan
    If gudtMachine.preScanStart < gudtMachine.preScanStop Then
        gudtMachine.scanStart = gudtReading(CHAN0).pedalFaceLoc - gudtMachine.offset4StartScan
        gudtMachine.scanEnd = gudtMachine.scanStart + gudtMachine.scanLength
    Else
        gudtMachine.scanStart = gudtReading(CHAN0).pedalFaceLoc + gudtMachine.offset4StartScan
        gudtMachine.scanEnd = gudtMachine.scanStart - gudtMachine.scanLength
    End If
'If we didn't program the part, search for the pedal face
ElseIf FindPedalFace Then
    gudtReading(CHAN0).pedalAtRestLoc = gudtMachine.blockOffset - gudtReading(CHAN0).pedalFaceLoc
    'Calculate ScanStart and ScanEnd based on direction of pre-scan
    If gudtMachine.preScanStart < gudtMachine.preScanStop Then
        gudtMachine.scanStart = gudtReading(CHAN0).pedalFaceLoc - gudtMachine.offset4StartScan
        gudtMachine.scanEnd = gudtMachine.scanStart + gudtMachine.scanLength
    Else
        gudtMachine.scanStart = gudtReading(CHAN0).pedalFaceLoc + gudtMachine.offset4StartScan
        gudtMachine.scanEnd = gudtMachine.scanStart - gudtMachine.scanLength
    End If
'If the pedal face isn't found, report the error
Else
    gintAnomaly = 151
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Unable to Find Pedal Face on Pre-Scan." & vbCrLf & _
                      "Check Force Sensing Equipment.", True, True)
End If                                                     '1.2ANM /\/\

End Sub

Public Sub FindEndScan()
'
'   PURPOSE: To find End Scan
'
'  INPUT(S): none
' OUTPUT(S): none
'1.4ANM new sub

Dim lsngScanWatchDogTimer As Single
Dim lsngScanWatchDogTimer2 As Single
Dim lblnTimeOut As Boolean
Dim lintStep As Integer

'Set the Scan Velocity
Call VIX500IE.SetVelocity(0.04)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(0.03)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(0.5)
'Set the Direction
Call VIX500IE.SetDirection(dtClockwise)
'Set the Servo Mode
Call VIX500IE.SetServoMode(mtModeContinuous)
'Delay so motor can catch all commands
Call frmDAQIO.KillTime(60)

'Initialize watchdog timer
lsngScanWatchDogTimer = Timer

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for force to be found, or for a timeout condition
Do
    'Reset the FOut control
    frmDAQIO.cwaiFOut.Reset
    'Read the force channel
    Call frmDAQIO.ReadFout(gsngEndForce)
    'See if we have waited too long
    lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until (gsngEndForce > gudtMachine.EndScanStopForce) Or lblnTimeOut

'Stop the motor
Call VIX500IE.StopMotor

If gintAnomaly = 0 Then
    'Read the end scan position
    gsngEndPos = Position
    'Set FPT
    gudtReading(CHAN0).fullPedalTravel.location = gsngEndPos - gudtReading(CHAN0).pedalFaceLoc
    Call frmDAQIO.ReadVout(gudtReading(CHAN0).fullPedalTravel.Value, gudtReading(CHAN1).fullPedalTravel.Value)
End If

'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)
'Delay so motor can catch all commands      '1.7ANM
Call frmDAQIO.KillTime(100)
'Set the Servo Mode
Call VIX500IE.SetServoMode(mtModeAbsolute)  '1.6ANM

End Sub

Public Sub FindPedalZeroAndTruncateData(ForwardOnly As Boolean)
'
'   PURPOSE: To find where the pedal face exists in the forward scan data,
'            then truncate all data prior to the pedal-at-rest location.
'            The pedal-at-rest location is the first data point before
'            machine.pedalAtRestLocForce is found.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lblnpedalAtRestLocationFound As Boolean
Dim lintDataPoint As Integer
Dim lsngForce As Single
Dim lintCountsToOffset As Integer
Dim lintIndex As Integer

'Initialize the boolean
lblnpedalAtRestLocationFound = False

'Look for gudtMachine.pedalAtRestLocForce Newtons
For lintDataPoint = 1 To gintMaxData - 1
    lsngForce = (CSng(gintForward(CHAN2, lintDataPoint)) * gsngNewtonsPerVolt * VOLTSPERLSB) + gsngForceAmplifierOffset
    If lsngForce > gudtMachine.pedalAtRestLocForce Then
        'Angle which pedal was found in reference from Datum Zero of Print
        gudtReading(CHAN0).pedalAtRestLoc = gudtMachine.blockOffset - (gudtMachine.scanStart + ((lintDataPoint - 1) / gsngResolution))
        'Define the number of data points to eliminate from the beginning of the scan
        lintCountsToOffset = lintDataPoint - 1
        lblnpedalAtRestLocationFound = True
        Exit For
    End If
Next lintDataPoint

'If the Pedal-At-Rest Location was not found, we have no reference, and cannot
'evaluate the scan data
If lblnpedalAtRestLocationFound = False Then
    gintAnomaly = 152
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Unable to Find Pedal Face on Forward Scan." & vbCrLf & _
                      "Check Force Sensing Equipment.", True, True)
Else
    If gblnForceOnly Then   '1.7ANM added if block
        'Truncate data based on where the pedal-at-rest location is found and
        'fill the rest with zero's
        For lintDataPoint = lintCountsToOffset To gintMaxData + lintCountsToOffset - 1
            lintIndex = lintDataPoint - lintCountsToOffset
            If lintDataPoint >= (gintMaxData - 1) Then
                'Force
                gintForward(CHAN2, lintIndex) = 0
                If Not ForwardOnly Then
                    'Force
                    gintReverse(CHAN2, lintIndex) = 0
                End If
            Else
                'Force
                gintForward(CHAN2, lintIndex) = gintForward(CHAN2, lintDataPoint)
                If Not ForwardOnly Then
                    'Force
                    gintReverse(CHAN2, lintIndex) = gintReverse(CHAN2, lintDataPoint)
                End If
            End If
        Next lintDataPoint
    Else
        'Truncate data based on where the pedal-at-rest location is found and
        'fill the rest with zero's
        For lintDataPoint = lintCountsToOffset To gintMaxData + lintCountsToOffset - 1
            lintIndex = lintDataPoint - lintCountsToOffset
            If lintDataPoint >= (gintMaxData - 1) Then
                'Output #1
                gintForward(CHAN0, lintIndex) = 0
                'Output #2
                gintForward(CHAN1, lintIndex) = 0
                'Force
                gintForward(CHAN2, lintIndex) = 0
                If Not ForwardOnly Then
                    'Output #1
                    gintReverse(CHAN0, lintIndex) = 0
                    'Output #2
                    gintReverse(CHAN1, lintIndex) = 0
                    'Force
                    gintReverse(CHAN2, lintIndex) = 0
                End If
            Else
                'Output #1
                gintForward(CHAN0, lintIndex) = gintForward(CHAN0, lintDataPoint)
                'Ouptut #2
                gintForward(CHAN1, lintIndex) = gintForward(CHAN1, lintDataPoint)
                'Force
                gintForward(CHAN2, lintIndex) = gintForward(CHAN2, lintDataPoint)
                If Not ForwardOnly Then
                    'Output #1
                    gintReverse(CHAN0, lintIndex) = gintReverse(CHAN0, lintDataPoint)
                    'Output #2
                    gintReverse(CHAN1, lintIndex) = gintReverse(CHAN1, lintDataPoint)
                    'Force
                    gintReverse(CHAN2, lintIndex) = gintReverse(CHAN2, lintDataPoint)
                End If
            End If
        Next lintDataPoint
    End If
End If

End Sub

Public Function GetDateCode() As String
'
'   PURPOSE: To make the date code and return it.
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lstrDateCode As String
Dim lintYear As Integer
Dim lintJulianDate As Integer
Dim lstrShiftLetter As String
Dim lintStation As Integer

On Error GoTo DateCodeError

'********** Date Code Format **********
'       XX              XXX        XX       XX
'Year beyond 2000 | Julian Date | Shift | Station

'Initialize the Date Code
lstrDateCode = ""

'Calculate how far beyond Year 2000 we are
lintYear = Year(Now) - BASEYEAR

'Get the Julian Date
lintJulianDate = CalcDayOfYear(DateTime.Month(DateTime.Now), DateTime.Day(DateTime.Now), IsLeapYear(DateTime.Year(DateTime.Now)))

'Get the shift
lstrShiftLetter = frmMain.GetShiftLetter '2.8ANM

'If GetShiftLetter did not return a letter, set an anomaly
If lstrShiftLetter = "" Then
    gintAnomaly = 20
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Error Generating Shift Letter for Date Code!", True, True)
End If

'Get the station number
lintStation = CInt(right(gstrSystemName, 1))

'Concatenate the Year, Date, and Shift
lstrDateCode = Format(lintYear, "00") & Format(lintJulianDate, "000") & lstrShiftLetter & CStr(lintStation)

'Return the date code
GetDateCode = lstrDateCode

Exit Function
DateCodeError:
    gintAnomaly = 20
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in Pedal.GetDateCode: " & Err.Description, True, True)
End Function

'2.8ANM moved to frmMain
'Public Function GetShiftLetter() As String
''
''   PURPOSE: To build the shift letter based on the time of day.
''
''  INPUT(S): None
'' OUTPUT(S): returns the Shift Letter (String)
'
'Dim lintHours As Integer
'Dim lintMinutes As Integer
'Dim lintMinuteOfDay As Integer
'Dim lstrShift As String
'
''Determine the minute of the day (x/1439)
'lintHours = DateTime.Hour(DateTime.Now)         'Get the hour
'lintMinutes = DateTime.Minute(DateTime.Now)     'Get the minutes
'lintMinuteOfDay = lintHours * 60 + lintMinutes  'Calculate the current minute of the day
'
''Select the shift based on what minute of the day it is
'Select Case lintMinuteOfDay
'    '12:00AM to 6:59AM, 11:00PM to 11:59PM
'    Case 0 To 419, 1380 To 1439
'        lstrShift = "C"
'    '7:00AM to 2:59PM
'    Case 420 To 899
'        lstrShift = "A"
'    '3:00PM to 10:59PM
'    Case 900 To 1379
'        lstrShift = "B"
'End Select
'
''Return the selected shift letter
'GetShiftLetter = lstrShift
'
'End Function

Public Sub HomeMotor()
'
'   PURPOSE: To initiate a home of the motor to the z-channel of the encoder
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintMotorErr As Integer

On Error GoTo HomeMotorError

'Initialize the System Anomaly to zero
gintAnomaly = 0

frmMain.staMessage.Panels(1).Text = "System Message:  Initializing Servo Motor ... Please Wait."

'Set the controller number = 1
Call VIX500IE.SetControllerNumber(1)

'Establish RS-232 communication with the motor
Call VIX500IE.InitializeCommunication(VIX500IEPORT)

'Check the status of the motor before the home move
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200        'Convert to system fault code
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Motor Error During Homing: Please Cycle Power to the Motor" & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, False)
    'Exit on System Error
    Exit Sub
End If

'Energize the motor
Call VIX500IE.EnergizeMotor

'Delay 100msec
Call frmDAQIO.KillTime(100)

'NOTE: Pedal Scanners are the exception to the rotary rule:
'      They do use limit proxes...  typically, this line
'      would disable limits for a rotary scanner
Call VIX500IE.EnableLimits

'Delay 100msec
Call frmDAQIO.KillTime(100)

'Let the motor controller know the motor's encoder resolution
Call VIX500IE.SetStepsPerRev(MOTORSTEPSPERREV)

'Delay 100msec
Call frmDAQIO.KillTime(100)

'Set the Gear Ratio
Call VIX500IE.SetGearRatio(gudtMachine.gearRatio)

'Set the PID
Call VIX500IE.SetPIDParameters(25, 35, 30, 5, 25, 10, 10, 10, 0)

'Send Reset to microprocessor for ZFind
Call frmDAQIO.OffPort1(PORT2, BIT2)     'Set Reset line low (RESET)
frmDAQIO.KillTime (100)                 'Delay 100 msec
Call frmDAQIO.OnPort1(PORT2, BIT2)      'Set Reset line high
frmDAQIO.KillTime (100)                 'Delay 100 msec

'Set the PT Board to HomeMode
If Not WritePTBoardHomeMode(True) Then
    gintAnomaly = 7
    frmMain.staMessage.Panels(1).Text = "System Message:  Home Incomplete; Position Trigger Board Problem."
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Error Homing Motor: Position-Trigger Board not responding." & vbCrLf & _
                      "Please cycle power to the Digitizer and Home the motor using the Function Menu", True, False)
    'Exit on System Error
    Exit Sub
Else
    MsgBox "Verify that drive arm can move to Home and click OK.", vbOKOnly, "Homing Motor"
End If
'Home Routine for Rotary Systems
If HomeRotarySystem Then

    frmMain.staMessage.Panels(1).Text = "System Message:  Home Complete; Moving to Part Load/Unload Position ... Please Wait."

    'Set servo to absolute position mode
    'NOTE: HomeMotor is the only routine in which continuous mode is used...
    '      Throughout the rest of the code, it is assumed that the motor is in
    '      Absolute Position mode.
    Call VIX500IE.SetServoMode(mtModeAbsolute)

    'Move the drive arm to the Part Load/Unload Location
    If gintAnomaly = 0 Then Call MoveToLoadLocation

    'Let the PLC know that initialization is complete
    If gintAnomaly = 0 Then
        If gudtMachine.PLCCommType = pctDDE Then
            Call frmDDE.WriteDDEOutput(ScannerInit, 1)
        ElseIf gudtMachine.PLCCommType = pctTTL Then
            Call frmDAQIO.OffPort2(PORT8, BIT5)
        End If
    Else
        MsgBox "There was an error Homing the Motor." & vbCrLf & _
               "NOTE:  You will be unable to scan.", _
               vbOKOnly + vbCritical, "HomeMotor Error"
    End If

Else
    'Don't allow further motor movements
    Call VIX500IE.SetLinkStatus(False)
    'Display an error message to the user
    MsgBox "There was an error Homing the Motor." & _
           "NOTE:  You will be unable to run the motor.", _
           vbOKOnly + vbCritical, "HomeMotor Error"
    frmMain.staMessage.Panels(1).Text = "System Message:  Home Incomplete; System must be re-homed!"
End If

'Take the PT Board out of HomeMode
Call WritePTBoardHomeMode(False)

Exit Sub
HomeMotorError:

    'Don't allow further motor movements
    Call VIX500IE.SetLinkStatus(False)
    MsgBox "There was an error Homing the Motor." & _
           "NOTE:  You will be unable to run the motor.", _
           vbOKOnly + vbCritical, "HomeMotor Error"
    frmMain.staMessage.Panels(1).Text = "System Message:  Home Incomplete; System must be re-homed!"
End Sub

Public Function HomeRotarySystem() As Boolean
'
'   PURPOSE: To initiate a home sequence for a rotary system
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lsngFirstPosition As Variant
Dim lsngPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngStartTimer As Single
Dim lblnMoveDone As Boolean
Dim lblnTimeOut As Boolean

'Initialize the routine to return an unsuccessful completion
HomeRotarySystem = False

'Set servo motor into continuous move mode
Call VIX500IE.SetServoMode(mtModeContinuous)
'Set scan arm Velocity to 0.05 rps
Call VIX500IE.SetVelocity(0.05)
'Set scan arm Acceleration to 0.25 rps²
Call VIX500IE.SetAcceleration(0.25)
'Set scan arm Deceleration to 0.55 rps²
Call VIX500IE.SetDeceleration(0.5)

frmMain.staMessage.Panels(1).Text = "System Message:  Homing Servo Motor ... Please Wait."

'Set the direction of the motor for Homing
If gudtMachine.gearRatio > 0 Then
    'CCW direction move to home
    Call VIX500IE.SetDirection(dtCounterClockwise)
Else
    'CW direction move to home
    Call VIX500IE.SetDirection(dtClockwise)
End If

If VIX500IE.GetLinkStatus Then
    'Find the Home Marker on the encoder
    If ZFind Then
        'Start the watchdog timer
        lsngStartTimer = Timer
        'Wait for the motor to stop
        Do
            lsngFirstPosition = Position
            'Delay 10 msec for movement
            Call frmDAQIO.KillTime(10)
            'Check to see if the move has completed
            lsngCurrentPosition = Position
            lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
            'Exit the loop if the motor has stopped
            If lblnMoveDone Then Exit Do
            'Check for timeout
            lblnTimeOut = (lsngStartTimer - Timer > ZFINDTIMEOUT)
        Loop Until lblnTimeOut
        If lblnMoveDone Then
            'Set servo mode to absolute position mode
            Call VIX500IE.SetServoMode(mtModeAbsolute)
            'Set the motor controller counter to zero
            Call VIX500IE.ZeroPosition
            'Get the position to find out how much we overtravelled
            lsngPosition = Position
            'Move the motor back to encoder 0° by moving the inverse of the overtravel
            Call VIX500IE.DefineMovement(-lsngPosition)
            'Set scan arm Velocity to 0.01 rps
            Call VIX500IE.SetVelocity(0.01)
            'Set scan arm Acceleration to 0.1 rps²
            Call VIX500IE.SetAcceleration(0.1)
            'Set scan arm Deceleration to 0.1 rps²
            Call VIX500IE.SetDeceleration(0.1)
            'Make the movement
            Call VIX500IE.StartMotor
            'Delay 500 mSec for motor to start moving
            Call frmDAQIO.KillTime(500)
            'Start the watchdog timer
            lsngStartTimer = Timer
            'Re-Initialize the MoveDone boolean
            lblnMoveDone = False
            'Wait for the motor to stop
            Do
                lsngFirstPosition = Position
                'Delay 10 msec for movement
                Call frmDAQIO.KillTime(10)
                'Check to see if the move has completed
                lsngCurrentPosition = Position
                lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
                'Exit the loop if the motor has stopped
                If lblnMoveDone Then Exit Do
                'Check for timeout
                lblnTimeOut = (lsngStartTimer - Timer > ZFINDTIMEOUT)
            Loop Until lblnTimeOut
            If lblnMoveDone Then
                'Set the motor controller counter to zero
                Call VIX500IE.ZeroPosition
                'Set Trigger Counter to Zero
                Call ClearCounter
                'Homing completed successfully
                HomeRotarySystem = True
            End If
        End If
    End If
Else
    MsgBox "Communication with Motor Controller Lost"
End If

End Function

Public Sub InitializeAndMaskProgFailures()
'
'   PURPOSE: To initialize all programming failures to failed and mask
'            out failures which are not checked.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintProgrammerNum As Integer
Dim lintFailureNum As Integer

'Initialize Boolean
gblnProgFailure = False         'Initialize failure to False

'Initialize failure arrays to all parameters failed
For lintProgrammerNum = 1 To 2
    For lintFailureNum = 0 To PROGFAULTCNT
        If lintFailureNum = 0 Then       'First element reserved for...
            gintProgFailure(lintProgrammerNum, lintFailureNum) = PROGFAULTCNT     'Number of failures checked
        Else                                'Otherwise...
            gintProgFailure(lintProgrammerNum, lintFailureNum) = True            'Initialize all failures to failed
        End If
    Next lintFailureNum
Next lintProgrammerNum

End Sub

Public Sub InitializeSensotec()
'
'   PURPOSE:    To initialize communication with the Sensotec SC2000 and set up the
'               limits, as well as define the coefficent between voltage and force.
'
'  INPUT(S):    None.
' OUTPUT(S):    None.
'2.0ANM updated to use CONST

Dim lstrSerialNum As String

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    'If lost RS-232 communication to the motor
    If Not (Sensotec.GetLinkStatus) Then

        'Reset the Sensotec and initialize communication
        Call Sensotec.InitializeCommunication(SENSOTECPORT)

    End If

    'Call Sensotec.ReadSerialNumber(1, lstrSerialNum)

'    If (lstrSerialNum <> gudtMachine.forceCell) Then
'        MsgBox "Incorrect Force Cell Attached; correct actuator for the current" _
'            & vbCrLf & "parameter file is not connected.  Serial # " & lstrSerialNum & " Found." _
'            & vbCrLf & "Please attach the correct force cell or contact Electronics.", vbOKOnly + vbCritical, _
'            "Force Cell Error"
'            Unload frmMain      'Close the main form
'    End If

    'Set DAC full scale
    Call Sensotec.SetDACFullScale(1, gudtMachine.maxLBF)

    'Set Frequency Response
    Call Sensotec.SetFreqResponse(1, 800)

    'Calculate the coefficient for Force -> Voltage from the force cell
    'Call CalcLBFPerVolt

    'Reset sensotec
    Call Sensotec.Reset
    
    'Assign the gain
    gsngNewtonsPerVolt = gsngForceGain
    'Assign the offset
    gsngForceAmplifierOffset = gsngForceOffset

End If

End Sub

Private Function IsLeapYear(Year As Long) As Boolean
'
'   PURPOSE:    To check to see if current year is a leap year.
'
'  INPUT(S):    Year = Year to be evaluated for leap year / not leap year
' OUTPUT(S):    IsLeapYear = whether or not it is a leap year

If (Year Mod 4 = 0 And Year Mod 100 <> 0) Or Year Mod 400 = 0 Then
    IsLeapYear = True
Else
    IsLeapYear = False
End If

End Function

Public Sub LoadGraphArray(chanNum As Integer, numberOfGraphs As Integer, subTitle As String, subsetTitle() As String, xLabel As String, ylabel As String, xStart As Variant, xStop As Variant, yHigh As Integer, yLow As Integer, evaluateStart As Single, evaluateStop As Single, increment As Single, graphFooterHigh As String, graphFooterLow As String, DataArray() As Single, HighLimit() As Single, LowLimit() As Single)
'
'   PURPOSE: To fill the graph array with data which is used for graphing
'
'  INPUT(S): chanNum = Channel number, used to create output name
'            numberOfGraphs = number of outputs per graph
'            subTitle = sub title on graph
'            subsetTitle() = an array of the outputs labels
'            xLabel = label on the x-axis
'            ylabel = label on the y-axis
'            xStart = the start location of the x-axis
'            xStop = the stop location of the x-axis
'            yHigh = the high value on the y-axis
'            yLow = the low value on the y-axis
'            evaluateStart = the start location of the data
'            evaluateStop = the stop location of the data
'            increment = the discrete step between data points
'            graphFooterHigh = graph footer
'            graphFooterLow = graph footer
'            dataArray = raw data array
'            highLimit = high limit data array
'            lowLimit = low limit data array
' OUTPUT(S): none

Dim i As Integer
Dim NumberOfDataPoints As Long

NumberOfDataPoints = (evaluateStop - evaluateStart) / increment

If chanNum = -1 Then
    gvntGraph(gintPointer, 0) = ""
Else
    gvntGraph(gintPointer, 0) = " Output " & Format(chanNum + 1, 0)
End If
gvntGraph(gintPointer, 1) = gudtMachine.seriesID & " Final Scanner System"   '1.6ANM
gvntGraph(gintPointer, 2) = xLabel
gvntGraph(gintPointer, 3) = ylabel
gvntGraph(gintPointer, 4) = xStart
gvntGraph(gintPointer, 5) = xStop
gvntGraph(gintPointer, 6) = yHigh
gvntGraph(gintPointer, 7) = yLow
gvntGraph(gintPointer, 8) = evaluateStart + gudtMachine.graphZeroOffset
gvntGraph(gintPointer, 9) = evaluateStop + gudtMachine.graphZeroOffset
gvntGraph(gintPointer, 10) = increment
gvntGraph(gintPointer, 11) = subTitle                                        '1.6ANM
'2.3ANM use sample if force test only
If gblnForceOnly Then
    gvntGraph(gintPointer, 12) = frmMain.ctrSetupInfo1.Sample
Else
    gvntGraph(gintPointer, 12) = frmMain.ctrSetupInfo1.PartNum
End If

If numberOfGraphs >= 2 Then
    gvntGraph(gintPointer, 13) = numberOfGraphs
Else
    gvntGraph(gintPointer, 13) = 0
End If

If numberOfGraphs >= 2 Then
    For i = 0 To numberOfGraphs - 1
        gvntGraph(gintPointer, i + 14) = subsetTitle(i)
    Next i
End If

For i = 0 To NumberOfDataPoints - 1
    gvntGraph(gintPointer + 1, i) = DataArray(i + evaluateStart / increment)
    gvntGraph(gintPointer + 2, i) = HighLimit(i + evaluateStart / increment)
    gvntGraph(gintPointer + 3, i) = LowLimit(i + evaluateStart / increment)
Next i

'If more than two graphs on one output
If numberOfGraphs >= 2 Then
    For i = NumberOfDataPoints To numberOfGraphs * NumberOfDataPoints
        gvntGraph(gintPointer + 1, i) = DataArray(i)
    Next i
End If

'Keep track of column to save data
gintPointer = gintPointer + 4

End Sub

Public Sub LoadMultipleGraphArray(graphNumber As Integer, maxData As Long, DataArray() As Single, multipleGraphArray() As Single)
'
'   PURPOSE: To create an array for multiple outputs to be display on one screen

Dim i As Integer

For i = 0 To maxData - 1
    multipleGraphArray(i + (graphNumber * maxData)) = DataArray(i)
Next i

End Sub

Public Sub MachineSetupCode()
'
'   PURPOSE:   This code is used to communicate the correct setup code to
'              PLC for a specific part number.  This code is used in cases
'              where there are several part numbers to run on the same
'              station and the PC communicates to the PLC via TTL.
'
'  INPUT(S):    None.
'
' OUTPUT(S):    None.

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    Select Case gudtMachine.BOMNumber                  'Set machine setup code BIT 0
       Case 1, 3, 5, 7, 9, 11, 13, 15
          Call frmDAQIO.OffPort2(PORT7, BIT0)
       Case Else
          Call frmDAQIO.OnPort2(PORT7, BIT0)
    End Select
          
    Select Case gudtMachine.BOMNumber                  'Set machine setup code BIT 1
       Case 2, 3, 6, 7, 10, 11, 14, 15
          Call frmDAQIO.OffPort2(PORT7, BIT1)
       Case Else
          Call frmDAQIO.OnPort2(PORT7, BIT1)
    End Select
          
    Select Case gudtMachine.BOMNumber                  'Set machine setup code BIT 2
       Case 4 To 7, 12 To 15
          Call frmDAQIO.OffPort2(PORT7, BIT2)
       Case Else
          Call frmDAQIO.OnPort2(PORT7, BIT2)
    End Select
          
    Select Case gudtMachine.BOMNumber                  'Set machine setup code BIT 3
       Case 8 To 15
          Call frmDAQIO.OffPort2(PORT7, BIT3)
       Case Else
          Call frmDAQIO.OnPort2(PORT7, BIT3)
    End Select
End If

End Sub

Public Sub MoveToLoadLocation()
'
'   PURPOSE:    This code returns the scanner to position where a part is loaded.
'
'  INPUT(S):    None
' OUTPUT(S):    None
'

Dim lintMotorErr As Integer
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngStartTimer As Single
Dim lblnMoveDone As Boolean
Dim lblnTimeOut As Boolean

frmMain.staMessage.Panels(1).Text = "System Message: Moving to Part Load/Unload Position..."

'Check the status of the motor before the home move
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200        'Convert to system fault code
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Motor Error while moving to the Load Location: Please Cycle Power to" & vbCrLf & _
                      "the Motor Control Box and Home the motor using the Function Menu", True, True)
    'Update the system status
    frmMain.staMessage.Panels(1).Text = "System Message: Drive Arm NOT at Part Load/Unload Position: Home Motor!"
    'Exit on System Error
    Exit Sub
End If

'Set servo motor Velocity
Call VIX500IE.SetVelocity(gudtMachine.preScanVelocity)
'Set servo motor Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.preScanAcceleration)
'Set servo motor Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.preScanAcceleration)

'Define the location to move to
Call VIX500IE.DefineMovement(gudtMachine.loadLocation)

'Check the status of scanner home complete before the offset move
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Moving to the Load Location: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If

'Start motor
Call VIX500IE.StartMotor

'Delay 50 msec for the motor to start
Call frmDAQIO.KillTime(50)

'Start the watchdog timer for the move
lsngStartTimer = Timer

'Wait for the motor to stop
Do
    lsngFirstPosition = Position
    'Delay 10 msec for movement
    Call frmDAQIO.KillTime(10)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check for timeout
    lblnTimeOut = (lsngStartTimer - Timer > MOVETOLOADLOCATIONTIMEOUT)
Loop Until lblnTimeOut

If lblnTimeOut Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Error while moving to the Load Location: Please Cycle Power to" & vbCrLf & _
                      "the Motor Control Box and Home the motor using the Function Menu", True, True)
    'Update the system status
    frmMain.staMessage.Panels(1).Text = "System Message: Drive Arm NOT at Part Load/Unload Position: Home Motor!"
Else
    'Update the system status
    frmMain.staMessage.Panels(1).Text = "System Message: Drive Arm at Part Load/Unload Position."
End If

End Sub

Public Function MoveToPosition(LocationInDegrees As Single, TimeOutInSeconds As Single) As Boolean
'
'     PURPOSE:  To move to a specified position while checking that a timeout
'               condition has not occured.
'
'    INPUT(S):  LocationInDegrees
'               TimeOutInSeconds
'   OUTPUT(S):  None.
        
Dim lsngStartTime As Single
Dim lsngTimeOutTime As Single
Dim lsngCurrentTime As Single
Dim lsngPast As Single
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single

Dim lblnTimeOut As Single
Dim lblnMoveDone As Single
Dim lintMotorErr As Integer

'Define the position to move to
Call VIX500IE.DefineMovement(LocationInDegrees)
'Start the motor
Call VIX500IE.StartMotor

'Delay 50 msec for the motor to start
Call frmDAQIO.KillTime(50)

'Start the watchdog timer for the move
lsngStartTime = Timer
'Determine when the timeout condition occurs
lsngTimeOutTime = lsngStartTime + TimeOutInSeconds
lsngPast = lsngStartTime

'Wait for the motor to stop
Do
    lsngFirstPosition = Position
    'Delay 10 msec for movement
    Call frmDAQIO.KillTime(10)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check to see if the System Clock rolled over
    lsngCurrentTime = Timer
    If lsngCurrentTime < lsngPast Then
        'Timer rolled over at midnight; reset the Time to timeout at:
        lsngTimeOutTime = lsngTimeOutTime - SECONDSPERDAY
    Else
        lsngPast = lsngCurrentTime
    End If
    'Check for a timeout condition
    lblnTimeOut = lsngCurrentTime > lsngTimeOutTime
Loop Until lblnTimeOut Or lblnMoveDone

If lblnTimeOut Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Error while moving to a specified position: Please Cycle Power to" & vbCrLf & _
                      "the Motor Control Box and Home the motor using the Function Menu", True, True)
    'Update the system status
    frmMain.staMessage.Panels(1).Text = "System Message: Drive Arm NOT at Correct Position: Home Motor!"
Else
    'Routine completed successfully
    MoveToPosition = True
End If

End Function

Public Function Position() As Single
'
'     PURPOSE:  To return the current position from the Position Trigger Board
'
'    INPUT(S):  None.
'   OUTPUT(S):  Value => Position of the encoder

Dim LSBAddr As Integer, MSBAddr As Integer
Dim LSBData As Variant, MIDData As Variant, MSBData As Variant
Dim lsngPosition As Single

'Get Position Address
LSBAddr = PTBASEADDRESS And &HFF                  'Get LSB address
MSBAddr = (PTBASEADDRESS \ BIT8) And &HFF         'Get MSB address
                                            
'Read Position Data
Call frmDAQIO.ReadPTBoardData(LSBAddr, MSBAddr, LSBData, MIDData, MSBData)

MIDData = MIDData * BIT8
MSBData = MSBData * BIT16

'Convert position to the proper units
lsngPosition = ((MSBData + MIDData + LSBData) / gudtMachine.encReso) * DEGPERREV

'Any numbers beyond half the PT Board's max count are considered negative counts
'NOTE: The maximum count from the PT Board is 2^24 - 1 (24-bit counter)
If lsngPosition >= (2 ^ 23 / gudtMachine.encReso) * DEGPERREV Then
    lsngPosition = lsngPosition - (2 ^ 24 / gudtMachine.encReso) * DEGPERREV
End If

'Display the position
frmMain.txtPosition = Format(lsngPosition, "##0.00°")

'Return the position
Position = lsngPosition

End Function

Public Sub ReadSerialNumberAndDateCode90293()
'
'   PURPOSE: To Read the 90293 MLX IC's for the S/N & Date Code
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lintMLXID(6) As Integer
Dim SN As String
Dim d As Double

'Enable Programming paths
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2

Call MyDev(lintDev1).DeviceReplaced
Call MyDev(lintDev1).ReadFullDevice

lintMLXID(0) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID0)
lintMLXID(1) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID1)
lintMLXID(2) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID2)
lintMLXID(3) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID3)
lintMLXID(4) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID4)

gudtMLX90277(1).Read.X = (Int(lintMLXID(0) / BIT5) And &H7) + ((lintMLXID(1) And &H1F) * BIT3)
gudtMLX90277(1).Read.Y = (Int(lintMLXID(1) / BIT5) And &H7) + ((lintMLXID(2) And &H1F) * BIT3)
gudtMLX90277(1).Read.Wafer = (lintMLXID(0) And &H1F)
d = CDbl(lintMLXID(4)) * BIT11
gudtMLX90277(1).Read.Lot = (Int(lintMLXID(2) / BIT5) And &H7) + (lintMLXID(3) * BIT3) + d

SN = MLX90293.EncodePartID
gstrSerialNumber = SN

'Display the S/N of the DUT
frmMain.ctrSetupInfo1.PartNum = gstrSerialNumber
'Calculate the date code from the DUT
glngCUSTID1 = MyDev(lintDev1).GetEEParameterCode(CodeUSERID1)
glngCUSTID2 = MyDev(lintDev1).GetEEParameterCode(CodeUSERID2)

'Get the current readings
gudtReading(0).mlxCurrent = MyDev(lintDev1).GetIdd
gudtReading(1).mlxCurrent = MyDev(lintDev2).GetIdd

gblnGoodDateCode = True

gstrDateCode = MLX90293.DecodeCustomerID90293
    
End Sub

Public Sub ReadSerialNumberAndDateCode()
'
'   PURPOSE: To Read the MLX IC's for the S/N & Date Code
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lblnVotingError As Boolean
Dim lintAttemptNum As Integer
Dim lstrSN As String '3.0aANM

'Enable Programming paths
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2
'Proceed if communication with the PTC-04's is active
If gblnGoodPTC04Link Then
    'Try to read the chip twice (if necessary)
    For lintAttemptNum = 1 To 2
        frmMain.staMessage.Panels(1).Text = "System Message:  Reading Contents of IC's"
        'Read values back from EEprom
        If Not MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError) Then
            If lintAttemptNum = 2 Then                          'Only assign error after two attempts
                gintAnomaly = 164
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", True, True)
            End If
        End If
        'Verify that there were no voting errors
        If lblnVotingError Then
            If lintAttemptNum = 2 Then                          'Only assign error after two attempts
                gintAnomaly = 167
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & _
                                  "Verify correct Revision of ASIC is in use.", True, True)
            End If
        End If
        'Calculate the serial number from the DUT
        gstrSerialNumber = MLX90277.EncodePartID
        
        '3.0aANM MLX Checks
        gblnMLXOk = False
        
        If gstrSerialNumber = "00000000000" Then
            lstrSN = MLX90277.EncodePartID2 & " (#2)"
        Else
            lstrSN = gstrSerialNumber
        End If
        
        If (gudtMLX90277(1).Read.Lot = 0) Or (gudtMLX90277(2).Read.Lot = 0) Then
            gintAnomaly = 169 'Bad SN
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & _
                              "Verify Checkhead Connections.", True, True)
            Exit Sub
        Else
            If (gudtMLX90277(1).Read.MelexisLock = False) Or (gudtMLX90277(2).Read.MelexisLock = False) Or ((gudtMLX90277(1).Read.TC = 0) And (gudtMLX90277(1).Read.TCWin = 0) And (gudtMLX90277(1).Read.TC2nd = 0)) Or ((gudtMLX90277(2).Read.TC = 0) And (gudtMLX90277(2).Read.TCWin = 0) And (gudtMLX90277(2).Read.TC2nd = 0)) Then
                gintAnomaly = 181
                Call ErrorLogFile("Severe Programmer Error: MLX Lock or TC not set! Verify MLX ICs." & vbCrLf & "Please Tag " & lstrSN & " Part as Bad MLX Chip.", True, True)
            End If
            
            If gudtMLX90277(1).Read.Lot <> gudtMLX90277(2).Read.Lot Then
                gintAnomaly = 182
                Call ErrorLogFile("Severe Programmer Error: MLX Lot #s don't match! Verify MLX ICs." & vbCrLf & "Please Tag " & lstrSN & " Part as Bad MLX Chip.", True, True)
            End If
        End If
        If gintAnomaly = 0 Then gblnMLXOk = True
        
        'Display the S/N of the DUT
        frmMain.ctrSetupInfo1.PartNum = gstrSerialNumber
        'Calculate the date code from the DUT
        gstrDateCode = MLX90277.DecodeCustomerID(gudtMLX90277(1).Read.CustID)
        'Verify that the Date Code is non-zero (and specifically that the shift letter is non-zero)
        If gstrDateCode = "0000000" Or (Mid(gstrDateCode, 6, 1) = "0") Then
            If lintAttemptNum = 2 Then                          'Only assign error after two attempts
                gintAnomaly = 170
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: No Date Code Found!" & vbCrLf & _
                                  "Verify That Part Was Programmed.", True, True)
            End If
        Else
            gblnGoodDateCode = True
        End If
        'Verify that the serial number is non-zero
        If gstrSerialNumber = "00000000000" Then
            If lintAttemptNum = 2 Then                          'Only assign error after two attempts
                gintAnomaly = 169
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & _
                                  "Verify Checkhead Connections.", True, True)
            End If
        Else
            gblnGoodSerialNumber = True
        End If
        'Exit the For...Next Loop if both date code and serial number were read successfully
        If gblnGoodDateCode And gblnGoodSerialNumber Then Exit For
    Next lintAttemptNum
Else
    gintAnomaly = 168
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & _
                           "Verify Connections to Programmer.", True, True)
End If

'Get the current readings '2.9ANM
If Not gblnProgramStart Then Call MLX90277.GetCurrent

End Sub

Public Sub RunSolver()
'
'   PURPOSE: To run through the steps of the Solver
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lintProgrammerNum As Integer
Dim lintCycleNum As Integer
Dim lsngIndexLoc(1 To 2) As Single
Dim lblnGainAndOffsetGood As Boolean

'Zero the current part counter if it needs to be reset
If gudtProgSummary.currentTotal >= gudtMachine.currentPartCount Then
    gudtProgSummary.currentTotal = 0
    gudtProgSummary.currentGood = 0
End If

'Initialize for the first step of the solver
Call Solver90277.SolverInitialization

'The starting position is our Index 1 position:
lsngIndexLoc(1) = Position

'2.0ANM \/\/
'Show users that we're Solving for Clamp Codes
frmMain.staMessage.Panels(1).Text = "System Message:  Solving for Low Clamp Codes"

'Find low saturation levels
Call Solver90277.FindSaturationLevels(1)

'Find high saturation levels
Call Solver90277.FindSaturationLevels(2)
        
'2.4ANM moved high sat above and if below
If gintAnomaly Then Exit Sub

'Solve for the low clamp codes
Call Solver90277.ClampSolver(1)
'2.0ANM /\/\

'Two iterations through the Solver, if necessary
For lintCycleNum = 1 To 2

    '*** Solver Step #1 ***
    'Display the current Status
    frmMain.staMessage.Panels(1).Text = "System Message:  Performing Solver Cycle #" & CStr(lintCycleNum) & ", Step #1..."

    'Index one is at pedal-at-rest location degrees
    For lintProgrammerNum = 1 To 2
        gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(1).MeasuredLocation(1) = 0
    Next lintProgrammerNum

    'Make measurements: Cycle number, Step 1, Index 1
    If Not Solver90277.MakeSolverMeasurements(lintCycleNum, 1, 1) Then
        gintAnomaly = 161
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Solver Measurement Failure on Cycle #" & CStr(lintCycleNum) & ", Step #1, Index #1.", True, True)
        Exit Sub
    End If

    'Find the appropriate position to move to if this is the first cycle
    'Otherwise, the second index location has already been defined
    If (lintCycleNum = 1) Then
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
    End If

    'Move to Index 2
    If Not (MoveToPosition(lsngIndexLoc(2), 1.5)) Then
        gintAnomaly = 163
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Motor Movement Failed While Solving.", True, True)
        Exit Sub
    End If

    'Measure the current location (Index 2)
    For lintProgrammerNum = 1 To 2
        gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(1).MeasuredLocation(2) = Position - gudtReading(CHAN0).pedalFaceLoc
    Next lintProgrammerNum

    '2.0ANM \/\/
    If (lintCycleNum = 1) Then
        'Show users that we're Solving for Clamp Codes
        frmMain.staMessage.Panels(1).Text = "System Message:  Solving for High Clamp Codes"
        
        '2.4ANM moved sat check up to idle location
        
        'Solve for the high clamp codes
        Call Solver90277.ClampSolver(2)
    End If
    '2.0ANM /\/\

    'Make measurements: Cycle number, Step 1, Index 2
    If Not Solver90277.MakeSolverMeasurements(lintCycleNum, 1, 2) Then
        gintAnomaly = 161
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Solver Measurement Failure on Cycle #" & CStr(lintCycleNum) & ", Step #1, Index #2.", True, True)
        Exit Sub
    End If

    'Calculate the Slopes and Intercepts for this step
    Call Solver90277.CalculateSlopesAndIntercepts(lintCycleNum, 1)

    'Count the number of good measurement pairs
    Call CountGoodMeasurements(lintCycleNum, 1)

    'Evaluate the results of the tests and calculate codes for Step 2
    If Not Solver90277.EvaluateTests(lintCycleNum, 1) Then
        gintAnomaly = 160
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Solver Calculation Failure on Cycle #" & CStr(lintCycleNum) & ", Step #1.", True, True)
        Exit Sub
    End If

    '*** Solver Step #2 ***
    'Display the current Status
    frmMain.staMessage.Panels(1).Text = "System Message:  Performing Solver Cycle #" & CStr(lintCycleNum) & ", Step #2..."

    'Measure the current location (Index 2)
    For lintProgrammerNum = 1 To 2
        gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(2).MeasuredLocation(2) = Position - gudtReading(CHAN0).pedalFaceLoc
    Next lintProgrammerNum

    'Make measurements: Cycle number, Step 2, Index 2
    If Not Solver90277.MakeSolverMeasurements(lintCycleNum, 2, 2) Then
        gintAnomaly = 161
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Solver Measurement Failure on Cycle #" & CStr(lintCycleNum) & ", Step #2, Index #2.", True, True)
        Exit Sub
    End If

    'Move to Index 1
    If Not (MoveToPosition(lsngIndexLoc(1), 1.5)) Then
        gintAnomaly = 163
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Motor Movement Failed While Solving.", True, True)
        Exit Sub
    End If

    'Index one is at pedal-at-rest degrees
    For lintProgrammerNum = 1 To 2
        gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(2).MeasuredLocation(1) = 0
    Next lintProgrammerNum

    'Make measurements: Cycle number, Step 2, Index 1
    If Not Solver90277.MakeSolverMeasurements(lintCycleNum, 2, 1) Then
        gintAnomaly = 161
        'Log the error to the error log and display the error message
        Call ErrorLogFile("Solver Measurement Failure on Cycle #" & CStr(lintCycleNum) & ", Step #2, Index #1.", True, True)
        Exit Sub
    End If

    'Calculate the Slopes and Intercepts for this step
    Call Solver90277.CalculateSlopesAndIntercepts(lintCycleNum, 2)

    'Count the number of good measurement pairs
    Call CountGoodMeasurements(lintCycleNum, 2)

    'Attempt to fine-tune offset
    gblnGoodOffsetAndGainCodes = AdjustOffset(lintCycleNum)

    'If adjustment worked, we're done cycling
    If gblnGoodOffsetAndGainCodes Then
        'If adjustment worked, we're done cycling
        Exit For
    Else
        'If it didn't work, check if this was the last cycle
        If lintCycleNum < MAXCYCLENUM Then
            'Evaluate the results of the tests and calculate codes for Step 1 of the last cycle
            If Not Solver90277.EvaluateTests(lintCycleNum, 2) Then
                gintAnomaly = 160
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Solver Calculation Failure on Cycle #" & CStr(lintCycleNum) & ", Step #2.", True, True)
                Exit Sub
            End If
        Else '2.1ANM added else
            gintAnomaly = 160
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Solver Calculation Failure on Cycle #" & CStr(lintCycleNum) & ", Step #2.", True, True)
            Exit Sub
        End If
    End If

Next lintCycleNum

'Check for pass/fail
If gblnUseNewAmad Then '2.6ANM
    Call Pedal.CheckForProgrammingFaultsTestDynamicMPC
Else
    Call CheckForProgrammingFaults                             '2.0ANM moved up
End If

'Only Solve for Clamps if Offset And Gain were correctly calculated
If gblnGoodOffsetAndGainCodes And (Not gblnProgFailure) Then '2.0ANM moved clamp solver up and added gblnprogfailure
    'Reset the Offset and Clamp parameters
    For lintProgrammerNum = 1 To 2
        gudtMLX90277(lintProgrammerNum).Write.offset = gudtMLX90277(lintProgrammerNum).Read.offset
        gudtMLX90277(lintProgrammerNum).Write.clampLow = gudtMLX90277(lintProgrammerNum).Read.clampLow
        gudtMLX90277(lintProgrammerNum).Write.clampHigh = gudtMLX90277(lintProgrammerNum).Read.clampHigh
        gudtMLX90277(lintProgrammerNum).Write.InvertSlope = gudtSolver(lintProgrammerNum).InvertSlope   'V1.4.0
    Next lintProgrammerNum
Else
    'Set the clamp codes to force to the low diagnostic region
    gudtMLX90277(1).Write.clampLow = MINCLAMPCODE
    gudtMLX90277(2).Write.clampLow = MINCLAMPCODE
    gudtMLX90277(1).Write.clampHigh = MINCLAMPCODE
    gudtMLX90277(2).Write.clampHigh = MINCLAMPCODE
End If

End Sub

Public Sub RunTest()
'
'   PURPOSE: The executive which handles a scan and data processing
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim lsngPosition As Single
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngStartTimer As Single
Dim lblnTimeOut As Boolean
Dim lblnMoveDone As Boolean
Dim lintMotorErr As Integer
Dim lintProgrammerNum As Integer
Dim lblnVotingError As Boolean

'Zero the current part counter if it needs to be reset
If gudtScanSummary.currentTotal >= gudtMachine.currentPartCount Then
    gudtScanSummary.currentTotal = 0
    gudtScanSummary.currentGood = 0
End If

'Initialize all failures to true and mask failures not used
Call InitializeAndMaskScanFailures

'If hardware is not present bypass logic
If InStr(command$, "NOHARDWARE") = 0 Then

    'Continue testing if no errors
    If gintAnomaly = 0 Then

        'Activate the Sensotec SC2000's Tare function to remove any offset
        If Sensotec.GetLinkStatus Then
            Call Sensotec.ActivateTare(1)
        Else
            'Error Communicating with the Sensotec SC2000
            gintAnomaly = 50
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
        End If

        'Find Start Scan
        Call FindStartScan

        'De-Activate the Sensotec SC2000's Tare Function
        If Sensotec.GetLinkStatus Then
            Call Sensotec.DeActivateTare(1)
        Else
            'Error Communicating with the Sensotec SC2000
            gintAnomaly = 50
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
        End If

        'Continue testing if no errors
        If gintAnomaly = 0 Then

            'NOTE: Time is saved here by starting the movement, switching
            '      relays, then checking for the movement to be done.
            'Define the location to move to before starting the scan
            Call VIX500IE.DefineMovement(gudtMachine.scanStart - gudtMachine.overTravel)
            'Start the motor
            Call VIX500IE.StartMotor

            'Update the position
            lsngPosition = Position

            'Update the D/A output to 0V '2.1ANM
            Call frmDAQIO.cwaoVRef.SingleWrite(0)
            
            'Enable VRef
            Call frmDAQIO.OnPort1(PORT4, BIT1)

            'Enable the proper filter
            For lintChanNum = CHAN0 To CHAN3
                Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
            Next lintChanNum

            'Disable Programming paths
            Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
            Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

            Call Position

            'Delay 50 msec
            Call frmDAQIO.KillTime(50)

            'Update the D/A output to Vref '2.1ANM
            Call frmMain.RampToVref

            'Start the watchdog timer
            lsngStartTimer = Timer
            'Wait for the motor to stop
            Do
                lsngFirstPosition = Position
                'Delay 10 msec for movement
                Call frmDAQIO.KillTime(10)
                'Check to see if the move has completed
                lsngCurrentPosition = Position
                lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
                'Exit the loop if the motor has stopped
                If lblnMoveDone Then Exit Do
                'Check for timeout
                lblnTimeOut = (lsngStartTimer - Timer > 2)  '2 second timeout
            Loop Until lblnTimeOut
            'Display an error if there was a timeout
            If lblnTimeOut Then
                lintMotorErr = VIX500IE.ReadDriveFault
                gintAnomaly = lintMotorErr + 200                'Convert to system fault code
                Call ErrorLogFile("Motor Error while moving to a Start Scan: Please Cycle Power to" & vbCrLf & _
                                  "the Motor Control Box and Home the motor using the Function Menu", True, True)
                'Update the system status
                frmMain.staMessage.Panels(1).Text = "System Message: Drive Arm NOT at Correct Position: Home Motor!"
            End If

            'Clear the previous peak/valley force
            If Sensotec.GetLinkStatus Then
                Call Sensotec.ClearPeakAndValley(1)
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If

            'Activate the Sensotec SC2000's Tare function to remove any offset
            If Sensotec.GetLinkStatus Then
                Call Sensotec.ActivateTare(1)
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If

            'Check the Supply Voltage
            If (gintAnomaly = 0) Then Call VerifySupplyVoltage

            'Measure Static Index 1 (Idle) for Track #1 & #2  '1.7ANM
            If Not gblnForceOnly Then Call frmDAQIO.ReadVout(gudtReading(CHAN0).Index(1).Value, gudtReading(CHAN1).Index(1).Value)

            'Scan the DUT   '1.7ANM added if block
            If gblnForceOnly Then
                If (gintAnomaly = 0) Then Call ScanForce
            Else
                If (gintAnomaly = 0) Then Call ScanUnit
            End If
            
            '1.6ANM Measure Static Index 1 (Idle) for Track #1 & #2 after scan for 705
            If gudtMachine.seriesID = "705" Then
                Call frmDAQIO.ReadVout(gudtReading(CHAN0).Index(4).Value, gudtReading(CHAN1).Index(4).Value)
            End If
            
            'Read the peak force
            If Sensotec.GetLinkStatus Then
                Call Sensotec.ReadValley(1, gudtReading(CHAN0).peakForce)
                'The force read back is negative; invert the sign here
                'Also force read back is in LBF, so scale by N/LBF
                gudtReading(CHAN0).peakForce = -gudtReading(CHAN0).peakForce * NEWTONSPERLBF
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If

            'De-Activate the Sensotec SC2000's Tare Function
            If Sensotec.GetLinkStatus Then
                Call Sensotec.DeActivateTare(1)
            Else
                'Error Communicating with the Sensotec SC2000
                gintAnomaly = 50
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
            End If

            'Disable the filters
            For lintChanNum = CHAN0 To CHAN3
                Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
            Next lintChanNum

            'Enable Programming paths
            Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
            Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

            'Disable VRef
            Call frmDAQIO.OffPort1(PORT4, BIT1)
        End If

    End If

End If

If gblnForceOnly Then    '1.7ANM added if block
    'Find the Pedal At Rest Location and Manage the Data Arrays Accordingly
    If (gintAnomaly = 0) Then Call FindPedalZeroAndTruncateData(False)
    
    'Evaluate the Scan Data
    If (gintAnomaly = 0) Then Call DataEvaluation
Else
    'Check the Supply Voltage Data Arrays
    If (gintAnomaly = 0) Then Call CheckSupplyArray
    
    'Find the Pedal At Rest Location and Manage the Data Arrays Accordingly
    If (gintAnomaly = 0) Then Call FindPedalZeroAndTruncateData(False)
    
    'Evaluate the Scan Data
    If (gintAnomaly = 0) Then Call DataEvaluation
End If

'Set the part status
If (gintAnomaly = 0) Then Call SetFailureStatus

'Save the results data if called for
'1.6ANM If ((gintAnomaly = 0) And gblnSaveScanResultsToFile) Then Call SaveScanResultsToFile

End Sub

Public Sub SaveProgResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
'2.8ANM added zg

Dim lintFileNum As Integer
Dim lstrFileName As String

'Make the results file name
lstrFileName = gstrLotName + " Programming Results" & DATAEXT
'Get a file
lintFileNum = FreeFile

'If file does not exist then add a header
If Not gfsoFileSystemObject.FileExists(PARTPROGDATAPATH + lstrFileName) Then
    Open PARTPROGDATAPATH + lstrFileName For Append As #lintFileNum
    'Part S/N, Date Code, Date/Time, Software Revision, Parameter File Name, & Pallet Number
    Print #lintFileNum, _
        "Part Number,"; _
        "Date Code,"; _
        "Date/Time,"; _
        "S/W Revision,"; _
        "Parameter File Name,"; _
        "Pallet Number,";
    'Programming Results Vout #1
    Print #lintFileNum, _
        "Final Idle Output Vout #1 [%],"; _
        "Final Idle Location Vout #1 [°],"; _
        "Final WOT Output Vout #1 [%],"; _
        "Final WOT Location Vout #1 [°],"; _
        "Final Clamp Low Value Vout #1 [%],"; _
        "Final Clamp High Value Vout #1 [%],"; _
        "Final Offset Code Vout #1,"; _
        "Final Rough Gain Code Vout #1,"; _
        "Final Fine Gain Code Vout #1,"; _
        "Final Clamp Low Code Vout #1,"; _
        "Final Clamp High Code Vout #1,";
        '1.9ANM "Offset Seed Code Vout #1,";
        '1.9ANM "Rough Gain Seed Code Vout #1,";
        '1.9ANM "Fine Gain Seed Code Vout #1,";
    'Programming Process Variables Vout #1
    Print #lintFileNum, _
        "Vout #1 Cycle 1 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #1 Cycle 1 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #1 Cycle 1 Step 1 Measured Slope 3 [%/°],"; _
        "Vout #1 Cycle 1 Step 1 Measured Slope 4 [%/°],"; _
        "Vout #1 Cycle 1 Step 2 Measured Slope 1 [%/°],"; _
        "Vout #1 Cycle 1 Step 2 Measured Slope 2 [%/°],"; _
        "Vout #1 Cycle 1 Step 2 Measured Slope 3 [%/°],"; _
        "Vout #1 Cycle 1 Offset Adjusted Index 1 [%],"; _
        "Vout #1 Cycle 1 Offset Adjusted Index 2 [%],"; _
        "Vout #1 Cycle 2 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #1 Cycle 2 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #1 Cycle 2 Step 1 Measured Slope 3 [%/°],"; _
        "Vout #1 Cycle 2 Step 2 Measured Slope 1 [%/°],"; _
        "Vout #1 Cycle 2 Step 2 Measured Slope 2 [%/°],"; _
        "Vout #1 Cycle 2 Step 2 Measured Slope 3 [%/°],"; _
        "Vout #1 Cycle 2 Offset Adjusted Index 1 [%],"; _
        "Vout #1 Cycle 2 Offset Adjusted Index 2 [%],"; _
        "Vout #1 Zero Gauss Position [°],";
    'Programming Results Vout #2
    Print #lintFileNum, _
        "Final Idle Output Vout #2 [%],"; _
        "Final Idle Location Vout #2 [°],"; _
        "Final WOT Output Vout #2 [%],"; _
        "Final WOT Location Vout #2 [°],"; _
        "Final Clamp Low Value Vout #2 [%],"; _
        "Final Clamp High Value Vout #2 [%],"; _
        "Final Offset Code Vout #2,"; _
        "Final Rough Gain Code Vout #2,"; _
        "Final Fine Gain Code Vout #2,"; _
        "Final Clamp Low Code Vout #2,"; _
        "Final Clamp High Code Vout #2,";
        '1.9ANM "Offset Seed Code Vout #2,";
        '1.9ANM "Rough Gain Seed Code Vout #2,";
        '1.9ANM "Fine Gain Seed Code Vout #2,";
    'Programming Process Variables Vout #1
    Print #lintFileNum, _
        "Vout #2 Cycle 1 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #2 Cycle 1 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #2 Cycle 1 Step 1 Measured Slope 3 [%/°],"; _
        "Vout #2 Cycle 1 Step 1 Measured Slope 4 [%/°],"; _
        "Vout #2 Cycle 1 Step 2 Measured Slope 1 [%/°],"; _
        "Vout #2 Cycle 1 Step 2 Measured Slope 2 [%/°],"; _
        "Vout #2 Cycle 1 Step 2 Measured Slope 3 [%/°],"; _
        "Vout #2 Cycle 1 Offset Adjusted Index 1 [%],"; _
        "Vout #2 Cycle 1 Offset Adjusted Index 2 [%],"; _
        "Vout #2 Cycle 2 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #2 Cycle 2 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #2 Cycle 2 Step 1 Measured Slope 3 [%/°],"; _
        "Vout #2 Cycle 2 Step 2 Measured Slope 1 [%/°],"; _
        "Vout #2 Cycle 2 Step 2 Measured Slope 2 [%/°],"; _
        "Vout #2 Cycle 2 Step 2 Measured Slope 3 [%/°],"; _
        "Vout #2 Cycle 2 Offset Adjusted Index 1 [%],"; _
        "Vout #2 Cycle 2 Offset Adjusted Index 2 [%],"; _
        "Vout #2 Zero Gauss Position [°],";
    'Part Status, Comment, Operator Initials, & Temperature
    Print #lintFileNum, _
        "Status,"; _
        "Comment,"; _
        "Operator,"; _
        "Temperature,"; _
        "Part Locked,"
Else
    Open PARTPROGDATAPATH + lstrFileName For Append As #lintFileNum
End If
'Part S/N, Date Code, Date/Time, Software Revision, Parameter File Name, & Pallet Number
Print #lintFileNum, _
    gstrSerialNumber; ","; _
    gstrDateCode; ","; _
    DateTime.Now; ","; _
    App.Major & "." & App.Minor & "." & App.Revision; ","; _
    gudtMachine.parameterName; ","; _
    gintPalletNumber; ",";
'Programming Results Vout #1
Print #lintFileNum, _
    Format(Round(gudtSolver(1).FinalIndexVal(1), 3), "##0.000"); ","; _
    Format(Round(gudtSolver(1).FinalIndexLoc(1), 2), "##0.00"); ","; _
    Format(Round(gudtSolver(1).FinalIndexVal(2), 3), "##0.000"); ","; _
    Format(Round(gudtSolver(1).FinalIndexLoc(2), 2), "##0.00", 2); ","; _
    Format(Round(gudtSolver(1).FinalClampLowVal, 3), "##0.000"); ","; _
    Format(Round(gudtSolver(1).FinalClampHighVal, 3), "##0.000"); ","; _
    Format(Round(gudtSolver(1).FinalOffsetCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(1).FinalRGCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(1).FinalFGCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(1).FinalClampLowCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(1).FinalClampHighCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(1).OffsetSeedCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(1).RoughGainSeedCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(1).FineGainSeedCode, 2), "##0.00"); ",";
'Programming Process Variables Vout#1
Print #lintFileNum, _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(4).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(2).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(2).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(2).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).OffsetAdjustedOutput(1), 3), "0.000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).OffsetAdjustedOutput(2), 3), "0.000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(2).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(2).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).Step(2).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).OffsetAdjustedOutput(1), 3), "0.000"); ","; _
    Format(Round(gudtSolver(1).Cycle(2).OffsetAdjustedOutput(2), 3), "0.000"); ","; _
    Format(Round(gudtSolver(1).ZeroGXPos(1, 1), 3), "0.000"); ",";
'Programming Results Vout #2
Print #lintFileNum, _
    Format(Round(gudtSolver(2).FinalIndexVal(1), 3), "##0.000"); ","; _
    Format(Round(gudtSolver(2).FinalIndexLoc(1), 2), "##0.00"); ","; _
    Format(Round(gudtSolver(2).FinalIndexVal(2), 3), "##0.000"); ","; _
    Format(Round(gudtSolver(2).FinalIndexLoc(2), 2), "##0.00", 2); ","; _
    Format(Round(gudtSolver(2).FinalClampLowVal, 3), "##0.000"); ","; _
    Format(Round(gudtSolver(2).FinalClampHighVal, 3), "##0.000"); ","; _
    Format(Round(gudtSolver(2).FinalOffsetCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(2).FinalRGCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(2).FinalFGCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(2).FinalClampLowCode, 2), "##0.00"); ","; _
    Format(Round(gudtSolver(2).FinalClampHighCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(2).OffsetSeedCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(2).RoughGainSeedCode, 2), "##0.00"); ",";
    '1.9ANM Format(Round(gudtSolver(2).FineGainSeedCode, 2), "##0.00"); ",";
'Programming Process Variables Vout#2
Print #lintFileNum, _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(4).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(2).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(2).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(2).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).OffsetAdjustedOutput(1), 3), "0.000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).OffsetAdjustedOutput(2), 3), "0.000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(2).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(2).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).Step(2).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).OffsetAdjustedOutput(1), 3), "0.000"); ","; _
    Format(Round(gudtSolver(2).Cycle(2).OffsetAdjustedOutput(2), 3), "0.000"); ","; _
    Format(Round(gudtSolver(2).ZeroGXPos(1, 1), 3), "0.000"); ",";
'Part Status, Comment, Temperature, & Operator Initials
If gblnProgFailure Then
    Print #lintFileNum, "REJECT,";
Else
    Print #lintFileNum, "PASS,";
End If
Print #lintFileNum, _
    frmMain.ctrSetupInfo1.Comment; ","; _
    frmMain.ctrSetupInfo1.Operator; ","; _
    frmMain.ctrSetupInfo1.Temperature; ",";

'2.0ANM \/\/
If gblnLockedPart Then
    Print #lintFileNum, "LOCKED,"
Else
    Print #lintFileNum, "UNLOCKED,"
End If
'2.0ANM /\/\

'Close the file
Close #lintFileNum

End Sub

Public Sub SaveRawDataToFile(maxData As Long, fwdGradient1() As Single, fwdGradient2() As Single, revGradient1() As Single, revGradient2() As Single, fwdForce() As Single, revForce() As Single)
'
'   PURPOSE: To save the raw data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintFileNum As Integer
Dim lsngPosition As Single
Dim i As Integer

'Get a file and open it
lintFileNum = FreeFile
Open PARTRAWDATAPATH & gstrSerialNumber & DATAEXT For Append As #lintFileNum

'Part Number & Temperature
Print #lintFileNum, "Data for part #: "; gstrSerialNumber; ",Temperature: "; frmMain.ctrSetupInfo1.Temperature

'Header for Data
Print #lintFileNum, "Position[°],Forward Vout1[%],Forward Vout2[%],Reverse Vout1[%],Reverse Vout2[%],Forward Force[N], Reverse Force[N]"

'Save the Raw Data File
For i = 0 To (maxData - 1)
    lsngPosition = (i / gsngResolution) + gudtMachine.graphZeroOffset
    Write #lintFileNum, Format(lsngPosition, "##0.00"); Format(fwdGradient1(i), "##0.00"), Format(fwdGradient2(i), "##0.00"), Format(revGradient1(i), "##0.00"), Format(revGradient2(i), "##0.00"), Format(fwdForce(i), "##0.00"), Format(revForce(i), "##0.00")
Next i

Close #lintFileNum

End Sub

Public Sub SaveTLRawDataToFile(maxData As Long, fwdGradient1() As Single, fwdGradient2() As Single, revGradient1() As Single, revGradient2() As Single, fwdForce() As Single, revForce() As Single)
'
'   PURPOSE: To save the raw data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintFileNum As Integer
Dim lsngPosition As Single
Dim i As Integer
Dim DateTime As String        '1.7ANM

'Check if lot name folder exists, if not create it
If Not gfsoFileSystemObject.FolderExists(PARTRAWDATAPATH & gstrLotName) Then
    gfsoFileSystemObject.CreateFolder (PARTRAWDATAPATH & gstrLotName)
End If

'Get a file and open it
lintFileNum = FreeFile
DateTime = Format(Date, " MM-DD-YY") & Format(Time, " HH.MM.SSAM/PM")  '1.7ANM added date/time to filename
If (frmMain.mnuOptionsSSN.Checked = True) Or (frmMain.mnuOptionsSSD.Checked = True) Then '3.0aANM
    Open PARTRAWDATAPATH & gstrLotName & "\SS " & gstrSerialNumber & DateTime & DATAEXT For Append As #lintFileNum
Else
    Open PARTRAWDATAPATH & gstrLotName & "\" & gstrSerialNumber & DateTime & DATAEXT For Append As #lintFileNum
End If

'Part Number & Temperature
Print #lintFileNum, "Data for part #: "; gstrSerialNumber; ",Temperature: "; frmMain.ctrSetupInfo1.Temperature

'Header for Data
Print #lintFileNum, "Position[°],Forward Vout1[%],Forward Vout2[%],Reverse Vout1[%],Reverse Vout2[%],Forward Force[N], Reverse Force[N]"

'Save the Raw Data File
For i = 0 To (maxData - 1)
    lsngPosition = (i / gsngResolution) + gudtMachine.graphZeroOffset
    Write #lintFileNum, Format(lsngPosition, "##0.00"); Format(fwdGradient1(i), "##0.00"), Format(fwdGradient2(i), "##0.00"), Format(revGradient1(i), "##0.00"), Format(revGradient2(i), "##0.00"), Format(fwdForce(i), "##0.00"), Format(revForce(i), "##0.00")
Next i

Close #lintFileNum

End Sub

Public Sub SaveFORawDataToFile(maxData As Long, fwdForce() As Single, revForce() As Single)
'
'   PURPOSE: To save the force raw data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
'3.0ANM new sub

Dim lintFileNum As Integer
Dim lsngPosition As Single
Dim i As Integer
Dim DateTime As String

'Check if lot name folder exists, if not create it
If Not gfsoFileSystemObject.FolderExists(PARTRAWDATAPATH & gstrLotName) Then
    gfsoFileSystemObject.CreateFolder (PARTRAWDATAPATH & gstrLotName)
End If

'Get a file and open it
lintFileNum = FreeFile
DateTime = Format(Date, " MM-DD-YY") & Format(Time, " HH.MM.SSAM/PM")
Open PARTRAWDATAPATH & gstrLotName & "\" & gstrSampleNum & DateTime & DATAEXT For Append As #lintFileNum

'Part Number & Temperature
Print #lintFileNum, "Data for part #: "; gstrSampleNum; ",Temperature: "; frmMain.ctrSetupInfo1.Temperature

'Header for Data
Print #lintFileNum, "Position[°],Forward Force[N], Reverse Force[N]"

'Save the Raw Data File
For i = 0 To (maxData - 1)
    lsngPosition = (i / gsngResolution) + gudtMachine.graphZeroOffset
    Write #lintFileNum, Format(lsngPosition, "##0.00"); Format(fwdForce(i), "##0.00"), Format(revForce(i), "##0.00")
Next i

Close #lintFileNum

End Sub

Public Sub ScanForPedalAtRestLocation()
'
'   PURPOSE: To find the pedal face location rest location in order to determine
'            where StartScan should be.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintMotorErr As Integer
Dim lsngScanWatchDogTimer As Single
Dim lblnTimeOut As Boolean

'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.preScanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.preScanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.preScanAcceleration)

'*** Scan For Pedal Face ***
frmMain.staMessage.Panels(1).Text = "System Message:  Scanning For Pedal Face..."

'Check to be sure that the Pre-Scan Scan start position is not within the scan region
Call CheckStartPosition("Pre-Scan")

If gintAnomaly Then Exit Sub                    'Exit on system error

'Check the status of the motor before the forward scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'Check the status of scanner home complete before the forward scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning for Pedal Face: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'Clear trigger counts before each scan
Call ClearCounter

'The following data is sent to the PT Board before the FORWARD scan:
'               Start Scan          : Begin data acquisition
'               End Scan            : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.preScanStart, gudtMachine.preScanStop, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.preScanStart < gudtMachine.preScanStop Then
    Call VIX500IE.DefineMovement(gudtMachine.preScanStop + gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.preScanStop - gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Initialize watchdog timer
lsngScanWatchDogTimer = Timer

'Calculate number of data points
gintMaxData = (gudtMachine.preScanStop - gudtMachine.preScanStart) * gsngResolution

'Setup the Forward Data Aqcuisition Control
frmDAQIO.cwaiPreScanDAQ.NScans = gintMaxData                                'Number of data points
frmDAQIO.cwaiPreScanDAQ.NScansPerBuffer = gintMaxData * 3                   'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert        'Hardware trigger line
frmDAQIO.cwaiPreScanDAQ.Configure                                           'Implement new settings
frmDAQIO.cwaiPreScanDAQ.start                                               'Start Data Acquisition

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    frmMain.txtTrigCount.Text = TriggerCnt
    Call Position

    'Check the Scan Watchdog Timer
    lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'Stop the motor
        gintAnomaly = 101                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Pre-Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiPreScanDAQ.Reset               'Reset the DAQ control

End Sub

Public Function ScanHomeIsComplete() As Boolean
'
'   PURPOSE: To deterimine if the Position Trigger Board is homed
'
'  INPUT(S): none
' OUTPUT(S): none

'Check to see if the Home Marker has been found
ScanHomeIsComplete = frmDAQIO.ReadDIOLine1(PORT5, 0)

End Function

Public Sub ScanForwardOnly()
'
'   PURPOSE: To scan the part in the forward direction only
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintMotorErr As Integer
Dim lblnTimeOut As Boolean
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngScanWatchDogTimer As Single
Dim lblnMoveDone As Boolean

Call ClearCounter                               'Clear trigger counts before each scan

'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)

'*** Forward Scan ***
frmMain.staMessage.Panels(1).Text = "System Message:  Acquiring Forward Data ..."

'Check to be sure that the Forward Scan start position is not within the scan region
Call CheckStartPosition("Forward")

If gintAnomaly Then Exit Sub                    'Exit on system error

'Check the status of the motor before the forward scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'Check the status of scanner home complete before the forward scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning Forward: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'The following data is sent to the PT Board before the FORWARD scan:
'               Start Scan          : Begin data acquisition
'               End Scan            : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.scanStart, gudtMachine.scanEnd, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.scanStart < gudtMachine.scanEnd Then
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd + gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd - gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Initialize watchdog timer`
lsngScanWatchDogTimer = Timer

'Calculate number of data points
gintMaxData = (gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution

'Setup the Forward Data Aqcuisition Control
frmDAQIO.cwaiFDAQ.NScans = gintMaxData                                  'Number of data points
frmDAQIO.cwaiFDAQ.NScansPerBuffer = gintMaxData * 3                     'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert    'Hardware trigger line
frmDAQIO.cwaiFDAQ.Configure                                             'Implement new settings
frmDAQIO.cwaiFDAQ.start                                                 'Start Data Acquisition

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    'Update the Trigger Counts
    frmMain.txtTrigCount.Text = TriggerCnt
    'Update the position
    lsngCurrentPosition = Position
    'Check the Scan Watchdog Timer
    lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiFDAQ.Reset                     'Reset the DAQ control

If gintAnomaly Then Exit Sub                'Exit on system error

'Verify that the motor has stopped
Do
    lsngFirstPosition = Position
    'Delay 10 msec for movement
    Call frmDAQIO.KillTime(10)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check the Scan Watchdog Timer
    lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until lblnTimeOut

If gintAnomaly Then Exit Sub                'Exit on system error

'1.6ANM if 700 series then find end scan and delay 100 ms
If gudtMachine.seriesID = "700" Then
    Call FindEndScan
    Call frmDAQIO.KillTime(100)
End If

'*** Update the Status Bar ***
frmMain.staMessage.Panels(1).Text = "System Message:  Forward DAQ Complete ..."

End Sub

Public Sub ScanUnit()
'
'   PURPOSE: To scan the part for electrical output
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintMotorErr As Integer
Dim lblnTimeOut As Boolean
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngScanWatchDogTimer As Single
Dim lblnMoveDone As Boolean

Call ClearCounter                               'Clear trigger counts before each scan

If (frmMain.mnuOptionsSSN.Checked = True) Or (frmMain.mnuOptionsSSD.Checked = True) Then   '3.0aANM
    'Set the Scan Velocity
    Call VIX500IE.SetVelocity(gudtMachine.scanVelocityB)
    'Set the Scan Acceleration
    Call VIX500IE.SetAcceleration(gudtMachine.scanAccelerationB)
    'Set the Scan Deceleration
    Call VIX500IE.SetDeceleration(gudtMachine.scanAccelerationB)
Else
    'Set the Scan Velocity
    Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
    'Set the Scan Acceleration
    Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
    'Set the Scan Deceleration
    Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)
End If

'*** Forward Scan ***
frmMain.staMessage.Panels(1).Text = "System Message:  Acquiring Forward Data ..."

'Check to be sure that the Forward Scan start position is not within the scan region
Call CheckStartPosition("Forward")

If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of the motor before the forward scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of scanner home complete before the forward scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning Forward: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'The following data is sent to the PT Board before the FORWARD scan:
'               Start Scan          : Begin data acquisition
'               End Scan            : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.scanStart, gudtMachine.scanEnd, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.scanStart < gudtMachine.scanEnd Then
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd + gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd - gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Initialize watchdog timer`
lsngScanWatchDogTimer = Timer

'Calculate number of data points
gintMaxData = (gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution

'Setup the Forward Data Aqcuisition Control
frmDAQIO.cwaiFDAQ.NScans = gintMaxData                                  'Number of data points
frmDAQIO.cwaiFDAQ.NScansPerBuffer = gintMaxData * 3                     'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert    'Hardware trigger line
frmDAQIO.cwaiFDAQ.Configure                                             'Implement new settings
frmDAQIO.cwaiFDAQ.start                                                 'Start Data Acquisition

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    'Update the Trigger Counts
    frmMain.txtTrigCount.Text = TriggerCnt
    'Update the position
    lsngCurrentPosition = Position
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiFDAQ.Reset                     'Reset the DAQ control

If gintAnomaly Then Exit Sub                'Exit on system error

'Verify that the motor has stopped
Do
    lsngFirstPosition = Position
    'Delay 30 msec for movement
    Call frmDAQIO.KillTime(30)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until lblnTimeOut

'1.4ANM if 700 series then find end scan and delay 100 ms
If (gudtMachine.seriesID = "700") And (gintAnomaly = 0) Then  '1.6ANM moved up and added Anomaly check
    Call FindEndScan
    Call frmDAQIO.KillTime(100)
End If

If gintAnomaly Then Exit Sub                        'Exit on system error

'Get the current readings '3.2ANM
'Switch to prog
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2
Call frmDAQIO.KillTime(50)

'Call MLX90277.GetCurrentW
gudtReading(0).mlxWCurrent = MyDev(lintDev1).GetIdd
gudtReading(1).mlxWCurrent = MyDev(lintDev2).GetIdd

'Switch to DAQ
Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2
Call frmDAQIO.KillTime(50)

'*** Reverse Scan ***
frmMain.staMessage.Panels(1).Text = "System Message:  Forward DAQ Complete ... Acquiring Reverse Data ..."

'Check to be sure that the Reverse Scan start position is not within the scan region
Call CheckStartPosition("Reverse")

'Check the status of the motor before the reverse scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of scanner home complete before the reverse scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning Reverse: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'The following data is sent to the PT Board before the REVERSE scan:
'               End Scan            : Begin data acquisition
'               Start Scan          : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.scanEnd, gudtMachine.scanStart, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.scanStart < gudtMachine.scanEnd Then
    Call VIX500IE.DefineMovement(gudtMachine.scanStart - gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.scanStart + gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Calculate number of data points
gintMaxData = (gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution

'Setup the Reverse Data Aqcuisition Control
frmDAQIO.cwaiRDAQ.NScans = gintMaxData                                  'Number of data points
frmDAQIO.cwaiRDAQ.NScansPerBuffer = gintMaxData * 3                     'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert    'Hardware trigger line
frmDAQIO.cwaiRDAQ.Configure                                             'Implement the new settings
frmDAQIO.cwaiRDAQ.start                                                 'Start the Data Acquisition

'Delay 50msec       '1.2ANM
Call frmDAQIO.KillTime(50)

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    'Update the Trigger Counts
    frmMain.txtTrigCount.Text = TriggerCnt
    'Update the position
    lsngCurrentPosition = Position
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 103                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Reverse Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiRDAQ.Reset                     'Reset the DAQ control

If gintAnomaly Then Exit Sub                'Exit on system error

'Verify that the motor has stopped
Do
    lsngFirstPosition = Position
    'Delay 30 msec for movement
    Call frmDAQIO.KillTime(30)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 103                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Reverse Scan.", True, True)
    End If
Loop Until lblnTimeOut

'Reset speed '3.0aANM
'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)

If gintAnomaly Then Exit Sub                'Exit on system error

frmMain.staMessage.Panels(1).Text = "System Message:  Reverse DAQ Complete."

End Sub

Public Sub ScanForce()
'
'   PURPOSE: To scan the part for force output only
'
'  INPUT(S): none
' OUTPUT(S): none
'1.7ANM new sub

Dim lintMotorErr As Integer
Dim lblnTimeOut As Boolean
Dim lsngFirstPosition As Single
Dim lsngCurrentPosition As Single
Dim lsngScanWatchDogTimer As Single
Dim lblnMoveDone As Boolean

Call ClearCounter                               'Clear trigger counts before each scan

If (frmMain.mnuOptionsSSN.Checked = True) Or (frmMain.mnuOptionsSSD.Checked = True) Then   '3.0aANM
    'Set the Scan Velocity
    Call VIX500IE.SetVelocity(gudtMachine.scanVelocityB)
    'Set the Scan Acceleration
    Call VIX500IE.SetAcceleration(gudtMachine.scanAccelerationB)
    'Set the Scan Deceleration
    Call VIX500IE.SetDeceleration(gudtMachine.scanAccelerationB)
Else
    'Set the Scan Velocity
    Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
    'Set the Scan Acceleration
    Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
    'Set the Scan Deceleration
    Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)
End If

'*** Forward Force Scan ***
frmMain.staMessage.Panels(1).Text = "System Message:  Acquiring Forward Force Data ..."

'Check to be sure that the Forward Scan start position is not within the scan region
Call CheckStartPosition("Forward")

If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of the motor before the forward scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of scanner home complete before the forward scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning Forward: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'The following data is sent to the PT Board before the FORWARD scan:
'               Start Scan          : Begin data acquisition
'               End Scan            : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.scanStart, gudtMachine.scanEnd, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.scanStart < gudtMachine.scanEnd Then
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd + gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.scanEnd - gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Initialize watchdog timer`
lsngScanWatchDogTimer = Timer

'Calculate number of data points
gintMaxData = (gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution

'Setup the Forward Data Aqcuisition Control
frmDAQIO.cwaiFTO.NScans = gintMaxData                                   'Number of data points
frmDAQIO.cwaiFTO.NScansPerBuffer = gintMaxData * 3                      'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert    'Hardware trigger line
frmDAQIO.cwaiFTO.Configure                                              'Implement new settings
frmDAQIO.cwaiFTO.start                                                  'Start Data Acquisition

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    'Update the Trigger Counts
    frmMain.txtTrigCount.Text = TriggerCnt
    'Update the position
    lsngCurrentPosition = Position
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiFTO.Reset                     'Reset the DAQ control

If gintAnomaly Then Exit Sub                'Exit on system error

'Verify that the motor has stopped
Do
    lsngFirstPosition = Position
    'Delay 30 msec for movement
    Call frmDAQIO.KillTime(30)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 102                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Forward Scan.", True, True)
    End If
Loop Until lblnTimeOut

'1.4ANM if 700 series then find end scan and delay 100 ms
If (gudtMachine.seriesID = "700") And (gintAnomaly = 0) Then  '1.6ANM moved up and added Anomaly check
    Call FindEndScan
    Call frmDAQIO.KillTime(100)
End If

If gintAnomaly Then Exit Sub                        'Exit on system error

'*** Reverse Force Scan ***
frmMain.staMessage.Panels(1).Text = "System Message:  Forward DAQ Complete ... Acquiring Reverse Force Data ..."

'Check to be sure that the Reverse Scan start position is not within the scan region
Call CheckStartPosition("Reverse")

'Check the status of the motor before the reverse scan
If Not VIX500IE.GetLinkStatus Then
    lintMotorErr = VIX500IE.ReadDriveFault
    gintAnomaly = lintMotorErr + 200                'Convert to system fault code
    Call ErrorLogFile("Motor Not Responding: Please Cycle Power to the Motor " & vbCrLf & _
                      "Control Box and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                        'Exit on system error

'Check the status of scanner home complete before the reverse scan
If Not (ScanHomeIsComplete) Then
    gintAnomaly = 301
    Call ErrorLogFile("Error Scanning Reverse: Position-Trigger Board has lost the Home Complete" & vbCrLf & _
                      "Signal.  Please cycle power to the Digitizer and Home the motor using the Function Menu", True, True)
End If
If gintAnomaly Then Exit Sub                'Exit on system error

'The following data is sent to the PT Board before the REVERSE scan:
'               End Scan            : Begin data acquisition
'               Start Scan          : Finish data acquisition
'               Counts per Trigger  : # pulses between triggers
Call WriteScanDataToPTBoard(gudtMachine.scanEnd, gudtMachine.scanStart, gudtMachine.countsPerTrigger, DEGPERREV)

'Define the position to move to for the forward scan
If gudtMachine.scanStart < gudtMachine.scanEnd Then
    Call VIX500IE.DefineMovement(gudtMachine.scanStart - gudtMachine.overTravel)
Else
    Call VIX500IE.DefineMovement(gudtMachine.scanStart + gudtMachine.overTravel)
End If

'Initialize the DAQ done control variable
gblnAnalogDone = False

'Calculate number of data points
gintMaxData = (gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution

'Setup the Reverse Data Aqcuisition Control
frmDAQIO.cwaiRFTO.NScans = gintMaxData                                  'Number of data points
frmDAQIO.cwaiRFTO.NScansPerBuffer = gintMaxData * 3                     'Total number of samples
frmDAQIO.cwDAQTools1.RouteSignal 1, cwrsPinPFI2, cwrsSourceAIConvert    'Hardware trigger line
frmDAQIO.cwaiRFTO.Configure                                             'Implement the new settings
frmDAQIO.cwaiRFTO.start                                                 'Start the Data Acquisition

'Delay 50msec       '1.2ANM
Call frmDAQIO.KillTime(50)

'Start the motor
Call VIX500IE.StartMotor

'Loop waiting for DAQ to be complete, or for a timeout condition
Do
    DoEvents
    'Update the Trigger Counts
    frmMain.txtTrigCount.Text = TriggerCnt
    'Update the position
    lsngCurrentPosition = Position
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 103                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Reverse Scan.", True, True)
    End If
Loop Until gblnAnalogDone Or lblnTimeOut

frmDAQIO.cwaiRFTO.Reset                     'Reset the DAQ control

If gintAnomaly Then Exit Sub                'Exit on system error

'Verify that the motor has stopped
Do
    lsngFirstPosition = Position
    'Delay 30 msec for movement
    Call frmDAQIO.KillTime(30)
    'Check to see if the move has completed
    lsngCurrentPosition = Position
    lblnMoveDone = (lsngFirstPosition = lsngCurrentPosition)
    'Exit the loop if the motor has stopped
    If lblnMoveDone Then Exit Do
    'Check the Scan Watchdog Timer '3.0aANM
    If gblnTLScanner Then
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUTTL
    Else
        lblnTimeOut = (Timer - lsngScanWatchDogTimer) > SCANTIMEOUT
    End If
    If lblnTimeOut Then
        Call VIX500IE.StopMotor                         'STOP THE MOTOR!
        gintAnomaly = 103                               'Set the anomaly number
        'Log the error to the error log and display the error message
        Call ErrorLogFile("The Data Acquisition was unable to complete during the Reverse Scan.", True, True)
    End If
Loop Until lblnTimeOut

'Reset speed '3.0aANM
'Set the Scan Velocity
Call VIX500IE.SetVelocity(gudtMachine.scanVelocity)
'Set the Scan Acceleration
Call VIX500IE.SetAcceleration(gudtMachine.scanAcceleration)
'Set the Scan Deceleration
Call VIX500IE.SetDeceleration(gudtMachine.scanAcceleration)

If gintAnomaly Then Exit Sub                'Exit on system error

frmMain.staMessage.Panels(1).Text = "System Message:  Reverse DAQ Complete."

End Sub

Public Sub SendProgrammingResultsToPLC()
'
'   PURPOSE:    To send the results of the Programming Operation to the PLC.
'
'  INPUT(S):    None.
' OUTPUT(S):    None.

Dim lblnPartGood As Boolean
Dim lsngStartTimer As Single
Dim lsngWatchDogTimer As Single
Dim lintLoopCnt As Integer
Dim lintReadBack As Integer
Dim lblnReadBackOK As Boolean

If InStr(command$, "NOHARDWARE") = 0 Then

    'Note:
    '  For Systems using DDE,
    '   ResultsCode=900 represents a part that was programmed successfully
    '   ResultsCode=999 represents a part that was scanned successfully.
    '   Any other resultsCode represents a bad part.
    '  For Systems using TTL,
    '   Programming Results are not sent to the PLC (Scan Results Only)
    '   A good part is represented by both results lines turned OFF.
    '   A bad part is represented by any results lines turned ON.

    If (gintAnomaly = 0) And (Not gblnProgFailure) Then
        'Remember that the part is good
        lblnPartGood = True
        If gudtMachine.PLCCommType = pctDDE Then
            'Send 900 (Programming Successful) to the PLC Results Register
            Call frmDDE.WriteDDEOutput(ResultsCode, 900)
        ElseIf gudtMachine.PLCCommType = pctTTL Then
            'No action for TTL systems
        End If
    Else
        'Remember that the part is bad
        lblnPartGood = False
    End If

    'Initialize watchdog timer
    lsngStartTimer = Timer
    'Initialize Readback to bad
    lblnReadBackOK = False

    'Verify the correct results were written if using DDE communication
    If gudtMachine.PLCCommType = pctDDE Then
        Do
            'Record number of loop iterations (for debug use only)
            lintLoopCnt = lintLoopCnt + 1

            'Readback the part results code
            lintReadBack = frmDDE.ReadDDEOutput(ResultsCode)

            If lblnPartGood Then
                'Verify that the ResultsCode = 900 (Programming Successful)
                If lintReadBack = 900 Then
                    lblnReadBackOK = True   'ResultsCode Correct
                    Exit Do                 'Exit Loop
                Else
                    'Send 900 (Programming Successful) to the PLC Results Register
                    Call frmDDE.WriteDDEOutput(ResultsCode, 900)
                End If
            Else
                'Verify that the ResultsCode <> 900 (Programming Not Successful)
                If lintReadBack <> 900 Then
                    lblnReadBackOK = True   'ResultsCode Correct
                    Exit Do                 'Exit Loop
                Else
                    'Send 300 (Programming Not Successful) to the PLC Results Register
                    Call frmDDE.WriteDDEOutput(ResultsCode, 300)
                End If
            End If

            'Get timer value
            lsngWatchDogTimer = Timer - lsngStartTimer

            If (lsngWatchDogTimer) > 0.1 Then       'Timeout in 100 msec
                gintAnomaly = 33                    'gintAnomaly = 33 will force the loop to finish
                'Log the error to the error log and display the error message
                Call ErrorLogFile("PLC Communication Timeout." & vbCrLf & _
                                  "Programming Results not received by PLC.", True, True)
            End If
        Loop While gintAnomaly <> 33

        'Only Send the Calc Complete signal if we've ensured that the proper
        'Good/Bad code is in the ResultsCode PLC register
        If lblnReadBackOK Then
            'Send Calc Complete to PLC via DDE
            Call frmDDE.WriteDDEOutput(CalcComplete, 1)
        End If
    'If using TTL, complete handshaking
    ElseIf gudtMachine.PLCCommType = pctTTL Then
        'No action for TTL systems
    End If

End If

End Sub

Public Sub SendScanResultsToPLC()
'
'   PURPOSE:    To send the results of the Scanning Operation to the PLC.
'
'  INPUT(S):    None.
' OUTPUT(S):    None.

Dim lblnPartGood As Boolean
Dim lsngStartTimer As Single
Dim lsngWatchDogTimer As Single
Dim lintLoopCnt As Integer
Dim lintReadBack As Integer
Dim lblnReadBackOK As Boolean

If InStr(command$, "NOHARDWARE") = 0 Then

    'Note:
    '  For Systems using DDE,
    '   ResultsCode=900 represents a part that was programmed successfully
    '   ResultsCode=999 represents a part that was scanned successfully.
    '   Any other resultsCode represents a bad part.
    '  For Systems using TTL,
    '   Programming Results are not sent to the PLC (Scan Results Only)
    '   A good part is represented by both results lines turned OFF.
    '   A bad part is represented by any results lines turned ON.

    If (gintAnomaly = 0) And (Not gblnProgFailure) And (Not gblnScanFailure) And (Not gblnSevere) Then    '1.8ANM added ProgFail
        'Remember that the part is good
        lblnPartGood = True
        If gudtMachine.PLCCommType = pctDDE Then
            'Send 999 (Scanning Successful) to the PLC Results Register
            If gudtMachine.seriesID <> "705" Then
                Call frmDDE.WriteDDEOutput(ResultsCode, 999)
            Else
                Call frmDDE.WriteDDEOutput(ResultsCode, 1)
            End If
        ElseIf gudtMachine.PLCCommType = pctTTL Then
            'Turn off appropriate bits
            Call frmDAQIO.OffPort2(PORT7, BIT4 + BIT5)
        End If
    Else
        'Remember that the part is bad
        lblnPartGood = False
    End If

    'Verify the correct results were written if using DDE communication
    If gudtMachine.PLCCommType = pctDDE Then
        'Initialize watchdog timer
        lsngStartTimer = Timer
        'Initialize Readback to bad
        lblnReadBackOK = False

        Do
            'Record number of loop iterations (for debug use only)
            lintLoopCnt = lintLoopCnt + 1

            'Readback the part results code
            lintReadBack = frmDDE.ReadDDEOutput(ResultsCode)
            If gudtMachine.seriesID = "705" Then lintReadBack = lintReadBack + 998
            
            If lblnPartGood Then
                'Verify that the ResultsCode = 999 (Scanning Successful)
                If lintReadBack = 999 Then
                    lblnReadBackOK = True   'ResultsCode Correct
                    Exit Do                 'Exit Loop
                Else
                    'Send 999 (Scanning Successful) to the PLC Results Register
                    If gudtMachine.seriesID <> "705" Then
                        Call frmDDE.WriteDDEOutput(ResultsCode, 999)
                    Else
                        Call frmDDE.WriteDDEOutput(ResultsCode, 1)
                    End If
                End If
            Else
                'Verify that the ResultsCode <> 999 (Scanning Not Successful)
                If lintReadBack <> 999 Then
                    lblnReadBackOK = True   'ResultsCode Correct
                    Exit Do                 'Exit Loop
                Else
                    'Send 301 (Scanning Not Successful) to the PLC Results Register
                    If gudtMachine.seriesID <> "705" Then
                        Call frmDDE.WriteDDEOutput(ResultsCode, 301)
                    Else
                        Call frmDDE.WriteDDEOutput(ResultsCode, 0)
                    End If
                End If
            End If

            'Get timer value
            lsngWatchDogTimer = Timer - lsngStartTimer
    
            If (lsngWatchDogTimer) > 0.1 Then       'Timeout in 100 msec
                gintAnomaly = 34                    'gintAnomaly = 34 will force the loop to finish
                'Log the error to the error log and display the error message
                Call ErrorLogFile("PLC Communication Timeout." & vbCrLf & _
                                  "Scanning Results not received by PLC.", True, True)
            End If
        Loop While gintAnomaly <> 34

        'Only Send the Calc Complete signal if we've ensured that the proper
        'Good/Bad code is in the ResultsCode PLC register
        If lblnReadBackOK Then
            'Send Calc Complete to PLC via DDE
            Call frmDDE.WriteDDEOutput(CalcComplete, 1)
        End If

    'If using TTL, complete handshaking
    ElseIf gudtMachine.PLCCommType = pctTTL Then

        If (gintAnomaly <> 0) Or gblnSevere Then
            'Severe failure or anomaly
            Call frmDAQIO.OnPort2(PORT8, BIT0)    'Set scan complete low (Complete)
            Call frmDAQIO.OffPort2(PORT8, BIT4)   'Set severe failure high (Severe)
        End If

        'Send Calc Complete to PLC
        Call frmDAQIO.OffPort2(PORT8, BIT1)     'Set calc complete high(0V)
        Call frmDAQIO.KillTime(400)             'Delay 400 msec
        Call frmDAQIO.OnPort2(PORT8, BIT1)      'Set calc complete low(5V)

        'Reset handshaking lines
        Call frmDAQIO.OnPort2(PORT7, BIT4)      'Reset Part #1 Results Bit #1 (Bad Part)
        Call frmDAQIO.OnPort2(PORT7, BIT5)      'Reset Part #1 Results Bit #2 (Bad Part)
        Call frmDAQIO.OnPort2(PORT7, BIT6)      'Reset Part #2 Results Bit #1 (Bad Part)
        Call frmDAQIO.OnPort2(PORT7, BIT7)      'Reset Part #2 Results Bit #2 (Bad Part)
        Call frmDAQIO.OnPort2(PORT8, BIT4)      'Reset Severe bit (Not Severe)
    End If

End If

End Sub

Public Sub SendSerialNumberToPLC()
'
'   PURPOSE:    To send the serial number and date code of the part to the PLC
'               via DDE.  The data must be parsed into the desired format.
'               This data is sent to the PLC in five (5) words, as follows:
'
'               Part ID #1 => MLX Lot Bits 0 - 15
'               Part ID #2 => MLX Wafer Bits 0 - 4 & MLX Lot Bits 16 - 17
'               Part ID #3 => MLX X Bits 0 - 6 & MLX Y Bits 0 - 6
'               Part ID #4 => Date Code Julian Date
'               Part ID #5 => Date Code Year & Shift
'
'   INPUT(S):   None
'  OUTPUT(S):   None

Dim lintSerialNumberAndDateCode(1 To 5) As Integer
Dim lintWordNum As Integer
Dim lsngStartTimer As Single
Dim lsngWatchDogTimer As Single
Dim lintLoopCnt As Integer
Dim lintReadBack As Integer
Dim lblnReadBackOK As Boolean

If InStr(command$, "NOHARDWARE") = 0 Then

    'Create the Encoded Serial Number & Date Code Words for transmission to the PLC
    Call EncodeSerialNumberAndDateCode(gudtMLX90277(1).Read.Lot, gudtMLX90277(1).Read.Wafer, gudtMLX90277(1).Read.X, gudtMLX90277(1).Read.Y, gstrDateCode, lintSerialNumberAndDateCode())

    For lintWordNum = 1 To 5
        'PartInfo 1 = 6, Part Info 2 = 7, etc.
        'So offset the WordNum by 6
        Call frmDDE.WriteDDEOutput(lintWordNum + 6, lintSerialNumberAndDateCode(lintWordNum))
    Next lintWordNum

    If (gintAnomaly = 0) Then

        For lintWordNum = 1 To 5
            'Initialize watchdog timer
            lsngStartTimer = Timer
            'Initialize Readback to bad
            lblnReadBackOK = False
            Do
                '***For Debug Only***
                lintLoopCnt = lintLoopCnt + 1               'Number of loop iterations

                'Readback the part number of the serial number
                'PartInfo 1 = 6, Part Info 2 = 7, etc.
                'So offset the WordNum by 5
                lintReadBack = frmDDE.ReadDDEOutput(lintWordNum + 6)

                'Get timer value
                lsngWatchDogTimer = Timer - lsngStartTimer

                'Compare data written to data read back
                If lintReadBack = lintSerialNumberAndDateCode(lintWordNum) Then
                    'Exit loop if data read back correctly
                    lblnReadBackOK = True
                    Exit Do
                ElseIf (lsngWatchDogTimer) > 0.05 Then      'Timeout in 50 msec
                    gintAnomaly = 35                        'gintAnomaly = 35 will force the loop to finish
                    'Log the error to the error log and display the error message
                    Call ErrorLogFile("PLC Communication Timeout." & vbCrLf & _
                                      "Serial Number not received by PLC.", True, True)
                End If
            Loop Until gintAnomaly = 35
        Next lintWordNum
    End If

End If

End Sub

Public Sub SelectFilter(channelNum As Integer, filterNum As Integer, Enable As Boolean)
'
'   PURPOSE:

'
'  INPUT(S):    channelNum : which channel filter to change, 0 to 3
'               filterNum : which filterNum to select, 1 to 6
'                           1 = Filter 1
'                           2 = Filter 2
'                           3 = Filter 3
'                           4 = Filter 4
'                           5 = Load 1
'                           6 = Load 2
'               Enable : TRUE  = Enable the filter
'                      : FALSE = Disable the filter
'
' OUTPUT(S):    None.
'
Dim lintPortNum As Integer
Dim lintBitNum As Integer

'Determine which port to write to
Select Case channelNum
    Case CHAN0  'Output #1
        lintPortNum = PORT3
    Case CHAN1  'Output #2
        lintPortNum = PORT4
    Case CHAN2  'Output #3  <-Programmer #1 Vdd for EE897
        lintPortNum = PORT5
    Case CHAN3  'Output #4  <-Programmer #2 Vdd for EE897
        lintPortNum = PORT1
End Select

'Determine which bit the filter is
Select Case filterNum
    Case 0    'External Load Card                '3.2aANM
        lintBitNum = BIT0                        '3.2aANM
    Case 1    'Filter 1   (Straight-Thru Path)
        lintBitNum = BIT2
    Case 2    'Filter 2
        lintBitNum = BIT3
    Case 3  'Filter 3
        lintBitNum = BIT4
    Case 4   'Filter 4
        lintBitNum = BIT5
    Case 5   'Load 1
        lintBitNum = BIT6
    Case 6    'Load 2
        lintBitNum = BIT7
End Select

'If enable is true, first turn off all filters other than the selected filter,
'then turn on the selected filter filter (this makes it so there is no loss
'of continuity if selecting the same filter as already selected)
If Enable Then
    If filterNum <> 0 Then                                  '3.2aANM
        'Disable all other paths
        Call frmDAQIO.OffPort2(lintPortNum, (BIT0 + BIT1 + BIT2 + BIT3 + BIT4 + BIT5 + BIT6 + BIT7) - lintBitNum)
        'Enable the selected path
        Call frmDAQIO.OnPort2(lintPortNum, lintBitNum)
        If (lintPortNum <> 1) And (lintPortNum <> 5) Then Call frmDAQIO.OffPort2(PORT2, BIT0)                 '3.2aANM
    Else                                                    '3.2aANM
        Call frmDAQIO.OffPort2(lintPortNum, (BIT0 + BIT1 + BIT2 + BIT3 + BIT4 + BIT5 + BIT6 + BIT7)) '3.2aANM
        If (lintPortNum <> 1) And (lintPortNum <> 5) Then Call frmDAQIO.OnPort2(PORT2, lintBitNum)            '3.2aANM
    End If                                                  '3.2aANM
'If enable is false, disable the selected filter
Else
    'Disable the selected path
    If filterNum <> 0 Then                                  '3.2aANM
        Call frmDAQIO.OffPort2(lintPortNum, lintBitNum)
    Else                                                    '3.2aANM
        Call frmDAQIO.OffPort2(PORT2, lintBitNum)           '3.2aANM
    End If                                                  '3.2aANM
End If

End Sub

Public Sub SetFailureStatus()
'
'   PURPOSE:    To set the global boolean variables for failures and
'               severe failures based on the global integer variables.
'
'
'  INPUT(S):    None.
' OUTPUT(S):    None.

Dim lintChanNum As Integer
Dim lintFaultNum As Integer
Dim lintSevereChanNum As Integer
Dim lintSevereFaultNum As Integer

'Initialize the variables
lintSevereChanNum = -1
lintSevereFaultNum = -1

'*** Check part parameters for failures ***
For lintChanNum = 0 To MAXCHANNUM     'Check every output of each part
    For lintFaultNum = 1 To MAXFAULTCNT             'Check every fault on each output of each part
        If gintSevere(lintChanNum, lintFaultNum) And Not gblnSevere Then
            gblnSevere = True                       'Severe failure occured
            gblnScanFailure = True                  'Failure occured
            lintSevereChanNum = lintChanNum         'Store channel number for severe message (subtract lintFirstChan to get track # of current part)
            lintSevereFaultNum = lintFaultNum       'Store first fault number for severe message
        End If
        If gintFailure(lintChanNum, lintFaultNum) Then
            gblnScanFailure = True                  'Failure occured
        End If
    Next lintFaultNum
Next lintChanNum

'Set the part status control
If gblnScanFailure Then
    frmMain.ctrStatus1.StatusOnText(2) = "REJECT"
    frmMain.ctrStatus1.StatusOnColor(2) = vbRed
Else
    frmMain.ctrStatus1.StatusOnText(2) = "GOOD"
    frmMain.ctrStatus1.StatusOnColor(2) = vbGreen
End If

'Turn the status control on
frmMain.ctrStatus1.StatusValue(2) = True

'Display a severe failure message if appropriate
If InStr(command$, "NOHARDWARE") = 0 Then               'If hardware is not present bypass logic
    If (lintSevereChanNum <> -1) Then
        If Not gblnReScanEnable Or gblnReScanRun Then
            If gudtMachine.PLCCommType = pctDDE Then
                Call frmDDE.WriteDDEOutput(StationFault, 1)     'Display a Station Fault on the PLC
            ElseIf gudtMachine.PLCCommType = pctTTL Then
                'Disable the Watchdog Timer
                Call frmDAQIO.OffPort2(PORT8, BIT7)
            End If
            'Display the Severe Message Box
            Call SevereMessage(lintSevereChanNum, lintSevereFaultNum)
            If gudtMachine.PLCCommType = pctDDE Then
                Call frmDDE.WriteDDEOutput(StationFault, 0)     'Clear the Station Fault on the PLC
            ElseIf gudtMachine.PLCCommType = pctTTL Then
                'Enable the Watchdog Timer
                Call frmDAQIO.OnPort2(PORT8, BIT7)
            End If
        End If
    End If
End If

End Sub

Public Sub StatsUpdateProgCounts()
'
'   PURPOSE:   To increment count values based how a unit passed or failed.
'
'              NOTE:  A priority for failures is established by the order of
'                     the elseif's below, as only the first true statement
'                     will be executed.  This means that only the first
'                     decoded failure will have its count incremented.  This
'                     is done so the sum of the individual failure counts
'                     will be equal to the total number of failures.
'
'  INPUT(S):
' OUTPUT(S):
'2.1ANM removed offset drift

Dim lintProgrammerNum As Integer
Dim lintFaultNum As Integer

If (gintAnomaly <> 0) Then              'Count No Test Units
    gudtProgSummary.totalNoTest = gudtProgSummary.totalNoTest + 1
Else
    If gblnProgFailure Then    'Count Failure if one occurred
        'Loop through the two programmers
        For lintProgrammerNum = 1 To 2
            'Increment the failure count (prioritized)
            If gintProgFailure(lintProgrammerNum, HIGHPROGINDEX1) Then
                gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, LOWPROGINDEX1) Then
                gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low = gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, HIGHPROGINDEX2) Then
                gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, LOWPROGINDEX2) Then
                gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low = gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, HIGHCLAMPLOW) Then
                gudtProgStats(lintProgrammerNum).clampLow.failCount.high = gudtProgStats(lintProgrammerNum).clampLow.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, LOWCLAMPLOW) Then
                gudtProgStats(lintProgrammerNum).clampLow.failCount.low = gudtProgStats(lintProgrammerNum).clampLow.failCount.low + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, HIGHCLAMPHIGH) Then
                gudtProgStats(lintProgrammerNum).clampHigh.failCount.high = gudtProgStats(lintProgrammerNum).clampHigh.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, LOWCLAMPHIGH) Then
                gudtProgStats(lintProgrammerNum).clampHigh.failCount.low = gudtProgStats(lintProgrammerNum).clampHigh.failCount.low + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, AGNDFAILURE) Then
                gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high = gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, FCKADJFAILURE) Then
                gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high = gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, CKANACHFAILURE) Then
                gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high = gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, CKDACCHFAILURE) Then
                gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high = gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high + 1
                Exit For
            ElseIf gintProgFailure(lintProgrammerNum, SLOWMODEFAILURE) Then
                gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high = gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high + 1
                Exit For
            End If
        Next lintProgrammerNum
    Else
        'Good Part counted for Stats: Update Running Average Codes
        '1.9ANM Call Solver90277.UpdateHistoryCodes
        gudtProgSummary.currentGood = gudtProgSummary.currentGood + 1   'XXX parts good
        gudtProgSummary.totalGood = gudtProgSummary.totalGood + 1       'Total parts good
    End If
End If
gudtProgSummary.currentTotal = gudtProgSummary.currentTotal + 1         'XXX part total
gudtProgSummary.totalUnits = gudtProgSummary.totalUnits + 1             'Total part total

End Sub

Public Sub StatsUpdateProgSums()
'
'   PURPOSE: To update the statistical sum information.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintProgrammerNum As Integer

'Loop through the two programmers
For lintProgrammerNum = 1 To 2

    'Index 1 (Idle) Values
    If gudtSolver(lintProgrammerNum).FinalIndexVal(1) > gudtProgStats(lintProgrammerNum).indexVal(1).max Then
        gudtProgStats(lintProgrammerNum).indexVal(1).max = gudtSolver(lintProgrammerNum).FinalIndexVal(1)
    End If
    If gudtSolver(lintProgrammerNum).FinalIndexVal(1) < gudtProgStats(lintProgrammerNum).indexVal(1).min Then
        gudtProgStats(lintProgrammerNum).indexVal(1).min = gudtSolver(lintProgrammerNum).FinalIndexVal(1)
    End If
    gudtProgStats(lintProgrammerNum).indexVal(1).sigma = gudtProgStats(lintProgrammerNum).indexVal(1).sigma + gudtSolver(lintProgrammerNum).FinalIndexVal(1)
    gudtProgStats(lintProgrammerNum).indexVal(1).sigma2 = gudtProgStats(lintProgrammerNum).indexVal(1).sigma2 + gudtSolver(lintProgrammerNum).FinalIndexVal(1) ^ 2
    gudtProgStats(lintProgrammerNum).indexVal(1).n = gudtProgStats(lintProgrammerNum).indexVal(1).n + 1

    'Index 1 (Idle) Locations
    If gudtSolver(lintProgrammerNum).FinalIndexLoc(1) > gudtProgStats(lintProgrammerNum).indexLoc(1).max Then
        gudtProgStats(lintProgrammerNum).indexLoc(1).max = gudtSolver(lintProgrammerNum).FinalIndexLoc(1)
    End If
    If gudtSolver(lintProgrammerNum).FinalIndexLoc(1) < gudtProgStats(lintProgrammerNum).indexLoc(1).min Then
        gudtProgStats(lintProgrammerNum).indexLoc(1).min = gudtSolver(lintProgrammerNum).FinalIndexLoc(1)
    End If
    gudtProgStats(lintProgrammerNum).indexLoc(1).sigma = gudtProgStats(lintProgrammerNum).indexLoc(1).sigma + gudtSolver(lintProgrammerNum).FinalIndexLoc(1)
    gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2 = gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2 + gudtSolver(lintProgrammerNum).FinalIndexLoc(1) ^ 2
    gudtProgStats(lintProgrammerNum).indexLoc(1).n = gudtProgStats(lintProgrammerNum).indexLoc(1).n + 1

    'Index 2 (WOT) Values
    If gudtSolver(lintProgrammerNum).FinalIndexVal(2) > gudtProgStats(lintProgrammerNum).indexVal(2).max Then
        gudtProgStats(lintProgrammerNum).indexVal(2).max = gudtSolver(lintProgrammerNum).FinalIndexVal(2)
    End If
    If gudtSolver(lintProgrammerNum).FinalIndexVal(2) < gudtProgStats(lintProgrammerNum).indexVal(2).min Then
        gudtProgStats(lintProgrammerNum).indexVal(2).min = gudtSolver(lintProgrammerNum).FinalIndexVal(2)
    End If
    gudtProgStats(lintProgrammerNum).indexVal(2).sigma = gudtProgStats(lintProgrammerNum).indexVal(2).sigma + gudtSolver(lintProgrammerNum).FinalIndexVal(2)
    gudtProgStats(lintProgrammerNum).indexVal(2).sigma2 = gudtProgStats(lintProgrammerNum).indexVal(2).sigma2 + gudtSolver(lintProgrammerNum).FinalIndexVal(2) ^ 2
    gudtProgStats(lintProgrammerNum).indexVal(2).n = gudtProgStats(lintProgrammerNum).indexVal(2).n + 1

    'Index 2 (WOT) Locations
    If gudtSolver(lintProgrammerNum).FinalIndexLoc(2) > gudtProgStats(lintProgrammerNum).indexLoc(2).max Then
        gudtProgStats(lintProgrammerNum).indexLoc(2).max = gudtSolver(lintProgrammerNum).FinalIndexLoc(2)
    End If
    If gudtSolver(lintProgrammerNum).FinalIndexLoc(2) < gudtProgStats(lintProgrammerNum).indexLoc(2).min Then
        gudtProgStats(lintProgrammerNum).indexLoc(2).min = gudtSolver(lintProgrammerNum).FinalIndexLoc(2)
    End If
    gudtProgStats(lintProgrammerNum).indexLoc(2).sigma = gudtProgStats(lintProgrammerNum).indexLoc(2).sigma + gudtSolver(lintProgrammerNum).FinalIndexLoc(2)
    gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2 = gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2 + gudtSolver(lintProgrammerNum).FinalIndexLoc(2) ^ 2
    gudtProgStats(lintProgrammerNum).indexLoc(2).n = gudtProgStats(lintProgrammerNum).indexLoc(2).n + 1

    'Clamp Low
    If gudtSolver(lintProgrammerNum).FinalClampLowVal > gudtProgStats(lintProgrammerNum).clampLow.max Then
        gudtProgStats(lintProgrammerNum).clampLow.max = gudtSolver(lintProgrammerNum).FinalClampLowVal
    End If
    If gudtSolver(lintProgrammerNum).FinalClampLowVal < gudtProgStats(lintProgrammerNum).clampLow.min Then
        gudtProgStats(lintProgrammerNum).clampLow.min = gudtSolver(lintProgrammerNum).FinalClampLowVal
    End If
    gudtProgStats(lintProgrammerNum).clampLow.sigma = gudtProgStats(lintProgrammerNum).clampLow.sigma + gudtSolver(lintProgrammerNum).FinalClampLowVal
    gudtProgStats(lintProgrammerNum).clampLow.sigma2 = gudtProgStats(lintProgrammerNum).clampLow.sigma2 + gudtSolver(lintProgrammerNum).FinalClampLowVal ^ 2
    gudtProgStats(lintProgrammerNum).clampLow.n = gudtProgStats(lintProgrammerNum).clampLow.n + 1

    'Clamp High
    If gudtSolver(lintProgrammerNum).FinalClampHighVal > gudtProgStats(lintProgrammerNum).clampHigh.max Then
        gudtProgStats(lintProgrammerNum).clampHigh.max = gudtSolver(lintProgrammerNum).FinalClampHighVal
    End If
    If gudtSolver(lintProgrammerNum).FinalClampHighVal < gudtProgStats(lintProgrammerNum).clampHigh.min Then
        gudtProgStats(lintProgrammerNum).clampHigh.min = gudtSolver(lintProgrammerNum).FinalClampHighVal
    End If
    gudtProgStats(lintProgrammerNum).clampHigh.sigma = gudtProgStats(lintProgrammerNum).clampHigh.sigma + gudtSolver(lintProgrammerNum).FinalClampHighVal
    gudtProgStats(lintProgrammerNum).clampHigh.sigma2 = gudtProgStats(lintProgrammerNum).clampHigh.sigma2 + gudtSolver(lintProgrammerNum).FinalClampHighVal ^ 2
    gudtProgStats(lintProgrammerNum).clampHigh.n = gudtProgStats(lintProgrammerNum).clampHigh.n + 1

    'Offset Code
    If gudtSolver(lintProgrammerNum).FinalOffsetCode > gudtProgStats(lintProgrammerNum).offsetCode.max Then
        gudtProgStats(lintProgrammerNum).offsetCode.max = gudtSolver(lintProgrammerNum).FinalOffsetCode
    End If
    If gudtSolver(lintProgrammerNum).FinalOffsetCode < gudtProgStats(lintProgrammerNum).offsetCode.min Then
        gudtProgStats(lintProgrammerNum).offsetCode.min = gudtSolver(lintProgrammerNum).FinalOffsetCode
    End If
    gudtProgStats(lintProgrammerNum).offsetCode.sigma = gudtProgStats(lintProgrammerNum).offsetCode.sigma + gudtSolver(lintProgrammerNum).FinalOffsetCode
    gudtProgStats(lintProgrammerNum).offsetCode.sigma2 = gudtProgStats(lintProgrammerNum).offsetCode.sigma2 + gudtSolver(lintProgrammerNum).FinalOffsetCode ^ 2
    gudtProgStats(lintProgrammerNum).offsetCode.n = gudtProgStats(lintProgrammerNum).offsetCode.n + 1

    'Rough Gain Code
    If gudtSolver(lintProgrammerNum).FinalRGCode > gudtProgStats(lintProgrammerNum).roughGainCode.max Then
        gudtProgStats(lintProgrammerNum).roughGainCode.max = gudtSolver(lintProgrammerNum).FinalRGCode
    End If
    If gudtSolver(lintProgrammerNum).FinalRGCode < gudtProgStats(lintProgrammerNum).roughGainCode.min Then
        gudtProgStats(lintProgrammerNum).roughGainCode.min = gudtSolver(lintProgrammerNum).FinalRGCode
    End If
    gudtProgStats(lintProgrammerNum).roughGainCode.sigma = gudtProgStats(lintProgrammerNum).roughGainCode.sigma + gudtSolver(lintProgrammerNum).FinalRGCode
    gudtProgStats(lintProgrammerNum).roughGainCode.sigma2 = gudtProgStats(lintProgrammerNum).roughGainCode.sigma2 + gudtSolver(lintProgrammerNum).FinalRGCode ^ 2
    gudtProgStats(lintProgrammerNum).roughGainCode.n = gudtProgStats(lintProgrammerNum).roughGainCode.n + 1

    'Fine Gain Code
    If gudtSolver(lintProgrammerNum).FinalFGCode > gudtProgStats(lintProgrammerNum).fineGainCode.max Then
        gudtProgStats(lintProgrammerNum).fineGainCode.max = gudtSolver(lintProgrammerNum).FinalFGCode
    End If
    If gudtSolver(lintProgrammerNum).FinalFGCode < gudtProgStats(lintProgrammerNum).fineGainCode.min Then
        gudtProgStats(lintProgrammerNum).fineGainCode.min = gudtSolver(lintProgrammerNum).FinalFGCode
    End If
    gudtProgStats(lintProgrammerNum).fineGainCode.sigma = gudtProgStats(lintProgrammerNum).fineGainCode.sigma + gudtSolver(lintProgrammerNum).FinalFGCode
    gudtProgStats(lintProgrammerNum).fineGainCode.sigma2 = gudtProgStats(lintProgrammerNum).fineGainCode.sigma2 + gudtSolver(lintProgrammerNum).FinalFGCode ^ 2
    gudtProgStats(lintProgrammerNum).fineGainCode.n = gudtProgStats(lintProgrammerNum).fineGainCode.n + 1

    'Clamp Low Code
    If gudtSolver(lintProgrammerNum).FinalClampLowCode > gudtProgStats(lintProgrammerNum).clampLowCode.max Then
        gudtProgStats(lintProgrammerNum).clampLowCode.max = gudtSolver(lintProgrammerNum).FinalClampLowCode
    End If
    If gudtSolver(lintProgrammerNum).FinalClampLowCode < gudtProgStats(lintProgrammerNum).clampLowCode.min Then
        gudtProgStats(lintProgrammerNum).clampLowCode.min = gudtSolver(lintProgrammerNum).FinalClampLowCode
    End If
    gudtProgStats(lintProgrammerNum).clampLowCode.sigma = gudtProgStats(lintProgrammerNum).clampLowCode.sigma + gudtSolver(lintProgrammerNum).FinalClampLowCode
    gudtProgStats(lintProgrammerNum).clampLowCode.sigma2 = gudtProgStats(lintProgrammerNum).clampLowCode.sigma2 + gudtSolver(lintProgrammerNum).FinalClampLowCode ^ 2
    gudtProgStats(lintProgrammerNum).clampLowCode.n = gudtProgStats(lintProgrammerNum).clampLowCode.n + 1

    'Clamp High Code
    If gudtSolver(lintProgrammerNum).FinalClampHighCode > gudtProgStats(lintProgrammerNum).clampHighCode.max Then
        gudtProgStats(lintProgrammerNum).clampHighCode.max = gudtSolver(lintProgrammerNum).FinalClampHighCode
    End If
    If gudtSolver(lintProgrammerNum).FinalClampHighCode < gudtProgStats(lintProgrammerNum).clampHighCode.min Then
        gudtProgStats(lintProgrammerNum).clampHighCode.min = gudtSolver(lintProgrammerNum).FinalClampHighCode
    End If
    gudtProgStats(lintProgrammerNum).clampHighCode.sigma = gudtProgStats(lintProgrammerNum).clampHighCode.sigma + gudtSolver(lintProgrammerNum).FinalClampHighCode
    gudtProgStats(lintProgrammerNum).clampHighCode.sigma2 = gudtProgStats(lintProgrammerNum).clampHighCode.sigma2 + gudtSolver(lintProgrammerNum).FinalClampHighCode ^ 2
    gudtProgStats(lintProgrammerNum).clampHighCode.n = gudtProgStats(lintProgrammerNum).clampHighCode.n + 1

    'Offset Seed Code
    '1.9ANM \/\/
'    If gudtSolver(lintProgrammerNum).OffsetSeedCode > gudtProgStats(lintProgrammerNum).OffsetSeedCode.max Then
'        gudtProgStats(lintProgrammerNum).OffsetSeedCode.max = gudtSolver(lintProgrammerNum).OffsetSeedCode
'    End If
'    If gudtSolver(lintProgrammerNum).OffsetSeedCode < gudtProgStats(lintProgrammerNum).OffsetSeedCode.min Then
'        gudtProgStats(lintProgrammerNum).OffsetSeedCode.min = gudtSolver(lintProgrammerNum).OffsetSeedCode
'    End If
'    gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma = gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma + gudtSolver(lintProgrammerNum).OffsetSeedCode
'    gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2 = gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2 + gudtSolver(lintProgrammerNum).OffsetSeedCode ^ 2
'    gudtProgStats(lintProgrammerNum).OffsetSeedCode.n = gudtProgStats(lintProgrammerNum).OffsetSeedCode.n + 1
'
'    'Rough Gain Seed Code
'    If gudtSolver(lintProgrammerNum).RoughGainSeedCode > gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max Then
'        gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max = gudtSolver(lintProgrammerNum).RoughGainSeedCode
'    End If
'    If gudtSolver(lintProgrammerNum).RoughGainSeedCode < gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min Then
'        gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min = gudtSolver(lintProgrammerNum).RoughGainSeedCode
'    End If
'    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma = gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma + gudtSolver(lintProgrammerNum).RoughGainSeedCode
'    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2 = gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2 + gudtSolver(lintProgrammerNum).RoughGainSeedCode ^ 2
'    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n = gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n + 1
'
'    'Fine Gain Seed Code
'    If gudtSolver(lintProgrammerNum).FineGainSeedCode > gudtProgStats(lintProgrammerNum).FineGainSeedCode.max Then
'        gudtProgStats(lintProgrammerNum).FineGainSeedCode.max = gudtSolver(lintProgrammerNum).FineGainSeedCode
'    End If
'    If gudtSolver(lintProgrammerNum).FineGainSeedCode < gudtProgStats(lintProgrammerNum).FineGainSeedCode.min Then
'        gudtProgStats(lintProgrammerNum).FineGainSeedCode.min = gudtSolver(lintProgrammerNum).FineGainSeedCode
'    End If
'    gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma = gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma + gudtSolver(lintProgrammerNum).FineGainSeedCode
'    gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2 = gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2 + gudtSolver(lintProgrammerNum).FineGainSeedCode ^ 2
'    gudtProgStats(lintProgrammerNum).FineGainSeedCode.n = gudtProgStats(lintProgrammerNum).FineGainSeedCode.n + 1
    '1.9ANM /\/\
     
Next lintProgrammerNum

End Sub

Public Sub SummaryInitialization()
'
'   PURPOSE: To initialize the display of the Lot Summaries.
'
'  INPUT(S): none
' OUTPUT(S): none

'Define the frame captions
frmMain.ctrProgSummary.FrameCaption = "Programming Summary"
frmMain.ctrScanSummary.FrameCaption = "Scanning Summary"

'Define the Programming label captions
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbTotalUnits) = "Total"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbGoodUnits) = "Good"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbRejectedUnits) = "Reject"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbSevereUnits) = "Severe"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbSystemErrors) = "Error"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbCurrentYield) = "Yield"
frmMain.ctrProgSummary.LabelCaption(SummaryTextBox.stbLotYield) = "Lot Yield"

'Reset the Programming background colors
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbTotalUnits) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbGoodUnits) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbRejectedUnits) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbSevereUnits) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbSystemErrors) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbWhite
frmMain.ctrProgSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbWhite

'Define the Scanning label captions
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbTotalUnits) = "Total"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbGoodUnits) = "Good"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbRejectedUnits) = "Reject"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbSevereUnits) = "Severe"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbSystemErrors) = "Error"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbCurrentYield) = "Yield"
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbLotYield) = "Lot Yield"

'Reset the Scanning background colors
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbTotalUnits) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbGoodUnits) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbRejectedUnits) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbSevereUnits) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbSystemErrors) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbWhite
frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbWhite

End Sub

Public Sub TabulateFile(fileName, lstrRc() As String)
'
'   PURPOSE:    To convert file into table format.
'
'
'  INPUT(S):    fileName: String representing name of parameter file
'
' OUTPUT(S):    lstrRC(row,column): file converted to table

Dim lintFileNum As Integer          'Free file number
Dim lstrLine(999) As String         'Entire line read from parameter file
Dim lintRow As Integer              'Current row of parameter file being evaluated
Dim lintNumberOfRows As Integer     'Number of rows in parameter file
Dim lintColumn As Integer           'Column of parameter file
Dim i As Integer                    'Increment variable
Dim lintComma As Integer
Dim lintLastComma As Integer        'Last comma location
    
lintFileNum = FreeFile              'Next free file available
Open fileName For Input As #lintFileNum 'Open file for reading

'Read in every line of file
Do While Not EOF(lintFileNum)       'Read until end of file
    Line Input #lintFileNum, lstrLine(i): i = i + 1 'read in entire line of file
Loop

lintNumberOfRows = i - 1            'Number of rows in file

Close #lintFileNum                  'Close file

frmParamViewer.MSHFlexGrid1.Cols = 6                    'Set up 6 columns
frmParamViewer.MSHFlexGrid1.Rows = lintNumberOfRows + 1 'Set up number of rows
frmParamViewer.MSHFlexGrid1.ColWidth(0, 0) = 4500
frmParamViewer.MSHFlexGrid1.ColWidth(1, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(2, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(3, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(4, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(5, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(6, 0) = 1300

For lintRow = 0 To lintNumberOfRows     'Evaluate each row of file
lintLastComma = 0                       'Initialize to zero
    For lintColumn = 0 To 5             'Evaluate each column of file
        lintComma = InStr(lintLastComma + 1, lstrLine(lintRow), ",")
        If lintComma <> 0 Then lstrRc(lintRow, lintColumn) = Mid(lstrLine(lintRow), lintLastComma + 1, (lintComma - lintLastComma - 1))
        frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = lstrRc(lintRow, lintColumn)
        lintLastComma = lintComma
    Next lintColumn
Next lintRow

End Sub

Public Function TriggerCnt() As Long
'
'     PURPOSE:  To read one 24-bit counter from the Position Trigger Board and return
'               the number of trigger counts.
'
'    INPUT(S):  None.
'   OUTPUT(S):  Value => Number of trigger counts

Dim Address As Long
Dim Data As Variant
Dim LSBAddr As Integer, MSBAddr As Integer
Dim LSBData As Variant, MIDData As Variant, MSBData As Variant

'Get Trigger Counter Address
Address = (PTBASEADDRESS) + &H2         'Address of trigger cnt register
LSBAddr = Address And &HFF              'Get LSB address
MSBAddr = (Address \ BIT8) And &HFF     'Get MSB address
                                            
'Read Trigger Data
Call frmDAQIO.ReadPTBoardData(LSBAddr, MSBAddr, LSBData, MIDData, MSBData)

MIDData = MIDData * BIT8
MSBData = MSBData * BIT16

TriggerCnt = (MSBData + MIDData + LSBData)

End Function

Public Sub TunePedal()
'
'   PURPOSE: The executive which runs through the programming steps
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lintProgrammerNum As Integer
Dim lintChanNum As Integer
Dim lblnVotingError As Boolean
Dim lintAttemptNum As Integer
Dim lblnPLCStart As Boolean       '2.2ANM \/\/
Dim lblnTimeOut As Boolean
Dim lsngTimer As Single           '2.2ANM /\/\
Dim lstrSN As String              '3.0aANM
         
'Initialize the pedal failures
Call InitializeAndMaskProgFailures

'Enable VRef
Call frmDAQIO.OnPort1(PORT4, BIT1)

'Enable Programming paths
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2

'Enable the proper filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
Next lintChanNum

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

'Build the current date code
gstrDateCode = GetDateCode

'Proceed if communication with the PTC-04's is active
If gblnGoodPTC04Link Then
    'Try to read the chip twice (if necessary)
    For lintAttemptNum = 1 To 2
        frmMain.staMessage.Panels(1).Text = "System Message:  Reading Contents of IC's"
        'Read values back from EEprom
        If Not MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError) Then
            If lintAttemptNum = 2 Then      'Only assign error after two attempts
                gintAnomaly = 164
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", False, False)
                End If
            End If
        End If
        'Verify that there were no voting errors
        If lblnVotingError Then
            If lintAttemptNum = 2 Then      'Only assign error after two attempts
                gintAnomaly = 167
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & _
                                      "Verify correct Revision of ASIC is in use.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & "Verify correct Revision of ASIC is in use.", False, False)
                End If
            End If
        End If
        'Calculate the serial number from the DUT
        gstrSerialNumber = MLX90277.EncodePartID
        
        '3.0aANM MLX Checks
        gblnMLXOk = False
        
        If gstrSerialNumber = "00000000000" Then
            lstrSN = MLX90277.EncodePartID2 & " (#2)"
        Else
            lstrSN = gstrSerialNumber
        End If
        
        If (gudtMLX90277(1).Read.Lot = 0) Or (gudtMLX90277(2).Read.Lot = 0) Then
            gintAnomaly = 169 'Bad SN
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & _
                              "Verify Checkhead Connections.", True, True)
            Exit Sub
        Else
            If (gudtMLX90277(1).Read.MelexisLock = False) Or (gudtMLX90277(2).Read.MelexisLock = False) Or ((gudtMLX90277(1).Read.TC = 0) And (gudtMLX90277(1).Read.TCWin = 0) And (gudtMLX90277(1).Read.TC2nd = 0)) Or ((gudtMLX90277(2).Read.TC = 0) And (gudtMLX90277(2).Read.TCWin = 0) And (gudtMLX90277(2).Read.TC2nd = 0)) Then
                gintAnomaly = 181
                Call ErrorLogFile("Severe Programmer Error: MLX Lock or TC not set! Verify MLX ICs." & vbCrLf & "Please Tag " & lstrSN & " Part as Bad MLX Chip.", True, True)
            End If
            
            If gudtMLX90277(1).Read.Lot <> gudtMLX90277(2).Read.Lot Then
                gintAnomaly = 182
                Call ErrorLogFile("Severe Programmer Error: MLX Lot #s don't match! Verify MLX ICs." & vbCrLf & "Please Tag " & lstrSN & " Part as Bad MLX Chip.", True, True)
            End If
        End If
        If gintAnomaly = 0 Then gblnMLXOk = True
        
        'Display the S/N of the DUT
        frmMain.ctrSetupInfo1.PartNum = gstrSerialNumber
        'Verify that the serial number is non-zero
        If gstrSerialNumber = "00000000000" Then
            If lintAttemptNum = 2 Then          'Only assign error after two attempts
                gintAnomaly = 169
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & _
                                   "Verify Checkhead Connections.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & "Verify Checkhead Connections.", False, False)
                End If
            End If
        Else
            gblnGoodSerialNumber = True
        End If
        'Moved up from below '2.2ANM \/\/
        If gudtMLX90277(1).Read.MemoryLock Or gudtMLX90277(2).Read.MemoryLock Then
            gblnLockSkip = True
            GoTo LockSkip
            'gintAnomaly = 171
            ''Log the error to the error log and display the error message
            'Call ErrorLogFile("Programmer Error: EEPROM Locked!" & vbCrLf & _
            '                  "Cannot be Re-Programmed.", True, True)
            'Exit For
        Else '2.8ANM added all lockskip items
            gblnLockSkip = False
        End If
        '                     2.2ANM /\/\
        'Verify MLX clamp values '2.2ANM \/\/
        If (gudtMLX90277(1).Read.clampLow <> MLXCLAMP1) Or (gudtMLX90277(1).Read.clampHigh <> MLXCLAMP1) Or (gudtMLX90277(2).Read.clampLow <> MLXCLAMP2) Or (gudtMLX90277(2).Read.clampHigh <> MLXCLAMP2) Then
            If (gstrMLX90277Revision = "Cx") Then
                If ((gudtMLX90277(1).Read.clampLow = MLXCLAMPC) And (gudtMLX90277(1).Read.clampHigh = MLXCLAMPC) And (gudtMLX90277(2).Read.clampLow = MLXCLAMPC) And (gudtMLX90277(2).Read.clampHigh = MLXCLAMPC)) Then
                    GoTo SkipA
                End If
            End If
                
            If (Not gblnReClampEnable) Or (gblnReClampEnable And (gintAnomaly = 0) And gblnReClamp) Then       '2.8ANM Added if else block
                If gblnTLScanner Then
                    Dim lintResp As Integer
                    lintResp = MsgBox("The EEPROM Clamp Values Are Not Correct." & vbCrLf & _
                    "There are several possible problems." & vbCrLf & _
                    "Press OK to continue programming the part." & vbCrLf & _
                    "Press CANCEL to stop programming process.", vbOKCancel, "MLX EEPROM Error!")
                    If lintResp = vbCancel Then
                        gintAnomaly = 174
                        Call ErrorLogFile("Programmer Error: MLX Clamp Values Do NOT Match Default Values!", False, False)
                        Exit For
                    End If
                Else
                    gintAnomaly = 174
                    Call ErrorLogFile("Programmer Error: MLX Clamp Values Do NOT Match Default Values!", True, True)
                    gblnAdministrator = False
                    Do
                        Call MsgBox("The EEPROM Clamp Values Are Not Correct." & vbCrLf & _
                        "There are several potential problems:" & vbCrLf & _
                        " (1) The unit has been programmed before" & vbCrLf & _
                        " (2) The unit has a bad circuit board component." & vbCrLf & _
                        " (3) There was a problem with the pogo pin connections." & vbCrLf & _
                        " (4) There is a problem with the checkhead wiring." & vbCrLf & vbCrLf & _
                        "HAS MAINTENENCE BEEN PERFORMED RECENTLY ON THE CHECKHEAD CONNECTION?" & vbCrLf & _
                        " VERIFY OUTPUT #1 & OUTPUT #2 ARE NOT SWAPPED" & vbCrLf & vbCrLf & _
                        "After you press OK you must enter the password to continue.", vbOKOnly, "MLX EEPROM Error!")
                        'Show the password form
                        Beep
                        frmPassword.Show vbModal
                    Loop While Not gblnAdministrator
                    Exit For
                End If
            Else
                gintAnomaly = 174
                Call ErrorLogFile("Programmer Error: MLX Clamp Values Do NOT Match Default Values!", False, False)
            End If
        End If
SkipA:
        '                         2.2ANM /\/\
        'Exit the For...Next Loop if serial number was read successfully
        If gblnGoodSerialNumber Then Exit For
    Next lintAttemptNum
    'Get datecode from chip  '1.5ANM \/\/
    gstrDateCode2 = MLX90277.DecodeCustomerID(gudtMLX90277(1).Read.CustID)
    If (right(gstrDateCode, 1) <> "A") And (right(gstrDateCode, 1) <> "B") Then
        gstrPalletLoad = right(gstrDateCode2, 1)
        If (gstrPalletLoad = "A") Or (gstrPalletLoad = "B") Then
            gstrDateCode = left(gstrDateCode, (Len(gstrDateCode) - 1)) & gstrPalletLoad
        End If
    End If                   '1.5ANM /\/\
Else
    gintAnomaly = 168
    'Log the error to the error log and display the error message
    If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & _
                                "Verify Connections to Programmer.", True, True)
    Else
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & "Verify Connections to Programmer.", False, False)
    End If
End If

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Make sure that the read variables get transferred into the write variables
    Call MLX90277.CopyMLXReadsToMLXWrites(lintProgrammerNum)
Next lintProgrammerNum

'2.6ANM Setup UDB if needed \/\/
If gblnUseNewAmad Then
    gstrSubProcess = PROG
    
    'New AMAD calls
    gdbkDbKeys.DeviceInProcessID = Pedal.GetDipID
    
    'Insert Programming Record
    gdbkDbKeys.ProgrammingID = Pedal.GetProgrammingIdWithInsert
End If

'Proceed if there is no anomaly
If gintAnomaly = 0 Then
    'Run the Solver
    gblnLockedPart = False  '2.0ANM
    Call RunSolver
    If (gintAnomaly = 0) Or (gintAnomaly = 160) Then '2.0ANM allow prog errors to write/lock
        'Transfer the chosen codes to the MLX Write variables
        For lintProgrammerNum = 1 To 2
            gudtMLX90277(lintProgrammerNum).Write.Filter = gudtSolver(lintProgrammerNum).Filter
            gudtMLX90277(lintProgrammerNum).Write.InvertSlope = gudtSolver(lintProgrammerNum).InvertSlope
            gudtMLX90277(lintProgrammerNum).Write.Mode = gudtSolver(lintProgrammerNum).Mode
            gudtMLX90277(lintProgrammerNum).Write.FaultLevel = gudtSolver(lintProgrammerNum).FaultLevel
            gudtMLX90277(lintProgrammerNum).Write.offset = gudtSolver(lintProgrammerNum).FinalOffsetCode
            gudtMLX90277(lintProgrammerNum).Write.RGain = gudtSolver(lintProgrammerNum).FinalRGCode
            gudtMLX90277(lintProgrammerNum).Write.FGain = gudtSolver(lintProgrammerNum).FinalFGCode
            If gintAnomaly = 0 And (Not gblnProgFailure) Then '2.0ANM
                gudtMLX90277(lintProgrammerNum).Write.clampLow = gudtSolver(lintProgrammerNum).FinalClampLowCode
                gudtMLX90277(lintProgrammerNum).Write.clampHigh = gudtSolver(lintProgrammerNum).FinalClampHighCode
            Else
                gudtMLX90277(lintProgrammerNum).Write.clampLow = MINCLAMPCODE
                gudtMLX90277(lintProgrammerNum).Write.clampHigh = MINCLAMPCODE
            End If
            If gudtMachine.seriesID = "705" Then  '2.1ANM added if block
                gudtMLX90277(lintProgrammerNum).Write.CustID = MLX90277.Encode705CustomerID(gstrDateCode)
            Else
                gudtMLX90277(lintProgrammerNum).Write.CustID = MLX90277.EncodeCustomerID(gstrDateCode)
            End If
            'Lock the part if it is called for (and it almost always should be)
            If gblnLockICs Then '2.0ANM \/\/
                If (Not gblnLockRejects And (gintAnomaly <> 0)) Then gblnProgFailure = True '3.0aANM
                If gblnProgFailure Then
                    If gblnLockRejects Then
                        'Set MemLock to True
                        gudtMLX90277(lintProgrammerNum).Write.MemoryLock = True
                        gblnLockedPart = True
                    Else
                        If lintProgrammerNum = 1 Then MsgBox "Part was not Locked!", vbOKOnly, "Melexis Status"
                    End If
                Else
                    'Set MemLock to True
                    gudtMLX90277(lintProgrammerNum).Write.MemoryLock = True
                    gblnLockedPart = True
                End If
            Else
                If lintProgrammerNum = 1 Then MsgBox "Part was not Locked!", vbOKOnly, "Melexis Status"
            End If              '2.0ANM /\/\
            'Encode the new Write variables
            Call MLX90277.EncodeEEpromWrite(lintProgrammerNum)
        Next lintProgrammerNum
        'Write new codes to the EEPROM
        frmMain.staMessage.Panels(1).Text = "System Message:  Writing Calculate Codes to ICs"
        If MLX90277.WriteEEPROMBlockByRows(0, 7) Then
            'Delay 100 msec
            Call frmDAQIO.KillTime(100)
            'Read back the contents of the EEPROM
            frmMain.staMessage.Panels(1).Text = "System Message:  Verifying Contents of ICs"
            If Not MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError) Then
                gintAnomaly = 164
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", True, True)
            End If
            'Verify that there were no voting errors
            If lblnVotingError Then
                gintAnomaly = 167
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & _
                                  "Verify correct Revision of ASIC is in use.", True, True)
            End If
            For lintProgrammerNum = 1 To 2
                'Verify that the reads & writes match
                If Not MLX90277.CompareReadsAndWrites(lintProgrammerNum) Then
                    gintAnomaly = 166
                    'Log the error to the error log and display the error message
                    Call ErrorLogFile("Programmer Error: Reads do not match Writes." & vbCrLf & _
                                      "Verify correct Revision of ASIC is in use.", True, True)
                    Exit For
                 End If
            Next lintProgrammerNum
        Else
            gintAnomaly = 165
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Programmer Error: Error Writing to ASIC EEPROM.", True, True)
        End If
    Else
        If gblnReClampEnable And Not gblnReClamp Then
            'Send ReClamp
            Call frmDDE.WriteDDEOutput(ReClamp, 1)
            
            'Clear SSA
            Call frmDDE.WriteDDEOutput(StartScanAck, 0)
            
            'Delay 500ms
            Call frmDAQIO.KillTime(500)
            
            'Set timer
            lsngTimer = Timer
            
            'Wait for SS
            Do
                DoEvents
                lblnPLCStart = frmDDE.ReadDDEInput(StartScan)
                If (Timer - lsngTimer > 15) Then lblnTimeOut = True
            Loop Until lblnPLCStart Or lblnTimeOut
            
            If lblnTimeOut Then
                gintAnomaly = 169
                Call ErrorLogFile("PLC Error: PLC did not respond to ReClamp request!", True, True)
                GoTo LockSkip
            End If
            
            'Send SSA
            Call frmDDE.WriteDDEOutput(StartScanAck, 1)
            
            'Set boolean so we know we did it
            gblnReClamp = True
            gintAnomaly = 0
        Else
            MsgBox "Part was not Locked!", vbOKOnly, "Melexis Status"  '2.0ANM
        End If
    End If
Else '2.8ANM for reclamp when needed
    If gblnReClampEnable And Not gblnReClamp Then
        'Send ReClamp
        Call frmDDE.WriteDDEOutput(ReClamp, 1)
        
        'Clear SSA
        Call frmDDE.WriteDDEOutput(StartScanAck, 0)
        
        'Delay 500ms
        Call frmDAQIO.KillTime(500)
                
        'Set timer
        lsngTimer = Timer
        
        'Wait for SS
        Do
            DoEvents
            lblnPLCStart = frmDDE.ReadDDEInput(StartScan)
            If (Timer - lsngTimer > 15) Then lblnTimeOut = True
        Loop Until lblnPLCStart Or lblnTimeOut
        
        If lblnTimeOut Then
            gintAnomaly = 169
            Call ErrorLogFile("PLC Error: PLC did not respond to ReClamp request!", True, True)
            GoTo LockSkip
        End If
        
        'Send SSA
        Call frmDDE.WriteDDEOutput(StartScanAck, 1)
        
        'Set boolean so we know we did it
        gblnReClamp = True
        gintAnomaly = 0
    End If
End If

LockSkip:

'Get the current readings '2.9ANM
Call MLX90277.GetCurrent

frmMain.staMessage.Panels(1).Text = "System Message:  Programming Complete"

'Disable filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
Next lintChanNum

'Disable Programming paths
Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

'Disable VRef
Call frmDAQIO.OffPort1(PORT4, BIT1)

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

End Sub

Public Sub TunePedal90293()
'
'   PURPOSE: The executive which runs through the programming steps of 90293
'
'  INPUT(S): None
' OUTPUT(S): None
'3.6*ANM New Sub

Dim lintProgrammerNum As Integer
Dim lintChanNum As Integer
Dim lblnVotingError As Boolean
Dim lblnCRC As Boolean
Dim llngCRC As Long
Dim llngChip As Long
Dim lintAttemptNum As Integer
Dim lblnPLCStart As Boolean
Dim lblnTimeOut As Boolean
Dim lsngTimer As Single
Dim lstrSN As String
Dim lintMLXID(6) As Integer
Dim lintMLXID2(6) As Integer
Dim i As Long
Dim SN As String
Dim d As Double

On Error GoTo TP90293

'Initialize the pedal failures
Call InitializeAndMaskProgFailures

'Enable VRef
Call frmDAQIO.OnPort1(PORT4, BIT1)

'Enable Programming paths
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2

'Enable the proper filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
Next lintChanNum

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

'Build the current date code
gstrDateCode = GetDateCode

'Proceed if communication with the PTC-04's is active
If gblnGoodPTC04Link Then

    Call MyDev(lintDev1).DeviceReplaced
    llngChip = MyDev(lintDev1).Advanced.ReadChipVersion
    
    'Log the error to the error log and display the error message
    If llngChip <> 2 Then
        gintAnomaly = 175
        Call ErrorLogFile("Programmer Error: Invalid Chip Version Output 1!" & vbCrLf & "Verify Checkhead Connections.", True, True)
        Exit Sub
    End If
    
    Call MyDev(lintDev1).ReadFullDevice
    lblnCRC = MyDev(lintDev1).Advanced.CheckCRC(llngCRC)
            
    'Log the error to the error log and display the error message
    If Not lblnCRC Then
        gintAnomaly = 175
        Call ErrorLogFile("Programmer Error: Invalid CRC Output 1!" & vbCrLf & "Verify Checkhead Connections.", True, True)
        Exit Sub
    End If
        
    Call MyDev(lintDev2).DeviceReplaced
    llngChip = MyDev(lintDev2).Advanced.ReadChipVersion
    
    'Log the error to the error log and display the error message
    If llngChip <> 2 Then
        gintAnomaly = 175
        Call ErrorLogFile("Programmer Error: Invalid Chip Version Output 2!" & vbCrLf & "Verify Checkhead Connections.", True, True)
        Exit Sub
    End If
    
    Call MyDev(lintDev2).ReadFullDevice
    lblnCRC = MyDev(lintDev2).Advanced.CheckCRC(llngCRC)
            
    'Log the error to the error log and display the error message
    If Not lblnCRC Then
        gintAnomaly = 175
        Call ErrorLogFile("Programmer Error: Invalid CRC Output 2!" & vbCrLf & "Verify Checkhead Connections.", True, True)
        Exit Sub
    End If
            
    lintMLXID(0) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID0)
    lintMLXID(1) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID1)
    lintMLXID(2) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID2)
    lintMLXID(3) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID3)
    lintMLXID(4) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID4)
    
    gudtMLX90277(1).Read.X = (Int(lintMLXID(0) / BIT5) And &H7) + ((lintMLXID(1) And &H1F) * BIT3)
    gudtMLX90277(1).Read.Y = (Int(lintMLXID(1) / BIT5) And &H7) + ((lintMLXID(2) And &H1F) * BIT3)
    gudtMLX90277(1).Read.Wafer = (lintMLXID(0) And &H1F)
    d = CDbl(lintMLXID(4)) * BIT11
    gudtMLX90277(1).Read.Lot = (Int(lintMLXID(2) / BIT5) And &H7) + (lintMLXID(3) * BIT3) + d

    SN = MLX90293.EncodePartID
    'Calculate the serial number from the DUT
    gstrSerialNumber = SN

    'Set DC bits
    Call MLX90293.EncodeCustomerID90293(gstrDateCode)
    Call MyDev(lintDev1).SetEEParameterCode(CodeUSERID1, glngCUSTID1)
    Call MyDev(lintDev1).SetEEParameterCode(CodeUSERID2, glngCUSTID2)

    gblnMLXOk = True
    
    If gstrSerialNumber = "000000000000" Then
        lintMLXID(0) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID0)
        lintMLXID(1) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID1)
        lintMLXID(2) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID2)
        lintMLXID(3) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID3)
        lintMLXID(4) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID4)
    
        gudtMLX90277(1).Read.X = (Int(lintMLXID(0) / BIT5) And &H7) + ((lintMLXID(1) And &H1F) * BIT3)
        gudtMLX90277(1).Read.Y = (Int(lintMLXID(1) / BIT5) And &H7) + ((lintMLXID(2) And &H1F) * BIT3)
        gudtMLX90277(1).Read.Wafer = (lintMLXID(0) And &H1F)
        d = CDbl(lintMLXID(4)) * BIT11
        gudtMLX90277(1).Read.Lot = (Int(lintMLXID(2) / BIT5) And &H7) + (lintMLXID(3) * BIT3) + d
    
        lstrSN = MLX90293.EncodePartID
    Else
        lstrSN = gstrSerialNumber
    End If
    
    If gintAnomaly = 0 Then gblnMLXOk = True
    
    'Display the S/N of the DUT
    frmMain.ctrSetupInfo1.PartNum = gstrSerialNumber
    'Verify that the serial number is non-zero
    If gstrSerialNumber = "000000000000" Then
        gintAnomaly = 169
        'Log the error to the error log and display the error message
        If gblnReClamp Or Not gblnReClampEnable Then
            Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & "Verify Checkhead Connections.", True, True)
        Else
            Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & "Verify Checkhead Connections.", False, False)
        End If
    Else
        gblnGoodSerialNumber = True
    End If
    'Moved up from below
    gudtMLX90277(1).Read.MemoryLock = MyDev(lintDev1).GetEEParameterCode(CodeMEMLOCK)
    gudtMLX90277(2).Read.MemoryLock = MyDev(lintDev2).GetEEParameterCode(CodeMEMLOCK)
    If gudtMLX90277(1).Read.MemoryLock Or gudtMLX90277(2).Read.MemoryLock Then
        gblnLockSkip = True
        GoTo LockSkip
    Else 'added all lockskip items
        gblnLockSkip = False
    End If
    
    'Get datecode from chip
    'gstrDateCode2 = MLX90277.DecodeCustomerID(gudtMLX90277(1).Read.CustID)
Else
    gintAnomaly = 168
    'Log the error to the error log and display the error message
    If gblnReClamp Or Not gblnReClampEnable Then
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & "Verify Connections to Programmer.", True, True)
    Else
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & "Verify Connections to Programmer.", False, False)
    End If
End If

'Setup UDB if needed \/\/
If gblnUseNewAmad Then
    gstrSubProcess = PROG
    
    'New AMAD calls
    gdbkDbKeys.DeviceInProcessID = Pedal.GetDipID
    
    'Insert Programming Record
    gdbkDbKeys.ProgrammingID = Pedal.GetProgrammingIdWithInsert
End If

'Proceed if there is no anomaly
If gintAnomaly = 0 Then
    'Run the Solver
    Call MLX90293.RunSolver90293
End If

LockSkip:

'Get the current readings
gudtReading(0).mlxCurrent = MyDev(lintDev1).GetIdd
gudtReading(1).mlxCurrent = MyDev(lintDev2).GetIdd

frmMain.staMessage.Panels(1).Text = "System Message:  Programming Complete"

'Disable filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
Next lintChanNum

'Disable Programming paths
Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

'Disable VRef
Call frmDAQIO.OffPort1(PORT4, BIT1)

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

Exit Sub
TP90293:
    gintAnomaly = 175
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in Pedal.TunePedal90293: " & Err.Description, True, True)

End Sub

Public Sub ReWorkPedal()
'
'   PURPOSE: The executive which runs through the rework steps
'
'  INPUT(S): None
' OUTPUT(S): None
'1.5ANM new sub

Dim lintProgrammerNum As Integer
Dim lintChanNum As Integer
Dim lblnVotingError As Boolean
Dim lintAttemptNum As Integer

'Enable VRef
Call frmDAQIO.OnPort1(PORT4, BIT1)

'Enable Programming paths
Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2

'Enable the proper filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
Next lintChanNum

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

'Build the current date code
gstrDateCode = GetDateCode

'Proceed if communication with the PTC-04's is active
If gblnGoodPTC04Link Then
    'Try to read the chip twice (if necessary)
    For lintAttemptNum = 1 To 2
        frmMain.staMessage.Panels(1).Text = "System Message:  Reading Contents of IC's"
        'Read values back from EEprom
        If Not MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError) Then
            If lintAttemptNum = 2 Then      'Only assign error after two attempts
                gintAnomaly = 164
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", False, False)
                End If
            End If
        End If
        'Verify that there were no voting errors
        If lblnVotingError Then
            If lintAttemptNum = 2 Then      'Only assign error after two attempts
                gintAnomaly = 167
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & _
                                      "Verify correct Revision of ASIC is in use.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & "Verify correct Revision of ASIC is in use.", False, False)
                End If
            End If
        End If
        'Calculate the serial number from the DUT
        gstrSerialNumber = MLX90277.EncodePartID
        'Verify that the serial number is non-zero
        If gstrSerialNumber = "00000000000" Then
            If lintAttemptNum = 2 Then          'Only assign error after two attempts
                gintAnomaly = 169
                'Log the error to the error log and display the error message
                If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
                    Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & _
                                      "Verify Checkhead Connections.", True, True)
                Else
                    Call ErrorLogFile("Programmer Error: No Serial Number Found!" & vbCrLf & "Verify Checkhead Connections.", False, False)
                End If
            End If
        Else
            gblnGoodSerialNumber = True
        End If
        'Exit the For...Next Loop if serial number was read successfully
        If gblnGoodSerialNumber Then Exit For
    Next lintAttemptNum
Else
    gintAnomaly = 168
    'Log the error to the error log and display the error message
    If gblnReClamp Or Not gblnReClampEnable Then '2.2ANM
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & _
                                "Verify Connections to Programmer.", True, True)
    Else
        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & "Verify Connections to Programmer.", False, False)
    End If
End If

'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Make sure that the read variables get transferred into the write variables
    Call MLX90277.CopyMLXReadsToMLXWrites(lintProgrammerNum)
Next lintProgrammerNum

If gudtMLX90277(1).Read.MemoryLock Or gudtMLX90277(2).Read.MemoryLock Then
    gblnLockSkip = True
    GoTo LockSkip
    'gintAnomaly = 171
    ''Log the error to the error log and display the error message
    'Call ErrorLogFile("Programmer Error: EEPROM Locked!" & vbCrLf & _
    '                  "Cannot be Re-Programmed.", True, True)
Else '2.8ANM added all lockskip items
    gblnLockSkip = False
End If

'Proceed if there is no anomaly
If gintAnomaly = 0 Then
    If gintAnomaly = 0 Then
        'Transfer the chosen codes to the MLX Write variables
        For lintProgrammerNum = 1 To 2
            If gudtMachine.seriesID = "705" Then  '2.1ANM added if block
                gudtMLX90277(lintProgrammerNum).Write.CustID = MLX90277.Encode705CustomerID(gstrDateCode)
            Else
                gudtMLX90277(lintProgrammerNum).Write.CustID = MLX90277.EncodeCustomerID(gstrDateCode)
            End If
            'Encode the new Write variables
            Call MLX90277.EncodeEEpromWrite(lintProgrammerNum)
        Next lintProgrammerNum
        'Write new codes to the EEPROM
        frmMain.staMessage.Panels(1).Text = "System Message:  Writing Codes to ICs"
        If MLX90277.WriteEEPROMBlockByRows(0, 7) Then
            'Delay 100 msec
            Call frmDAQIO.KillTime(100)
            'Read back the contents of the EEPROM
            frmMain.staMessage.Panels(1).Text = "System Message:  Verifying Contents of ICs"
            If Not MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError) Then
                gintAnomaly = 164
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: Error Reading ASIC EEPROM.", True, True)
            End If
            'Verify that there were no voting errors
            If lblnVotingError Then
                gintAnomaly = 167
                'Log the error to the error log and display the error message
                Call ErrorLogFile("Programmer Error: EEPROM Voting Error." & vbCrLf & _
                                  "Verify correct Revision of ASIC is in use.", True, True)
            End If
            For lintProgrammerNum = 1 To 2
                'Verify that the reads & writes match
                If Not MLX90277.CompareReadsAndWrites(lintProgrammerNum) Then
                    gintAnomaly = 166
                    'Log the error to the error log and display the error message
                    Call ErrorLogFile("Programmer Error: Reads do not match Writes." & vbCrLf & _
                                      "Verify correct Revision of ASIC is in use.", True, True)
                    Exit For
                 End If
            Next lintProgrammerNum
        Else
            gintAnomaly = 165
            'Log the error to the error log and display the error message
            Call ErrorLogFile("Programmer Error: Error Writing to ASIC EEPROM.", True, True)
        End If
    End If
Else '2.8ANM for reclamp when needed
    If gblnReClampEnable And Not gblnReClamp Then
        'Send ReClamp
        Call frmDDE.WriteDDEOutput(ReClamp, 1)
        
        'Clear SSA
        Call frmDDE.WriteDDEOutput(StartScanAck, 0)
        
        'Delay 500ms
        Call frmDAQIO.KillTime(500)
        
        'Wait for start scan vars
        Dim lblnPLCStart As Boolean
        Dim lblnTimeOut As Boolean
        Dim lsngTimer As Single
        
        'Set timer
        lsngTimer = Timer
        
        'Wait for SS
        Do
            DoEvents
            lblnPLCStart = frmDDE.ReadDDEInput(StartScan)
            If (Timer - lsngTimer > 10) Then lblnTimeOut = True
        Loop Until lblnPLCStart Or lblnTimeOut
        
        If lblnTimeOut Then
            gintAnomaly = 169
            Call ErrorLogFile("PLC Error: PLC did not respond to ReClamp request!", True, True)
            GoTo LockSkip
        End If
        
        'Send SSA
        Call frmDDE.WriteDDEOutput(StartScanAck, 1)
        
        'Set boolean so we know we did it
        gblnReClamp = True
        gintAnomaly = 0
    End If
End If

LockSkip:

frmMain.staMessage.Panels(1).Text = "System Message:  ReWork Programming Complete"

'Disable filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
Next lintChanNum

'Disable Programming paths
Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

'Disable VRef
Call frmDAQIO.OffPort1(PORT4, BIT1)

'Delay for the relays to debounce
Call frmDAQIO.KillTime(50)

End Sub

Public Sub UpdateResultsCounts(GridNum As Integer, RowNum As Long, high As Variant, low As Variant)
'
'   PURPOSE: To write count information to the results tab
'
'  INPUT(S): gridNum = grid number to write to
'            rowNum = row number to write
'            high = number of times this parameter failed high
'            low = number of time this parameter failed low
' OUTPUT(S): none

frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 3) = high
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 4) = low

End Sub

Public Sub UpdateResultsData(GridNum As Integer, RowNum As Long, ValueAndLocation As Variant, results As ParameterResultsDisplay)
'
'   PURPOSE: To write analysis information to the results tab
'
'  INPUT(S): gridNum = grid number to write to
'            rowNum = row number to write
'            valueAndLocation = analysis results value and location
'            results = Good or Reject status for parameter
' OUTPUT(S): none

Dim lstrStatusText As String
Dim lccBackColor As ColorConstants

'Select the appropriate text and back color
If results = prdGood Then           'Good part
    lstrStatusText = "GOOD"
    lccBackColor = vbGreen
ElseIf results = prdReject Then     'Bad part
    lstrStatusText = "REJECT"
    lccBackColor = vbRed
ElseIf results = prdNotChecked Then 'Not Checked
    lstrStatusText = "REPORT ONLY"
    lccBackColor = vbWhite
ElseIf results = prdEmpty Then      'Empty
    lstrStatusText = ""
    lccBackColor = vbWhite
End If

'Set the value and location
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 1) = ValueAndLocation
'Set the part status text
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 2) = lstrStatusText
'Set the back color
frmMain.ctrResultsTabs1.CellColor(GridNum, RowNum, 2) = lccBackColor
'Set the Font to Bold
frmMain.ctrResultsTabs1.BoldText(GridNum, RowNum, 2) = True

End Sub

Public Sub UpdateName(GridNum As Integer, RowNum As Long, Name As String, Bold As Boolean, Alignment As AlignmentSettings)
'
'   PURPOSE: To write a row name (parameter name) to the results tab
'
'  INPUT(S): GridNum   = grid number to write to
'            RowNum    = row number to write
'            Name      = parameter type (name)
'            Bold      = whether or not to bold the text
'            Alignment = Alignment of that cell
'
' OUTPUT(S): none

'Set the Parameter Name
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 0) = Name
'Bold the label if called for
frmMain.ctrResultsTabs1.BoldText(GridNum, RowNum, 0) = Bold
'Set the alignment of the label
frmMain.ctrResultsTabs1.TextAlignment(GridNum, RowNum, 0) = Alignment

End Sub

Public Sub UpdateStatisticsCounts(GridNum As Integer, RowNum As Long, high As Variant, low As Variant)
'
'   PURPOSE: To write count information to the scan stats tab
'
'  INPUT(S): GridNum = grid number to write to
'            RowNum = row number to write
'            high = number of times this parameter failed high
'            low = number of time this parameter failed low
' OUTPUT(S): none

frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 7) = high
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 8) = low

End Sub

Public Sub UpdateStatisticsData(GridNum As Integer, RowNum As Long, Avg As Variant, Std As Variant, Cpk As Variant, Cp As Variant, RangeHigh As Variant, RangeLow As Variant)
'
'   PURPOSE: To write statistics data
'
'  INPUT(S): gridNum = grid number to write to
'            rowNum = row number to write
'            Avg = Average
'            Std = Standard Deviation
'            Cpk = Cpk
'            Cp = Cp
'            RangeHigh = Highest value for particular parameter
'            RangeLow = Lowest value for particular parameter
' OUTPUT(S): none

frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 1) = Avg
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 2) = Std
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 3) = Cpk
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 4) = Cp
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 5) = RangeHigh
frmMain.ctrResultsTabs1.Data(GridNum, RowNum, 6) = RangeLow

End Sub

Public Sub VerifySupplyVoltage()
'
'   PURPOSE: Measures the supply voltage and verifies that it is within the acceptable
'            range.
'
'  INPUT(S): None
'
' OUTPUT(S): None

Dim lvntVoltage As Variant
Dim llngSupply As Long

On Error GoTo VRef_Err

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    'Read VRef
    gsngVRef = frmDAQIO.ReadVRef

    'Check voltage against tolerance and adjust as required
    If (gsngVRef >= SUPPLYMAX) Or (gsngVRef <= SUPPLYMIN) Then  'Check voltage limits
        If gudtMachine.VRefMode = vrmSWControlled Then
            'If the system is setup for SW controlled VRef,
            'then adjust it now: gsngVRef will be reset and checked below.
            Call AdjustVRef
        End If
        'Check to see if we are outside of the VRef limits
        If gsngVRef >= SUPPLYMAX Then
            gintAnomaly = 4
            'Log the error to the error log and display the error message
            Call ErrorLogFile("The measured Reference Voltage is too high.  If this " & vbCrLf & _
                              "problem persists, reset the Reference Voltage using" & vbCrLf & _
                              "the Function Menu.", True, True)
        ElseIf gsngVRef <= SUPPLYMIN Then
            gintAnomaly = 5
            'Log the error to the error log and display the error message
            Call ErrorLogFile("The measured Reference Voltage is too low.  If this " & vbCrLf & _
                              "problem persists, reset the Reference Voltage using" & vbCrLf & _
                              "the Function Menu.", True, True)
        End If
    End If
End If

Exit Sub

VRef_Err:
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error while Verifing Supply Voltage: " & Err.Description, True, True)
End Sub

Public Function WritePTBoardHomeMode(HomeMode As Boolean) As Boolean
'
'     PURPOSE:  To put the PT Board in HomeMode
'
'    INPUT(S):  HomeMode = Whether or not to set the PT Board in HomeMode
'   OUTPUT(S):  Returns whether or not the PT board is in the proper mode

Dim llngAddress As Long
Dim lintLSBAddr As Integer
Dim lintMSBAddr As Integer
Dim lintLSBData As Integer

'Define the HomeMode Address
llngAddress = (PTBASEADDRESS) + &H3                 'HomeMode Address
lintLSBAddr = llngAddress And &HFF                  'LSB Address
lintMSBAddr = (llngAddress \ BIT8) And &HFF         'MSB Address

'A "1" in the LSB sets the PT Board to HomeMode, "0" is Normal Mode
If HomeMode Then
    lintLSBData = 1
Else
    lintLSBData = 0
End If

Call frmDAQIO.WritePTBoardData(lintLSBAddr, lintMSBAddr, lintLSBData, 0, 0)

'Delay 50 msec
Call frmDAQIO.KillTime(50)

'Return whther or not the PT board is in the proper Home/NoHome Mode
WritePTBoardHomeMode = (frmDAQIO.ReadDIOLine1(PORT5, 1) = HomeMode)

End Function

Public Sub WriteScanDataToPTBoard(StartScan As Single, EndScan As Single, countsPerTrigger As Integer, UnitsPerRev As Single)
'
'     PURPOSE:  To write StartScan, EndScan, & Counts/Trigger to the PT Board
'
'    INPUT(S):  StartScan        = StartScan position
'               EndScan          = EndScan position
'               CountsPerTrigger = Encoder counts per A/D trigger
'               UnitsPerRev      = Units per encoder revolution
'   OUTPUT(S):  None.

Dim llngAddress As Long
Dim llngData As Long
Dim lintLSBAddr As Integer
Dim lintMSBAddr As Integer
Dim lintLSBData As Integer
Dim lintMIDData As Integer
Dim lintMSBData As Integer

'Define the Start Scan Address
llngAddress = (PTBASEADDRESS) + &H1                         'End Scan Address on PT Board
lintLSBAddr = llngAddress And &HFF                          'LSB of the address
lintMSBAddr = (llngAddress \ BIT8) And &HFF                 'MSB of the address

llngData = (StartScan / UnitsPerRev) * gudtMachine.encReso  'Convert StartScan to Count Value
lintLSBData = (llngData And &HFF)                           'LSB of data
lintMIDData = (llngData \ BIT8) And &HFF                    'MID of data
lintMSBData = (llngData \ BIT16) And &HFF                   'MSB of data

'Send the Start Scan Data to the PT Board
Call frmDAQIO.WritePTBoardData(lintLSBAddr, lintMSBAddr, lintLSBData, lintMIDData, lintMSBData)

'Define the End Scan Address
llngAddress = (PTBASEADDRESS)                               'End Scan Address on PT Board
lintLSBAddr = llngAddress And &HFF                          'LSB of the address
lintMSBAddr = (llngAddress \ BIT8) And &HFF                 'MSB of the address

llngData = (EndScan / UnitsPerRev) * gudtMachine.encReso    'Convert EndScan to Count Value
lintLSBData = (llngData And &HFF)                           'LSB of data
lintMIDData = (llngData \ BIT8) And &HFF                    'MID of data
lintMSBData = (llngData \ BIT16) And &HFF                   'MSB of data

'Send the Data to the PT Board
Call frmDAQIO.WritePTBoardData(lintLSBAddr, lintMSBAddr, lintLSBData, lintMIDData, lintMSBData)

'Define the Counts/Trigger Address
llngAddress = (PTBASEADDRESS) + &H2                         'Counts/Trigger Address on PT Board
lintLSBAddr = llngAddress And &HFF                          'LSB of the address
lintMSBAddr = (llngAddress \ BIT8) And &HFF                 'MSB of the address

llngData = countsPerTrigger                                 'Counts/Trigger
lintLSBData = (llngData And &HFF)                           'LSB of data
lintMIDData = (llngData \ BIT8) And &HFF                    'MID of data

lintMSBData = (llngData \ BIT16) And &HFF                   'MSB of data

'Send the Data to the PT Board
Call frmDAQIO.WritePTBoardData(lintLSBAddr, lintMSBAddr, lintLSBData, lintMIDData, lintMSBData)

End Sub

Public Function ZFind() As Boolean
'
'   PURPOSE: To locate the z-channel of the encoder
'
'  INPUT(S): none
' OUTPUT(S): function returns whether or not the Z-Channel was found

Dim lsngStartTimer As Single
Dim lblnFoundZChannel As Boolean
Dim lblnTimeOut As Boolean

ZFind = False                                   'Default to False
Call frmDAQIO.OnPort1(PORT2, BIT1)              'Command search for Home Marker

lsngStartTimer = Timer                          'Start the "watchdog timer"

'Start the motor movement
If VIX500IE.GetLinkStatus Then Call VIX500IE.StartMotor

Do
    'Check the current position to see if the Home Marker has been found
    lblnFoundZChannel = ScanHomeIsComplete
    If lblnFoundZChannel Then                   'If the Home Marker is found...
        Call VIX500IE.StopMotor                 'Stop the motor immediately
        Call frmDAQIO.KillTime(100)             '1.5ANM
        ZFind = True
        frmMain.CWPosition.Pointers(1).Visible = True   'Position is now meaningful
    Else
        'Check for Timeout Condition
        lblnTimeOut = ((Timer - lsngStartTimer) > ZFINDTIMEOUT)
        If lblnTimeOut Then
            Call VIX500IE.StopMotor                'Stop the motor immediately
            MsgBox "The motor failed to find the Z-Channel of the encoder" _
                   & vbCrLf & "in the allotted time.  Click OK to continue.", _
                   vbOKOnly + vbCritical, "Z-Marker Not Found"
        End If
    End If
    Call Position                               'Update the position
    Call frmDAQIO.KillTime(10)                  'Needed to see correct position
Loop Until lblnFoundZChannel Or lblnTimeOut

Call frmDAQIO.OffPort1(PORT2, BIT1)             'Disable search for Home Marker

End Function

Public Sub CacheDipID()
'
'   PURPOSE: To Cache the DIP ID stored function
'
'  INPUT(S): none
' OUTPUT(S): none
'2.8ANM \/\/ new sub

On Error GoTo ERROR_CacheDipID

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetDipID"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 100
    End With
    
    cmd.Parameters(1) = "AAF24E4D-FF75-4BFB-BBD5-04D99B8B3F79"
    cmd.Parameters(2) = "12345678910"
    
    cmd.Execute
    
EXIT_CacheDipID:
    cmd.CommandTimeout = 30
    
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    
    Set cmd = Nothing
    
    Exit Sub
ERROR_CacheDipID:
    MsgBox "Error in CacheDipID:" & Err.number & "- " & Err.Description
    Resume EXIT_CacheDipID
    
End Sub

Public Function GetDipID() As String
'
'   PURPOSE: To return the DIP ID
'
'  INPUT(S): none
' OUTPUT(S): returns the the DIP ID

On Error GoTo ERROR_GetDipID

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetDipID"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProductID
    cmd.Parameters(2) = gstrSerialNumber    '2.7ANM
    
    cmd.Execute
    
    If IsNull(cmd.Parameters(3)) Then
        gconnAmad.Close
        GetDipID = InsertDeviceInProcess
    Else
        GetDipID = UpdateDipAttributeValues(cmd.Parameters(3))
    End If

EXIT_GetDipID:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Function
ERROR_GetDipID:
    MsgBox "Error in GetDipID:" & Err.number & "- " & Err.Description
        
    gintAnomaly = 1 '2.8ANM \/
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Run-Time Error in Pedal.GetDipID: " & Err.Description, False, True)

    Resume EXIT_GetDipID
    
End Function

Public Function GetProgrammingIdWithInsert() As String
On Error GoTo ERROR_GetProgrammingIdWithInsert

    'Test Variables
    Dim lstrOperator As String
    Dim lstrTemperature As String
    Dim lstrComment As String

    'Test Values
    lstrOperator = frmMain.ctrSetupInfo1.Operator
    lstrTemperature = frmMain.ctrSetupInfo1.Temperature
    lstrComment = frmMain.ctrSetupInfo1.Comment

    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter

    gconnAmad.Open

    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsProgrammingRecord"
        .CommandType = adCmdStoredProc
    End With

    cmd.Parameters(1) = gdbkDbKeys.DeviceInProcessID
    cmd.Parameters(2) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(3) = gdbkDbKeys.LotID
    cmd.Parameters(4) = gdbkDbKeys.TsopStartupID
    cmd.Parameters(5) = gdbkDbKeys.TSOP_ModeID
    cmd.Parameters(6) = DateTime.Now
    cmd.Parameters(7) = App.Major & "." & App.Minor & "." & App.Revision
    cmd.Parameters(8) = lstrOperator
    cmd.Parameters(9) = lstrTemperature
    cmd.Parameters(10) = lstrComment

    cmd.Execute

    GetProgrammingIdWithInsert = cmd.Parameters(11)

    gconnAmad.Close
    Set cmd = Nothing

EXIT_GetProgrammingIdWithInsert:
    Exit Function
ERROR_GetProgrammingIdWithInsert:
    MsgBox "Error in GetProgrammingIdWithInsert:" & Err.number & "- " & Err.Description
    Resume EXIT_GetProgrammingIdWithInsert
End Function

Public Function NewErrorLogFile(MessageText As String) As String
'
'   PURPOSE:   To output error and message to UDB
'
'  INPUT(S):   MessageText = Error message text
'
' OUTPUT(S):   NewErrorLogFile = Error Description from UDB
'

    Dim lstrType As String               'Error Type
    Dim lstrErrorDescription As String   'Error Description
    
    Dim rsAnomaly As New ADODB.Recordset 'TER_7/9/07 need recordset variable for anomaly info
    Dim bUndefinedAnomaly As Boolean     'TER_7/9/07 flag to indicate undefined anomaly
    
    gconnAmad.Open
    Call GetAnomalyInfo(rsAnomaly)
    If rsAnomaly.BOF And rsAnomaly.EOF Then
        gdbkDbKeys.AnomalyID = GetUndefinedAnomalyID
        bUndefinedAnomaly = True
        lstrType = "Undefined Anomaly Type"
        lstrErrorDescription = "Undefined Anomaly"
    Else
        bUndefinedAnomaly = False
        gdbkDbKeys.AnomalyID = rsAnomaly!AnomalyID
        lstrType = rsAnomaly!AnomalyType
        lstrErrorDescription = rsAnomaly!AnomalyDescription
    End If
    gconnAmad.Close
    
    If grsTsopAnomaly.RecordCount > 0 Then
        grsTsopAnomaly.MoveFirst
        Do Until grsTsopAnomaly.EOF
            grsTsopAnomaly.Delete
            grsTsopAnomaly.MoveNext
        Loop
    End If
    grsTsopAnomaly.AddNew
    grsTsopAnomaly!AnomalyMessage = lstrErrorDescription
    grsTsopAnomaly!AnomalyDateTime = Now()
    grsTsopAnomaly!Operator = frmMain.ctrSetupInfo1.Operator
    If bUndefinedAnomaly Then
        grsTsopAnomaly!UndefinedAnomalyNumber = gintAnomaly
    Else
        grsTsopAnomaly!UndefinedAnomalyNumber = Null
    End If
    grsTsopAnomaly.Update
    
    Call InsertTsopAnomaly
    
    grsTsopAnomaly.Delete
    
    NewErrorLogFile = lstrErrorDescription
    
End Function

Public Sub CheckForProgrammingFaultsTestDynamicMPC()
'
'     PURPOSE:  To check for programming faults and set the pass/fail boolean
'
'    INPUT(S):  None.
'   OUTPUT(S):  None.

Dim lintProgrammerNum As Integer
Dim lintFaultNum As Integer
Dim lsngIdealSlope As Single

'Check the Solver outputs for pass/fail
For lintProgrammerNum = 1 To 2
    'NOTE: The Index checks are based on the actual position WOT was programmed at:
    'Calculate the ideal slope to use in calculating Index limits based on actual locations
    lsngIdealSlope = (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(1).IdealValue) / (gudtSolver(lintProgrammerNum).Index(2).IdealLocation - gudtSolver(lintProgrammerNum).Index(1).IdealLocation)
    
    'Check Index 1 (Idle)
    'old
    'Call Calc.CheckFault(intProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, LOWPROGINDEX1, HIGHPROGINDEX1, gintProgFailure())
    
    'New variables to store High and Low Limits (Dynamic MPCs)
    gudtSolver(lintProgrammerNum).Index(1).low = gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope
    gudtSolver(lintProgrammerNum).Index(1).high = gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).Index(1).low, gudtSolver(lintProgrammerNum).Index(1).high, LOWPROGINDEX1, HIGHPROGINDEX1, gintProgFailure())
    
    'Check Index 2 (WOT)
    'old
    'Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, LOWPROGINDEX2, HIGHPROGINDEX2, gintProgFailure())
    
    'New variables to store High and Low Limits (Dynamic MPCs)
    gudtSolver(lintProgrammerNum).Index(2).low = gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope
    gudtSolver(lintProgrammerNum).Index(2).high = gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).Index(2).low, gudtSolver(lintProgrammerNum).Index(2).high, LOWPROGINDEX2, HIGHPROGINDEX2, gintProgFailure())
    
    'Check the Low Clamp
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue - gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue + gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, LOWCLAMPLOW, HIGHCLAMPLOW, gintProgFailure())
    'Check the High Clamp
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue - gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue + gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, LOWCLAMPHIGH, HIGHCLAMPHIGH, gintProgFailure())
    'Check Offset Drift Code
    '2.1ANM gintProgFailure(lintProgrammerNum, HIGHOFFSETDRIFT) = (gudtMLX90277(lintProgrammerNum).Read.Drift > gudtSolver(lintProgrammerNum).MaxOffsetDrift)
    'Check AGND Code
    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).MinAGND, gudtSolver(lintProgrammerNum).MaxAGND, gudtSolver(lintProgrammerNum).MinAGND, gudtSolver(lintProgrammerNum).MaxAGND, AGNDFAILURE, AGNDFAILURE, gintProgFailure())
    'Check Oscillator Adjust Code
    gintProgFailure(lintProgrammerNum, FCKADJFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.FCKADJ <> gudtSolver(lintProgrammerNum).FCKADJ)
    'Check Capacitor Frequency Adjust Code
    gintProgFailure(lintProgrammerNum, CKANACHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKANACH <> gudtSolver(lintProgrammerNum).CKANACH)
    'Check DAC Code Frequency Adjust Code
    gintProgFailure(lintProgrammerNum, CKDACCHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKDACCH <> gudtSolver(lintProgrammerNum).CKDACCH)
    'Check Slow Code
    gintProgFailure(lintProgrammerNum, SLOWMODEFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.SlowMode <> gudtSolver(lintProgrammerNum).SlowMode)
Next lintProgrammerNum

'Check each output
For lintProgrammerNum = 1 To 2
    'Check every fault on each output
    For lintFaultNum = 1 To PROGFAULTCNT
        If gintProgFailure(lintProgrammerNum, lintFaultNum) Then
            gblnProgFailure = True         'Failure occured
        End If
    Next lintFaultNum
Next lintProgrammerNum

'Set the part status control
If gblnProgFailure Then
    frmMain.ctrStatus1.StatusOnText(1) = "REJECT"
    frmMain.ctrStatus1.StatusOnColor(1) = vbRed
Else
    frmMain.ctrStatus1.StatusOnText(1) = "GOOD"
    frmMain.ctrStatus1.StatusOnColor(1) = vbGreen
End If

'Turn the status control on
frmMain.ctrStatus1.StatusValue(1) = True

End Sub

