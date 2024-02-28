Attribute VB_Name = "Series705"
'*********************************Series705.BAS*********************************
'
'   705 Series Specific Software, supplemental to Pedal.Bas.
'   This module should handle all 705 series production programmer/scanners,
'   test lab programmer/scanners, and database recall software.
'   The software is to be kept in the pedal software library, EE947.
'
'VER    DATE      BY   PURPOSE OF MODIFICATION                          TAG
'1.0  10/21/2005  ANM  First release per SCN 704T-001. (3102)         1.0ANM
'1.1  01/13/2006  ANM  Change for SCN# 704T-003 (3258).               1.1ANM
'1.2  01/31/2006  ANM  Adjusted excel file headers per SRC.           1.2ANM
'1.3  02/22/2006  ANM  Added testlab features per PR 11801-K.         1.3ANM
'1.4  03/02/2006  ANM  Update for TL/Customer prints.                 1.4ANM
'1.5  05/03/2006  ANM  Update per SCN# MISC-092 (3365). Moved         1.5ANM
'                      TL subs to testlab.bas
'1.6  05/04/2006  ANM  Updates per SCN# MISC-094 (3423).              1.6ANM
'1.7  06/01/2006  ANM  Updates per SCN# 705T-005 (3481).              1.7ANM
'1.8  08/17/2006  ANM  Updates per SCN# MISC-100 (3521).              1.8ANM
'1.9  12/05/2006  ANM  Updates per SCN# MISC-101 (3636).              1.9ANM
'2.0  01/18/2007  ANM  Updates for TR 8501-E (705 Prod.) and          2.0ANM
'                      SCN# MISC-102 (3702).
'2.0* 02/07/2007  ANM  Changed fwd force @ force knee to save only.   2.0*ANM
'2.1  02/16/2007  ANM  Changed VG limits at WOT to open.              2.1ANM
'2.2  08/05/2007  ANM  Update for new AMAD & SSPSSs.                  2.2ANM
'2.3  08/30/2007  ANM  Removed force knew per SCN# 705F-007 (3973).   2.3ANM
'                      Fixed v2.2 bug with force test only.
'2.4  11/02/2007  ANM  Update to remove/rename solver items per       2.4ANM
'                      SCN# 705F-011 (4018).
'2.5  04/29/2008  ANM  Update for stat fix and MLX current check.     2.5ANM
'2.6  05/02/2008  ANM  Add FO raw data save per SCN# 4139.            2.6ANM
'2.7  07/29/2008  ANM  Add ABS Lin per SCN# 4167.                     2.7ANM
'2.8  10/31/2008  ANM  Add P@R per SCN# 4236.                         2.8ANM
'2.8a 09/17/2009  ANM  Add ABSLin M/MLX I per SCN# 4392/4401.         2.8aANM
'2.8b 11/10/2009  ANM  Update slow speed & Bnmk Idd per 4420/4428.    2.8bANM
'2.8c 06/22/2010  ANM  Update MLX Idds per SCN# 4585.                 2.8cANM
'2.8d 01/07/2020  ANM  Update for KD.                                 2.8dANM
'

Option Explicit

'*****************
'*   Constants   *
'*****************

'File I/O Constants
Public Const STATEXT = ".lot"
Public Const STATFILEPATH = "D:\Data\705\LotData\"
Public Const DATAPATH = "D:\Data\"
Public Const DATA705PATH = "D:\Data\705\"
Public Const ERRORPATH = "D:\Data\705\Errors\"
Public Const ERRORLOG = "Error.log"
Public Const DATAEXT = ".csv"
Public Const PARTDATAPATH = "D:\Data\705\PartData\"
Public Const PARTPROGDATAPATH = "D:\Data\705\PartData\ProgrammingData\"
Public Const PARTSCANDATAPATH = "D:\Data\705\PartData\ScanData\"
Public Const PARTRAWDATAPATH = "D:\Data\705\PartData\ScanData\RawData\"
Public Const PAREXT = ".csv"
Public Const PARPATH = "\Parameter\"
Public Const PARTMLXDATAPATH = "D:\Data\705\PartData\ProgrammingData\MLXData\" '2.0ANM

'Display Constants
Public Const NUMROWSSCANRESULTSDISPLAY = 52     'Maximum number of rows in the Scan Results display
Public Const NUMROWSSCANSTATSDISPLAY = 40       'Maximum number of rows in the Scan Stats display
Public Const MAXGRAPHS = 3                      'Max number of outputs per graph

'Serial Port Contants
Public Const VIX500IEPORT = 4                   'Com Port for VIX500IE Motor Controller
Public Const SENSOTECPORT = 3                   'Com Port for Sensotec SC2000
Public Const PTC04PORT1 = 1                     'Com Port for the 1st PTC-03 Melexis Programmer
Public Const PTC04PORT2 = 7                     'Com Port for the 2nd PTC-03 Melexis Programmer

Public Const MAXCHANNUM = 1                     'Identifies outputs

'******************** Failure Definitions ********************
'
'   The failure arrays (gintFailure and gintSevere) are
'   defined as two dimensional INTEGER arrays.  The first
'   index in the two dimensional array identifies each
'   output (i.e. channel number). The second index
'   identifies the failure for that channel (i.e. fault
'   number).  This structure allows expandability for
'   adding failures in a project.
'
'   It is also important to note that the data contained in
'   the first element (i.e. faultNum = 0) contains the total
'   number of failures for that channel.
'
'   Hence,
'   the syntax for the array definitions would be as follows:
'
'               gintFailure(chanNum, faultNum)
'               gintSevere (chanNum, faultNum)
'
'Examples:
'   This sets a high index failure on channel 1:
'
'           gintFailure(1, HIGHINDEXPT1) = True
'
'   While this clears a high index failure on channel 1:
'
'           gintFailure(1, HIGHINDEXPT2) = False
'

Public Const HIGHINDEXPT1 = 1
Public Const LOWINDEXPT1 = 2

Public Const HIGHOUTPUTATFORCEKNEE = 3
Public Const LOWOUTPUTATFORCEKNEE = 4

Public Const HIGHINDEXPT2 = 5
Public Const LOWINDEXPT2 = 6

Public Const HIGHINDEXPT3 = 7
Public Const LOWINDEXPT3 = 8

Public Const HIGHMAXOUTPUT = 9
Public Const LOWMAXOUTPUT = 10

Public Const HIGHSINGLEPOINTLIN = 11
Public Const LOWSINGLEPOINTLIN = 12

Public Const HIGHSLOPE = 13
Public Const LOWSLOPE = 14

Public Const HIGHFWDOUTPUTCOR = 15
Public Const LOWFWDOUTPUTCOR = 16

Public Const HIGHREVOUTPUTCOR = 17
Public Const LOWREVOUTPUTCOR = 18

Public Const HIGHFORCEKNEELOC = 19
Public Const LOWFORCEKNEELOC = 20

Public Const HIGHFORCEKNEEFWDFORCE = 21
Public Const LOWFORCEKNEEFWDFORCE = 22

Public Const HIGHFWDFORCEPT1 = 23
Public Const LOWFWDFORCEPT1 = 24

Public Const HIGHFWDFORCEPT2 = 25
Public Const LOWFWDFORCEPT2 = 26

Public Const HIGHFWDFORCEPT3 = 27
Public Const LOWFWDFORCEPT3 = 28

Public Const HIGHREVFORCEPT1 = 29
Public Const LOWREVFORCEPT1 = 30

Public Const HIGHREVFORCEPT2 = 31
Public Const LOWREVFORCEPT2 = 32

Public Const HIGHREVFORCEPT3 = 33
Public Const LOWREVFORCEPT3 = 34

Public Const HIGHPEAKFORCE = 35
Public Const LOWPEAKFORCE = 36

Public Const HIGHMECHHYSTPT1 = 37
Public Const LOWMECHHYSTPT1 = 38

Public Const HIGHMECHHYSTPT2 = 39
Public Const LOWMECHHYSTPT2 = 40

Public Const HIGHMECHHYSTPT3 = 41
Public Const LOWMECHHYSTPT3 = 42

Public Const HIGHFCHYS = 43
Public Const LOWFCHYS = 44

Public Const HIGHMLXI = 45 '2.5ANM
Public Const LOWMLXI = 46  '2.5ANM

Public Const HIGHABSLIN = 47 '2.7ANM
Public Const LOWABSLIN = 48  '2.7ANM

Public Const HIGHPEDALATREST = 49 '2.8ANM
Public Const LOWPEDALATREST = 50  '2.8ANM

Public Const HIGHMLXI2 = 51 '2.8cANM
Public Const LOWMLXI2 = 52  '2.8cANM

Public Const HIGHKDSTART = 53 '2.8dANM
Public Const LOWKDSTART = 54  '2.8dANM

Public Const HIGHKDSTOP = 55  '2.8dANM
Public Const LOWKDSTOP = 56   '2.8dANM

Public Const HIGHKDSPAN = 57  '2.8dANM
Public Const LOWKDSPAN = 58   '2.8dANM

'NOTE:  MAXFAULTCNT must be set equal to the total number of faults defined:
Public Const MAXFAULTCNT = 58

'******************************************************************
'*                         TYPE DEFINTIONS                        *
'******************************************************************
'
'NOTE(s):   The arrays defined for the modular user-defined types
'           (i.e. prefix = gudt) represent each output (where
'           output = chanNum). Also, the arrays defined for the
'           parameter values in the ExtremeType and TestParameters
'           are used to evaluate multiple test regions (where
'           test region = regionNum).
'
'           Hence,
'           the syntax for the array definitions would be as follows:
'
'                gudtTest(chanNum).singlePtLin(regionNum).high
'
'           This type definition gives added flexiblity for index:
'                gudtTest(chanNum).index(1).ideal     => ideal FullClose on channel 1
'                gudtTest(chanNum).index(2).ideal     => ideal Midpoint on channel 1
'                gudtTest(chanNum).index(3).ideal     => ideal FullOpen on channel 1

'*** Last Level Type Definitions ***
Type HighLowInteger
    high            As Integer                  'High count value
    low             As Integer                  'Low  count value
End Type

Type HighLowSingle
    high            As Single                   'High single value
    low             As Single                   'Low single value
End Type

Type HighLowAndLocSingle
    high            As Single                   'High single value
    low             As Single                   'Low single value
    location        As Single                   'Location single value
End Type

Type ValueAndLocation
    Value           As Single                   'Single value
    location        As Single                   'Single location
End Type

Type StartAndStop
    start           As Single                   'Parameter start location
    stop            As Single                   'Parameter stop location
End Type

Type Statistics
    max             As Single                   'Maximum measured value
    min             As Single                   'Minimum measured value
    sigma           As Double                   'Sum of measured values
    sigma2          As Double                   'Sum of square of measured values
    n               As Integer                  'Number of samples measured values
    failCount       As HighLowInteger           'Prioritized failure counts
    totalFailCount  As HighLowInteger           'Non-Prioritized failure counts
End Type

Type PointTestWithRange
    location        As Single                   'Parameter location
    ideal           As Single                   'Parameter ideal value
    ideal2          As Single                   'Parameter ideal value
    high            As Single                   'Parameter high limit
    low             As Single                   'Parameter low limit
    start           As Single                   'Parameter start location
    stop            As Single                   'Parameter stop location
End Type

Type PointTest
    location        As Single                   'Parameter location
    ideal           As Single                   'Parameter ideal value
    high            As Single                   'Parameter high limit
    low             As Single                   'Parameter low limit
End Type

Type RangeTest
    start           As HighLowAndLocSingle      'Parameter start location
    stop            As HighLowAndLocSingle      'Parameter stop location
    ideal           As Single                   'Parameter ideal value
End Type

Type ExtremeHighLow
    high            As ValueAndLocation        'Extreme high values
    low             As ValueAndLocation        'Extreme low  values
End Type

Type ExtremeType
    absoluteLin             As ExtremeHighLow       'Calculated absolute linearity deviation values  '1.5ANM
    SinglePointLin          As ExtremeHighLow       'Calculated SinglePoint linearity deviation values
    AbsLin                  As ExtremeHighLow       'Calculated Absolute linearity deviation values '2.7ANM
    linDevPerTol(1 To 2)    As ValueAndLocation     'Calculated linearity % tol values '2.7ANM 1 = SPL 2 = ABS
    slope                   As ExtremeHighLow       'Calculated slope values
    hysteresis              As ValueAndLocation     'Calculated hysteresis value
    fwdOutputCor            As ExtremeHighLow       'Calculated forward output correlation values
    revOutputCor            As ExtremeHighLow       'Calculated reverse output correlation values
    outputCorPerTol(1 To 2) As ValueAndLocation     'Calculated output correlation % tol values
    mechHysteresis          As ExtremeHighLow       'Calculated mechanical hysteresis range values
End Type
Public gudtExtreme(MAXCHANNUM)   As ExtremeType

Type ReadingType
    pedalAtRestLoc          As Single               'Measured pedal-at-rest location referenced to part datum zero
    pedalFaceLoc            As Single               'Measured pedal-at-rest location referenced to encoder zero
    forceKnee               As ValueAndLocation     'Measured force knee force and location
    fullPedalTravel         As ValueAndLocation     'Measured full-pedal-travel force and location '1.5ANM
    outputAtForceKnee       As Single               'Measured output at the force knee location
    Index(1 To 4)           As ValueAndLocation     'Measured index output
    FullCloseHys            As ValueAndLocation     'Measured full-close hysteresis value
    maxOutput               As ValueAndLocation     'Measured maximum output
    aveForcePt()            As ValueAndLocation     'Measured average force points
    fwdForcePt(1 To 3)      As ValueAndLocation     'Measured forward force points
    revForcePt(1 To 3)      As ValueAndLocation     'Measured reverse force points
    peakForce               As Single               'Measured peak force
    mechHystPt(1 To 3)      As ValueAndLocation     'Measured mechanical hysteresis points
    mlxCurrent              As Single               'Measured MLX current value '2.5ANM
    mlxWCurrent             As Single               'Measured MLX WOT current value '2.8cANM
    mlxSupply               As Single               'Measured MLX supply value  '2.8aANM
    KDStart                 As ValueAndLocation     'Measured KD Start value '2.8dANM
    KDStop                  As ValueAndLocation     'Measured KD Stop value  '2.8dANM
    KDPeak                  As ValueAndLocation     'Measured KD Peak value  '2.8dANM
    KDSpan                  As Single               'Measured KD Span value  '2.8dANM

End Type
Public gudtReading(MAXCHANNUM)   As ReadingType

Type ScanParameterStats
    pedalAtRestLoc          As Statistics           'Stat counts for pedal-at-rest location
    forceKneeLoc            As Statistics           'Stat counts for force knee location
    outputAtForceKnee       As Statistics           'Stat counts for output at the force knee location
    Index(1 To 3)           As Statistics           'Stat counts for index
    maxOutput               As Statistics           'Stat counts for maximum output
    linDevPerTol(1 To 2)    As Statistics           'Stat counts for max % of tol '2.7ANM 1 = SPL 2 = ABS
    slopeMax                As Statistics           'Stat counts for slope deviation max
    slopeMin                As Statistics           'Stat counts for slope deviation min
    hysteresis              As Statistics           'Stat counts for hysteresis
    FullCloseHys            As Statistics           'Stat counts for full-close hysteresis
    outputCorPerTol(1 To 2) As Statistics           'Stat counts for output correlation
    forceKneeForce          As Statistics           'Stat counts for force knee forward force
    aveForcePt()            As Statistics           'Stat counts for average force points
    fwdForcePt(1 To 3)      As Statistics           'Stat counts for forward force points
    revForcePt(1 To 3)      As Statistics           'Stat counts for reverse force points
    peakForce               As Statistics           'Stat counts for peak force
    mechHystPt(1 To 3)      As Statistics           'Stat counts for mechanical hysteresis points
    mechHysteresis          As Statistics           'Stat counts for mechanical hysteresis range
    mlxCurrent              As Statistics           'Stat counts for MLX I '2.5ANM
    mlxWCurrent             As Statistics           'Stat counts for WOT MLX I '2.8cANM
    KDStart                 As Statistics           'Stat counts for KD start      '2.8dANM
    KDStop                  As Statistics           'Stat counts for KD stop       '2.8dANM
    KDPeak                  As Statistics           'Stat counts for KD peak       '2.8dANM
    KDPeakForce             As Statistics           'Stat counts for KD peak force '2.8dANM
    KDSpan                  As Statistics           'Stat counts for KD span       '2.8dANM
End Type
Public gudtScanStats(MAXCHANNUM)      As ScanParameterStats

Type TestParameters
    evaluate                As StartAndStop         'Specified evaluation regions
    pedalAtRestLoc          As PointTest            'Specified pedal-at-rest location values
    riseTarget              As Single               'Specified rising point target value         '1.5ANM
    forceKneeLoc            As PointTest            'Specified force knee location limits
    outputAtForceKnee       As PointTest            'Specified values for output at the force knee location
    Index(1 To 3)           As PointTest            'Specified index values
    maxOutput               As PointTest            'Specified maximum output values
    SinglePointLin(1 To 5)  As RangeTest            'Specified SinglePoint linearity deviation values
    AbsLin(1 To 5)          As RangeTest            'Specified Absolute linearity deviation values '2.7ANM
    slope                   As PointTestWithRange   'Specified slope values
    hysteresis              As RangeTest            'Specified hysteresis values
    FullCloseHys            As PointTest            'Specified full-close hysteresis values
    fwdOutputCor(1 To 5)    As RangeTest            'Specified forward output correlation values
    revOutputCor(1 To 5)    As RangeTest            'Specified reverse output correlation values
    forceKneeForce          As PointTest            'Specified force knee forward force values
    aveForcePt()            As PointTest            'Specified average force point values
    fwdForcePt(1 To 3)      As PointTest            'Specified forward force point values
    revForcePt(1 To 3)      As PointTest            'Specified reverse force point values
    peakForce               As HighLowSingle        'Specified peak force values
    mechHystPt(1 To 3)      As PointTest            'Specified mechanical hysteresis point values
    mechHysteresis          As RangeTest            'Specified mechanical hysteresis range values
    kickdownStartLoc        As PointTestWithRange   'Specified kickdown start location values
    kickdownPeakLoc         As PointTestWithRange   'Specified kickdown peak locations
    kickdownPeakForce       As PointTestWithRange   'Specified kickdown peak force values
    kickdownForceSpan       As PointTestWithRange   'Specified kickdown peak force span values
    kickdownOnLoc           As PointTestWithRange   'Specified kickdown on location values
    kickdownOnSpan          As PointTestWithRange   'Specified kickdown on span values
    kickdownEndLoc          As PointTestWithRange   'Specified kickdown end loc values '2.8dANM
    fullPedalTravel         As PointTest            'Specified full-pedal-travel location values
    kickdownFullOpenSpan    As PointTestWithRange   'Specified kickdown FullOpen span values
    mlxCurrent              As PointTest            'Specified MLX I values '2.5ANM
    mlxWCurrent             As PointTest            'Specified WOT MLX I values '2.8cANM
End Type
Public gudtTest(MAXCHANNUM)      As TestParameters

Type ControlLimits
    pedalAtRestLoc          As HighLowSingle        'Severe pedal-zero location limits
    forceKneeLoc            As HighLowSingle        'Severe force knee location limits
    outputAtForceKnee       As HighLowSingle        'Severe output at the force knee location limits
    Index(1 To 3)           As HighLowSingle        'Severe index limits
    maxOutput               As HighLowSingle        'Severe max output limits
    linDevPerTol(1 To 2)    As HighLowSingle        'Severe max % of tolerance limits
    slope                   As HighLowSingle        'Severe slope deviation limits
    hysteresis              As HighLowSingle        'Severe hysteresis limits
    outputCorPerTol(1 To 2) As HighLowSingle        'Severe output correlation limits
    forceKneeForce          As HighLowSingle        'Severe force knee forward force limits
    aveForcePt()            As HighLowSingle        'Severe average force point limits
    fwdForcePt(1 To 3)      As HighLowSingle        'Severe forward force point limits
    revForcePt(1 To 3)      As HighLowSingle        'Severe reverse force point limits
    peakForce               As HighLowSingle        'Severe peak force limits
    mechHystPt(1 To 3)      As HighLowSingle        'Severe mechanical hysteresis point limits
    mechHysteresis          As HighLowSingle        'Severe mechanical hysteresis range limits
    mlxCurrent              As HighLowSingle        'Severe mlx current range limits '2.8aANM
End Type
Public gudtControl(MAXCHANNUM)   As ControlLimits

Type CustomerTestParameters
    pedalAtRestLoc          As HighLowSingle        'Specified pedal-at-rest location values for cp & cpk calculations
    forceKneeLoc            As HighLowSingle        'Specified force knee location values for cp & cpk calculations
    outputAtForceKnee       As HighLowSingle        'Specified output at the force knee location values for cp & cpk calculations
    Index(1 To 3)           As HighLowSingle        'Specified index values for cp & cpk calculations
    forceKneeForce          As HighLowSingle        'Specified force knee forward force values for cp & cpk calculations
    aveForcePt()            As HighLowSingle        'Specified average force values for cp & cpk calculations
    fwdForcePt(1 To 3)      As HighLowSingle        'Specified forward force values for cp & cpk calculations
    revForcePt(1 To 3)      As HighLowSingle        'Specified reverse force values for cp & cpk calculations
    mechHystPt(1 To 3)      As HighLowSingle        'Specified mechanical hysteresis values for cp & cpk calculations
End Type
Public gudtCustomerSpec(MAXCHANNUM)      As CustomerTestParameters

Public gsngIdeal() As Single             '1.7ANM
Public glngBOM As Long                   '2.0ANM
Public gblnMLXVI As Boolean              '2.8aANM
Public gblnKD As Boolean                 '2.8dANM

'*******************************
'*  Module-Level Declarations  *
'*******************************

'Limit Arrays
Private msngMaxVoltageGradientLimit()   As Single
Private msngMinVoltageGradientLimit()   As Single
Private msngMaxLinearityLimit()         As Single
Private msngMinLinearityLimit()         As Single
Private msngMaxAbsLinearityLimit()      As Single '2.7ANM
Private msngMinAbsLinearityLimit()      As Single '2.7ANM
Private msngMaxSlopeDeviationLimit()    As Single
Private msngMinSlopeDeviationLimit()    As Single
Private msngMaxHysteresisLimit()        As Single
Private msngMinHysteresisLimit()        As Single
Private msngMaxFwdOutputCorLimit()      As Single
Private msngMinFwdOutputCorLimit()      As Single
Private msngMaxRevOutputCorLimit()      As Single
Private msngMinRevOutputCorLimit()      As Single
Private msngMaxForceGradientLimit()     As Single
Private msngMinForceGradientLimit()     As Single
Private msngMaxMechHysteresisLimit()    As Single
Private msngMinMechHysteresisLimit()    As Single

Public Sub DataEvaluation()
'
'   PURPOSE:   Executive which calls the routines that calculates the
'              linearity, hysteresis, and slope values.
'
'  INPUT(S):
' OUTPUT(S):

'Data Arrays
Dim lsngFwdVoltageGradient() As Single  'Calculated Forward Voltage Gradient data array
Dim lsngRevVoltageGradient() As Single  'Calculated Reverse Voltage Gradient data array
Dim lsngFwdVoltageGradient2() As Single 'Calculated Forward Voltage Gradient data array 2
Dim lsngRevVoltageGradient2() As Single 'Calculated Reverse Voltage Gradient data array 2
Dim lsngSinglePointLinDev() As Single   'Calculated SinglePoint Linearity Deviation data array
Dim lsngAbsLinDev() As Single           'Calculated Absolute Linearity Deviation data array '2.7ANM
Dim lsngLinDevPerTol() As Single        'Calculated Linearity Deviation % of Tolerance data array
Dim lsngLinDevPerTol2() As Single       'Calculated Linearity Deviation % of Tolerance data array '2.7ANM
Dim lsngSlopeDev() As Single            'Calculated Slope Deviation data array
Dim lsngHysteresis() As Single          'Calculated Hysteresis data array
Dim lsngFwdOutputCor() As Single        'Calculated Rorward data array for Output Correlation
Dim lsngFwdOutputCorPerTol() As Single  'Calculated Rorward data array for Output Correlation % of Tolerance
Dim lsngRevOutputCor() As Single        'Calculated Reverse data array for Output Correlation
Dim lsngRevOutputCorPerTol() As Single  'Calculated Reverse data array for Output Correlation % of Tolerance
Dim lsngFwdForce() As Single            'Calculated Forward Force data array
Dim lsngRevForce() As Single            'Calculated Reverse Force data array
Dim lsngMechHysteresis() As Single      'Calculated Mechanical Hysteresis data array
Dim lintChanNum As Integer              'Identifies Channel Number
Dim lintRegion As Integer               'Identifies Region Number
Dim lintDataPoint As Integer            'Identifies data point in loops
Dim lstrGraphFooterhigh As String       'Graph Footer Information
Dim lstrGraphFooterLow As String        'Graph Footer Information
Dim lsngIncrement As Single             'Graph Increment size
Dim lstrTitle(MAXGRAPHS) As String      'Subset titles for graphs if displaying more than one output on a graph
Dim lsngNotUsed1 As Single              'Dummy variable
Dim lsngNotUsed2 As Single              'Dummy variable
Dim lsngMinVal As Single                'Temp variable
Dim lsngMinLoc As Single                'Temp variable
Dim lsngMaxVal As Single                'Temp variable
Dim lsngMaxLoc As Single                'Temp variable
'2.3ANM Dim lblnForceKneeFound As Boolean       'Whether or not the Force Knee was found
Dim lblnFullPedalTravelFound As Boolean 'Whether or not Full-Pedal-Travel was found
Dim lintRegionCount As Integer          'Count number             '2.2ANM
Dim lintRegionCount2 As Integer         'Graph Count number       '2.7ANM
Dim lstrParameterName As String         'Parameter or Metric Name '2.2ANM
Dim lblnKDStartFound As Boolean         'Whether or not KD Start was found '2.8dANM
Dim lblnKDStopFound As Boolean          'Whether or not KD Stop was found  '2.8dANM

'Set the error trap
On Error GoTo DataEvaluationError

lsngIncrement = 1 / gsngResolution      'Define the x-axis increment size for graphs
gintPointer = 0                         'Initialize graph pointer to zero

'Create Forward Force data array
ReDim lsngFwdForce(gintMaxData)
Call Calc.CalcScaledDataArray(CHAN2, gintForward(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, VOLTSPERLSB * gsngNewtonsPerVolt, gsngForceAmplifierOffset, gsngResolution, lsngFwdForce())
If gintAnomaly Then Exit Sub                'Exit on system error

'Create Reverse Force data array
ReDim lsngRevForce(gintMaxData)
Call Calc.CalcScaledDataArray(CHAN2, gintReverse(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, VOLTSPERLSB * gsngNewtonsPerVolt, gsngForceAmplifierOffset, gsngResolution, lsngRevForce())
If gintAnomaly Then Exit Sub                'Exit on system error

'2.3ANM 'Calculate Force Knee Location & Force
'2.3ANM Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.FKSlope, False, gudtMachine.FKPercentage, gudtMachine.FKWindow, gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, gsngResolution, gudtReading(CHAN0).forceKnee.location, gudtReading(CHAN0).forceKnee.Value, lblnForceKneeFound)
'2.3ANM If gintAnomaly Then Exit Sub                'Exit on system error
'2.3ANM If Not lblnForceKneeFound Then
'2.3ANM     gintAnomaly = 157
'2.3ANM     'Log the error to the error log and display the error message
'2.3ANM     Call Pedal.ErrorLogFile("Force Knee Not Found: Verify that force cell and amplifier" & vbCrLf & _
'2.3ANM                             "                      are functioning properly and the pedal" & vbCrLf & _
'2.3ANM                             "                      is properly clamped in the fixture.", True, True)
'2.3ANM     'Exit on system error
'2.3ANM     Exit Sub
'2.3ANM End If

'2.2ANM 'Check Force Knee Location
'2.2ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).forceKnee.location, gudtReading(CHAN0).forceKnee.location, gudtTest(CHAN0).forceKneeLoc.low, gudtTest(CHAN0).forceKneeLoc.high, LOWFORCEKNEELOC, HIGHFORCEKNEELOC, gintFailure())
'2.2ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).forceKnee.location, gudtReading(CHAN0).forceKnee.location, gudtControl(CHAN0).forceKneeLoc.low, gudtControl(CHAN0).forceKneeLoc.high, LOWFORCEKNEELOC, HIGHFORCEKNEELOC, gintSevere())

'2.3ANM 'Check Force Knee Forward Force
'2.3ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).forceKnee.Value, gudtReading(CHAN0).forceKnee.Value, gudtTest(CHAN0).forceKneeForce.low, gudtTest(CHAN0).forceKneeForce.high, LOWFORCEKNEEFWDFORCE, HIGHFORCEKNEEFWDFORCE, gintFailure())
'2.3ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).forceKnee.Value, gudtReading(CHAN0).forceKnee.Value, gudtControl(CHAN0).forceKneeForce.low, gudtControl(CHAN0).forceKneeForce.high, LOWFORCEKNEEFWDFORCE, HIGHFORCEKNEEFWDFORCE, gintSevere())

'Check Pedal at Rest Loc '2.8ANM
Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).pedalAtRestLoc, gudtReading(CHAN0).pedalAtRestLoc, gudtTest(CHAN0).pedalAtRestLoc.low, gudtTest(CHAN0).pedalAtRestLoc.high, LOWPEDALATREST, HIGHPEDALATREST, gintFailure())
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).pedalAtRestLoc, gudtReading(CHAN0).pedalAtRestLoc, gudtControl(CHAN0).pedalAtRestLoc.low, gudtControl(CHAN0).pedalAtRestLoc.high, LOWPEDALATREST, HIGHPEDALATREST, gintSevere())

'Calculate and Check Forward Force Point 1
Call Calc.CalcRefPointByLoc(lsngFwdForce(), gudtTest(CHAN0).fwdForcePt(1).location, gsngResolution, gudtReading(CHAN0).fwdForcePt(1).Value, gudtReading(CHAN0).fwdForcePt(1).location)
If gintAnomaly Then Exit Sub                'Exit on system error
Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).fwdForcePt(1).Value, gudtReading(CHAN0).fwdForcePt(1).Value, gudtTest(CHAN0).fwdForcePt(1).low, gudtTest(CHAN0).fwdForcePt(1).high, LOWFWDFORCEPT1, HIGHFWDFORCEPT1, gintFailure())
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).fwdForcePt(1).Value, gudtReading(CHAN0).fwdForcePt(1).Value, gudtControl(CHAN0).fwdForcePt(1).low, gudtControl(CHAN0).fwdForcePt(1).high, LOWFWDFORCEPT1, HIGHFWDFORCEPT1, gintSevere())

'Calculate and Check Forward Force Point 2
Call Calc.CalcRefPointByLoc(lsngFwdForce(), gudtTest(CHAN0).fwdForcePt(2).location, gsngResolution, gudtReading(CHAN0).fwdForcePt(2).Value, gudtReading(CHAN0).fwdForcePt(2).location)
If gintAnomaly Then Exit Sub                'Exit on system error
'1.5ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).fwdForcePt(2).Value, gudtReading(CHAN0).fwdForcePt(2).Value, gudtTest(CHAN0).fwdForcePt(2).low, gudtTest(CHAN0).fwdForcePt(2).high, LOWFWDFORCEPT2, HIGHFWDFORCEPT2, gintFailure())
'1.5ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).fwdForcePt(2).Value, gudtReading(CHAN0).fwdForcePt(2).Value, gudtControl(CHAN0).fwdForcePt(2).low, gudtControl(CHAN0).fwdForcePt(2).high, LOWFWDFORCEPT2, HIGHFWDFORCEPT2, gintSevere())

'Calculate and Check Forward Force Point 3
Call Calc.CalcRefPointByLoc(lsngFwdForce(), gudtTest(CHAN0).fwdForcePt(3).location, gsngResolution, gudtReading(CHAN0).fwdForcePt(3).Value, gudtReading(CHAN0).fwdForcePt(3).location)
If gintAnomaly Then Exit Sub                'Exit on system error
Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).fwdForcePt(3).Value, gudtReading(CHAN0).fwdForcePt(3).Value, gudtTest(CHAN0).fwdForcePt(3).low, gudtTest(CHAN0).fwdForcePt(3).high, LOWFWDFORCEPT3, HIGHFWDFORCEPT3, gintFailure())
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).fwdForcePt(3).Value, gudtReading(CHAN0).fwdForcePt(3).Value, gudtControl(CHAN0).fwdForcePt(3).low, gudtControl(CHAN0).fwdForcePt(3).high, LOWFWDFORCEPT3, HIGHFWDFORCEPT3, gintSevere())

'Calculate and Check Reverse Force Point 1
Call Calc.CalcRefPointByLoc(lsngRevForce(), gudtTest(CHAN0).revForcePt(1).location, gsngResolution, gudtReading(CHAN0).revForcePt(1).Value, gudtReading(CHAN0).revForcePt(1).location)
If gintAnomaly Then Exit Sub                'Exit on system error
Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).revForcePt(1).Value, gudtReading(CHAN0).revForcePt(1).Value, gudtTest(CHAN0).revForcePt(1).low, gudtTest(CHAN0).revForcePt(1).high, LOWREVFORCEPT1, HIGHREVFORCEPT1, gintFailure())
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).revForcePt(1).Value, gudtReading(CHAN0).revForcePt(1).Value, gudtControl(CHAN0).revForcePt(1).low, gudtControl(CHAN0).revForcePt(1).high, LOWREVFORCEPT1, HIGHREVFORCEPT1, gintSevere())

'Calculate and Check Reverse Force Point 2
Call Calc.CalcRefPointByLoc(lsngRevForce(), gudtTest(CHAN0).revForcePt(2).location, gsngResolution, gudtReading(CHAN0).revForcePt(2).Value, gudtReading(CHAN0).revForcePt(2).location)
If gintAnomaly Then Exit Sub                'Exit on system error
'1.5ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).revForcePt(2).Value, gudtReading(CHAN0).revForcePt(2).Value, gudtTest(CHAN0).revForcePt(2).low, gudtTest(CHAN0).revForcePt(2).high, LOWREVFORCEPT2, HIGHREVFORCEPT2, gintFailure())
'1.5ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).revForcePt(2).Value, gudtReading(CHAN0).revForcePt(2).Value, gudtControl(CHAN0).revForcePt(2).low, gudtControl(CHAN0).revForcePt(2).high, LOWREVFORCEPT2, HIGHREVFORCEPT2, gintSevere())

'Calculate and Check Pedal Reverse Force Point 3
Call Calc.CalcRefPointByLoc(lsngRevForce(), gudtTest(CHAN0).revForcePt(3).location, gsngResolution, gudtReading(CHAN0).revForcePt(3).Value, gudtReading(CHAN0).revForcePt(3).location)
If gintAnomaly Then Exit Sub                'Exit on system error
Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).revForcePt(3).Value, gudtReading(CHAN0).revForcePt(3).Value, gudtTest(CHAN0).revForcePt(3).low, gudtTest(CHAN0).revForcePt(3).high, LOWREVFORCEPT3, HIGHREVFORCEPT3, gintFailure())
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).revForcePt(3).Value, gudtReading(CHAN0).revForcePt(3).Value, gudtControl(CHAN0).revForcePt(3).low, gudtControl(CHAN0).revForcePt(3).high, LOWREVFORCEPT3, HIGHREVFORCEPT3, gintSevere())

'Check Peak Force   (NOTE that all peak force failures are Severe)
'NOTE: Peak Force is measured using a serial command to the SC2000 after scanning is complete
Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).peakForce, gudtReading(CHAN0).peakForce, gudtTest(CHAN0).peakForce.low, gudtTest(CHAN0).peakForce.high, LOWPEAKFORCE, HIGHPEAKFORCE, gintSevere())

'Create Mechanical Hysteresis data array
ReDim lsngMechHysteresis(gintMaxData)
Call Calc.CalcMechanicalHysteresisPercentage(lsngFwdForce(), lsngRevForce(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, gsngResolution, lsngMechHysteresis())
If gintAnomaly Then Exit Sub                'Exit on system error

'Calculate and Check Mechanical Hysteresis Point 1
Call Calc.CalcRefPointByLoc(lsngMechHysteresis(), gudtTest(CHAN0).mechHystPt(1).location, gsngResolution, gudtReading(CHAN0).mechHystPt(1).Value, gudtReading(CHAN0).mechHystPt(1).location)
If gintAnomaly Then Exit Sub                'Exit on system error
'2.2ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).mechHystPt(1).Value, gudtReading(CHAN0).mechHystPt(1).Value, gudtTest(CHAN0).mechHystPt(1).low, gudtTest(CHAN0).mechHystPt(1).high, LOWMECHHYSTPT1, HIGHMECHHYSTPT1, gintFailure())
'2.2ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).mechHystPt(1).Value, gudtReading(CHAN0).mechHystPt(1).Value, gudtControl(CHAN0).mechHystPt(1).low, gudtControl(CHAN0).mechHystPt(1).high, LOWMECHHYSTPT1, HIGHMECHHYSTPT1, gintSevere())

'Calculate and Check Mechanical Hysteresis Point 2
Call Calc.CalcRefPointByLoc(lsngMechHysteresis(), gudtTest(CHAN0).mechHystPt(2).location, gsngResolution, gudtReading(CHAN0).mechHystPt(2).Value, gudtReading(CHAN0).mechHystPt(2).location)
If gintAnomaly Then Exit Sub                'Exit on system error
'1.5ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).mechHystPt(2).Value, gudtReading(CHAN0).mechHystPt(2).Value, gudtTest(CHAN0).mechHystPt(2).low, gudtTest(CHAN0).mechHystPt(2).high, LOWMECHHYSTPT2, HIGHMECHHYSTPT2, gintFailure())
'1.5ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).mechHystPt(2).Value, gudtReading(CHAN0).mechHystPt(2).Value, gudtControl(CHAN0).mechHystPt(2).low, gudtControl(CHAN0).mechHystPt(2).high, LOWMECHHYSTPT2, HIGHMECHHYSTPT2, gintSevere())

'Calculate and Check Mechanical Hysteresis Point 3
Call Calc.CalcRefPointByLoc(lsngMechHysteresis(), gudtTest(CHAN0).mechHystPt(3).location, gsngResolution, gudtReading(CHAN0).mechHystPt(3).Value, gudtReading(CHAN0).mechHystPt(3).location)
If gintAnomaly Then Exit Sub                'Exit on system error
'2.2ANM Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).mechHystPt(3).Value, gudtReading(CHAN0).mechHystPt(3).Value, gudtTest(CHAN0).mechHystPt(3).low, gudtTest(CHAN0).mechHystPt(3).high, LOWMECHHYSTPT3, HIGHMECHHYSTPT3, gintFailure())
'2.2ANM Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).mechHystPt(3).Value, gudtReading(CHAN0).mechHystPt(3).Value, gudtControl(CHAN0).mechHystPt(3).low, gudtControl(CHAN0).mechHystPt(3).high, LOWMECHHYSTPT3, HIGHMECHHYSTPT3, gintSevere())

'2.8dANM \/\/
If gblnKD Then
    'Calculate and Check the Kickdown Start Location
    Call Calc.CalcKneeLoc(lsngFwdForce(), gudtMachine.KDStartSlope, True, gudtMachine.KDStartPercentage, gudtMachine.KDStartWindow, gudtTest(CHAN0).kickdownStartLoc.start, gudtTest(CHAN0).kickdownStartLoc.stop, gsngResolution, gudtReading(CHAN0).KDStart.location, gudtReading(CHAN0).KDStart.Value, lblnKDStartFound)
    If Not lblnKDStartFound Then '2.8hANM \/\/
'        gintAnomaly = 156
'        'Log the error to the error log and display the error message
        Call MsgBox("Kickdown Start Location Not Found." & vbCrLf & _
                    "Verify Kickdown Module Installed Correctly" & vbCrLf & _
                    "and Check Force Sensing Equipment.", vbOKOnly, "KD Error")
'        'Exit on system error
'        Exit Sub
        gudtReading(CHAN0).KDStart.location = 0
        gudtReading(CHAN0).KDPeak.location = 0
        gudtReading(CHAN0).KDPeak.Value = 0
        gudtReading(CHAN0).KDSpan = 0
        gudtReading(CHAN0).KDStop.location = 0
    Else
        Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).KDStart.location, gudtReading(CHAN0).KDStart.location, gudtTest(CHAN0).kickdownStartLoc.low, gudtTest(CHAN0).kickdownStartLoc.high, LOWKDSTART, HIGHKDSTART, gintFailure())
        
        'Calculate Kickdown Peak Location and Force
        Call CalcMinMax(lsngFwdForce(), gudtTest(CHAN0).kickdownPeakLoc.start, gudtTest(CHAN0).kickdownPeakLoc.stop, gsngResolution, lsngNotUsed1, lsngNotUsed2, gudtReading(CHAN0).KDPeak.Value, gudtReading(CHAN0).KDPeak.location)
    
        'Calculate Kickdown Force Span
        gudtReading(CHAN0).KDSpan = gudtReading(CHAN0).KDPeak.Value - gudtReading(CHAN0).KDStart.Value
        Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).KDSpan, gudtReading(CHAN0).KDSpan, gudtTest(CHAN0).kickdownForceSpan.low, gudtTest(CHAN0).kickdownForceSpan.high, LOWKDSPAN, HIGHKDSPAN, gintFailure()) '2.8dANM
    
        'Calculate and Check the Kickdown Stop Location
        Call Calc.CalcBkKneeLoc(lsngFwdForce(), 0, False, 80, 0.5, gudtReading(CHAN0).KDPeak.location, gudtTest(CHAN0).kickdownEndLoc.stop, gsngResolution, gudtReading(CHAN0).KDStop.location, gudtReading(CHAN0).KDStop.Value, lblnKDStopFound)
        Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).KDStop.location, gudtReading(CHAN0).KDStop.location, gudtTest(CHAN0).kickdownEndLoc.low, gudtTest(CHAN0).kickdownEndLoc.high, LOWKDSTOP, HIGHKDSTOP, gintFailure())
    End If
Else
    gudtReading(CHAN0).KDStart.location = 0
    gudtReading(CHAN0).KDPeak.location = 0
    gudtReading(CHAN0).KDPeak.Value = 0
    gudtReading(CHAN0).KDSpan = 0
    gudtReading(CHAN0).KDStop.location = 0
End If
'2.8dANM /\/\

If Not gblnForceOnly Then       '1.6ANM added if block
    'Iterate through all channel numbers
    For lintChanNum = CHAN0 To MAXCHANNUM
    
        'Calculate the Forward Voltage Gradient
        ReDim lsngFwdVoltageGradient(gintMaxData)
        Call Calc.CalcRatiometricDataArray(lintChanNum, gintForward(), gintForSupply(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gsngResolution, lsngFwdVoltageGradient())
        If gintAnomaly Then Exit Sub                'Exit on system error
    
        'Calculate the Reverse Voltage Gradient
        ReDim lsngRevVoltageGradient(gintMaxData)
        Call Calc.CalcRatiometricDataArray(lintChanNum, gintReverse(), gintRevSupply(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gsngResolution, lsngRevVoltageGradient())
        If gintAnomaly Then Exit Sub                'Exit on system error
    
        'Check Index 1 (FullClose)
        'NOTE: FullClose is measured statically, prior to the scan
        Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).Index(1).Value, gudtReading(lintChanNum).Index(1).Value, gudtTest(lintChanNum).Index(1).low, gudtTest(lintChanNum).Index(1).high, LOWINDEXPT1, HIGHINDEXPT1, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).Index(1).Value, gudtReading(lintChanNum).Index(1).Value, gudtControl(lintChanNum).Index(1).low, gudtControl(lintChanNum).Index(1).high, LOWINDEXPT1, HIGHINDEXPT1, gintSevere())
    
        'Calculate and Check Output at Force Knee (Note that high limits are index 1 reading + test.high
        Call Calc.CalcRefPointByLoc(lsngFwdVoltageGradient(), gudtReading(CHAN0).forceKnee.location, gsngResolution, gudtReading(lintChanNum).outputAtForceKnee, lsngNotUsed1)
        '1.5ANM Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).outputAtForceKnee, gudtReading(lintChanNum).outputAtForceKnee, gudtTest(lintChanNum).outputAtForceKnee.low, gudtReading(lintChanNum).Index(1).Value + gudtTest(lintChanNum).outputAtForceKnee.high, LOWOUTPUTATFORCEKNEE, HIGHOUTPUTATFORCEKNEE, gintFailure())
        '1.5ANM Call Calc.CalcControlLimits(gudtTest(lintChanNum).outputAtForceKnee.ideal, gudtTest(lintChanNum).outputAtForceKnee.low, gudtReading(lintChanNum).Index(1).Value + gudtTest(lintChanNum).outputAtForceKnee.high, gudtControl(lintChanNum).outputAtForceKnee.low, gudtControl(lintChanNum).outputAtForceKnee.high)
        '1.5ANM Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).outputAtForceKnee, gudtReading(lintChanNum).outputAtForceKnee, gudtControl(lintChanNum).outputAtForceKnee.low, gudtControl(lintChanNum).outputAtForceKnee.high, LOWOUTPUTATFORCEKNEE, HIGHOUTPUTATFORCEKNEE, gintSevere())
    
        'Calculate and Check Index 2 (Midpoint)
        Call Calc.CalcRefPointByLoc(lsngFwdVoltageGradient(), gudtTest(lintChanNum).Index(2).location, gsngResolution, gudtReading(lintChanNum).Index(2).Value, gudtReading(lintChanNum).Index(2).location)
        '1.5ANM Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).Index(2).Value, gudtReading(lintChanNum).Index(2).Value, gudtTest(lintChanNum).Index(2).low, gudtTest(lintChanNum).Index(2).high, LOWINDEXPT2, HIGHINDEXPT2, gintFailure())
        '1.5ANM Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).Index(2).Value, gudtReading(lintChanNum).Index(2).Value, gudtControl(lintChanNum).Index(2).low, gudtControl(lintChanNum).Index(2).high, LOWINDEXPT2, HIGHINDEXPT2, gintSevere())
    
        'Calculate and Check Index 3 (FullOpen)
        Call Calc.CalcRefPointByLoc(lsngFwdVoltageGradient(), gudtTest(lintChanNum).Index(3).location, gsngResolution, gudtReading(lintChanNum).Index(3).Value, gudtReading(lintChanNum).Index(3).location)
        Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).Index(3).Value, gudtReading(lintChanNum).Index(3).Value, gudtTest(lintChanNum).Index(3).low, gudtTest(lintChanNum).Index(3).high, LOWINDEXPT3, HIGHINDEXPT3, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).Index(3).Value, gudtReading(lintChanNum).Index(3).Value, gudtControl(lintChanNum).Index(3).low, gudtControl(lintChanNum).Index(3).high, LOWINDEXPT3, HIGHINDEXPT3, gintSevere())
    
        'Calculate and Check Index 3 by value (FullOpen)
        'NOTE1: Ideal FullOpen Voltage = test.index(3).location.
        '       location is used to represent value; ideal,
        '       high, and low are used to represent locations
        'Call Calc.CalcRefPointByVal(lsngFwdVoltageGradient(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gudtTest(lintChanNum).Index(3).location, gudtTest(lintChanNum).slope.ideal, gsngResolution, gudtReading(lintChanNum).Index(3).Value, gudtReading(lintChanNum).Index(3).location)
        'If gintAnomaly Then Exit Sub                'Exit on system error
        'Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).Index(3).location, gudtReading(lintChanNum).Index(3).location, gudtTest(lintChanNum).Index(3).low, gudtTest(lintChanNum).Index(3).high, LOWINDEXPT3, HIGHINDEXPT3, gintFailure())
        'Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).Index(3).location, gudtReading(lintChanNum).Index(3).location, gudtControl(lintChanNum).Index(3).low, gudtControl(lintChanNum).Index(3).high, LOWINDEXPT3, HIGHINDEXPT3, gintSevere())
    
        'Calculate and Check Maximum Output
        Call Calc.CalcMinMax(lsngFwdVoltageGradient(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gsngResolution, lsngNotUsed1, lsngNotUsed2, gudtReading(lintChanNum).maxOutput.Value, gudtReading(lintChanNum).maxOutput.location)
        If gintAnomaly Then Exit Sub                'Exit on system error
        Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).maxOutput.Value, gudtReading(lintChanNum).maxOutput.Value, gudtTest(lintChanNum).maxOutput.low, gudtTest(lintChanNum).maxOutput.high, LOWMAXOUTPUT, HIGHMAXOUTPUT, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtReading(lintChanNum).maxOutput.Value, gudtReading(lintChanNum).maxOutput.Value, gudtControl(lintChanNum).maxOutput.low, gudtControl(lintChanNum).maxOutput.high, LOWMAXOUTPUT, HIGHMAXOUTPUT, gintSevere())
    
        'Calculate SinglePoint Linearity Deviation across all regions
        ReDim lsngSinglePointLinDev(gintMaxData)
        
        '2.2ANM \/\/
        If gblnUseNewAmad Then
            Dim lintOutput As Integer
            If lintChanNum = 0 Then
                lintOutput = 1
            Else
                lintOutput = 2
            End If
            lstrParameterName = "SingLinDevVal"
            lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, lintOutput, MPCTYPE_IDEAL)
        Else
            lintRegionCount = 5
        End If
        lintRegionCount2 = lintRegionCount '2.7ANM
        '2.2ANM /\/\
        
        Call Calc.CalcLinSinglePointwBend(lsngFwdVoltageGradient(), gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(lintRegionCount).stop.location, gudtReading(lintChanNum).Index(1).Value, gudtTest(lintChanNum).Index(1).location, gudtTest(lintChanNum).slope.ideal, gudtTest(lintChanNum).slope.ideal2, gsngResolution, lsngSinglePointLinDev()) '1.1aANM
        If gintAnomaly Then Exit Sub                'Exit on system error
        Call Calc.CalcMinMax(lsngSinglePointLinDev(), gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(lintRegionCount).stop.location, gsngResolution, gudtExtreme(lintChanNum).SinglePointLin.low.Value, gudtExtreme(lintChanNum).SinglePointLin.low.location, gudtExtreme(lintChanNum).SinglePointLin.high.Value, gudtExtreme(lintChanNum).SinglePointLin.high.location)
        If gintAnomaly Then Exit Sub                'Exit on system error
    
        'Calculate Limit Arrays
        Call EvaluateLimits(lintChanNum)
    
        'Calculate and Check Linearity Percentage of Tolerance
        ReDim lsngLinDevPerTol(gintMaxData)
        Call Calc.CalcPercentTol(lsngSinglePointLinDev(), msngMinLinearityLimit(), msngMaxLinearityLimit(), gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(lintRegionCount).stop.location, gsngResolution, lsngLinDevPerTol())
        If gintAnomaly Then Exit Sub                'Exit on system error
        Call Calc.CalcMinMax(lsngLinDevPerTol(), gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(lintRegionCount).stop.location, gsngResolution, lsngMinVal, lsngMinLoc, lsngMaxVal, lsngMaxLoc)
        If gintAnomaly Then Exit Sub                'Exit on system error
        If Abs(lsngMaxVal) > Abs(lsngMinVal) Then
            gudtExtreme(lintChanNum).linDevPerTol(1).Value = lsngMaxVal
            gudtExtreme(lintChanNum).linDevPerTol(1).location = lsngMaxLoc
        Else
            gudtExtreme(lintChanNum).linDevPerTol(1).Value = lsngMinVal
            gudtExtreme(lintChanNum).linDevPerTol(1).location = lsngMinLoc
        End If
        Call Calc.CheckFault(lintChanNum, gudtExtreme(lintChanNum).linDevPerTol(1).Value, gudtExtreme(lintChanNum).linDevPerTol(1).Value, -HUNDREDPERCENT, HUNDREDPERCENT, LOWSINGLEPOINTLIN, HIGHSINGLEPOINTLIN, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtExtreme(lintChanNum).linDevPerTol(1).Value, gudtExtreme(lintChanNum).linDevPerTol(1).Value, gudtControl(lintChanNum).linDevPerTol(1).low, gudtControl(lintChanNum).linDevPerTol(1).high, LOWSINGLEPOINTLIN, HIGHSINGLEPOINTLIN, gintSevere())
    
        '2.7ANM \/\/
        'Calculate Absolute Linearity Deviation across all regions
        ReDim lsngAbsLinDev(gintMaxData)
        
        If gblnUseNewAmad Then
            If lintChanNum = 0 Then
                lintOutput = 1
            Else
                lintOutput = 2
            End If
            lstrParameterName = "AbsLinDevVal"
            lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, lintOutput, MPCTYPE_IDEAL)
        Else
            lintRegionCount = 5
        End If
        
        '2.8aANM \/\/
        Dim lsngSlope As Single
        If gudtTest(lintChanNum).AbsLin(lintRegionCount).ideal <> 0 Then
            lsngSlope = gudtTest(lintChanNum).AbsLin(lintRegionCount).ideal
        Else
            lsngSlope = gudtTest(lintChanNum).slope.ideal2
        End If
        
        Call Calc.CalcLinAbsoluteSegment(lsngFwdVoltageGradient(), gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(lintRegionCount).stop.location, gudtTest(lintChanNum).Index(1).ideal, gudtTest(lintChanNum).Index(1).location, lsngSlope, gsngResolution, lsngAbsLinDev()) '2.8aANM slope set to var
        If gintAnomaly Then Exit Sub                'Exit on system error
        Call Calc.CalcMinMax(lsngAbsLinDev(), gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(lintRegionCount).stop.location, gsngResolution, gudtExtreme(lintChanNum).AbsLin.low.Value, gudtExtreme(lintChanNum).AbsLin.low.location, gudtExtreme(lintChanNum).AbsLin.high.Value, gudtExtreme(lintChanNum).AbsLin.high.location)
        If gintAnomaly Then Exit Sub                'Exit on system error
    
        'Calculate and Check Linearity Percentage of Tolerance
        ReDim lsngLinDevPerTol2(gintMaxData)
        Call Calc.CalcPercentTol(lsngAbsLinDev(), msngMinAbsLinearityLimit(), msngMaxAbsLinearityLimit(), gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(lintRegionCount).stop.location, gsngResolution, lsngLinDevPerTol2())
        If gintAnomaly Then Exit Sub                'Exit on system error
        Call Calc.CalcMinMax(lsngLinDevPerTol2(), gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(lintRegionCount).stop.location, gsngResolution, lsngMinVal, lsngMinLoc, lsngMaxVal, lsngMaxLoc)
        If gintAnomaly Then Exit Sub                'Exit on system error
        If Abs(lsngMaxVal) > Abs(lsngMinVal) Then
            gudtExtreme(lintChanNum).linDevPerTol(2).Value = lsngMaxVal
            gudtExtreme(lintChanNum).linDevPerTol(2).location = lsngMaxLoc
        Else
            gudtExtreme(lintChanNum).linDevPerTol(2).Value = lsngMinVal
            gudtExtreme(lintChanNum).linDevPerTol(2).location = lsngMinLoc
        End If
        Call Calc.CheckFault(lintChanNum, gudtExtreme(lintChanNum).linDevPerTol(2).Value, gudtExtreme(lintChanNum).linDevPerTol(2).Value, -HUNDREDPERCENT, HUNDREDPERCENT, LOWABSLIN, HIGHABSLIN, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtExtreme(lintChanNum).linDevPerTol(2).Value, gudtExtreme(lintChanNum).linDevPerTol(2).Value, gudtControl(lintChanNum).linDevPerTol(2).low, gudtControl(lintChanNum).linDevPerTol(2).high, LOWABSLIN, HIGHABSLIN, gintSevere())
        '2.7ANM /\/\
    
        'Calculate and Check Slope Deviation (Ratio Method)
        'NOTE: The ideal slope is defined as the slope between the measured output at the force knee
        '      and the ideal output at FullOpen.  This is the slope used to define ideal linearity
        ReDim lsngSlopeDev((gudtTest(lintChanNum).slope.start * gsngResolution) / gudtMachine.slopeIncrement To (gudtTest(lintChanNum).slope.stop * gsngResolution) / gudtMachine.slopeIncrement)
        Call Calc.CalcSlopeDev(lsngFwdVoltageGradient(), gudtTest(lintChanNum).slope.start, gudtTest(lintChanNum).slope.stop, gudtMachine.slopeInterval, gudtMachine.slopeIncrement, gudtTest(lintChanNum).slope.ideal2, gsngResolution, True, lsngSlopeDev())
        If gintAnomaly Then Exit Sub                'Exit on system error
        'Calculate Min/Max
        Call Calc.CalcMinMax(lsngSlopeDev(), gudtTest(lintChanNum).slope.start, (gudtTest(lintChanNum).slope.stop - (gudtMachine.slopeInterval / gsngResolution)), gsngResolution / gudtMachine.slopeIncrement, gudtExtreme(lintChanNum).slope.low.Value, gudtExtreme(lintChanNum).slope.low.location, gudtExtreme(lintChanNum).slope.high.Value, gudtExtreme(lintChanNum).slope.high.location)
        If gintAnomaly Then Exit Sub                'Exit on system error
        'Check for Slope Deviation for Failures
        Call Calc.CheckFault(lintChanNum, gudtExtreme(lintChanNum).slope.low.Value, gudtExtreme(lintChanNum).slope.high.Value, gudtTest(lintChanNum).slope.low, gudtTest(lintChanNum).slope.high, LOWSLOPE, HIGHSLOPE, gintFailure())
        Call Calc.CheckSevere(lintChanNum, gudtExtreme(lintChanNum).slope.low.Value, gudtExtreme(lintChanNum).slope.high.Value, gudtControl(lintChanNum).slope.low, gudtControl(lintChanNum).slope.high, LOWSLOPE, HIGHSLOPE, gintSevere())
    
        'Calculate Hysteresis
        ReDim lsngHysteresis(gintMaxData)
        Call Calc.CalcHysteresis(lsngFwdVoltageGradient(), lsngRevVoltageGradient(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gsngResolution, lsngHysteresis())
        Call Calc.CalcMinMax(lsngHysteresis(), gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, gsngResolution, lsngMinVal, lsngMinLoc, lsngMaxVal, lsngMaxLoc)
        If gintAnomaly Then Exit Sub                'Exit on system error
        'Select the peak Hysteresis
        If Abs(lsngMaxVal) > Abs(lsngMinVal) Then
            gudtExtreme(lintChanNum).hysteresis.Value = lsngMaxVal
            gudtExtreme(lintChanNum).hysteresis.location = lsngMaxLoc
        Else
            gudtExtreme(lintChanNum).hysteresis.Value = lsngMinVal
            gudtExtreme(lintChanNum).hysteresis.location = lsngMinLoc
        End If
    
        'Calculate Full-Close Hysteresis
        gudtReading(lintChanNum).FullCloseHys.Value = gudtReading(lintChanNum).Index(1).Value - gudtReading(lintChanNum).Index(4).Value
        Call Calc.CheckFault(lintChanNum, gudtReading(lintChanNum).FullCloseHys.Value, gudtReading(lintChanNum).FullCloseHys.Value, gudtTest(lintChanNum).FullCloseHys.low, gudtTest(lintChanNum).FullCloseHys.high, LOWFCHYS, HIGHFCHYS, gintFailure())
    
        'Graphing for Current Channel
        If gblnGraphEnable = True Then
            'Initialize the Graph for Forward Voltage Gradient
            Call LoadMultipleGraphArray(0, (gudtTest(lintChanNum).evaluate.stop - gudtTest(lintChanNum).evaluate.start) * gsngResolution, lsngFwdVoltageGradient(), gsngMultipleGraphArray())
            If gintAnomaly Then Exit Sub                'Exit on system error
            'Initialize the Graph for Reverse Voltage Gradient
            Call LoadMultipleGraphArray(1, (gudtTest(lintChanNum).evaluate.stop - gudtTest(lintChanNum).evaluate.start) * gsngResolution, lsngRevVoltageGradient(), gsngMultipleGraphArray())
            'Display the Voltage Gradient Graphs
            lstrTitle(0) = "Forward": lstrTitle(1) = "Reverse"
            Call LoadGraphArray(lintChanNum, 2, "Voltage Gradient", lstrTitle(), "Position (°)", "Output (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 110, -10, gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, gsngMultipleGraphArray(), msngMaxVoltageGradientLimit(), msngMinVoltageGradientLimit())
            'Initialize and Display the Graph for SinglePoint Linearity Deviation
            Call LoadGraphArray(lintChanNum, 1, "Single Point Linearity Deviation", lstrTitle(), "Position (°)", "Deviation from Nominal (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 8, -8, gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(lintRegionCount2).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngSinglePointLinDev(), msngMaxLinearityLimit(), msngMinLinearityLimit()) '2.7ANM
            'Initialize and Display the Graph for Absolute Linearity Deviation
            Call LoadGraphArray(lintChanNum, 1, "Absolute Linearity Deviation", lstrTitle(), "Position (°)", "Deviation from Nominal (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 8, -8, gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(lintRegionCount).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngAbsLinDev(), msngMaxAbsLinearityLimit(), msngMinAbsLinearityLimit()) '2.7ANM
            'Initialize and Display the Graph for Slope Deviation
            Call LoadGraphArray(lintChanNum, 1, "Slope Deviation", lstrTitle(), "Position (°)", "Ratio (Actual Slope / Ideal Slope)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 2, 0, gudtTest(lintChanNum).slope.start, (gudtTest(lintChanNum).slope.stop - (gudtMachine.slopeInterval / gsngResolution)), (lsngIncrement * gudtMachine.slopeIncrement), lstrGraphFooterhigh, lstrGraphFooterLow, lsngSlopeDev(), msngMaxSlopeDeviationLimit(), msngMinSlopeDeviationLimit())
            'Initialize and Display the Graph for Hysteresis
            Call LoadGraphArray(lintChanNum, 1, "Hysteresis", lstrTitle(), "Position (°)", "Hysteresis (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 6, -6, gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngHysteresis(), msngMaxHysteresisLimit(), msngMinHysteresisLimit())
        End If
    
    Next lintChanNum
    
    'Re-Calculate Voltage Gradients for Output Correlation
    ReDim lsngFwdVoltageGradient(gintMaxData)
    Call Calc.CalcRatiometricDataArray(CHAN0, gintForward(), gintForSupply(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, gsngResolution, lsngFwdVoltageGradient())
    If gintAnomaly Then Exit Sub            'Exit on system error
    ReDim lsngFwdVoltageGradient2(gintMaxData)
    Call Calc.CalcRatiometricDataArray(CHAN1, gintForward(), gintForSupply(), gudtTest(CHAN1).evaluate.start, gudtTest(CHAN1).evaluate.stop, gsngResolution, lsngFwdVoltageGradient2())
    If gintAnomaly Then Exit Sub            'Exit on system error
    
    ReDim lsngRevVoltageGradient(gintMaxData)
    Call Calc.CalcRatiometricDataArray(CHAN0, gintReverse(), gintRevSupply(), gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, gsngResolution, lsngRevVoltageGradient())
    If gintAnomaly Then Exit Sub            'Exit on system error
    ReDim lsngRevVoltageGradient2(gintMaxData)
    Call Calc.CalcRatiometricDataArray(CHAN1, gintReverse(), gintRevSupply(), gudtTest(CHAN1).evaluate.start, gudtTest(CHAN1).evaluate.stop, gsngResolution, lsngRevVoltageGradient2())
    If gintAnomaly Then Exit Sub            'Exit on system error
    
    'Calculate Forward Output Correlation
    ReDim lsngFwdOutputCor(gintMaxData)
    
    '2.2ANM \/\/
    If gblnUseNewAmad Then
        lstrParameterName = "FwdOutputCorrelation"
        lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
    Else
        lintRegionCount = 5
    End If
    
    For lintRegion = 1 To lintRegionCount
        Call Calc.CalcOutputCor705(gudtTest(CHAN0).fwdOutputCor(lintRegion).ideal, lsngFwdVoltageGradient(), lsngFwdVoltageGradient2(), gudtTest(CHAN0).fwdOutputCor(lintRegion).start.location, gudtTest(CHAN0).fwdOutputCor(lintRegion).stop.location, gsngResolution, lsngFwdOutputCor())
        If gintAnomaly Then Exit Sub                    'Exit on system error
    Next lintRegion
    '2.2ANM /\/\
    
    'Find Min/Max Output Correlation
    Call Calc.CalcMinMax(lsngFwdOutputCor(), gudtTest(CHAN0).fwdOutputCor(1).start.location, gudtTest(CHAN0).fwdOutputCor(lintRegionCount).stop.location, gsngResolution, gudtExtreme(CHAN0).fwdOutputCor.low.Value, gudtExtreme(CHAN0).fwdOutputCor.low.location, gudtExtreme(CHAN0).fwdOutputCor.high.Value, gudtExtreme(CHAN0).fwdOutputCor.high.location)
    If gintAnomaly Then Exit Sub                        'Exit on system error
    
    'Calculate and Check Forward Output Correlation Percentage of Tolerance
    ReDim lsngFwdOutputCorPerTol(gintMaxData)
    Call Calc.CalcPercentTol(lsngFwdOutputCor(), msngMinFwdOutputCorLimit(), msngMaxFwdOutputCorLimit(), gudtTest(CHAN0).fwdOutputCor(1).start.location, gudtTest(CHAN0).fwdOutputCor(lintRegionCount).stop.location, gsngResolution, lsngFwdOutputCorPerTol())
    If gintAnomaly Then Exit Sub                        'Exit on system error
    Call Calc.CalcMinMax(lsngFwdOutputCorPerTol(), gudtTest(CHAN0).fwdOutputCor(1).start.location, gudtTest(CHAN0).fwdOutputCor(lintRegionCount).stop.location, gsngResolution, lsngMinVal, lsngMinLoc, lsngMaxVal, lsngMaxLoc)
    If gintAnomaly Then Exit Sub                        'Exit on system error
    'Select the peak percentage of tolerance
    If Abs(lsngMaxVal) > Abs(lsngMinVal) Then
        gudtExtreme(CHAN0).outputCorPerTol(1).Value = lsngMaxVal
        gudtExtreme(CHAN0).outputCorPerTol(1).location = lsngMaxLoc
    Else
        gudtExtreme(CHAN0).outputCorPerTol(1).Value = lsngMinVal
        gudtExtreme(CHAN0).outputCorPerTol(1).location = lsngMinLoc
    End If
    Call Calc.CheckFault(CHAN0, gudtExtreme(CHAN0).outputCorPerTol(1).Value, gudtExtreme(CHAN0).outputCorPerTol(1).Value, -HUNDREDPERCENT, HUNDREDPERCENT, LOWFWDOUTPUTCOR, HIGHFWDOUTPUTCOR, gintFailure())
    Call Calc.CheckSevere(CHAN0, gudtExtreme(CHAN0).outputCorPerTol(1).Value, gudtExtreme(CHAN0).outputCorPerTol(1).Value, gudtControl(CHAN0).outputCorPerTol(1).low, gudtControl(CHAN0).outputCorPerTol(1).high, LOWFWDOUTPUTCOR, HIGHFWDOUTPUTCOR, gintSevere())
    
    'Calculate Reverse Output Correlation
    ReDim lsngRevOutputCor(gintMaxData)
    
    '2.2ANM \/\/
    If gblnUseNewAmad Then
        lstrParameterName = "RevOutputCorrelation"
        lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
    Else
        lintRegionCount = 5
    End If
    
    For lintRegion = 1 To lintRegionCount
        Call Calc.CalcOutputCor705(gudtTest(CHAN0).revOutputCor(lintRegion).ideal, lsngRevVoltageGradient(), lsngRevVoltageGradient2(), gudtTest(CHAN0).revOutputCor(lintRegion).start.location, gudtTest(CHAN0).revOutputCor(lintRegion).stop.location, gsngResolution, lsngRevOutputCor())
        If gintAnomaly Then Exit Sub                    'Exit on system error
    Next lintRegion
    '2.2ANM /\/\
    
    'Find Min/Max Output Correlation
    Call Calc.CalcMinMax(lsngRevOutputCor(), gudtTest(CHAN0).revOutputCor(1).start.location, gudtTest(CHAN0).revOutputCor(lintRegionCount).stop.location, gsngResolution, gudtExtreme(CHAN0).revOutputCor.low.Value, gudtExtreme(CHAN0).revOutputCor.low.location, gudtExtreme(CHAN0).revOutputCor.high.Value, gudtExtreme(CHAN0).revOutputCor.high.location)
    If gintAnomaly Then Exit Sub                        'Exit on system error
    
    'Calculate and Check Forward Output Correlation Percentage of Tolerance
    ReDim lsngRevOutputCorPerTol(gintMaxData)
    Call Calc.CalcPercentTol(lsngRevOutputCor(), msngMinRevOutputCorLimit(), msngMaxRevOutputCorLimit(), gudtTest(CHAN0).revOutputCor(1).start.location, gudtTest(CHAN0).revOutputCor(lintRegionCount).stop.location, gsngResolution, lsngRevOutputCorPerTol())
    If gintAnomaly Then Exit Sub                        'Exit on system error
    Call Calc.CalcMinMax(lsngRevOutputCorPerTol(), gudtTest(CHAN0).revOutputCor(1).start.location, gudtTest(CHAN0).revOutputCor(lintRegionCount).stop.location, gsngResolution, lsngMinVal, lsngMinLoc, lsngMaxVal, lsngMaxLoc)
    If gintAnomaly Then Exit Sub                        'Exit on system error
    'Select the peak percentage of tolerance
    If Abs(lsngMaxVal) > Abs(lsngMinVal) Then
        gudtExtreme(CHAN0).outputCorPerTol(2).Value = lsngMaxVal
        gudtExtreme(CHAN0).outputCorPerTol(2).location = lsngMaxLoc
    Else
        gudtExtreme(CHAN0).outputCorPerTol(2).Value = lsngMinVal
        gudtExtreme(CHAN0).outputCorPerTol(2).location = lsngMinLoc
    End If
    Call Calc.CheckFault(CHAN0, gudtExtreme(CHAN0).outputCorPerTol(2).Value, gudtExtreme(CHAN0).outputCorPerTol(2).Value, -HUNDREDPERCENT, HUNDREDPERCENT, LOWREVOUTPUTCOR, HIGHREVOUTPUTCOR, gintFailure())
    Call Calc.CheckSevere(CHAN0, gudtExtreme(CHAN0).outputCorPerTol(2).Value, gudtExtreme(CHAN0).outputCorPerTol(2).Value, gudtControl(CHAN0).outputCorPerTol(2).low, gudtControl(CHAN0).outputCorPerTol(2).high, LOWREVOUTPUTCOR, HIGHREVOUTPUTCOR, gintSevere())

    'MLX Current '2.5ANM \/\/
    If Not gblnBnmkTest Then '2.8bANM
        Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).mlxCurrent, gudtReading(CHAN0).mlxCurrent, gudtTest(CHAN0).mlxCurrent.low, gudtTest(CHAN0).mlxCurrent.high, LOWMLXI, HIGHMLXI, gintFailure())
        Call Calc.CheckFault(CHAN1, gudtReading(CHAN1).mlxCurrent, gudtReading(CHAN1).mlxCurrent, gudtTest(CHAN1).mlxCurrent.low, gudtTest(CHAN1).mlxCurrent.high, LOWMLXI, HIGHMLXI, gintFailure())
        Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).mlxCurrent, gudtReading(CHAN0).mlxCurrent, gudtControl(CHAN0).mlxCurrent.low, gudtControl(CHAN0).mlxCurrent.high, LOWMLXI, HIGHMLXI, gintSevere()) '2.8aANM
        Call Calc.CheckSevere(CHAN1, gudtReading(CHAN1).mlxCurrent, gudtReading(CHAN1).mlxCurrent, gudtControl(CHAN1).mlxCurrent.low, gudtControl(CHAN1).mlxCurrent.high, LOWMLXI, HIGHMLXI, gintSevere()) '2.8aANM
        '2.8cANM \/\/
        Call Calc.CheckFault(CHAN0, gudtReading(CHAN0).mlxWCurrent, gudtReading(CHAN0).mlxWCurrent, gudtReading(CHAN0).mlxCurrent + gudtTest(CHAN0).mlxWCurrent.low, gudtReading(CHAN0).mlxCurrent + gudtTest(CHAN0).mlxWCurrent.high, LOWMLXI2, HIGHMLXI2, gintFailure())
        Call Calc.CheckFault(CHAN1, gudtReading(CHAN1).mlxWCurrent, gudtReading(CHAN1).mlxWCurrent, gudtReading(CHAN1).mlxCurrent + gudtTest(CHAN1).mlxWCurrent.low, gudtReading(CHAN1).mlxCurrent + gudtTest(CHAN1).mlxWCurrent.high, LOWMLXI2, HIGHMLXI2, gintFailure())
        Call Calc.CheckSevere(CHAN0, gudtReading(CHAN0).mlxWCurrent, gudtReading(CHAN0).mlxWCurrent, gudtReading(CHAN0).mlxCurrent + gudtTest(CHAN0).mlxWCurrent.low, gudtReading(CHAN0).mlxCurrent + gudtTest(CHAN0).mlxWCurrent.high, LOWMLXI2, HIGHMLXI2, gintSevere())
        Call Calc.CheckSevere(CHAN1, gudtReading(CHAN1).mlxWCurrent, gudtReading(CHAN1).mlxWCurrent, gudtReading(CHAN1).mlxCurrent + gudtTest(CHAN1).mlxWCurrent.low, gudtReading(CHAN1).mlxCurrent + gudtTest(CHAN1).mlxWCurrent.high, LOWMLXI2, HIGHMLXI2, gintSevere())
    End If
Else
    'Reset values to zero each scan
    For lintChanNum = CHAN0 To MAXCHANNUM
        gudtReading(lintChanNum).Index(1).Value = 0
        gudtReading(lintChanNum).outputAtForceKnee = 0
        gudtReading(lintChanNum).Index(2).Value = 0
        gudtReading(lintChanNum).Index(3).Value = 0
        gudtReading(lintChanNum).Index(3).location = 0
        gudtReading(lintChanNum).maxOutput.Value = 0
        gudtReading(lintChanNum).maxOutput.location = 0
        gudtReading(lintChanNum).mlxCurrent = 0 '2.5ANM
        gudtReading(lintChanNum).mlxWCurrent = 0 '2.8cANM
        gudtExtreme(lintChanNum).linDevPerTol(1).Value = 0
        gudtExtreme(lintChanNum).linDevPerTol(1).location = 0
        gudtExtreme(lintChanNum).SinglePointLin.low.Value = 0
        gudtExtreme(lintChanNum).SinglePointLin.low.location = 0
        gudtExtreme(lintChanNum).SinglePointLin.high.Value = 0
        gudtExtreme(lintChanNum).SinglePointLin.high.location = 0
        gudtExtreme(lintChanNum).linDevPerTol(2).Value = 0          '2.7ANM \/\/
        gudtExtreme(lintChanNum).linDevPerTol(2).location = 0
        gudtExtreme(lintChanNum).AbsLin.low.Value = 0
        gudtExtreme(lintChanNum).AbsLin.low.location = 0
        gudtExtreme(lintChanNum).AbsLin.high.Value = 0
        gudtExtreme(lintChanNum).AbsLin.high.location = 0           '2.7ANM /\/\
        gudtExtreme(lintChanNum).slope.low.Value = 0
        gudtExtreme(lintChanNum).slope.low.location = 0
        gudtExtreme(lintChanNum).slope.high.Value = 0
        gudtExtreme(lintChanNum).slope.high.location = 0
        gudtExtreme(lintChanNum).hysteresis.Value = 0
        gudtExtreme(lintChanNum).hysteresis.location = 0
        gudtReading(lintChanNum).FullCloseHys.Value = 0
        
        'Fill graphs with 0s
        ReDim lsngFwdVoltageGradient(gintMaxData)
        ReDim lsngRevVoltageGradient(gintMaxData)
        ReDim lsngFwdVoltageGradient2(gintMaxData)
        ReDim lsngRevVoltageGradient2(gintMaxData)
        ReDim lsngHysteresis(gintMaxData)
        ReDim lsngSlopeDev((gudtTest(lintChanNum).slope.start * gsngResolution) / gudtMachine.slopeIncrement To (gudtTest(lintChanNum).slope.stop * gsngResolution) / gudtMachine.slopeIncrement)
        ReDim lsngLinDevPerTol(gintMaxData)
        ReDim lsngSinglePointLinDev(gintMaxData)
        ReDim lsngLinDevPerTol2(gintMaxData) '2.7ANM
        ReDim lsngAbsLinDev(gintMaxData)     '2.7ANM
        ReDim lsngFwdOutputCorPerTol(gintMaxData)
        ReDim lsngFwdOutputCor(gintMaxData)
        ReDim lsngRevOutputCorPerTol(gintMaxData)
        ReDim lsngRevOutputCor(gintMaxData)
        
        'Calculate Limit Arrays
        Call EvaluateLimits(lintChanNum)
        
        'Graphing for Current Channel
        If gblnGraphEnable = True Then
            'Initialize the Graph for Forward Voltage Gradient
            Call LoadMultipleGraphArray(0, (gudtTest(lintChanNum).evaluate.stop - gudtTest(lintChanNum).evaluate.start) * gsngResolution, lsngFwdVoltageGradient(), gsngMultipleGraphArray())
            If gintAnomaly Then Exit Sub                'Exit on system error
            'Initialize the Graph for Reverse Voltage Gradient
            Call LoadMultipleGraphArray(1, (gudtTest(lintChanNum).evaluate.stop - gudtTest(lintChanNum).evaluate.start) * gsngResolution, lsngRevVoltageGradient(), gsngMultipleGraphArray())
            'Display the Voltage Gradient Graphs
            lstrTitle(0) = "Forward": lstrTitle(1) = "Reverse"
            Call LoadGraphArray(lintChanNum, 2, "Voltage Gradient", lstrTitle(), "Position (°)", "Output (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 110, -10, gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, gsngMultipleGraphArray(), msngMaxVoltageGradientLimit(), msngMinVoltageGradientLimit())
            'Initialize and Display the Graph for SinglePoint Linearity Deviation
            Call LoadGraphArray(lintChanNum, 1, "Single Point Linearity Deviation", lstrTitle(), "Position (°)", "Deviation from Nominal (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 8, -8, gudtTest(lintChanNum).SinglePointLin(1).start.location, gudtTest(lintChanNum).SinglePointLin(5).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngSinglePointLinDev(), msngMaxLinearityLimit(), msngMinLinearityLimit())
            'Initialize and Display the Graph for Absolute Linearity Deviation
            Call LoadGraphArray(lintChanNum, 1, "Absolute Linearity Deviation", lstrTitle(), "Position (°)", "Deviation from Nominal (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 8, -8, gudtTest(lintChanNum).AbsLin(1).start.location, gudtTest(lintChanNum).AbsLin(5).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngAbsLinDev(), msngMaxAbsLinearityLimit(), msngMinAbsLinearityLimit()) '2.7ANM
            'Initialize and Display the Graph for Slope Deviation
            Call LoadGraphArray(lintChanNum, 1, "Slope Deviation", lstrTitle(), "Position (°)", "Ratio (Actual Slope / Ideal Slope)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 2, 0, gudtTest(lintChanNum).slope.start, (gudtTest(lintChanNum).slope.stop - (gudtMachine.slopeInterval / gsngResolution)), (lsngIncrement * gudtMachine.slopeIncrement), lstrGraphFooterhigh, lstrGraphFooterLow, lsngSlopeDev(), msngMaxSlopeDeviationLimit(), msngMinSlopeDeviationLimit())
            'Initialize and Display the Graph for Hysteresis
            Call LoadGraphArray(lintChanNum, 1, "Hysteresis", lstrTitle(), "Position (°)", "Hysteresis (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 6, -6, gudtTest(lintChanNum).evaluate.start, gudtTest(lintChanNum).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngHysteresis(), msngMaxHysteresisLimit(), msngMinHysteresisLimit())
        End If
    Next lintChanNum
    
    gudtExtreme(CHAN0).outputCorPerTol(1).Value = 0
    gudtExtreme(CHAN0).outputCorPerTol(1).location = 0
    gudtExtreme(CHAN0).outputCorPerTol(2).Value = 0
    gudtExtreme(CHAN0).outputCorPerTol(2).location = 0
    lintRegionCount = 5
End If

If gblnGraphEnable Then
    'Display the Forward Output Correlation Graph
    Call LoadGraphArray(-1, 1, "Forward Output Correlation", lstrTitle(), "Position (°)", "Output Correlation (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 4, -4, gudtTest(CHAN0).fwdOutputCor(1).start.location, gudtTest(CHAN0).fwdOutputCor(lintRegionCount).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngFwdOutputCor(), msngMaxFwdOutputCorLimit(), msngMinFwdOutputCorLimit())

    'Display the Reverse Output Correlation Graph
    Call LoadGraphArray(-1, 1, "Reverse Output Correlation", lstrTitle(), "Position (°)", "Output Correlation (% of Applied)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 4, -4, gudtTest(CHAN0).revOutputCor(1).start.location, gudtTest(CHAN0).revOutputCor(lintRegionCount).stop.location, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngRevOutputCor(), msngMaxRevOutputCorLimit(), msngMinRevOutputCorLimit())

    'Initialize the Graph for Forward Force
    Call LoadMultipleGraphArray(0, (gudtTest(CHAN0).evaluate.stop - gudtTest(CHAN0).evaluate.start) * gsngResolution, lsngFwdForce(), gsngMultipleGraphArray())
    'Initialize the Graph for Reverse Force
    Call LoadMultipleGraphArray(1, (gudtTest(CHAN0).evaluate.stop - gudtTest(CHAN0).evaluate.start) * gsngResolution, lsngRevForce(), gsngMultipleGraphArray())

    'Display the Graph for Forward & Reverse
    lstrTitle(0) = "Forward": lstrTitle(1) = "Reverse"
    Call LoadGraphArray(-1, 2, "Force", lstrTitle(), "Position (°)", "Force (Newtons)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 100, 0, gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, gsngMultipleGraphArray(), msngMaxForceGradientLimit(), msngMinForceGradientLimit())

    'Display the Mechanical Hysteresis Graph Array
    Call LoadGraphArray(-1, 1, "Mechanical Hysteresis", lstrTitle(), "Position (°)", "Mechanical Hysteresis (% of Forward Scan Force)", gudtMachine.xAxisLow, gudtMachine.xAxisHigh, 110, -10, gudtTest(CHAN0).evaluate.start, gudtTest(CHAN0).evaluate.stop, lsngIncrement, lstrGraphFooterhigh, lstrGraphFooterLow, lsngMechHysteresis(), msngMaxMechHysteresisLimit(), msngMinMechHysteresisLimit())
End If

'NOTE: No GraphZeroOffset!

If gintAnomaly = 0 Then
    'Display the loaded graph arrays
    Call frmMain.ctrResultsTabs1.ExtractDataEvenXIntervals(gvntGraph())
    If Not gblnForceOnly Then   '1.6ANM
        'Save the raw data if called for   '2.1ANM
        If gblnSaveRawData Then Call Pedal.SaveTLRawDataToFile(CLng((gudtTest(CHAN0).evaluate.stop - gudtTest(CHAN0).evaluate.start) * gsngResolution), lsngFwdVoltageGradient(), lsngFwdVoltageGradient2(), lsngRevVoltageGradient(), lsngRevVoltageGradient2(), lsngFwdForce(), lsngRevForce())
    Else
        'Save the raw data if called for   '2.6ANM
        If gblnSaveRawData Then Call Pedal.SaveFORawDataToFile(CLng((gudtTest(CHAN0).evaluate.stop - gudtTest(CHAN0).evaluate.start) * gsngResolution), lsngFwdForce(), lsngRevForce())
    End If
End If

Exit Sub
DataEvaluationError:
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Software error in DataEvaluation: " & Err.Description, True, True)

End Sub

Public Sub DisplayScanResultsCountsPrioritized()
'
'   PURPOSE: To display the failure counts to the screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lvntHighCount(1 To NUMROWSSCANRESULTSDISPLAY) As Variant
Dim lvntLowCount(1 To NUMROWSSCANRESULTSDISPLAY) As Variant

'Rows:
'Pedal at Rest
'Output #1 Label (no counts)
'Index 1 (FullClose Output), output 1
'Index 3 (FullOpen Output), output 1
'Maximum Output, output 1
'SinglePoint Linearity Deviation Percentage of Tolerance, output 1
'SinglePoint Linearity Deviation Max, output 1 (no counts)
'SinglePoint Linearity Deviation Min, output 1 (no counts)
'Absolute Linearity Deviation Percentage of Tolerance, output 1
'Absolute Linearity Deviation Max, output 1 (no counts)
'Absolute Linearity Deviation Min, output 1 (no counts)
'Slope Deviation Max, output 1
'Slope Deviation Min, output 1
'Full-Close Hysteresis, output 1
'Output #2 Label (no counts)
'MLX Current, output 1
'MLX WOT Current, output 1
'Output #2 Label (no counts)
'Index 1 (FullClose Output), output 2
'Index 3 (FullOpen Output), output 2
'Maximum Output, output 2
'SinglePoint Linearity Deviation Percentage of Tolerance, output 2
'SinglePoint Linearity Deviation Max, output 2 (no counts)
'SinglePoint Linearity Deviation Min, output 2 (no counts)
'Absolute Linearity Deviation Percentage of Tolerance, output 2
'Absolute Linearity Deviation Max, output 2 (no counts)
'Absolute Linearity Deviation Min, output 2 (no counts)
'Slope Deviation Max, output 2
'Slope Deviation Min, output 2
'Full-Close Hysteresis, output 2
'MLX Current, output 2
'MLX WOT Current, output 2
'Correlation Label (no counts)
'Forward Output Correlation Percentage of Tolerance
'Forward Output Correlation Max (no counts)
'Forward Output Correlation Min (no counts)
'Reverse Output Correlation Percentage of Tolerance
'Reverse Output Correlation Max (no counts)
'Reverse Output Correlation Min (no counts)
'Force Label (no counts)
'Forward Force Point 1
'Forward Force Point 3
'Reverse Force Point 1
'Reverse Force Point 3
'Peak Force
'Mechanical Hysteresis Point 1
'Mechanical Hysteresis Point 3

lintRow = 1   'Initialize the Row Number

'Pedal at Rest Location '2.8ANM
lvntHighCount(lintRow) = gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low

lintRow = lintRow + 1

For lintChanNum = CHAN0 To MAXCHANNUM

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (FullClose Output)
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).Index(1).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).Index(1).failCount.low

    lintRow = lintRow + 1

    'Index 3 (FullOpen Location)
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).Index(3).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).Index(3).failCount.low

    lintRow = lintRow + 1

    'Maximum Output
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).maxOutput.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).maxOutput.failCount.low

    lintRow = lintRow + 1

    'SinglePoint Linearity Deviation % of Tolerance
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low

    lintRow = lintRow + 1

    'High SinglePoint Linearity Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Low SinglePoint Linearity Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    '2.7ANM \/\/
    'Absolute Linearity Deviation % of Tolerance
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low

    lintRow = lintRow + 1

    'High Absolute Linearity Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Low Absolute Linearity Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1
    '2.7ANM /\/\
    
    'High Slope Deviation
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).slopeMax.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Low Slope Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).slopeMin.failCount.low

    lintRow = lintRow + 1

    'Full-Close Hysteresis
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).FullCloseHys.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).FullCloseHys.failCount.low

    lintRow = lintRow + 1
    
    'MLX Current '2.5ANM
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).mlxCurrent.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).mlxCurrent.failCount.low
    
    lintRow = lintRow + 1

    'MLX WOT Current '2.8cANM
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).mlxWCurrent.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).mlxWCurrent.failCount.low
    
    lintRow = lintRow + 1
    
Next lintChanNum

'Correlation Label
lintRow = lintRow + 1

'Forward Output Correlation % of Tolerance
lvntHighCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low

lintRow = lintRow + 1

'High Forward Output Correlation
lvntHighCount(lintRow) = "N/A"
lvntLowCount(lintRow) = "N/A"

lintRow = lintRow + 1

'Low Forward Output Correlation
lvntHighCount(lintRow) = "N/A"
lvntLowCount(lintRow) = "N/A"

lintRow = lintRow + 1

'Reverse Output Correlation % of Tolerance
lvntHighCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low

lintRow = lintRow + 1

'High Reverse Output Correlation
lvntHighCount(lintRow) = "N/A"
lvntLowCount(lintRow) = "N/A"

lintRow = lintRow + 1

'Low Reverse Output Correlation
lvntHighCount(lintRow) = "N/A"
lvntLowCount(lintRow) = "N/A"

lintRow = lintRow + 1

'Force Label
lintRow = lintRow + 1

'Force Knee Location
'2.3ANM lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).forceKneeLoc.failCount.high
'2.3ANM lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).forceKneeLoc.failCount.low

'2.3ANM lintRow = lintRow + 1

'2.0*ANM
''Force Knee Forward Force
'lvntHighCount(lintRow) = gudtScanStats(CHAN0).forceKneeForce.failCount.high
'lvntLowCount(lintRow) = gudtScanStats(CHAN0).forceKneeForce.failCount.low
'
'lintRow = lintRow + 1

'Forward Force Point 1
lvntHighCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(1).failCount.low

lintRow = lintRow + 1

'Forward Force Point 3
lvntHighCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(3).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(3).failCount.low

lintRow = lintRow + 1

'Reverse Force Point 1
lvntHighCount(lintRow) = gudtScanStats(CHAN0).revForcePt(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).revForcePt(1).failCount.low

lintRow = lintRow + 1

'Reverse Force Point 3
lvntHighCount(lintRow) = gudtScanStats(CHAN0).revForcePt(3).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).revForcePt(3).failCount.low

lintRow = lintRow + 1

'Peak Force
lvntHighCount(lintRow) = gudtScanStats(CHAN0).peakForce.failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).peakForce.failCount.low

lintRow = lintRow + 1

'Mechanical Hysteresis Point 1
lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(1).failCount.high
lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(1).failCount.low

lintRow = lintRow + 1

'Mechanical Hysteresis Point 3
lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(3).failCount.high
lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(3).failCount.low

'2.8dANM \/\/
'Display Kickdown Counts if Appropriate
If gblnKD Then
    lintRow = lintRow + 1

    'Kickdown Label
    lintRow = lintRow + 1

    'Kickdown Start Location
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDStart.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDStart.failCount.low

    lintRow = lintRow + 1

    'Kickdown Force Span
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDSpan.failCount.high '2.8dANM
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDSpan.failCount.low   '2.8dANM

    lintRow = lintRow + 1
    
    'Kickdown Peak Location
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Kickdown Peak Force
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Kickdown End Location
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDStop.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDStop.failCount.low
End If
'2.8dANM /\/\

'Send the counts to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateResultsCounts(SCANRESULTSGRID, llngRowNum, lvntHighCount(llngRowNum), lvntLowCount(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayScanResultsData()

'
'   PURPOSE: To display the results Data to the screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lstrValueAndLocation(1 To NUMROWSSCANRESULTSDISPLAY) As String
Dim lprdParameterResults(1 To NUMROWSSCANRESULTSDISPLAY) As ParameterResultsDisplay

'Default all parameters to REJECT
For llngRowNum = 1 To NUMROWSSCANRESULTSDISPLAY
    lprdParameterResults(llngRowNum) = prdReject
Next llngRowNum

'Rows:
'Pedal at Rest
'Output #1 Label (no counts)
'Index 1 (FullClose Output), output 1
'Index 3 (FullOpen Output), output 1
'Maximum Output, output 1
'SinglePoint Linearity Deviation Percentage of Tolerance, output 1
'SinglePoint Linearity Deviation Max, output 1 (no counts)
'SinglePoint Linearity Deviation Min, output 1 (no counts)
'Absolute Linearity Deviation Percentage of Tolerance, output 1
'Absolute Linearity Deviation Max, output 1 (no counts)
'Absolute Linearity Deviation Min, output 1 (no counts)
'Slope Deviation Max, output 1
'Slope Deviation Min, output 1
'Full-Close Hysteresis, output 1
'Output #2 Label (no counts)
'MLX Current, output 1
'MLX WOT Current, output 1
'Output #2 Label (no counts)
'Index 1 (FullClose Output), output 2
'Index 3 (FullOpen Output), output 2
'Maximum Output, output 2
'SinglePoint Linearity Deviation Percentage of Tolerance, output 2
'SinglePoint Linearity Deviation Max, output 2 (no counts)
'SinglePoint Linearity Deviation Min, output 2 (no counts)
'Absolute Linearity Deviation Percentage of Tolerance, output 2
'Absolute Linearity Deviation Max, output 2 (no counts)
'Absolute Linearity Deviation Min, output 2 (no counts)
'Slope Deviation Max, output 2
'Slope Deviation Min, output 2
'Full-Close Hysteresis, output 2
'MLX Current, output 2
'MLX WOT Current, output 2
'Correlation Label (no counts)
'Forward Output Correlation Percentage of Tolerance
'Forward Output Correlation Max (no counts)
'Forward Output Correlation Min (no counts)
'Reverse Output Correlation Percentage of Tolerance
'Reverse Output Correlation Max (no counts)
'Reverse Output Correlation Min (no counts)
'Force Label (no counts)
'Forward Force Point 1
'Forward Force Point 3
'Reverse Force Point 1
'Reverse Force Point 3
'Peak Force
'Mechanical Hysteresis Point 1
'Mechanical Hysteresis Point 3

lintRow = 1   'Initialize the Row Number

'Pedal at Rest Location '2.8ANM
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).pedalAtRestLoc, "##0.00") & "° "
If Not (gintFailure(lintChanNum, HIGHPEDALATREST) Or gintFailure(lintChanNum, LOWPEDALATREST) Or gintSevere(lintChanNum, HIGHPEDALATREST) Or gintSevere(lintChanNum, LOWPEDALATREST)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

For lintChanNum = CHAN0 To MAXCHANNUM

    'Output # Label
    lprdParameterResults(lintRow) = prdEmpty
    lintRow = lintRow + 1

    'Index 1 (FullClose Output)
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).Index(1).Value, "##0.00") & "% "
    If Not (gintFailure(lintChanNum, HIGHINDEXPT1) Or gintFailure(lintChanNum, LOWINDEXPT1) Or gintSevere(lintChanNum, HIGHINDEXPT1) Or gintSevere(lintChanNum, LOWINDEXPT1)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Index 3 (FullOpen Output)
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).Index(3).Value, "##0.00") & "% " '2.2ANM fixed bug
    If Not (gintFailure(lintChanNum, HIGHINDEXPT3) Or gintFailure(lintChanNum, LOWINDEXPT3) Or gintSevere(lintChanNum, HIGHINDEXPT3) Or gintSevere(lintChanNum, LOWINDEXPT3)) Then
        lprdParameterResults(lintRow) = prdGood
    End If
    
    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Maximum Output
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).maxOutput.Value, "##0.00") & "% at " & Format(gudtReading(lintChanNum).maxOutput.location, "##0.00") & "° "
    If Not (gintFailure(lintChanNum, HIGHMAXOUTPUT) Or gintFailure(lintChanNum, LOWMAXOUTPUT) Or gintSevere(lintChanNum, HIGHMAXOUTPUT) Or gintSevere(lintChanNum, LOWMAXOUTPUT)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'SinglePoint Linearity Deviation % of Tolerance
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).linDevPerTol(1).Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).linDevPerTol(1).location, "##0.00") & "° "
    If Not (gintFailure(lintChanNum, HIGHSINGLEPOINTLIN) Or gintFailure(lintChanNum, LOWSINGLEPOINTLIN) Or gintSevere(lintChanNum, HIGHSINGLEPOINTLIN) Or gintSevere(lintChanNum, LOWSINGLEPOINTLIN)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Linearity Deviation Max
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).SinglePointLin.high.Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).SinglePointLin.high.location, "##0.00") & "° "
    lprdParameterResults(lintRow) = prdNotChecked

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Linearity Deviation Min
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).SinglePointLin.low.Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).SinglePointLin.low.location, "##0.00") & "° "
    lprdParameterResults(lintRow) = prdNotChecked

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    '2.7ANM \/\/
    'Absolute Linearity Deviation % of Tolerance
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).linDevPerTol(2).Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).linDevPerTol(2).location, "##0.00") & "° "
    If Not (gintFailure(lintChanNum, HIGHABSLIN) Or gintFailure(lintChanNum, LOWABSLIN) Or gintSevere(lintChanNum, HIGHABSLIN) Or gintSevere(lintChanNum, LOWABSLIN)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If

    lintRow = lintRow + 1

    'Linearity Deviation Max
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).AbsLin.high.Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).AbsLin.high.location, "##0.00") & "° "
    lprdParameterResults(lintRow) = prdNotChecked

    If gblnForceOnly Then
        lstrValueAndLocation(lintRow) = "N/A"
    End If

    lintRow = lintRow + 1

    'Linearity Deviation Min
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).AbsLin.low.Value, "##0.000") & "% at " & Format(gudtExtreme(lintChanNum).AbsLin.low.location, "##0.00") & "° "
    lprdParameterResults(lintRow) = prdNotChecked

    If gblnForceOnly Then
        lstrValueAndLocation(lintRow) = "N/A"
    End If

    lintRow = lintRow + 1
    '2.7ANM /\/\
    
    'Slope Deviation Max
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).slope.high.Value, "##0.000") & " at " & Format(gudtExtreme(lintChanNum).slope.high.location, "##0.00") & "° "
    If Not (gintFailure(lintChanNum, HIGHSLOPE) Or gintSevere(lintChanNum, HIGHSLOPE)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Slope Deviation Min
    lstrValueAndLocation(lintRow) = Format(gudtExtreme(lintChanNum).slope.low.Value, "##0.000") & " at " & Format(gudtExtreme(lintChanNum).slope.low.location, "##0.00") & "° "
    If Not (gintFailure(lintChanNum, LOWSLOPE) Or gintSevere(lintChanNum, LOWSLOPE)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'Full-Close Hysteresis
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).FullCloseHys.Value, "##0.000") & "% "
    If Not (gintFailure(lintChanNum, LOWFCHYS) Or gintFailure(lintChanNum, HIGHFCHYS)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Then                                 '1.6ANM \/\/
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If                                                '1.6ANM /\/\

    lintRow = lintRow + 1

    'MLX Current '2.5ANM
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).mlxCurrent, "##0.0") & " mA "
    If Not (gintFailure(lintChanNum, LOWMLXI) Or gintFailure(lintChanNum, HIGHMLXI)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    If gblnForceOnly Or gblnBnmkTest Then '2.8bANM
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If
    
    lintRow = lintRow + 1

    'MLX WOT Current '2.8cANM
    lstrValueAndLocation(lintRow) = Format(gudtReading(lintChanNum).mlxWCurrent, "##0.0") & " mA "
    If Not (gintFailure(lintChanNum, LOWMLXI2) Or gintFailure(lintChanNum, HIGHMLXI2)) Then
        lprdParameterResults(lintRow) = prdGood
    End If
 
    If gblnForceOnly Or gblnBnmkTest Then
        lstrValueAndLocation(lintRow) = "N/A"
        lprdParameterResults(lintRow) = prdNotChecked
    End If
    
    lintRow = lintRow + 1
    
Next lintChanNum

'Correlation Label
lprdParameterResults(lintRow) = prdEmpty

lintRow = lintRow + 1

'Forward Output Correlation % of Tolerance
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).outputCorPerTol(1).Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).outputCorPerTol(1).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHFWDOUTPUTCOR) Or gintFailure(CHAN0, LOWFWDOUTPUTCOR) Or gintSevere(CHAN0, HIGHFWDOUTPUTCOR) Or gintSevere(CHAN0, LOWFWDOUTPUTCOR)) Then
    lprdParameterResults(lintRow) = prdGood
End If

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
    lprdParameterResults(lintRow) = prdNotChecked
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Forward Output Correlation Max
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).fwdOutputCor.high.Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).fwdOutputCor.high.location, "##0.00") & "° "
lprdParameterResults(lintRow) = prdNotChecked

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Forward Output Correlation Min
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).fwdOutputCor.low.Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).fwdOutputCor.low.location, "##0.00") & "° "
lprdParameterResults(lintRow) = prdNotChecked

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Reverse Output Correlation % of Tolerance
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).outputCorPerTol(2).Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).outputCorPerTol(2).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHREVOUTPUTCOR) Or gintFailure(CHAN0, LOWREVOUTPUTCOR) Or gintSevere(CHAN0, HIGHREVOUTPUTCOR) Or gintSevere(CHAN0, LOWREVOUTPUTCOR)) Then
    lprdParameterResults(lintRow) = prdGood
End If

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
    lprdParameterResults(lintRow) = prdNotChecked
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Reverse Output Correlation Max
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).revOutputCor.high.Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).revOutputCor.high.location, "##0.00") & "° "
lprdParameterResults(lintRow) = prdNotChecked

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Reverse Output Correlation Min
lstrValueAndLocation(lintRow) = Format(gudtExtreme(CHAN0).revOutputCor.low.Value, "##0.00") & "% at " & Format(gudtExtreme(CHAN0).revOutputCor.low.location, "##0.00") & "° "
lprdParameterResults(lintRow) = prdNotChecked

If gblnForceOnly Then                                 '1.6ANM \/\/
    lstrValueAndLocation(lintRow) = "N/A"
End If                                                '1.6ANM /\/\

lintRow = lintRow + 1

'Force Label
lprdParameterResults(lintRow) = prdEmpty

lintRow = lintRow + 1

'Force Knee Location
'2.3ANM lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).forceKnee.location, "##0.00") & "° "
'2.2ANM If Not (gintFailure(CHAN0, HIGHFORCEKNEELOC) Or gintFailure(CHAN0, LOWFORCEKNEELOC) Or gintSevere(CHAN0, HIGHFORCEKNEELOC) Or gintSevere(CHAN0, LOWFORCEKNEELOC)) Then
'2.2ANM     lprdParameterResults(lintRow) = prdGood
'2.2ANM End If
'2.3ANM lprdParameterResults(lintRow) = prdNotChecked

'2.3ANM lintRow = lintRow + 1

'2.0*ANM
''Force Knee Forward Force
'lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).forceKnee.Value, "##0.00") & " Newtons "
'If Not (gintFailure(CHAN0, HIGHFORCEKNEEFWDFORCE) Or gintFailure(CHAN0, LOWFORCEKNEEFWDFORCE) Or gintSevere(CHAN0, HIGHFORCEKNEEFWDFORCE) Or gintSevere(CHAN0, LOWFORCEKNEEFWDFORCE)) Then
'    lprdParameterResults(lintRow) = prdGood
'End If
'
'lintRow = lintRow + 1

'Forward Force Point 1
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).fwdForcePt(1).Value, "##0.00") & " Newtons at " & Format(gudtTest(CHAN0).fwdForcePt(1).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHFWDFORCEPT1) Or gintFailure(CHAN0, LOWFWDFORCEPT1) Or gintSevere(CHAN0, HIGHFWDFORCEPT1) Or gintSevere(CHAN0, LOWFWDFORCEPT1)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

'Forward Force Point 3
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).fwdForcePt(3).Value, "##0.00") & " Newtons at " & Format(gudtTest(CHAN0).fwdForcePt(3).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHFWDFORCEPT3) Or gintFailure(CHAN0, LOWFWDFORCEPT3) Or gintSevere(CHAN0, HIGHFWDFORCEPT3) Or gintSevere(CHAN0, LOWFWDFORCEPT3)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

'Reverse Force Point 1
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).revForcePt(1).Value, "##0.00") & " Newtons at " & Format(gudtTest(CHAN0).revForcePt(1).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHREVFORCEPT1) Or gintFailure(CHAN0, LOWREVFORCEPT1) Or gintSevere(CHAN0, HIGHREVFORCEPT1) Or gintSevere(CHAN0, LOWREVFORCEPT1)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

'Reverse Force Point 3
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).revForcePt(3).Value, "##0.00") & " Newtons at " & Format(gudtTest(CHAN0).revForcePt(3).location, "##0.00") & "° "
If Not (gintFailure(CHAN0, HIGHREVFORCEPT3) Or gintFailure(CHAN0, LOWREVFORCEPT3) Or gintSevere(CHAN0, HIGHREVFORCEPT3) Or gintSevere(CHAN0, LOWREVFORCEPT3)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

'Peak Force
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).peakForce, "##0.00") & " Newtons "
If Not (gintFailure(CHAN0, HIGHPEAKFORCE) Or gintFailure(CHAN0, LOWPEAKFORCE) Or gintSevere(CHAN0, HIGHPEAKFORCE) Or gintSevere(CHAN0, LOWPEAKFORCE)) Then
    lprdParameterResults(lintRow) = prdGood
End If

lintRow = lintRow + 1

'Mechanical Hysteresis Point 1
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).mechHystPt(1).Value, "##0.00") & "% of Fwd Force at " & Format(gudtTest(CHAN0).mechHystPt(1).location, "##0.00") & "° "
'2.2ANM If Not (gintFailure(CHAN0, HIGHMECHHYSTPT1) Or gintFailure(CHAN0, LOWMECHHYSTPT1) Or gintSevere(CHAN0, HIGHMECHHYSTPT1) Or gintSevere(CHAN0, LOWMECHHYSTPT1)) Then
'2.2ANM     lprdParameterResults(lintRow) = prdGood
'2.2ANM End If
lprdParameterResults(lintRow) = prdNotChecked

lintRow = lintRow + 1

'Mechanical Hysteresis Point 3
lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).mechHystPt(3).Value, "##0.00") & "% of Fwd Force at " & Format(gudtTest(CHAN0).mechHystPt(3).location, "##0.00") & "° "
'2.2ANM If Not (gintFailure(CHAN0, HIGHMECHHYSTPT3) Or gintFailure(CHAN0, LOWMECHHYSTPT3) Or gintSevere(CHAN0, HIGHMECHHYSTPT3) Or gintSevere(CHAN0, LOWMECHHYSTPT3)) Then
'2.2ANM     lprdParameterResults(lintRow) = prdGood
'2.2ANM End If
lprdParameterResults(lintRow) = prdNotChecked

'2.8dANM \/\/
'Display Kickdown Data if Appropriate
If gblnKD Then
    lintRow = lintRow + 1

    'Kickdown Label
    lprdParameterResults(lintRow) = prdEmpty

    lintRow = lintRow + 1

    'Kickdown Start Location
    lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).KDStart.location, "##0.00") & "° "
    If Not (gintFailure(CHAN0, HIGHKDSTART) Or gintFailure(CHAN0, LOWKDSTART) Or gintSevere(CHAN0, HIGHKDSTART) Or gintSevere(CHAN0, LOWKDSTART)) Then
        lprdParameterResults(lintRow) = prdGood
    End If

    lintRow = lintRow + 1

    'Kickdown Force Span
    lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).KDSpan, "##0.00") & " Newtons "
    If Not (gintFailure(CHAN0, HIGHKDSPAN) Or gintFailure(CHAN0, LOWKDSPAN) Or gintSevere(CHAN0, HIGHKDSPAN) Or gintSevere(CHAN0, LOWKDSPAN)) Then '2.8dANM
        lprdParameterResults(lintRow) = prdGood
    End If
        
    lintRow = lintRow + 1

    'Kickdown Peak Location
    lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).KDPeak.location, "##0.00") & "° "
    lprdParameterResults(lintRow) = prdNotChecked
    
    lintRow = lintRow + 1

    'Kickdown Peak Force
    lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).KDPeak.Value, "##0.00") & "N "
    lprdParameterResults(lintRow) = prdNotChecked

    lintRow = lintRow + 1

    'Kickdown End Location
    lstrValueAndLocation(lintRow) = Format(gudtReading(CHAN0).KDStop.location, "##0.00") & "° "
    If Not (gintFailure(CHAN0, HIGHKDSTOP) Or gintFailure(CHAN0, LOWKDSTOP) Or gintSevere(CHAN0, HIGHKDSTOP) Or gintSevere(CHAN0, LOWKDSTOP)) Then
        lprdParameterResults(lintRow) = prdGood
    End If
End If
'2.8dANM /\/\

'Send the results to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateResultsData(SCANRESULTSGRID, llngRowNum, lstrValueAndLocation(llngRowNum), lprdParameterResults(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayScanResultsNames()
'
'   PURPOSE: To display the results parameter names to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.0*ANM removed FF@FK

'Output #1
Call UpdateName(SCANRESULTSGRID, 1, "Pedal at Rest Location", False, flexAlignLeftCenter) '2.8ANM
Call UpdateName(SCANRESULTSGRID, 2, "Output #1", True, flexAlignCenterCenter)
Call UpdateName(SCANRESULTSGRID, 3, "Full-Close Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 4, "Full-Open Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 5, "Maximum Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 6, "SinglePoint Linearity Deviation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 7, "SinglePoint Linearity Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 8, "SinglePoint Linearity Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 9, "Absolute Linearity Deviation % Tol", False, flexAlignLeftCenter) '2.7ANM
Call UpdateName(SCANRESULTSGRID, 10, "Absolute Linearity Deviation Max", False, flexAlignLeftCenter)   '2.7ANM
Call UpdateName(SCANRESULTSGRID, 11, "Absolute Linearity Deviation Min", False, flexAlignLeftCenter)  '2.7ANM
Call UpdateName(SCANRESULTSGRID, 12, "Slope Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 13, "Slope Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 14, "Full-Close Hysteresis", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 15, "Supply Current at Idle", False, flexAlignLeftCenter) '2.5ANM '2.8cANM
Call UpdateName(SCANRESULTSGRID, 16, "Supply Current at WOT", False, flexAlignLeftCenter)  '2.8cANM
'Output #2
Call UpdateName(SCANRESULTSGRID, 17, "Output #2", True, flexAlignCenterCenter)
Call UpdateName(SCANRESULTSGRID, 18, "Full-Close Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 19, "Full-Open Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 20, "Maximum Output", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 21, "SinglePoint Linearity Deviation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 22, "SinglePoint Linearity Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 23, "SinglePoint Linearity Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 24, "Absolute Linearity Deviation % Tol", False, flexAlignLeftCenter) '2.7ANM
Call UpdateName(SCANRESULTSGRID, 25, "Absolute Linearity Deviation Max", False, flexAlignLeftCenter)   '2.7ANM
Call UpdateName(SCANRESULTSGRID, 26, "Absolute Linearity Deviation Min", False, flexAlignLeftCenter)   '2.7ANM
Call UpdateName(SCANRESULTSGRID, 27, "Slope Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 28, "Slope Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 29, "Full-Close Hysteresis", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 30, "Supply Current at Idle", False, flexAlignLeftCenter) '2.5ANM '2.8cANM
Call UpdateName(SCANRESULTSGRID, 31, "Supply Current at WOT", False, flexAlignLeftCenter)  '2.8cANM
'Correlation
Call UpdateName(SCANRESULTSGRID, 32, "Correlation", True, flexAlignCenterCenter)
Call UpdateName(SCANRESULTSGRID, 33, "Fwd Output Correlation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 34, "Fwd Output Correlation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 35, "Fwd Output Correlation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 36, "Rev Output Correlation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 37, "Rev Output Correlation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 38, "Rev Output Correlation Min", False, flexAlignLeftCenter)
'Force
Call UpdateName(SCANRESULTSGRID, 39, "Force", True, flexAlignCenterCenter)
Call UpdateName(SCANRESULTSGRID, 40, "Forward Force Point @ " & Format(gudtTest(CHAN0).fwdForcePt(1).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANRESULTSGRID, 41, "Forward Force Point @ " & Format(gudtTest(CHAN0).fwdForcePt(3).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANRESULTSGRID, 42, "Reverse Force Point @ " & Format(gudtTest(CHAN0).revForcePt(1).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANRESULTSGRID, 43, "Reverse Force Point @ " & Format(gudtTest(CHAN0).revForcePt(3).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANRESULTSGRID, 44, "Peak Force", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 45, "Mechanical Hysteresis Point @ " & Format(gudtTest(CHAN0).mechHystPt(1).location, "#0.00°"), False, flexAlignLeftCenter) '1.5ANM
Call UpdateName(SCANRESULTSGRID, 46, "Mechanical Hysteresis Point @ " & Format(gudtTest(CHAN0).mechHystPt(3).location, "#0.00°"), False, flexAlignLeftCenter) '1.5ANM
'Kickdown
Call UpdateName(SCANRESULTSGRID, 47, "Kickdown", True, flexAlignCenterCenter)               '2.8dANM \/\/
Call UpdateName(SCANRESULTSGRID, 48, "Kickdown Start Location", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 49, "Kickdown Force Span", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 50, "Kickdown Peak Location", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 51, "Kickdown Peak Force", False, flexAlignLeftCenter)
Call UpdateName(SCANRESULTSGRID, 52, "Kickdown End Location", False, flexAlignLeftCenter)   '2.8dANM /\/\

End Sub

Public Sub DisplayScanStatisticsCountsPrioritized()
'
'   PURPOSE: To display the failure counts to the screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim llngRowNum As Long
Dim lintRow As Long
Dim lvntHighCount(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntLowCount(1 To NUMROWSSCANSTATSDISPLAY) As Variant

'Rows:
'Pedal at Rest
'Output #1 Label (no counts)
'Index 1 (FullClose Output), output 1
'Index 3 (FullOpen Output), output 1
'Maximum Output, output 1
'SP Linearity Percentage of Tolerance, output 1
'Abs Linearity Percentage of Tolerance, output 1
'Slope Deviation Max, output 1
'Slope Deviation Min, output 1
'Full-Close Hysteresis, output 1
'MLX Current, output 1
'MLX WOT Current, output 1
'Output #2 Label (no counts)
'Index 1 (FullClose Output), output 2
'Index 3 (FullOpen Output), output 2
'Maximum Output, output 2
'SP Linearity Percentage of Tolerance, output 2
'Abs Linearity Percentage of Tolerance, output 2
'Slope Deviation Max, output 2
'Slope Deviation Min, output 2
'Full-Close Hysteresis, output 2
'MLX Current, output 1
'MLX WOT Current, output 2
'Correlation Label (no counts)
'Forward Output Correlation Percentage of Tolerance
'Reverse Output Correlation Percentage of Tolerance
'Force Label (no counts)
'Forward Force Point 1
'Forward Force Point 3
'Reverse Force Point 1
'Reverse Force Point 3
'Peak Force
'Mechanical Hysteresis Point 1
'Mechanical Hysteresis Point 3

lintRow = 1   'Initialize the Row Number

'Pedal at Rest Location '2.8ANM
lvntHighCount(lintRow) = gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low

lintRow = lintRow + 1

For lintChanNum = CHAN0 To MAXCHANNUM

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (FullClose Output)
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).Index(1).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).Index(1).failCount.low

    lintRow = lintRow + 1

    'Index 3 (FullOpen Output)
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).Index(3).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).Index(3).failCount.low

    lintRow = lintRow + 1

    'Maximum Output
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).maxOutput.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).maxOutput.failCount.low

    lintRow = lintRow + 1

    'SinglePoint Linearity Deviation % of Tolerance
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low

    lintRow = lintRow + 1

    '2.7ANM \/\/
    'Absolute Linearity Deviation % of Tolerance
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low

    lintRow = lintRow + 1
    '2.7ANM /\/\
    
    'High Slope Deviation
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).slopeMax.failCount.high
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Low Slope Deviation
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).slopeMin.failCount.low

    lintRow = lintRow + 1

    'Full-Close Hysteresis
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).FullCloseHys.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).FullCloseHys.failCount.low

    lintRow = lintRow + 1
    
    'MLX Current '2.5ANM
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).mlxCurrent.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).mlxCurrent.failCount.low
    
    lintRow = lintRow + 1
    
    'MLX WOT Current '2.8cANM
    lvntHighCount(lintRow) = gudtScanStats(lintChanNum).mlxWCurrent.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(lintChanNum).mlxWCurrent.failCount.low
    
    lintRow = lintRow + 1
    
Next lintChanNum

'Correlation Label
lintRow = lintRow + 1

'Forward Output Correlation % of Tolerance
lvntHighCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low

lintRow = lintRow + 1

'Reverse Output Correlation % of Tolerance
lvntHighCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low

lintRow = lintRow + 1

'Force Label
lintRow = lintRow + 1

'Force Knee Location
'2.3ANM lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).forceKneeLoc.failCount.high
'2.3ANM lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).forceKneeLoc.failCount.low

'2.3ANM lintRow = lintRow + 1

'2.0*ANM
''Force Knee Forward Force
'lvntHighCount(lintRow) = gudtScanStats(CHAN0).forceKneeForce.failCount.high
'lvntLowCount(lintRow) = gudtScanStats(CHAN0).forceKneeForce.failCount.low
'
'lintRow = lintRow + 1

'Forward Force Point 1
lvntHighCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(1).failCount.low

lintRow = lintRow + 1

'Forward Force Point 3
lvntHighCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(3).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).fwdForcePt(3).failCount.low

lintRow = lintRow + 1

'Reverse Force Point 1
lvntHighCount(lintRow) = gudtScanStats(CHAN0).revForcePt(1).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).revForcePt(1).failCount.low

lintRow = lintRow + 1

'Reverse Force Point 3
lvntHighCount(lintRow) = gudtScanStats(CHAN0).revForcePt(3).failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).revForcePt(3).failCount.low

lintRow = lintRow + 1

'Peak Force
lvntHighCount(lintRow) = gudtScanStats(CHAN0).peakForce.failCount.high
lvntLowCount(lintRow) = gudtScanStats(CHAN0).peakForce.failCount.low

lintRow = lintRow + 1

'Mechanical Hysteresis Point 1
lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(1).failCount.high
lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(1).failCount.low

lintRow = lintRow + 1

'Mechanical Hysteresis Point 3
lvntHighCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(3).failCount.high
lvntLowCount(lintRow) = "N/A" '2.2ANM gudtScanStats(CHAN0).mechHystPt(3).failCount.low

'2.8dANM \/\/
'Display Kickdown Counts if Appropriate
If gblnKD Then

    lintRow = lintRow + 1

    'Kickdown Label
    lintRow = lintRow + 1

    'Kickdown Start Location
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDStart.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDStart.failCount.low

    lintRow = lintRow + 1

    'Kickdown Force Span
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDSpan.failCount.high '2.8dANM
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDSpan.failCount.low   '2.8dANM

    lintRow = lintRow + 1
    
    'Kickdown Peak Location
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Kickdown Peak Force
    lvntHighCount(lintRow) = "N/A"
    lvntLowCount(lintRow) = "N/A"

    lintRow = lintRow + 1

    'Kickdown End Location
    lvntHighCount(lintRow) = gudtScanStats(CHAN0).KDStop.failCount.high
    lvntLowCount(lintRow) = gudtScanStats(CHAN0).KDStop.failCount.low
End If
'2.8dANM /\/\

'Send the counts to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateStatisticsCounts(SCANSTATSGRID, llngRowNum, lvntHighCount(llngRowNum), lvntLowCount(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayScanStatisticsData()
'
'   PURPOSE: To display the statistics to the screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim lintRow As Integer
Dim llngRowNum As Long
Dim lvntAvg(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntStdDev(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntCpk(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntCp(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntRangehigh(1 To NUMROWSSCANSTATSDISPLAY) As Variant
Dim lvntRangeLow(1 To NUMROWSSCANSTATSDISPLAY) As Variant

'Rows:
'Pedal at Rest
'Output #1 Label (no counts)
'Index 1 (FullClose Output), output 1
'Index 3 (FullOpen Output), output 1
'Maximum Output, output 1
'SP Linearity Percentage of Tolerance, output 1
'Abs Linearity Percentage of Tolerance, output 1
'Slope Deviation Max, output 1
'Slope Deviation Min, output 1
'Full-Close Hysteresis, output 1
'MLX Current, output 1
'MLX WOT Current, output 1
'Output #2 Label (no counts)
'Index 1 (FullClose Output), output 2
'Index 3 (FullOpen Output), output 2
'Maximum Output, output 2
'SP Linearity Percentage of Tolerance, output 2
'Abs Linearity Percentage of Tolerance, output 2
'Slope Deviation Max, output 2
'Slope Deviation Min, output 2
'Full-Close Hysteresis, output 2
'MLX Current, output 1
'MLX WOT Current, output 2
'Correlation Label (no counts)
'Forward Output Correlation Percentage of Tolerance
'Reverse Output Correlation Percentage of Tolerance
'Force Label (no counts)
'Forward Force Point 1
'Forward Force Point 3
'Reverse Force Point 1
'Reverse Force Point 3
'Peak Force
'Mechanical Hysteresis Point 1
'Mechanical Hysteresis Point 3

lintRow = 1

'Pedal at Rest Location '2.8ANM
If gudtScanStats(CHAN0).pedalAtRestLoc.n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).pedalAtRestLoc.sigma / gudtScanStats(CHAN0).pedalAtRestLoc.n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).pedalAtRestLoc.sigma2 - gudtScanStats(CHAN0).pedalAtRestLoc.sigma ^ 2 / gudtScanStats(CHAN0).pedalAtRestLoc.n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).pedalAtRestLoc.sigma2 - gudtScanStats(CHAN0).pedalAtRestLoc.sigma ^ 2 / gudtScanStats(CHAN0).pedalAtRestLoc.n) / (gudtScanStats(CHAN0).pedalAtRestLoc.n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).pedalAtRestLoc.max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).pedalAtRestLoc.min, "##0.00")
End If

lintRow = lintRow + 1

For lintChanNum = 0 To MAXCHANNUM

    'Output # Label
    lintRow = lintRow + 1

    'Index 1 (FullClose Output)
    If gudtScanStats(lintChanNum).Index(1).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).Index(1).sigma / gudtScanStats(lintChanNum).Index(1).n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).Index(1).sigma2 - gudtScanStats(lintChanNum).Index(1).sigma ^ 2 / gudtScanStats(lintChanNum).Index(1).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).Index(1).sigma2 - gudtScanStats(lintChanNum).Index(1).sigma ^ 2 / gudtScanStats(lintChanNum).Index(1).n) / (gudtScanStats(lintChanNum).Index(1).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If (gudtCustomerSpec(lintChanNum).Index(1).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(lintChanNum).Index(1).low) Then
                lvntCpk(lintRow) = Format((gudtCustomerSpec(lintChanNum).Index(1).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(lintChanNum).Index(1).low) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((gudtCustomerSpec(lintChanNum).Index(1).high - gudtCustomerSpec(lintChanNum).Index(1).low) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).Index(1).max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).Index(1).min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Index 3 (FullOpen Output)
    If gudtScanStats(lintChanNum).Index(3).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).Index(3).sigma / gudtScanStats(lintChanNum).Index(3).n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).Index(3).sigma2 - gudtScanStats(lintChanNum).Index(3).sigma ^ 2 / gudtScanStats(lintChanNum).Index(3).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).Index(3).sigma2 - gudtScanStats(lintChanNum).Index(3).sigma ^ 2 / gudtScanStats(lintChanNum).Index(3).n) / (gudtScanStats(lintChanNum).Index(3).n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Calculate Cpk and Cp if Std <> 0
        If lvntStdDev(lintRow) <> 0 Then
            If (gudtCustomerSpec(lintChanNum).Index(3).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(lintChanNum).Index(3).low) Then
                lvntCpk(lintRow) = Format((gudtCustomerSpec(lintChanNum).Index(3).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
            Else
                lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(lintChanNum).Index(3).low) / (3 * lvntStdDev(lintRow)), "##0.00")
            End If
            lvntCp(lintRow) = Format((gudtCustomerSpec(lintChanNum).Index(3).high - gudtCustomerSpec(lintChanNum).Index(3).low) / (6 * lvntStdDev(lintRow)), "###0.00")
        Else
            lvntCpk(lintRow) = "0.00"
            lvntCp(lintRow) = "0.00"
        End If
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).Index(3).max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).Index(3).min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Maximum Output
    If gudtScanStats(lintChanNum).maxOutput.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).maxOutput.sigma / gudtScanStats(lintChanNum).maxOutput.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).maxOutput.sigma2 - gudtScanStats(lintChanNum).maxOutput.sigma ^ 2 / gudtScanStats(lintChanNum).maxOutput.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).maxOutput.sigma2 - gudtScanStats(lintChanNum).maxOutput.sigma ^ 2 / gudtScanStats(lintChanNum).maxOutput.n) / (gudtScanStats(lintChanNum).maxOutput.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).maxOutput.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).maxOutput.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'SinglePoint Linearity Deviation % of Tolerance
    If gudtScanStats(lintChanNum).linDevPerTol(1).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(1).sigma / gudtScanStats(lintChanNum).linDevPerTol(1).n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).linDevPerTol(1).sigma2 - gudtScanStats(lintChanNum).linDevPerTol(1).sigma ^ 2 / gudtScanStats(lintChanNum).linDevPerTol(1).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).linDevPerTol(1).sigma2 - gudtScanStats(lintChanNum).linDevPerTol(1).sigma ^ 2 / gudtScanStats(lintChanNum).linDevPerTol(1).n) / (gudtScanStats(lintChanNum).linDevPerTol(1).n - 1)), "###0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(1).max, "###0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(1).min, "###0.00")
    End If

    lintRow = lintRow + 1

    '2.7ANM \/\/
    'Absolute Linearity Deviation % of Tolerance
    If gudtScanStats(lintChanNum).linDevPerTol(2).n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(2).sigma / gudtScanStats(lintChanNum).linDevPerTol(2).n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).linDevPerTol(2).sigma2 - gudtScanStats(lintChanNum).linDevPerTol(2).sigma ^ 2 / gudtScanStats(lintChanNum).linDevPerTol(2).n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).linDevPerTol(2).sigma2 - gudtScanStats(lintChanNum).linDevPerTol(2).sigma ^ 2 / gudtScanStats(lintChanNum).linDevPerTol(2).n) / (gudtScanStats(lintChanNum).linDevPerTol(2).n - 1)), "###0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(2).max, "###0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).linDevPerTol(2).min, "###0.00")
    End If

    lintRow = lintRow + 1
    '2.7ANM /\/\
    
    'Slope Deviation Max
    If gudtScanStats(lintChanNum).slopeMax.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).slopeMax.sigma / gudtScanStats(lintChanNum).slopeMax.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).slopeMax.sigma2 - gudtScanStats(lintChanNum).slopeMax.sigma ^ 2 / gudtScanStats(lintChanNum).slopeMax.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).slopeMax.sigma2 - gudtScanStats(lintChanNum).slopeMax.sigma ^ 2 / gudtScanStats(lintChanNum).slopeMax.n) / (gudtScanStats(lintChanNum).slopeMax.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).slopeMax.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).slopeMax.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Slope Deviation Min
    If gudtScanStats(lintChanNum).slopeMin.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).slopeMin.sigma / gudtScanStats(lintChanNum).slopeMin.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).slopeMin.sigma2 - gudtScanStats(lintChanNum).slopeMin.sigma ^ 2 / gudtScanStats(lintChanNum).slopeMin.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).slopeMin.sigma2 - gudtScanStats(lintChanNum).slopeMin.sigma ^ 2 / gudtScanStats(lintChanNum).slopeMin.n) / (gudtScanStats(lintChanNum).slopeMin.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).slopeMin.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).slopeMin.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Full-Close Hysteresis
    If gudtScanStats(lintChanNum).FullCloseHys.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).FullCloseHys.sigma / gudtScanStats(lintChanNum).FullCloseHys.n, "##0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).FullCloseHys.sigma2 - gudtScanStats(lintChanNum).FullCloseHys.sigma ^ 2 / gudtScanStats(lintChanNum).FullCloseHys.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).FullCloseHys.sigma2 - gudtScanStats(lintChanNum).FullCloseHys.sigma ^ 2 / gudtScanStats(lintChanNum).FullCloseHys.n) / (gudtScanStats(lintChanNum).FullCloseHys.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).FullCloseHys.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).FullCloseHys.min, "##0.00")
    End If

    lintRow = lintRow + 1
    
    'MLX Current '2.5ANM
    If gudtScanStats(lintChanNum).mlxCurrent.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).mlxCurrent.sigma / gudtScanStats(lintChanNum).mlxCurrent.n, "##0.0")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).mlxCurrent.sigma2 - gudtScanStats(lintChanNum).mlxCurrent.sigma ^ 2 / gudtScanStats(lintChanNum).mlxCurrent.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).mlxCurrent.sigma2 - gudtScanStats(lintChanNum).mlxCurrent.sigma ^ 2 / gudtScanStats(lintChanNum).mlxCurrent.n) / (gudtScanStats(lintChanNum).mlxCurrent.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).mlxCurrent.max, "##0.0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).mlxCurrent.min, "##0.0")
    End If

    lintRow = lintRow + 1
    
    'MLX WOT Current '2.8cANM
    If gudtScanStats(lintChanNum).mlxWCurrent.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(lintChanNum).mlxWCurrent.sigma / gudtScanStats(lintChanNum).mlxWCurrent.n, "##0.0")
        'Calculate Standard Deviation
        If (gudtScanStats(lintChanNum).mlxWCurrent.sigma2 - gudtScanStats(lintChanNum).mlxWCurrent.sigma ^ 2 / gudtScanStats(lintChanNum).mlxWCurrent.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(lintChanNum).mlxWCurrent.sigma2 - gudtScanStats(lintChanNum).mlxWCurrent.sigma ^ 2 / gudtScanStats(lintChanNum).mlxWCurrent.n) / (gudtScanStats(lintChanNum).mlxWCurrent.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(lintChanNum).mlxWCurrent.max, "##0.0")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(lintChanNum).mlxWCurrent.min, "##0.0")
    End If

    lintRow = lintRow + 1
    
Next lintChanNum

'Correlation Label
lintRow = lintRow + 1

'Forward Output Correlation % of Tolerance
If gudtScanStats(CHAN0).outputCorPerTol(1).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(1).sigma / gudtScanStats(CHAN0).outputCorPerTol(1).n, "###0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).outputCorPerTol(1).sigma2 - gudtScanStats(CHAN0).outputCorPerTol(1).sigma ^ 2 / gudtScanStats(CHAN0).outputCorPerTol(1).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).outputCorPerTol(1).sigma2 - gudtScanStats(CHAN0).outputCorPerTol(1).sigma ^ 2 / gudtScanStats(CHAN0).outputCorPerTol(1).n) / (gudtScanStats(CHAN0).outputCorPerTol(1).n - 1)), "###0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(1).max, "###0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(1).min, "###0.00")
End If

lintRow = lintRow + 1

'Reverse Output Correlation % of Tolerance
If gudtScanStats(CHAN0).outputCorPerTol(2).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(2).sigma / gudtScanStats(CHAN0).outputCorPerTol(2).n, "###0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).outputCorPerTol(2).sigma2 - gudtScanStats(CHAN0).outputCorPerTol(2).sigma ^ 2 / gudtScanStats(CHAN0).outputCorPerTol(2).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).outputCorPerTol(2).sigma2 - gudtScanStats(CHAN0).outputCorPerTol(2).sigma ^ 2 / gudtScanStats(CHAN0).outputCorPerTol(2).n) / (gudtScanStats(CHAN0).outputCorPerTol(2).n - 1)), "###0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(2).max, "###0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).outputCorPerTol(2).min, "###0.00")
End If

lintRow = lintRow + 1

'Force Label
lintRow = lintRow + 1

'2.3ANM 'Force Knee Location
'2.3ANM If gudtScanStats(CHAN0).forceKneeLoc.n > 1 Then
'2.3ANM     'Calculate Average
'2.3ANM     lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).forceKneeLoc.sigma / gudtScanStats(CHAN0).forceKneeLoc.n, "##0.00")
'2.3ANM     'Calculate Standard Deviation
'2.3ANM     If (gudtScanStats(CHAN0).forceKneeLoc.sigma2 - gudtScanStats(CHAN0).forceKneeLoc.sigma ^ 2 / gudtScanStats(CHAN0).forceKneeLoc.n) > 0 Then
'2.3ANM         lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).forceKneeLoc.sigma2 - gudtScanStats(CHAN0).forceKneeLoc.sigma ^ 2 / gudtScanStats(CHAN0).forceKneeLoc.n) / (gudtScanStats(CHAN0).forceKneeLoc.n - 1)), "##0.000")
'2.3ANM     Else
'2.3ANM         lvntStdDev(lintRow) = 0
'2.3ANM     End If
'2.3ANM     'Calculate Cpk and Cp if Std <> 0
'2.2ANM     If lvntStdDev(lintRow) <> 0 Then
'2.2ANM         If (gudtCustomerSpec(CHAN0).forceKneeLoc.high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).forceKneeLoc.low) Then
'2.2ANM             lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).forceKneeLoc.high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
'2.2ANM         Else
'2.2ANM             lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).forceKneeLoc.low) / (3 * lvntStdDev(lintRow)), "##0.00")
'2.2ANM         End If
'2.2ANM         lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).forceKneeLoc.high - gudtCustomerSpec(CHAN0).forceKneeLoc.low) / (6 * lvntStdDev(lintRow)), "##0.00")
'2.2ANM     Else
'2.3ANM         lvntCpk(lintRow) = ""
'2.3ANM         lvntCp(lintRow) = ""
'2.2ANM     End If
'2.3ANM     'Range High
'2.3ANM     lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).forceKneeLoc.max, "##0.00")
'2.3ANM     'Range Low
'2.3ANM     lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).forceKneeLoc.min, "##0.00")
'2.3ANM End If
'2.3ANM
'2.3ANM lintRow = lintRow + 1

'2.0*ANM
''Forward Force at Force Knee Location
'If gudtScanStats(CHAN0).forceKneeForce.n > 1 Then
'    'Calculate Average
'    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).forceKneeForce.sigma / gudtScanStats(CHAN0).forceKneeForce.n, "##0.00")
'    'Calculate Standard Deviation
'    If (gudtScanStats(CHAN0).forceKneeForce.sigma2 - gudtScanStats(CHAN0).forceKneeForce.sigma ^ 2 / gudtScanStats(CHAN0).forceKneeForce.n) > 0 Then
'        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).forceKneeForce.sigma2 - gudtScanStats(CHAN0).forceKneeForce.sigma ^ 2 / gudtScanStats(CHAN0).forceKneeForce.n) / (gudtScanStats(CHAN0).forceKneeForce.n - 1)), "##0.000")
'    Else
'        lvntStdDev(lintRow) = "0.000"
'    End If
'    'Calculate Cpk and Cp if Std <> 0
'    If lvntStdDev(lintRow) <> 0 Then
'        If (gudtCustomerSpec(CHAN0).forceKneeForce.high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).forceKneeForce.low) Then
'            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).forceKneeForce.high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
'        Else
'            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).forceKneeForce.low) / (3 * lvntStdDev(lintRow)), "##0.00")
'        End If
'        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).forceKneeForce.high - gudtCustomerSpec(CHAN0).forceKneeForce.low) / (6 * lvntStdDev(lintRow)), "##0.00")
'    Else
'        lvntCpk(lintRow) = "0.00"
'        lvntCp(lintRow) = "0.00"
'    End If
'    'Range High
'    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).forceKneeForce.max, "##0.00")
'    'Range Low
'    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).forceKneeForce.min, "##0.00")
'End If
'
'lintRow = lintRow + 1

'Forward Force Point 1
If gudtScanStats(CHAN0).fwdForcePt(1).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(1).sigma / gudtScanStats(CHAN0).fwdForcePt(1).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).fwdForcePt(1).sigma2 - gudtScanStats(CHAN0).fwdForcePt(1).sigma ^ 2 / gudtScanStats(CHAN0).fwdForcePt(1).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).fwdForcePt(1).sigma2 - gudtScanStats(CHAN0).fwdForcePt(1).sigma ^ 2 / gudtScanStats(CHAN0).fwdForcePt(1).n) / (gudtScanStats(CHAN0).fwdForcePt(1).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
    If lvntStdDev(lintRow) <> 0 Then
        If (gudtCustomerSpec(CHAN0).fwdForcePt(1).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).fwdForcePt(1).low) Then
            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).fwdForcePt(1).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
        Else
            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).fwdForcePt(1).low) / (3 * lvntStdDev(lintRow)), "##0.00")
        End If
        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).fwdForcePt(1).high - gudtCustomerSpec(CHAN0).fwdForcePt(1).low) / (6 * lvntStdDev(lintRow)), "##0.00")
    Else
        lvntCpk(lintRow) = "0.00"
        lvntCp(lintRow) = "0.00"
    End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(1).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(1).min, "##0.00")
End If

lintRow = lintRow + 1

'Forward Force Point 3
If gudtScanStats(CHAN0).fwdForcePt(3).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(3).sigma / gudtScanStats(CHAN0).fwdForcePt(3).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).fwdForcePt(3).sigma2 - gudtScanStats(CHAN0).fwdForcePt(3).sigma ^ 2 / gudtScanStats(CHAN0).fwdForcePt(3).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).fwdForcePt(3).sigma2 - gudtScanStats(CHAN0).fwdForcePt(3).sigma ^ 2 / gudtScanStats(CHAN0).fwdForcePt(3).n) / (gudtScanStats(CHAN0).fwdForcePt(3).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
    If lvntStdDev(lintRow) <> 0 Then
        If (gudtCustomerSpec(CHAN0).fwdForcePt(3).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).fwdForcePt(3).low) Then
            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).fwdForcePt(3).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
        Else
            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).fwdForcePt(3).low) / (3 * lvntStdDev(lintRow)), "##0.00")
        End If
        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).fwdForcePt(3).high - gudtCustomerSpec(CHAN0).fwdForcePt(3).low) / (6 * lvntStdDev(lintRow)), "##0.00")
    Else
        lvntCpk(lintRow) = "0.00"
        lvntCp(lintRow) = "0.00"
    End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(3).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).fwdForcePt(3).min, "##0.00")
End If

lintRow = lintRow + 1

'Reverse Force Point 1
If gudtScanStats(CHAN0).revForcePt(1).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(1).sigma / gudtScanStats(CHAN0).revForcePt(1).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).revForcePt(1).sigma2 - gudtScanStats(CHAN0).revForcePt(1).sigma ^ 2 / gudtScanStats(CHAN0).revForcePt(1).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).revForcePt(1).sigma2 - gudtScanStats(CHAN0).revForcePt(1).sigma ^ 2 / gudtScanStats(CHAN0).revForcePt(1).n) / (gudtScanStats(CHAN0).revForcePt(1).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
    If lvntStdDev(lintRow) <> 0 Then
        If (gudtCustomerSpec(CHAN0).revForcePt(1).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).revForcePt(1).low) Then
            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).revForcePt(1).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
        Else
            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).revForcePt(1).low) / (3 * lvntStdDev(lintRow)), "##0.00")
        End If
        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).revForcePt(1).high - gudtCustomerSpec(CHAN0).revForcePt(1).low) / (6 * lvntStdDev(lintRow)), "##0.00")
    Else
        lvntCpk(lintRow) = "0.00"
        lvntCp(lintRow) = "0.00"
    End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(1).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(1).min, "##0.00")
End If

lintRow = lintRow + 1

'Reverse Force Point 3
If gudtScanStats(CHAN0).revForcePt(3).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(3).sigma / gudtScanStats(CHAN0).revForcePt(3).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).revForcePt(3).sigma2 - gudtScanStats(CHAN0).revForcePt(3).sigma ^ 2 / gudtScanStats(CHAN0).revForcePt(3).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).revForcePt(3).sigma2 - gudtScanStats(CHAN0).revForcePt(3).sigma ^ 2 / gudtScanStats(CHAN0).revForcePt(3).n) / (gudtScanStats(CHAN0).revForcePt(3).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
    If lvntStdDev(lintRow) <> 0 Then
        If (gudtCustomerSpec(CHAN0).revForcePt(3).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).revForcePt(3).low) Then
            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).revForcePt(3).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
        Else
            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).revForcePt(3).low) / (3 * lvntStdDev(lintRow)), "##0.00")
        End If
        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).revForcePt(3).high - gudtCustomerSpec(CHAN0).revForcePt(3).low) / (6 * lvntStdDev(lintRow)), "##0.00")
    Else
        lvntCpk(lintRow) = "0.00"
        lvntCp(lintRow) = "0.00"
    End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(3).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).revForcePt(3).min, "##0.00")
End If

lintRow = lintRow + 1

'Peak Force
If gudtScanStats(CHAN0).peakForce.n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).peakForce.sigma / gudtScanStats(CHAN0).peakForce.n, "###0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).peakForce.sigma2 - gudtScanStats(CHAN0).peakForce.sigma ^ 2 / gudtScanStats(CHAN0).peakForce.n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).peakForce.sigma2 - gudtScanStats(CHAN0).peakForce.sigma ^ 2 / gudtScanStats(CHAN0).peakForce.n) / (gudtScanStats(CHAN0).peakForce.n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Cpk & CP are N/A
    lvntCpk(lintRow) = "N/A"
    lvntCp(lintRow) = "N/A"
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).peakForce.max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).peakForce.min, "##0.00")
End If

lintRow = lintRow + 1

'Mechanical Hysteresis Point 1
If gudtScanStats(CHAN0).mechHystPt(1).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(1).sigma / gudtScanStats(CHAN0).mechHystPt(1).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).mechHystPt(1).sigma2 - gudtScanStats(CHAN0).mechHystPt(1).sigma ^ 2 / gudtScanStats(CHAN0).mechHystPt(1).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).mechHystPt(1).sigma2 - gudtScanStats(CHAN0).mechHystPt(1).sigma ^ 2 / gudtScanStats(CHAN0).mechHystPt(1).n) / (gudtScanStats(CHAN0).mechHystPt(1).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
'    If lvntStdDev(lintRow) <> 0 Then
'        If (gudtCustomerSpec(CHAN0).mechHystPt(1).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).mechHystPt(1).low) Then
'            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).mechHystPt(1).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
'        Else
'            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).mechHystPt(1).low) / (3 * lvntStdDev(lintRow)), "##0.00")
'        End If
'        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).mechHystPt(1).high - gudtCustomerSpec(CHAN0).mechHystPt(1).low) / (6 * lvntStdDev(lintRow)), "##0.00")
'    Else
        lvntCpk(lintRow) = ""
        lvntCp(lintRow) = ""
    '2.2ANM End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(1).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(1).min, "##0.00")
End If

lintRow = lintRow + 1

'Mechanical Hysteresis Point 3
If gudtScanStats(CHAN0).mechHystPt(3).n > 1 Then
    'Calculate Average
    lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(3).sigma / gudtScanStats(CHAN0).mechHystPt(3).n, "##0.00")
    'Calculate Standard Deviation
    If (gudtScanStats(CHAN0).mechHystPt(3).sigma2 - gudtScanStats(CHAN0).mechHystPt(3).sigma ^ 2 / gudtScanStats(CHAN0).mechHystPt(3).n) > 0 Then
        lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).mechHystPt(3).sigma2 - gudtScanStats(CHAN0).mechHystPt(3).sigma ^ 2 / gudtScanStats(CHAN0).mechHystPt(3).n) / (gudtScanStats(CHAN0).mechHystPt(3).n - 1)), "##0.000")
    Else
        lvntStdDev(lintRow) = "0.000"
    End If
    'Calculate Cpk and Cp if Std <> 0
'    If lvntStdDev(lintRow) <> 0 Then
'        If (gudtCustomerSpec(CHAN0).mechHystPt(3).high - lvntAvg(lintRow)) < (lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).mechHystPt(3).low) Then
'            lvntCpk(lintRow) = Format((gudtCustomerSpec(CHAN0).mechHystPt(3).high - lvntAvg(lintRow)) / (3 * lvntStdDev(lintRow)), "##0.00")
'        Else
'            lvntCpk(lintRow) = Format((lvntAvg(lintRow) - gudtCustomerSpec(CHAN0).mechHystPt(3).low) / (3 * lvntStdDev(lintRow)), "##0.00")
'        End If
'        lvntCp(lintRow) = Format((gudtCustomerSpec(CHAN0).mechHystPt(3).high - gudtCustomerSpec(CHAN0).mechHystPt(3).low) / (6 * lvntStdDev(lintRow)), "##0.00")
'    Else
        lvntCpk(lintRow) = ""
        lvntCp(lintRow) = ""
    '2.2ANM End If
    'Range High
    lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(3).max, "##0.00")
    'Range Low
    lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).mechHystPt(3).min, "##0.00")
End If

'2.8dANM \/\/
'Display Kickdown Stats if Appropriate
If gblnKD Then
    lintRow = lintRow + 1

    'Kickdown Label
    lintRow = lintRow + 1

    'Kickdown Start Location
    If gudtScanStats(CHAN0).KDStart.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).KDStart.sigma / gudtScanStats(CHAN0).KDStart.n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(CHAN0).KDStart.sigma2 - gudtScanStats(CHAN0).KDStart.sigma ^ 2 / gudtScanStats(CHAN0).KDStart.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).KDStart.sigma2 - gudtScanStats(CHAN0).KDStart.sigma ^ 2 / gudtScanStats(CHAN0).KDStart.n) / (gudtScanStats(CHAN0).KDStart.n - 1)), "###0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).KDStart.max, "###0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).KDStart.min, "###0.00")
    End If

    lintRow = lintRow + 1

    'Kickdown Force Span
    If gudtScanStats(CHAN0).KDSpan.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).KDSpan.sigma / gudtScanStats(CHAN0).KDSpan.n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(CHAN0).KDSpan.sigma2 - gudtScanStats(CHAN0).KDSpan.sigma ^ 2 / gudtScanStats(CHAN0).KDSpan.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).KDSpan.sigma2 - gudtScanStats(CHAN0).KDSpan.sigma ^ 2 / gudtScanStats(CHAN0).KDSpan.n) / (gudtScanStats(CHAN0).KDSpan.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).KDSpan.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).KDSpan.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Kickdown Peak Location
    If gudtScanStats(CHAN0).KDPeak.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).KDPeak.sigma / gudtScanStats(CHAN0).KDPeak.n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(CHAN0).KDPeak.sigma2 - gudtScanStats(CHAN0).KDPeak.sigma ^ 2 / gudtScanStats(CHAN0).KDPeak.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).KDPeak.sigma2 - gudtScanStats(CHAN0).KDPeak.sigma ^ 2 / gudtScanStats(CHAN0).KDPeak.n) / (gudtScanStats(CHAN0).KDPeak.n - 1)), "###0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).KDPeak.max, "###0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).KDPeak.min, "###0.00")
    End If

    lintRow = lintRow + 1

    'Kickdown Peak Force
    If gudtScanStats(CHAN0).KDPeakForce.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).KDPeakForce.sigma / gudtScanStats(CHAN0).KDPeakForce.n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(CHAN0).KDPeakForce.sigma2 - gudtScanStats(CHAN0).KDPeakForce.sigma ^ 2 / gudtScanStats(CHAN0).KDPeakForce.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).KDPeakForce.sigma2 - gudtScanStats(CHAN0).KDPeakForce.sigma ^ 2 / gudtScanStats(CHAN0).KDPeakForce.n) / (gudtScanStats(CHAN0).KDPeakForce.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).KDPeakForce.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).KDPeakForce.min, "##0.00")
    End If

    lintRow = lintRow + 1

    'Kickdown End Location
    If gudtScanStats(CHAN0).KDStop.n > 1 Then
        'Calculate Average
        lvntAvg(lintRow) = Format(gudtScanStats(CHAN0).KDStop.sigma / gudtScanStats(CHAN0).KDStop.n, "###0.00")
        'Calculate Standard Deviation
        If (gudtScanStats(CHAN0).KDStop.sigma2 - gudtScanStats(CHAN0).KDStop.sigma ^ 2 / gudtScanStats(CHAN0).KDStop.n) > 0 Then
            lvntStdDev(lintRow) = Format(Sqr((gudtScanStats(CHAN0).KDStop.sigma2 - gudtScanStats(CHAN0).KDStop.sigma ^ 2 / gudtScanStats(CHAN0).KDStop.n) / (gudtScanStats(CHAN0).KDStop.n - 1)), "##0.000")
        Else
            lvntStdDev(lintRow) = "0.000"
        End If
        'Cpk & CP are N/A
        lvntCpk(lintRow) = "N/A"
        lvntCp(lintRow) = "N/A"
        'Range High
        lvntRangehigh(lintRow) = Format(gudtScanStats(CHAN0).KDStop.max, "##0.00")
        'Range Low
        lvntRangeLow(lintRow) = Format(gudtScanStats(CHAN0).KDStop.min, "##0.00")
    End If
End If
'2.8dANM /\/\

'Send the stats to the control (start at row #1)
For llngRowNum = 1 To lintRow
    Call UpdateStatisticsData(SCANSTATSGRID, llngRowNum, lvntAvg(llngRowNum), lvntStdDev(llngRowNum), lvntCpk(llngRowNum), lvntCp(llngRowNum), lvntRangehigh(llngRowNum), lvntRangeLow(llngRowNum))
Next llngRowNum

End Sub

Public Sub DisplayScanStatisticsNames()
'
'   PURPOSE: To display the statistics parameter names to the screen
'
'  INPUT(S): none
' OUTPUT(S): none
'2.0*ANM removed FF@FK

'Output #1
Call UpdateName(SCANSTATSGRID, 1, "Pedal at Rest Location", False, flexAlignLeftCenter) '2.8ANM
Call UpdateName(SCANSTATSGRID, 2, "Output #1", True, flexAlignCenterCenter)
Call UpdateName(SCANSTATSGRID, 3, "Full-Close Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 4, "Full-Open Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 5, "Maximum Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 6, "SinglePoint Linearity Deviation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 7, "Absolute Linearity Deviation % Tol", False, flexAlignLeftCenter) '2.7ANM
Call UpdateName(SCANSTATSGRID, 8, "Slope Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 9, "Slope Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 10, "Full-Close Hysteresis", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 11, "Supply Current at Idle", False, flexAlignLeftCenter) '2.5ANM '2.8cANM
Call UpdateName(SCANSTATSGRID, 12, "Supply Current at WOT", False, flexAlignLeftCenter)  '2.8cANM
'Output #2
Call UpdateName(SCANSTATSGRID, 13, "Output #2", True, flexAlignCenterCenter)
Call UpdateName(SCANSTATSGRID, 14, "Full-Close Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 15, "Full-Open Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 16, "Maximum Output", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 17, "SinglePoint Linearity Deviation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 18, "Absolute Linearity Deviation % Tol", False, flexAlignLeftCenter) '2.7ANM
Call UpdateName(SCANSTATSGRID, 19, "Slope Deviation Max", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 20, "Slope Deviation Min", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 21, "Full-Close Hysteresis", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 22, "Supply Current at Idle", False, flexAlignLeftCenter) '2.5ANM '2.8cANM
Call UpdateName(SCANSTATSGRID, 23, "Supply Current at WOT", False, flexAlignLeftCenter)  '2.8cANM
'Correlation
Call UpdateName(SCANSTATSGRID, 24, "Correlation", True, flexAlignCenterCenter)
Call UpdateName(SCANSTATSGRID, 25, "Forward Output Correlation % Tol", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 26, "Reverse Output Correlation % Tol", False, flexAlignLeftCenter)
'Force
Call UpdateName(SCANSTATSGRID, 27, "Force", True, flexAlignCenterCenter)
Call UpdateName(SCANSTATSGRID, 28, "Forward Force Point @ " & Format(gudtTest(CHAN0).fwdForcePt(1).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANSTATSGRID, 29, "Forward Force Point @ " & Format(gudtTest(CHAN0).fwdForcePt(3).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANSTATSGRID, 30, "Reverse Force Point @ " & Format(gudtTest(CHAN0).revForcePt(1).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANSTATSGRID, 31, "Reverse Force Point @ " & Format(gudtTest(CHAN0).revForcePt(3).location, "#0.00°"), False, flexAlignLeftCenter)  '1.5ANM
Call UpdateName(SCANSTATSGRID, 32, "Peak Force", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 33, "Mechanical Hysteresis Point @ " & Format(gudtTest(CHAN0).mechHystPt(1).location, "#0.00°"), False, flexAlignLeftCenter) '1.5ANM
Call UpdateName(SCANSTATSGRID, 34, "Mechanical Hysteresis Point @ " & Format(gudtTest(CHAN0).mechHystPt(3).location, "#0.00°"), False, flexAlignLeftCenter) '1.5ANM
'Kickdown
Call UpdateName(SCANSTATSGRID, 35, "Kickdown", True, flexAlignCenterCenter)               '2.8dANM \/\/
Call UpdateName(SCANSTATSGRID, 36, "Kickdown Start Location", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 37, "Kickdown Force Span", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 38, "Kickdown Peak Location", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 39, "Kickdown Peak Force", False, flexAlignLeftCenter)
Call UpdateName(SCANSTATSGRID, 40, "Kickdown End Location", False, flexAlignLeftCenter)   '2.8dANM /\/\

End Sub

Public Sub DisplayScanSummary()
'
'   PURPOSE: To display the Scanning Lot Summary data to a screen
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lsngCurrentYield As Single
Dim lsngLotYield As Single

'Calculate the Current xxx part Yield
If gudtScanSummary.currentTotal <> 0 Then
    lsngCurrentYield = gudtScanSummary.currentGood / gudtScanSummary.currentTotal
Else
    lsngCurrentYield = 0
End If

'Calculate the Lot Yield
If gudtScanSummary.totalUnits <> 0 Then
    lsngLotYield = gudtScanSummary.totalGood / gudtScanSummary.totalUnits
Else
    lsngLotYield = 0
End If
    
'Display the number of Current parts in the Current Yield Label
frmMain.ctrScanSummary.LabelCaption(SummaryTextBox.stbCurrentYield) = Format(gudtScanSummary.currentTotal, "0") & " Yield"
 
'Display the values in the Scan Summary Boxes
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbTotalUnits) = gudtScanSummary.totalUnits
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbGoodUnits) = gudtScanSummary.totalGood
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbRejectedUnits) = gudtScanSummary.totalUnits - gudtScanSummary.totalGood
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbSevereUnits) = gudtScanSummary.totalSevere
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbSystemErrors) = gudtScanSummary.totalNoTest
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbCurrentYield) = Format(lsngCurrentYield, "0.00%")
frmMain.ctrScanSummary.TextBoxText(SummaryTextBox.stbLotYield) = Format(lsngLotYield, "0.00%")

'Set the appropriate Background color for Current xxx Part Yield
If gudtScanSummary.currentTotal = 0 Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbWhite
ElseIf lsngCurrentYield * HUNDREDPERCENT >= gudtMachine.yieldGreen Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbGreen
ElseIf lsngCurrentYield * HUNDREDPERCENT >= gudtMachine.yieldYellow Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbYellow
Else
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbCurrentYield) = vbRed
End If

'Set the appropriate Background color for Lot Yield
If gudtScanSummary.totalUnits = 0 Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbWhite
ElseIf lsngLotYield * HUNDREDPERCENT >= gudtMachine.yieldGreen Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbGreen
ElseIf lsngLotYield * HUNDREDPERCENT >= gudtMachine.yieldYellow Then
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbYellow
Else
    frmMain.ctrScanSummary.TextBackgroundColor(SummaryTextBox.stbLotYield) = vbRed
End If

End Sub

Public Sub EvaluateLimits(chanNum)
'
'   PURPOSE: Create limit arrays based on specified parameters
'
'  INPUT(S): ChanNum = channel number for evaluation
' OUTPUT(S): none

Dim lintRegionNum As Integer                'Region Number
Dim lsngHighLimitSlope As Single            'High Limit Line Slope
Dim lsngLowLimitSlope As Single             'Low Limit Line Slope
Dim lsngHighLimitYIntercept As Single       'High Limit Line Y-Intercept
Dim lsngLowLimitYIntercept As Single        'Low Limit Line Y-Intercept
Dim lsngStartHigh As Single                 'Voltage Gradient High Limit at Region Start
Dim lsngStartLow As Single                  'Voltage Gradient Low Limit at Region Start
Dim lsngStopHigh As Single                  'Voltage Gradient High Limit at Region Stop
Dim lsngStopLow As Single                   'Voltage Gradient Low Limit at Region Stop
Dim lsngSlopeToUse As Single                '1.1ANM slope to use in VG limit check based on position
Dim lsngOffset As Single                    '1.1ANM offset to use for VG limit
Dim lintRegionCount As Integer              'Count number             '2.2ANM
Dim lstrParameterName As String             'Parameter or Metric Name '2.2ANM
Dim lintChanNum As Integer                  'Channel Number           '2.2ANM

'Re-Dimension the Limit Arrays
ReDim msngMaxVoltageGradientLimit(gintMaxData)
ReDim msngMinVoltageGradientLimit(gintMaxData)
ReDim msngMaxLinearityLimit(gintMaxData)
ReDim msngMinLinearityLimit(gintMaxData)
ReDim msngMaxAbsLinearityLimit(gintMaxData) '2.7ANM
ReDim msngMinAbsLinearityLimit(gintMaxData) '2.7ANM
ReDim msngMaxSlopeDeviationLimit((gudtTest(chanNum).slope.start * gsngResolution) / gudtMachine.slopeIncrement To (gudtTest(chanNum).slope.stop * gsngResolution) / gudtMachine.slopeIncrement)
ReDim msngMinSlopeDeviationLimit((gudtTest(chanNum).slope.start * gsngResolution) / gudtMachine.slopeIncrement To (gudtTest(chanNum).slope.stop * gsngResolution) / gudtMachine.slopeIncrement)
ReDim msngMaxHysteresisLimit(gintMaxData)
ReDim msngMinHysteresisLimit(gintMaxData)

'Calculate SinglePoint Linearity Deviation Limit Arrays
'2.2ANM \/\/
If gblnUseNewAmad Then
    Dim lintOutput As Integer
    If lintChanNum = 0 Then
        lintOutput = 1
    Else
        lintOutput = 2
    End If
    lstrParameterName = "SingLinDevVal"
    lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, lintOutput, MPCTYPE_IDEAL)
Else
    lintRegionCount = 5
End If
'2.2ANM /\/\

For lintRegionNum = 1 To lintRegionCount
    'Calculate SinglePoint Linearity Limit Line Slopes & Intercepts
    Call CalcLimitLineMandB(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).start.high, gudtTest(chanNum).SinglePointLin(lintRegionNum).start.low, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.high, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.low, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
    'High SinglePoint Linearity Limit Array
    Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxLinearityLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
    'Low SinglePoint Linearity Limit Array
    Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinLinearityLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
Next lintRegionNum

'2.7ANM \/\/
'Calculate Absolute Linearity Deviation Limit Arrays
If gblnUseNewAmad Then
    If lintChanNum = 0 Then
        lintOutput = 1
    Else
        lintOutput = 2
    End If
    lstrParameterName = "AbsLinDevVal"
    lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, lintOutput, MPCTYPE_IDEAL)
Else
    lintRegionCount = 5
End If

For lintRegionNum = 1 To lintRegionCount
    'Calculate Absolute Linearity Limit Line Slopes & Intercepts
    Call CalcLimitLineMandB(gudtTest(chanNum).AbsLin(lintRegionNum).start.location, gudtTest(chanNum).AbsLin(lintRegionNum).start.high, gudtTest(chanNum).AbsLin(lintRegionNum).start.low, gudtTest(chanNum).AbsLin(lintRegionNum).stop.location, gudtTest(chanNum).AbsLin(lintRegionNum).stop.high, gudtTest(chanNum).AbsLin(lintRegionNum).stop.low, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
    'High Absolute Linearity Limit Array
    Call Calc.calcLimitArray(gudtTest(chanNum).AbsLin(lintRegionNum).start.location, gudtTest(chanNum).AbsLin(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxAbsLinearityLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
    'Low Absolute Linearity Limit Array
    Call Calc.calcLimitArray(gudtTest(chanNum).AbsLin(lintRegionNum).start.location, gudtTest(chanNum).AbsLin(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinAbsLinearityLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
Next lintRegionNum
'2.7ANM /\/\

'1.7ANM \/\/ fix VG limits to use actual not ideal index and correct limits
Dim i As Integer
ReDim gsngIdeal((gudtTest(chanNum).SinglePointLin(1).start.location * gsngResolution) To (gudtTest(chanNum).SinglePointLin(5).stop.location * gsngResolution)) As Single
For i = (gudtTest(chanNum).SinglePointLin(1).start.location * gsngResolution) To (gudtTest(chanNum).SinglePointLin(5).stop.location * gsngResolution)
    If (i / gsngResolution) <= gudtTest(chanNum).slope.start Then
        gsngIdeal(i) = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(1).start.location - gudtTest(CHAN0).Index(1).location) * gudtTest(chanNum).slope.ideal) + ((i / gsngResolution) * gudtTest(chanNum).slope.ideal)
    Else
        gsngIdeal(i) = (gsngIdeal(gudtTest(chanNum).slope.start * gsngResolution) - (gudtTest(chanNum).slope.start * gudtTest(chanNum).slope.ideal2)) + ((i / gsngResolution) * gudtTest(chanNum).slope.ideal2)
    End If
    msngMaxVoltageGradientLimit(i) = gsngIdeal(i) + msngMaxLinearityLimit(i)
    msngMinVoltageGradientLimit(i) = gsngIdeal(i) + msngMinLinearityLimit(i)
Next i
'1.7ANM /\/\

'1.7ANM '1.1ANM \/\/ Calculate Voltage Gradient Limit Arrays
'1.7ANM 'Set variables for bend
'1.7ANM lsngSlopeToUse = gudtTest(chanNum).slope.ideal
'1.7ANM lsngOffset = 0
'1.7ANM
'1.7ANM For lintRegionNum = 1 To 5
'1.7ANM     'Calculate Voltage Gradient Limit Line Start & Stop High & Low Locations
'1.7ANM     'NOTE: The high and low values at the start & stop of each region are calculated
'1.7ANM     '      as follows: Y-Value = A + B + C + D, where
'1.7ANM     '      A = ideal value at a known location (i.e. an index, in this case the ideal output at the FullClose location)
'1.7ANM     '      B = the Y-axis span from the known index value to the ideal value at the region start or stop
'1.7ANM     '      C = the high/low limit tolerance
'1.7ANM     '      D = offset due to the bend
'1.7ANM     '      B is calculated by multiplying the X-axis difference between the region
'1.7ANM     '      start/stop and the known index location by the ideal slope of the output
'1.7ANM     If gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location > gudtTest(chanNum).slope.start Then
'1.7ANM         lsngStartHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.high + lsngOffset
'1.7ANM         lsngStartLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.low + lsngOffset
'1.7ANM         lsngStopHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.high + lsngOffset
'1.7ANM         lsngStopLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.low + lsngOffset
'1.7ANM         'Calculate Voltage Gradient Limit Line Slopes & Intercepts
'1.7ANM         Call CalcLimitLineMandB(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, lsngStartHigh, lsngStartLow, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngStopHigh, lsngStopLow, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
'1.7ANM         'High Voltage Gradient Limit Array
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM         'Low Voltage Gradient Limit Array
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM     Else
'1.7ANM         lsngStartHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.high + lsngOffset
'1.7ANM         lsngStartLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.low + lsngOffset
'1.7ANM         lsngStopHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).slope.start - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.high + lsngOffset
'1.7ANM         lsngStopLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).slope.start - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.low + lsngOffset
'1.7ANM         'Calculate Voltage Gradient Limit Line Slopes & Intercepts to the bend
'1.7ANM         Call CalcLimitLineMandB(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, lsngStartHigh, lsngStartLow, gudtTest(CHAN0).slope.start, lsngStopHigh, lsngStopLow, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
'1.7ANM         'High Voltage Gradient Limit Array to the bend
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(CHAN0).slope.start, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM         'Low Voltage Gradient Limit Array to the bend
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(lintRegionNum).start.location, gudtTest(CHAN0).slope.start, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM         lsngStartHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).slope.start + -gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.high + lsngOffset
'1.7ANM         lsngStartLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).slope.start - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).start.low + lsngOffset
'1.7ANM         'Reset slope and offset for bend
'1.7ANM         lsngSlopeToUse = gudtTest(chanNum).slope.ideal2
'1.7ANM         lsngOffset = (gudtTest(chanNum).slope.start - gudtTest(CHAN0).Index(1).location) * (gudtTest(chanNum).slope.ideal - gudtTest(chanNum).slope.ideal2)
'1.7ANM         lsngStopHigh = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.high + lsngOffset
'1.7ANM         lsngStopLow = gudtReading(chanNum).Index(1).Value + ((gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location - gudtTest(CHAN0).Index(1).location) * lsngSlopeToUse) + gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.low + lsngOffset
'1.7ANM         'Calculate Voltage Gradient Limit Line Slopes & Intercepts after the bend
'1.7ANM         Call CalcLimitLineMandB(gudtTest(chanNum).slope.start, lsngStartHigh, lsngStartLow, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngStopHigh, lsngStopLow, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
'1.7ANM         'High Voltage Gradient Limit Array after the bend
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).slope.start, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM         'Low Voltage Gradient Limit Array after the bend
'1.7ANM         Call Calc.calcLimitArray(gudtTest(chanNum).slope.start, gudtTest(chanNum).SinglePointLin(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinVoltageGradientLimit())
'1.7ANM         If gintAnomaly Then Exit Sub        'Exit on system error
'1.7ANM     End If
'1.7ANM Next lintRegionNum
'1.7ANM '1.1ANM /\/\

'Voltage Gradient Limit Arrays Beyond End of Linearity Evaluation
'High Voltage Gradient Limit Array
lsngHighLimitSlope = 0
lsngHighLimitYIntercept = 100 '2.1ANM gudtTest(chanNum).maxOutput.high
Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(5).stop.location, gudtTest(chanNum).evaluate.stop, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxVoltageGradientLimit())
If gintAnomaly Then Exit Sub        'Exit on system error
'Low Voltage Gradient Limit Array
lsngLowLimitSlope = 0
lsngLowLimitYIntercept = 0 '2.1ANM gudtTest(chanNum).maxOutput.low
Call Calc.calcLimitArray(gudtTest(chanNum).SinglePointLin(5).stop.location, gudtTest(chanNum).evaluate.stop, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinVoltageGradientLimit())
If gintAnomaly Then Exit Sub        'Exit on system error

'Calculate Limit Arrays for Slope Deviation
'High Slope Limit Array
Call Calc.calcLimitArray(gudtTest(chanNum).slope.start, gudtTest(chanNum).slope.stop, 0, gudtTest(chanNum).slope.high, gsngResolution / gudtMachine.slopeIncrement, msngMaxSlopeDeviationLimit())
If gintAnomaly Then Exit Sub            'Exit on system error
'Low Slope Limit Array
Call Calc.calcLimitArray(gudtTest(chanNum).slope.start, gudtTest(chanNum).slope.stop, 0, gudtTest(chanNum).slope.low, gsngResolution / gudtMachine.slopeIncrement, msngMinSlopeDeviationLimit())
If gintAnomaly Then Exit Sub            'Exit on system error

'Calculate limit arrays for Hysteresis
'High Limit Array       'Use 100% as Hysteresis High Limit
Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, HUNDREDPERCENT, gsngResolution, msngMaxHysteresisLimit())
If gintAnomaly Then Exit Sub                'Exit on system error
'Low Limit Array        'Use -100% as Hysteresis Low Limit
Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, -HUNDREDPERCENT, gsngResolution, msngMinHysteresisLimit())
If gintAnomaly Then Exit Sub                'Exit on system error

'Calculate limit arrays for Output Correlation & Force Parameters
If chanNum = CHAN0 Then

    'Re-Dimension the Limit Arrays
    ReDim msngMaxFwdOutputCorLimit(gintMaxData)
    ReDim msngMinFwdOutputCorLimit(gintMaxData)
    ReDim msngMaxRevOutputCorLimit(gintMaxData)
    ReDim msngMinRevOutputCorLimit(gintMaxData)
    ReDim msngMaxForceGradientLimit(gintMaxData)
    ReDim msngMinForceGradientLimit(gintMaxData)
    ReDim msngMaxMechHysteresisLimit(gintMaxData)
    ReDim msngMinMechHysteresisLimit(gintMaxData)

    'Output Correlation Limit Arrays
    '2.2ANM \/\/
    If gblnUseNewAmad Then
        lstrParameterName = "FwdOutputCorrelation"
        lintRegionCount = AMAD705_2.MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
    Else
        lintRegionCount = 5
    End If
    '2.2ANM /\/\
    
    For lintRegionNum = 1 To lintRegionCount
        'Calculate Forward Output Correlation Limit Line Slopes & Intercepts
        Call CalcLimitLineMandB(gudtTest(chanNum).fwdOutputCor(lintRegionNum).start.location, gudtTest(chanNum).fwdOutputCor(lintRegionNum).start.high, gudtTest(chanNum).fwdOutputCor(lintRegionNum).start.low, gudtTest(chanNum).fwdOutputCor(lintRegionNum).stop.location, gudtTest(chanNum).fwdOutputCor(lintRegionNum).stop.high, gudtTest(chanNum).fwdOutputCor(lintRegionNum).stop.low, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
        'Calculate Forward Output Correlation high Limit Array
        Call Calc.calcLimitArray(gudtTest(chanNum).fwdOutputCor(lintRegionNum).start.location, gudtTest(chanNum).fwdOutputCor(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxFwdOutputCorLimit())
        If gintAnomaly Then Exit Sub        'Exit on system error
        'Calculate Forward Output Correlation Low Limit Array
        Call Calc.calcLimitArray(gudtTest(chanNum).fwdOutputCor(lintRegionNum).start.location, gudtTest(chanNum).fwdOutputCor(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinFwdOutputCorLimit())
        If gintAnomaly Then Exit Sub        'Exit on system error
        'Calculate Reverse Output Correlation Limit Line Slopes & Intercepts
        Call CalcLimitLineMandB(gudtTest(chanNum).revOutputCor(lintRegionNum).start.location, gudtTest(chanNum).revOutputCor(lintRegionNum).start.high, gudtTest(chanNum).revOutputCor(lintRegionNum).start.low, gudtTest(chanNum).revOutputCor(lintRegionNum).stop.location, gudtTest(chanNum).revOutputCor(lintRegionNum).stop.high, gudtTest(chanNum).revOutputCor(lintRegionNum).stop.low, lsngHighLimitSlope, lsngLowLimitSlope, lsngHighLimitYIntercept, lsngLowLimitYIntercept)
        'Calculate Reverse Output Correlation high Limit Array
        Call Calc.calcLimitArray(gudtTest(chanNum).revOutputCor(lintRegionNum).start.location, gudtTest(chanNum).revOutputCor(lintRegionNum).stop.location, lsngHighLimitSlope, lsngHighLimitYIntercept, gsngResolution, msngMaxRevOutputCorLimit())
        If gintAnomaly Then Exit Sub        'Exit on system error
        'Calculate Reverse Output Correlation Low Limit Array
        Call Calc.calcLimitArray(gudtTest(chanNum).revOutputCor(lintRegionNum).start.location, gudtTest(chanNum).revOutputCor(lintRegionNum).stop.location, lsngLowLimitSlope, lsngLowLimitYIntercept, gsngResolution, msngMinRevOutputCorLimit())
        If gintAnomaly Then Exit Sub        'Exit on system error
    Next lintRegionNum
    'Force Gradient High Limit Array        (Use 1000 Newtons as the High Limit Line)
    Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, 1000, gsngResolution, msngMaxForceGradientLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
    'Force Gradient Low Limit Array         (Use -1000 Newtons as the Low Limit Line)
    Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, -1000, gsngResolution, msngMinForceGradientLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
    'Mechanical Hysteresis High Limit Array (Use 1000 Newtons as the High Limit Line
    Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, 1000, gsngResolution, msngMaxMechHysteresisLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
    'Mechanical Hysteresis Low Limit Array  (Use -1000 Newtons as the Low Limit Line)
    Call Calc.calcLimitArray(gudtTest(chanNum).evaluate.start, gudtTest(chanNum).evaluate.stop, 0, -1000, gsngResolution, msngMinMechHysteresisLimit())
    If gintAnomaly Then Exit Sub        'Exit on system error
End If

End Sub

Public Sub InitializeAndMaskScanFailures()
'
'   PURPOSE: To initialize all failures to failed and mask out failures which are
'            not checked.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim lintFailureNum As Integer

'Initialize Booleans
gblnScanFailure = False     'Initialize failure to False
gblnSevere = False          'Initialize severe failure to False

'Initialize failure arrays to all parameters failed
For lintChanNum = CHAN0 To MAXCHANNUM
    For lintFailureNum = 0 To MAXFAULTCNT
        If lintFailureNum = 0 Then      'First element reserved for...
            gintFailure(lintChanNum, lintFailureNum) = MAXFAULTCNT     'Number of failures checked
            gintSevere(lintChanNum, lintFailureNum) = MAXFAULTCNT      'Number of severe failures checked
        Else                            'Otherwise...
            gintFailure(lintChanNum, lintFailureNum) = True            'Initialize all failures to failed
            gintSevere(lintChanNum, lintFailureNum) = True             'Initialize all severes to failed
        End If
    Next lintFailureNum
Next lintChanNum

'Mask all failures & severes not used

If gblnForceOnly Then                           '1.6ANM \/\/
    gintFailure(CHAN0, HIGHINDEXPT1) = False
    gintFailure(CHAN0, LOWINDEXPT1) = False
    gintFailure(CHAN0, HIGHOUTPUTATFORCEKNEE) = False
    gintFailure(CHAN0, LOWOUTPUTATFORCEKNEE) = False
    gintFailure(CHAN0, HIGHINDEXPT2) = False
    gintFailure(CHAN0, LOWINDEXPT2) = False
    gintFailure(CHAN0, HIGHINDEXPT3) = False
    gintFailure(CHAN0, LOWINDEXPT3) = False
    gintFailure(CHAN0, HIGHMAXOUTPUT) = False
    gintFailure(CHAN0, LOWMAXOUTPUT) = False
    gintFailure(CHAN0, HIGHSINGLEPOINTLIN) = False
    gintFailure(CHAN0, LOWSINGLEPOINTLIN) = False
    gintFailure(CHAN0, HIGHABSLIN) = False '2.7ANM
    gintFailure(CHAN0, LOWABSLIN) = False  '2.7ANM
    gintFailure(CHAN0, HIGHSLOPE) = False
    gintFailure(CHAN0, LOWSLOPE) = False
    gintFailure(CHAN0, HIGHFWDOUTPUTCOR) = False
    gintFailure(CHAN0, LOWFWDOUTPUTCOR) = False
    gintFailure(CHAN0, HIGHREVOUTPUTCOR) = False
    gintFailure(CHAN0, LOWREVOUTPUTCOR) = False
    gintFailure(CHAN0, HIGHFCHYS) = False
    gintFailure(CHAN0, LOWFCHYS) = False
    gintFailure(CHAN0, HIGHMLXI) = False '2.5ANM
    gintFailure(CHAN0, LOWMLXI) = False  '2.5ANM
    gintFailure(CHAN0, HIGHMLXI2) = False '2.8cANM
    gintFailure(CHAN0, LOWMLXI2) = False  '2.8cANM
    
    gintFailure(CHAN1, HIGHINDEXPT1) = False
    gintFailure(CHAN1, LOWINDEXPT1) = False
    gintFailure(CHAN1, HIGHOUTPUTATFORCEKNEE) = False
    gintFailure(CHAN1, LOWOUTPUTATFORCEKNEE) = False
    gintFailure(CHAN1, HIGHINDEXPT2) = False
    gintFailure(CHAN1, LOWINDEXPT2) = False
    gintFailure(CHAN1, HIGHINDEXPT3) = False
    gintFailure(CHAN1, LOWINDEXPT3) = False
    gintFailure(CHAN1, HIGHMAXOUTPUT) = False
    gintFailure(CHAN1, LOWMAXOUTPUT) = False
    gintFailure(CHAN1, HIGHSINGLEPOINTLIN) = False
    gintFailure(CHAN1, LOWSINGLEPOINTLIN) = False
    gintFailure(CHAN1, HIGHABSLIN) = False '2.7ANM
    gintFailure(CHAN1, LOWABSLIN) = False  '2.7ANM
    gintFailure(CHAN1, HIGHSLOPE) = False
    gintFailure(CHAN1, LOWSLOPE) = False
    gintFailure(CHAN1, HIGHFWDOUTPUTCOR) = False
    gintFailure(CHAN1, LOWFWDOUTPUTCOR) = False
    gintFailure(CHAN1, HIGHREVOUTPUTCOR) = False
    gintFailure(CHAN1, LOWREVOUTPUTCOR) = False
    gintFailure(CHAN1, HIGHFCHYS) = False
    gintFailure(CHAN1, LOWFCHYS) = False
    gintFailure(CHAN1, HIGHMLXI) = False '2.5ANM
    gintFailure(CHAN1, LOWMLXI) = False  '2.5ANM
    gintFailure(CHAN1, HIGHMLXI2) = False '2.8cANM
    gintFailure(CHAN1, LOWMLXI2) = False  '2.8cANM
    
    gintSevere(CHAN0, HIGHINDEXPT1) = False
    gintSevere(CHAN0, LOWINDEXPT1) = False
    gintSevere(CHAN0, HIGHOUTPUTATFORCEKNEE) = False
    gintSevere(CHAN0, LOWOUTPUTATFORCEKNEE) = False
    gintSevere(CHAN0, HIGHINDEXPT2) = False
    gintSevere(CHAN0, LOWINDEXPT2) = False
    gintSevere(CHAN0, HIGHINDEXPT3) = False
    gintSevere(CHAN0, LOWINDEXPT3) = False
    gintSevere(CHAN0, HIGHMAXOUTPUT) = False
    gintSevere(CHAN0, LOWMAXOUTPUT) = False
    gintSevere(CHAN0, HIGHSINGLEPOINTLIN) = False
    gintSevere(CHAN0, LOWSINGLEPOINTLIN) = False
    gintSevere(CHAN0, HIGHABSLIN) = False '2.7ANM
    gintSevere(CHAN0, LOWABSLIN) = False  '2.7ANM
    gintSevere(CHAN0, HIGHSLOPE) = False
    gintSevere(CHAN0, LOWSLOPE) = False
    gintSevere(CHAN0, HIGHFWDOUTPUTCOR) = False
    gintSevere(CHAN0, LOWFWDOUTPUTCOR) = False
    gintSevere(CHAN0, HIGHREVOUTPUTCOR) = False
    gintSevere(CHAN0, LOWREVOUTPUTCOR) = False
    gintSevere(CHAN0, HIGHFCHYS) = False
    gintSevere(CHAN0, LOWFCHYS) = False
    gintSevere(CHAN0, HIGHMLXI) = False                  '2.8aANM '2.5ANM
    gintSevere(CHAN0, LOWMLXI) = False                   '2.8aANM '2.5ANM
    gintSevere(CHAN0, HIGHMLXI2) = False '2.8cANM
    gintSevere(CHAN0, LOWMLXI2) = False  '2.8cANM

    gintSevere(CHAN1, HIGHINDEXPT1) = False
    gintSevere(CHAN1, LOWINDEXPT1) = False
    gintSevere(CHAN1, HIGHOUTPUTATFORCEKNEE) = False
    gintSevere(CHAN1, LOWOUTPUTATFORCEKNEE) = False
    gintSevere(CHAN1, HIGHINDEXPT2) = False
    gintSevere(CHAN1, LOWINDEXPT2) = False
    gintSevere(CHAN1, HIGHINDEXPT3) = False
    gintSevere(CHAN1, LOWINDEXPT3) = False
    gintSevere(CHAN1, HIGHMAXOUTPUT) = False
    gintSevere(CHAN1, LOWMAXOUTPUT) = False
    gintSevere(CHAN1, HIGHSINGLEPOINTLIN) = False
    gintSevere(CHAN1, LOWSINGLEPOINTLIN) = False
    gintSevere(CHAN1, HIGHABSLIN) = False '2.7ANM
    gintSevere(CHAN1, LOWABSLIN) = False  '2.7ANM
    gintSevere(CHAN1, HIGHSLOPE) = False
    gintSevere(CHAN1, LOWSLOPE) = False
    gintSevere(CHAN1, HIGHFWDOUTPUTCOR) = False
    gintSevere(CHAN1, LOWFWDOUTPUTCOR) = False
    gintSevere(CHAN1, HIGHREVOUTPUTCOR) = False
    gintSevere(CHAN1, LOWREVOUTPUTCOR) = False
    gintSevere(CHAN1, HIGHFCHYS) = False
    gintSevere(CHAN1, LOWFCHYS) = False
    gintSevere(CHAN1, HIGHMLXI) = False                  '2.8aANM '2.5ANM
    gintSevere(CHAN1, LOWMLXI) = False                   '2.8aANM '2.5ANM
    gintSevere(CHAN1, HIGHMLXI2) = False '2.8cANM
    gintSevere(CHAN1, LOWMLXI2) = False  '2.8cANM
End If                                          '1.6ANM /\/\

If gblnBnmkTest Then '2.8bANM
    gintFailure(CHAN0, HIGHMLXI) = False
    gintFailure(CHAN0, LOWMLXI) = False
    gintFailure(CHAN1, HIGHMLXI) = False
    gintFailure(CHAN1, LOWMLXI) = False
    gintSevere(CHAN0, HIGHMLXI) = False
    gintSevere(CHAN0, LOWMLXI) = False
    gintSevere(CHAN1, HIGHMLXI) = False
    gintSevere(CHAN1, LOWMLXI) = False
    '2.8cANM
    gintFailure(CHAN0, HIGHMLXI2) = False
    gintFailure(CHAN0, LOWMLXI2) = False
    gintFailure(CHAN1, HIGHMLXI2) = False
    gintFailure(CHAN1, LOWMLXI2) = False
    gintSevere(CHAN0, HIGHMLXI2) = False
    gintSevere(CHAN0, LOWMLXI2) = False
    gintSevere(CHAN1, HIGHMLXI2) = False
    gintSevere(CHAN1, LOWMLXI2) = False
End If

'Output 1 failures
gintFailure(CHAN0, HIGHPEAKFORCE) = False       'No Peak Force Failures, Severe Failure Only
gintFailure(CHAN0, LOWPEAKFORCE) = False        'No Peak Force Failures, Severe Failure Only

gintFailure(CHAN0, HIGHOUTPUTATFORCEKNEE) = False    '1.5ANM \/\/
gintFailure(CHAN0, LOWOUTPUTATFORCEKNEE) = False
gintFailure(CHAN0, HIGHINDEXPT2) = False
gintFailure(CHAN0, LOWINDEXPT2) = False

gintFailure(CHAN0, HIGHFWDFORCEPT2) = False
gintFailure(CHAN0, LOWFWDFORCEPT2) = False
gintFailure(CHAN0, HIGHREVFORCEPT2) = False
gintFailure(CHAN0, LOWREVFORCEPT2) = False
gintFailure(CHAN0, HIGHMECHHYSTPT2) = False
gintFailure(CHAN0, LOWMECHHYSTPT2) = False           '1.5ANM /\/\

gintFailure(CHAN0, HIGHMECHHYSTPT1) = False          '2.2ANM \/\/
gintFailure(CHAN0, LOWMECHHYSTPT1) = False
gintFailure(CHAN0, HIGHMECHHYSTPT3) = False
gintFailure(CHAN0, LOWMECHHYSTPT3) = False
gintFailure(CHAN0, HIGHFORCEKNEELOC) = False
gintFailure(CHAN0, LOWFORCEKNEELOC) = False
gintSevere(CHAN0, HIGHMECHHYSTPT1) = False
gintSevere(CHAN0, LOWMECHHYSTPT1) = False
gintSevere(CHAN0, HIGHMECHHYSTPT3) = False
gintSevere(CHAN0, LOWMECHHYSTPT3) = False
gintSevere(CHAN0, HIGHFORCEKNEELOC) = False
gintSevere(CHAN0, LOWFORCEKNEELOC) = False           '2.2ANM /\/\

gintFailure(CHAN0, HIGHFORCEKNEEFWDFORCE) = False    '2.3ANM \/\/
gintFailure(CHAN0, LOWFORCEKNEEFWDFORCE) = False
gintSevere(CHAN0, HIGHFORCEKNEEFWDFORCE) = False
gintSevere(CHAN0, LOWFORCEKNEEFWDFORCE) = False      '2.3ANM /\/\

'2.8dANM \/\/
If gblnKD = False Then
    gintFailure(CHAN0, HIGHKDSTART) = False
    gintFailure(CHAN0, LOWKDSTART) = False
    gintFailure(CHAN0, HIGHKDSTOP) = False
    gintFailure(CHAN0, LOWKDSTOP) = False
    gintFailure(CHAN0, HIGHKDSPAN) = False '2.8dANM
    gintFailure(CHAN0, LOWKDSPAN) = False  '2.8dANM
End If
'2.8dANM /\/\

'Output 1 severes
gintSevere(CHAN0, HIGHFCHYS) = False
gintSevere(CHAN0, LOWFCHYS) = False
gintSevere(CHAN0, HIGHOUTPUTATFORCEKNEE) = False     '1.5ANM \/\/
gintSevere(CHAN0, LOWOUTPUTATFORCEKNEE) = False
gintSevere(CHAN0, HIGHINDEXPT2) = False
gintSevere(CHAN0, LOWINDEXPT2) = False
    
gintSevere(CHAN0, HIGHFWDFORCEPT2) = False
gintSevere(CHAN0, LOWFWDFORCEPT2) = False
gintSevere(CHAN0, HIGHREVFORCEPT2) = False
gintSevere(CHAN0, LOWREVFORCEPT2) = False
gintSevere(CHAN0, HIGHMECHHYSTPT2) = False
gintSevere(CHAN0, LOWMECHHYSTPT2) = False            '1.5ANM /\/\
gintSevere(CHAN0, HIGHKDSTART) = False               '2.8dANM \/\/
gintSevere(CHAN0, LOWKDSTART) = False
gintSevere(CHAN0, HIGHKDSTOP) = False
gintSevere(CHAN0, LOWKDSTOP) = False
gintSevere(CHAN0, HIGHKDSPAN) = False
gintSevere(CHAN0, LOWKDSPAN) = False                 '2.8dANM /\/\

'Output 2 failures
gintFailure(CHAN1, HIGHFORCEKNEELOC) = False
gintFailure(CHAN1, LOWFORCEKNEELOC) = False
gintFailure(CHAN1, HIGHFORCEKNEEFWDFORCE) = False
gintFailure(CHAN1, LOWFORCEKNEEFWDFORCE) = False
gintFailure(CHAN1, HIGHFWDOUTPUTCOR) = False
gintFailure(CHAN1, LOWFWDOUTPUTCOR) = False
gintFailure(CHAN1, HIGHREVOUTPUTCOR) = False
gintFailure(CHAN1, LOWREVOUTPUTCOR) = False
gintFailure(CHAN1, HIGHFWDFORCEPT1) = False
gintFailure(CHAN1, LOWFWDFORCEPT1) = False
gintFailure(CHAN1, HIGHFWDFORCEPT2) = False
gintFailure(CHAN1, LOWFWDFORCEPT2) = False
gintFailure(CHAN1, HIGHFWDFORCEPT3) = False
gintFailure(CHAN1, LOWFWDFORCEPT3) = False
gintFailure(CHAN1, HIGHREVFORCEPT1) = False
gintFailure(CHAN1, LOWREVFORCEPT1) = False
gintFailure(CHAN1, HIGHREVFORCEPT2) = False
gintFailure(CHAN1, LOWREVFORCEPT2) = False
gintFailure(CHAN1, HIGHREVFORCEPT3) = False
gintFailure(CHAN1, LOWREVFORCEPT3) = False
gintFailure(CHAN1, HIGHPEAKFORCE) = False
gintFailure(CHAN1, LOWPEAKFORCE) = False
gintFailure(CHAN1, HIGHMECHHYSTPT1) = False
gintFailure(CHAN1, LOWMECHHYSTPT1) = False
gintFailure(CHAN1, HIGHMECHHYSTPT2) = False
gintFailure(CHAN1, LOWMECHHYSTPT2) = False
gintFailure(CHAN1, HIGHMECHHYSTPT3) = False
gintFailure(CHAN1, LOWMECHHYSTPT3) = False
gintFailure(CHAN1, HIGHOUTPUTATFORCEKNEE) = False     '1.5ANM \/\/
gintFailure(CHAN1, LOWOUTPUTATFORCEKNEE) = False
gintFailure(CHAN1, HIGHINDEXPT2) = False
gintFailure(CHAN1, LOWINDEXPT2) = False               '1.5ANM /\/\
gintFailure(CHAN1, HIGHPEDALATREST) = False           '2.8ANM
gintFailure(CHAN1, LOWPEDALATREST) = False            '2.8ANM
gintFailure(CHAN1, HIGHKDSTART) = False               '2.8dANM \/\/
gintFailure(CHAN1, LOWKDSTART) = False
gintFailure(CHAN1, HIGHKDSTOP) = False
gintFailure(CHAN1, LOWKDSTOP) = False
gintFailure(CHAN1, HIGHKDSPAN) = False
gintFailure(CHAN1, LOWKDSPAN) = False                 '2.8dANM /\/\

'Output 2 severe failures
gintSevere(CHAN1, HIGHFORCEKNEELOC) = False
gintSevere(CHAN1, LOWFORCEKNEELOC) = False
gintSevere(CHAN1, HIGHFORCEKNEEFWDFORCE) = False
gintSevere(CHAN1, LOWFORCEKNEEFWDFORCE) = False
gintSevere(CHAN1, HIGHFWDOUTPUTCOR) = False
gintSevere(CHAN1, LOWFWDOUTPUTCOR) = False
gintSevere(CHAN1, HIGHREVOUTPUTCOR) = False
gintSevere(CHAN1, LOWREVOUTPUTCOR) = False
gintSevere(CHAN1, HIGHFWDFORCEPT1) = False
gintSevere(CHAN1, LOWFWDFORCEPT1) = False
gintSevere(CHAN1, HIGHFWDFORCEPT2) = False
gintSevere(CHAN1, LOWFWDFORCEPT2) = False
gintSevere(CHAN1, HIGHFWDFORCEPT3) = False
gintSevere(CHAN1, LOWFWDFORCEPT3) = False
gintSevere(CHAN1, HIGHREVFORCEPT1) = False
gintSevere(CHAN1, LOWREVFORCEPT1) = False
gintSevere(CHAN1, HIGHREVFORCEPT2) = False
gintSevere(CHAN1, LOWREVFORCEPT2) = False
gintSevere(CHAN1, HIGHREVFORCEPT3) = False
gintSevere(CHAN1, LOWREVFORCEPT3) = False
gintSevere(CHAN1, HIGHPEAKFORCE) = False
gintSevere(CHAN1, LOWPEAKFORCE) = False
gintSevere(CHAN1, HIGHMECHHYSTPT1) = False
gintSevere(CHAN1, LOWMECHHYSTPT1) = False
gintSevere(CHAN1, HIGHMECHHYSTPT2) = False
gintSevere(CHAN1, LOWMECHHYSTPT2) = False
gintSevere(CHAN1, HIGHMECHHYSTPT3) = False
gintSevere(CHAN1, LOWMECHHYSTPT3) = False
gintSevere(CHAN1, HIGHFCHYS) = False
gintSevere(CHAN1, LOWFCHYS) = False
gintSevere(CHAN1, HIGHOUTPUTATFORCEKNEE) = False     '1.5ANM \/\/
gintSevere(CHAN1, LOWOUTPUTATFORCEKNEE) = False
gintSevere(CHAN1, HIGHINDEXPT2) = False
gintSevere(CHAN1, LOWINDEXPT2) = False               '1.5ANM /\/\
gintSevere(CHAN1, HIGHPEDALATREST) = False           '2.8ANM
gintSevere(CHAN1, LOWPEDALATREST) = False            '2.8ANM
gintSevere(CHAN1, HIGHKDSTART) = False               '2.8dANM \/\/
gintSevere(CHAN1, LOWKDSTART) = False
gintSevere(CHAN1, HIGHKDSTOP) = False
gintSevere(CHAN1, LOWKDSTOP) = False
gintSevere(CHAN1, HIGHKDSPAN) = False
gintSevere(CHAN1, LOWKDSPAN) = False                 '2.8dANM /\/\

End Sub

Public Sub LoadControlLimits()
'
'   PURPOSE:    To determine the control limits on critical parameters.
'
'  INPUT(S):    None.
' OUTPUT(S):    The associated control limit variables.

Dim lintChanNum As Integer
Dim lintIndexNum As Integer
Dim lintLinType As Integer
Dim lintSlopeCorrType As Integer
Dim lintForcePtNum As Integer
Dim lintMechHystPtNum As Integer

'Calculate the control limits for all channels
For lintChanNum = CHAN0 To MAXCHANNUM
    'Index
    For lintIndexNum = 1 To 3   '(FullClose Output, Midpoint Output, & FullOpen Output)
        Call Calc.CalcControlLimits(gudtTest(lintChanNum).Index(lintIndexNum).ideal, gudtTest(lintChanNum).Index(lintIndexNum).low, gudtTest(lintChanNum).Index(lintIndexNum).high, gudtControl(lintChanNum).Index(lintIndexNum).low, gudtControl(lintChanNum).Index(lintIndexNum).high)
    Next lintIndexNum
    'Maximum Output
    Call Calc.CalcControlLimits(gudtTest(lintChanNum).maxOutput.ideal, gudtTest(lintChanNum).maxOutput.low, gudtTest(lintChanNum).maxOutput.high, gudtControl(lintChanNum).maxOutput.low, gudtControl(lintChanNum).maxOutput.high)
    'Linearity Deviation    (Nominal is 0% of Tolerance)
    For lintLinType = 1 To 2    '(SngPt Lin & Abs Lin) '2.7ANM
        Call Calc.CalcControlLimits(0, -HUNDREDPERCENT, HUNDREDPERCENT, gudtControl(lintChanNum).linDevPerTol(lintLinType).low, gudtControl(lintChanNum).linDevPerTol(lintLinType).high)
    Next lintLinType
    'Slope Deviation        (Nominal is a ratio of 1)
    Call Calc.CalcControlLimits(1, gudtTest(lintChanNum).slope.low, gudtTest(lintChanNum).slope.high, gudtControl(lintChanNum).slope.low, gudtControl(lintChanNum).slope.high)
    'MLX Current '2.8aANM
    Call Calc.CalcControlLimits(gudtTest(lintChanNum).mlxCurrent.ideal, gudtTest(lintChanNum).mlxCurrent.low, gudtTest(lintChanNum).mlxCurrent.high, gudtControl(lintChanNum).mlxCurrent.low, gudtControl(lintChanNum).mlxCurrent.high)
Next

'Pedal at Rest Location '2.8ANM
Call Calc.CalcControlLimits(gudtTest(CHAN0).pedalAtRestLoc.ideal, gudtTest(CHAN0).pedalAtRestLoc.low, gudtTest(CHAN0).pedalAtRestLoc.high, gudtControl(CHAN0).pedalAtRestLoc.low, gudtControl(CHAN0).pedalAtRestLoc.high)

'Forward Output Correlation (Nominal is 0% of Tolerance)
Call Calc.CalcControlLimits(0, -HUNDREDPERCENT, HUNDREDPERCENT, gudtControl(CHAN0).outputCorPerTol(1).low, gudtControl(CHAN0).outputCorPerTol(1).high)

'Reverse Output Correlation (Nominal is 0% of Tolerance)
Call Calc.CalcControlLimits(0, -HUNDREDPERCENT, HUNDREDPERCENT, gudtControl(CHAN0).outputCorPerTol(2).low, gudtControl(CHAN0).outputCorPerTol(2).high)

'Force Knee Location
Call Calc.CalcControlLimits(gudtTest(CHAN0).forceKneeLoc.ideal, gudtTest(CHAN0).forceKneeLoc.low, gudtTest(CHAN0).forceKneeLoc.high, gudtControl(CHAN0).forceKneeLoc.low, gudtControl(CHAN0).forceKneeLoc.high)

'Forward Force Knee Force
Call Calc.CalcControlLimits(gudtTest(CHAN0).forceKneeForce.ideal, gudtTest(CHAN0).forceKneeForce.low, gudtTest(CHAN0).forceKneeForce.high, gudtControl(CHAN0).forceKneeForce.low, gudtControl(CHAN0).forceKneeForce.high)

'Forward & Reverse Force Points
For lintForcePtNum = 1 To 3   '(Points 1, 2, & 3)
    Call Calc.CalcControlLimits(gudtTest(CHAN0).fwdForcePt(lintForcePtNum).ideal, gudtTest(CHAN0).fwdForcePt(lintForcePtNum).low, gudtTest(CHAN0).fwdForcePt(lintForcePtNum).high, gudtControl(CHAN0).fwdForcePt(lintForcePtNum).low, gudtControl(CHAN0).fwdForcePt(lintForcePtNum).high)
    Call Calc.CalcControlLimits(gudtTest(CHAN0).revForcePt(lintForcePtNum).ideal, gudtTest(CHAN0).revForcePt(lintForcePtNum).low, gudtTest(CHAN0).revForcePt(lintForcePtNum).high, gudtControl(CHAN0).revForcePt(lintForcePtNum).low, gudtControl(CHAN0).revForcePt(lintForcePtNum).high)
Next lintForcePtNum

'Mechanical Hysteresis Points
For lintMechHystPtNum = 1 To 3   '(Points 1, 2, & 3)
    Call Calc.CalcControlLimits(gudtTest(CHAN0).mechHystPt(lintMechHystPtNum).ideal, gudtTest(CHAN0).mechHystPt(lintMechHystPtNum).low, gudtTest(CHAN0).mechHystPt(lintMechHystPtNum).high, gudtControl(CHAN0).mechHystPt(lintMechHystPtNum).low, gudtControl(CHAN0).mechHystPt(lintMechHystPtNum).high)
Next lintMechHystPtNum

End Sub

Public Sub LoadParameters()
'
'   PURPOSE:   To input operational parameters into the program from
'              a disk file.
'
'  INPUT(S):
' OUTPUT(S):

Dim lintRegionNum As Integer                'Region number
Dim lintRow As Integer                      'Row of table
Dim lintColumn As Integer                   'Column of table
Dim lstrParFileTable(999, 6) As String      'Parameter file converted to table format (MAX ROWS,MAX COLUMNS)
Dim lstrFileName As String                  'Name of parameter file from pull down combo box

'Get parameter file name which was selected
lstrFileName = App.Path + PARPATH + frmMain.cboParameterFileName

'Convert parameter file into a table
Call TabulateFile(lstrFileName, lstrParFileTable())

'*** Pedal at Rest Location *** '2.8ANM \/\/

'Pedal at Rest Location (Pedal at Rest Location)
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).pedalAtRestLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).pedalAtRestLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).pedalAtRestLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'*** Test Parameters Output #1 ***

'Index 1 - FullClose By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).Index(1).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(1).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'Output at Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
'2.3ANM gudtTest(CHAN0).outputAtForceKnee.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).outputAtForceKnee.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).outputAtForceKnee.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Index 2 - Midpoint By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).Index(2).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(2).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Index 3 - FullOpen By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).Index(3).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).Index(3).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Maximum Output
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).maxOutput.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).maxOutput.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).maxOutput.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'SinglePoint Linearity Deviation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Stop
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

'2.7ANM \/\/
'Absolute Linearity Deviation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).AbsLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Stop
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).AbsLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).AbsLin(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

'Slope Deviation Start
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).slope.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Slope Deviation Stop
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).slope.ideal2 = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).slope.stop = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Full-Close Hysteresis
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).FullCloseHys.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).FullCloseHys.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).FullCloseHys.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).FullCloseHys.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'MLX Current '2.5ANM \/\/
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).mlxCurrent.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mlxCurrent.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mlxCurrent.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'MLX WOT Current '2.8cANM \/\/
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).mlxWCurrent.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mlxWCurrent.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mlxWCurrent.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Evaluation Start
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).evaluate.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Evaluation Stop
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).evaluate.stop = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'*** Test Parameters Output #2 ***

'Index 1 - FullClose By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).Index(1).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(1).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'Output at Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
'2.3ANM gudtTest(CHAN1).outputAtForceKnee.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN1).outputAtForceKnee.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN1).outputAtForceKnee.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Index 2 - Midpoint By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).Index(2).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(2).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Index 3 - FullOpen By Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).Index(3).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).Index(3).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Maximum Output
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).maxOutput.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).maxOutput.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).maxOutput.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'SinglePoint Linearity Deviation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Stop
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

'2.7ANM \/\/
'Absolute Linearity Deviation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN1).AbsLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Stop
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN1).AbsLin(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN1).AbsLin(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).slope.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Slope Deviation Stop
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).slope.ideal2 = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).slope.stop = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Full-Close Hysteresis
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).FullCloseHys.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).FullCloseHys.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).FullCloseHys.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).FullCloseHys.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'MLX Current '2.5ANM \/\/
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).mlxCurrent.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).mlxCurrent.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).mlxCurrent.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'MLX WOT Current '2.8cANM \/\/
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN1).mlxWCurrent.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).mlxWCurrent.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN1).mlxWCurrent.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Evaluation Start
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN1).evaluate.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Evaluation Stop
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN1).evaluate.stop = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'*** Correlation Parameters ***

'Forward Output Correlation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

'Reverse Output Correlation
For lintRegionNum = 1 To 5      '5 Regions
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).revOutputCor(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    'Region Start
    lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
    gudtTest(CHAN0).revOutputCor(lintRegionNum).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
Next lintRegionNum

'*** Force Parameters ***

'Pedal at Rest Location
'2.8ANM lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
'2.8ANM gudtTest(CHAN0).pedalAtRestLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
'2.3ANM gudtTest(CHAN0).forceKneeLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).forceKneeLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).forceKneeLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'Forward Force at Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
'2.3ANM gudtTest(CHAN0).forceKneeForce.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).forceKneeForce.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtTest(CHAN0).forceKneeForce.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Forward Force Point #1
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).fwdForcePt(1).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(1).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Forward Force Point #2
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).fwdForcePt(2).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(2).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Forward Force Point #3
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).fwdForcePt(3).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).fwdForcePt(3).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Reverse Force Point #1
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).revForcePt(1).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(1).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Reverse Force Point #2
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).revForcePt(2).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(2).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Reverse Force Point #3
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).revForcePt(3).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).revForcePt(3).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Peak Force
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtTest(CHAN0).peakForce.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).peakForce.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Mechanical Hysteresis Point #1
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).mechHystPt(1).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(1).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Mechanical Hysteresis Point #2
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).mechHystPt(2).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(2).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'Mechanical Hysteresis Point #3
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).mechHystPt(3).ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).mechHystPt(3).location = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.8dANM \/\/
'*** Kickdown ***
'Kickdown Start Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).kickdownStartLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownStartLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownStartLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownStartLoc.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).kickdownStartLoc.stop = Val(lstrParFileTable(lintRow, lintColumn))

'Kickdown Force Span
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).kickdownForceSpan.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownForceSpan.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownForceSpan.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownForceSpan.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).kickdownForceSpan.stop = Val(lstrParFileTable(lintRow, lintColumn))

'Kickdown Peak Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).kickdownPeakLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakLoc.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).kickdownPeakLoc.stop = Val(lstrParFileTable(lintRow, lintColumn))

'Kickdown Peak Force
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).kickdownPeakForce.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakForce.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakForce.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownPeakForce.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).kickdownPeakForce.stop = Val(lstrParFileTable(lintRow, lintColumn))

'Kickdown End Location
lintRow = lintRow + 1: lintColumn = 1   'row & column start locations
gudtTest(CHAN0).kickdownEndLoc.ideal = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownEndLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownEndLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtTest(CHAN0).kickdownEndLoc.start = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
lintRow = lintRow + 1: lintColumn = 4   'row & column start locations
gudtTest(CHAN0).kickdownEndLoc.stop = Val(lstrParFileTable(lintRow, lintColumn))
'2.8dANM /\/\

'*** STATS Parameters for CP & CPK calculations ***

'STATS: Index 1 (FullClose) Output #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).Index(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).Index(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Index 2 (Midpoint) Output #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).Index(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).Index(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Index 3 (FullOpen) Output #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).Index(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).Index(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Index 1 (FullClose) Output #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN1).Index(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN1).Index(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Index 2 (Midpoint) Output #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN1).Index(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN1).Index(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Index 3 (FullOpen) Output #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN1).Index(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN1).Index(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'STATS: Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
'2.3ANM gudtCustomerSpec(CHAN0).forceKneeLoc.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtCustomerSpec(CHAN0).forceKneeLoc.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'2.3ANM 'STATS: Forward Force at Force Knee Location
'2.3ANM lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
'2.3ANM gudtCustomerSpec(CHAN0).forceKneeForce.high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
'2.3ANM gudtCustomerSpec(CHAN0).forceKneeForce.low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Forward Force Point #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).fwdForcePt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).fwdForcePt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Forward Force Point #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).fwdForcePt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).fwdForcePt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Forward Force Point #3
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).fwdForcePt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).fwdForcePt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Reverse Force Point #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).revForcePt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).revForcePt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Reverse Force Point #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).revForcePt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).revForcePt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Reverse Force Point #3
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).revForcePt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).revForcePt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Mechanical Hysteresis Point #1
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).mechHystPt(1).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).mechHystPt(1).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Mechanical Hysteresis Point #2
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).mechHystPt(2).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).mechHystPt(2).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'STATS: Mechanical Hysteresis Point #3
lintRow = lintRow + 1: lintColumn = 2   'row & column start locations
gudtCustomerSpec(CHAN0).mechHystPt(3).high = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtCustomerSpec(CHAN0).mechHystPt(3).low = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1

'*** PTC-04 Parameters ***
lintRow = lintRow + 2: lintColumn = 1  'row & column start locations
'Output #1
gudtPTC04(1).Tpuls = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).Tpor = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).Tprog = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).Thold = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).Tsynchro = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).TpulsMax = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).TpulsMin = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).SynchroDelay = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).ByteData = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).TSentTick = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).Baudrate = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).BaudrateSyncID = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).VDDLow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).VDDNorm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).VDDComm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).VBatLow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(1).VBatNorm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1

'Output #2
gudtPTC04(2).Tpuls = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).Tpor = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).Tprog = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).Thold = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).Tsynchro = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).TpulsMax = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).TpulsMin = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).SynchroDelay = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).ByteData = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).TSentTick = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).Baudrate = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).BaudrateSyncID = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).VDDLow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).VDDNorm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).VDDComm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).VBatLow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtPTC04(2).VBatNorm = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1

'MLX 90277 Chip Revision
gstrMLX90277Revision = lstrParFileTable(lintRow, lintColumn): lintRow = lintRow + 1

'*** Solver Parameters ***
lintRow = lintRow + 1: lintColumn = 1  'row & column start locations
'Output #1
gudtSolver(1).Index(1).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(1).IdealLocation = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(1).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(1).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(2).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(2).IdealLocation = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(2).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Index(2).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Filter = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).InvertSlope = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Mode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).FaultLevel = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).MaxOffsetDrift = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
lintColumn = 2  'column start location
gudtSolver(1).MaxAGND = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtSolver(1).MinAGND = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
lintColumn = 1  'column start location
gudtSolver(1).FCKADJ = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).CKANACH = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).CKDACCH = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).SlowMode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).InitialOffset = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).HighRGHighFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1  '1.8ANM
gudtSolver(1).HighRGLowFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1   '1.8ANM
gudtSolver(1).LowRGHighFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1   '1.8ANM
gudtSolver(1).LowRGLowFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1    '1.8ANM
gudtSolver(1).MinRG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).MaxRG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).OffsetStep = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(1).CodeRatio(1, 1) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(1).CodeRatio(1, 2) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(1).CodeRatio(1, 3) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).CodeRatio(2, 1) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).CodeRatio(2, 2) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).CodeRatio(2, 3) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(1).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(1).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(1).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(1).InitialCode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(2).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(2).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(2).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).Clamp(2).InitialCode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(1).ClampStep = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'Output #2
gudtSolver(2).Index(1).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(1).IdealLocation = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(1).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(1).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(2).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(2).IdealLocation = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(2).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Index(2).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Filter = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).InvertSlope = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Mode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).FaultLevel = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).MaxOffsetDrift = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
lintColumn = 2  'column start location
gudtSolver(2).MaxAGND = Val(lstrParFileTable(lintRow, lintColumn)): lintColumn = lintColumn + 1
gudtSolver(2).MinAGND = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
lintColumn = 1  'column start location
gudtSolver(2).FCKADJ = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).CKANACH = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).CKDACCH = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).SlowMode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).InitialOffset = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).HighRGHighFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1  '1.8ANM
gudtSolver(2).HighRGLowFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1   '1.8ANM
gudtSolver(2).LowRGHighFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1   '1.8ANM
gudtSolver(2).LowRGLowFG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1    '1.8ANM
gudtSolver(2).MinRG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).MaxRG = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).OffsetStep = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(2).CodeRatio(1, 1) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(2).CodeRatio(1, 2) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
'2.4ANM gudtSolver(2).CodeRatio(1, 3) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).CodeRatio(2, 1) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).CodeRatio(2, 2) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).CodeRatio(2, 3) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(1).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(1).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(1).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(1).InitialCode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(2).IdealValue = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(2).TargetTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(2).PassFailTolerance = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).Clamp(2).InitialCode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtSolver(2).ClampStep = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1

'*** Machine Parameters ***
lintRow = lintRow + 1: lintColumn = 1  'row & column start locations
gudtMachine.parameterName = lstrParFileTable(lintRow, lintColumn): lintRow = lintRow + 1
gudtMachine.parameterRev = lstrParFileTable(lintRow, lintColumn): lintRow = lintRow + 1
gudtMachine.BOMNumber = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.stationCode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.seriesID = LCase$(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.PLCCommType = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtTest(CHAN0).riseTarget = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1    '1.5ANM
gudtMachine.loadLocation = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.preScanStart = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.preScanStop = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.offset4StartScan = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.scanLength = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.overTravel = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.countsPerTrigger = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.gearRatio = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.encReso = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.pedalAtRestLocForce = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.FKSlope = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.FKWindow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.FKPercentage = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.slopeInterval = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.slopeIncrement = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.preScanVelocity = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.preScanAcceleration = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.scanVelocity = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.scanAcceleration = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.scanVelocityB = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1       '2.8bANM
gudtMachine.scanAccelerationB = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1   '2.8bANM
gudtMachine.progVelocity = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.progAcceleration = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.graphZeroOffset = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.currentPartCount = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.yieldGreen = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.yieldYellow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.xAxisLow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.xAxisHigh = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.filterLoc(CHAN0) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.filterLoc(CHAN1) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.filterLoc(CHAN2) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.filterLoc(CHAN3) = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.blockOffset = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.VRefMode = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.maxLBF = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gblnKD = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1                        '2.8dANM \/\/
'Don't want KD Programming so make sure normal enable is off
gudtMachine.kickdown = False
gudtMachine.KDStartSlope = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.KDStartWindow = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1
gudtMachine.KDStartPercentage = Val(lstrParFileTable(lintRow, lintColumn)): lintRow = lintRow + 1 '2.8dANM /\/\

'Calculate scan resolution
gsngResolution = (gudtMachine.encReso \ DEGPERREV) \ gudtMachine.countsPerTrigger

'Set the PT Board to 4X Resolution
Call frmDAQIO.OnPort1(PORT4, BIT0)

'Dimension Graph Arrays
ReDim gvntGraph(0 To 64, 0 To ((gudtMachine.scanLength * gsngResolution) * MAXGRAPHS))
ReDim gsngMultipleGraphArray(0 To (gudtMachine.scanLength * gsngResolution) * MAXGRAPHS)

'Set up the Default action on PLC Start
Select Case gudtMachine.stationCode
    Case 1
        'Program Only
        frmMain.mnuOptionsPLCStartProgram.Checked = True
        frmMain.mnuOptionsPLCStartScan.Checked = False
        frmMain.ctrResultsTabs1.ActiveTab = PROGRESULTSGRID
    Case 2
        'Scan Only
        frmMain.mnuOptionsPLCStartProgram.Checked = False
        frmMain.mnuOptionsPLCStartScan.Checked = True
        frmMain.ctrResultsTabs1.ActiveTab = SCANRESULTSGRID
    Case 3
        'Program and Scan
        frmMain.mnuOptionsPLCStartProgram.Checked = True
        frmMain.mnuOptionsPLCStartScan.Checked = True
        frmMain.ctrResultsTabs1.ActiveTab = SCANRESULTSGRID
End Select

'Send the BOM Setup Code to the PLC
If InStr(command$, "NOHARDWARE") = 0 Then
    '*** PLC DDE Setup ***
    If gudtMachine.PLCCommType Then
        'Setup PLC DDE Topics and Items
        Call frmDDE.PLCDDESetup
    
        'Clear Results Code
        Call frmDDE.WriteDDEOutput(ResultsCode, 0)
        'Clear BOM Number (Setup Code)
        Call frmDDE.WriteDDEOutput(BOMSetupCode, 0)
        'Clear start scan acknowledge, calc complete, scanner init and ready, and graphics mode
        Call frmDDE.WriteDDEOutput(StartScanAck, 0)
        Call frmDDE.WriteDDEOutput(CalcComplete, 0)
        Call frmDDE.WriteDDEOutput(ScannerInit, 0)
        Call frmDDE.WriteDDEOutput(WatchdogDisable, 0)
        Call frmDDE.WriteDDEOutput(StationFault, 0)

        'Set the BOM Number
        Call frmDDE.WriteDDEOutput(BOMSetupCode, gudtMachine.BOMNumber)
    End If

End If

End Sub

Public Sub SevereMessage(chanNum As Integer, faultNum As Integer)
'
'   PURPOSE:    To display a severe fault message to the screen.
'
'  INPUT(S):    ChanNum         : Channel number of the failure
'               FaultNum        : Identifies the severe failure
' OUTPUT(S):    None.

Dim lstrMsg1 As String
Dim lstrMsg2 As String
Dim lintOutputID As Integer

lintOutputID = chanNum + 1

lstrMsg2 = "Click OK to continue." '2.8aANM moved above case

Select Case faultNum
    Case HIGHINDEXPT1
        lstrMsg1 = "SEVERE HIGH FULL CLOSE OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWINDEXPT1
        lstrMsg1 = "SEVERE LOW FULL CLOSE OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHOUTPUTATFORCEKNEE
        lstrMsg1 = "SEVERE HIGH OUTPUT AT FORCE KNEE on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWOUTPUTATFORCEKNEE
        lstrMsg1 = "SEVERE LOW OUTPUT AT FORCE KNEE on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHINDEXPT2
        lstrMsg1 = "SEVERE HIGH MIDPOINT OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWINDEXPT2
        lstrMsg1 = "SEVERE LOW MIDPOINT OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHINDEXPT3
        lstrMsg1 = "SEVERE HIGH FULL OPEN LOCATION on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWINDEXPT3
        lstrMsg1 = "SEVERE LOW FULL OPEN LOCATION on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHMAXOUTPUT
        lstrMsg1 = "SEVERE HIGH MAX OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWMAXOUTPUT
        lstrMsg1 = "SEVERE LOW MAX OUTPUT on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHSINGLEPOINTLIN
        lstrMsg1 = "SEVERE HIGH SINGLE-POINT LINEARITY DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWSINGLEPOINTLIN
        lstrMsg1 = "SEVERE LOW SINGLE-POINT LINEARITY DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHSLOPE
        lstrMsg1 = "SEVERE HIGH SLOPE DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWSLOPE
        lstrMsg1 = "SEVERE LOW SLOPE DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHFWDOUTPUTCOR
        lstrMsg1 = "SEVERE HIGH FORWARD OUTPUT CORRELATION"
    Case LOWFWDOUTPUTCOR
        lstrMsg1 = "SEVERE LOW FORWARD OUTPUT CORRELATION"
    Case HIGHREVOUTPUTCOR
        lstrMsg1 = "SEVERE HIGH REVERSE OUTPUT CORRELATION"
    Case LOWREVOUTPUTCOR
        lstrMsg1 = "SEVERE LOW REVERSE OUTPUT CORRELATION"
    Case HIGHFORCEKNEELOC
        lstrMsg1 = "SEVERE HIGH FORCE KNEE LOCATION"
    Case LOWFORCEKNEELOC
        lstrMsg1 = "SEVERE LOW FORCE KNEE LOCATION"
    Case HIGHFORCEKNEEFWDFORCE
        lstrMsg1 = "SEVERE HIGH FORCE KNEE FORWARD FORCE"
    Case LOWFORCEKNEEFWDFORCE
        lstrMsg1 = "SEVERE LOW FORCE KNEE FORWARD FORCE"
    Case HIGHFWDFORCEPT1
        lstrMsg1 = "SEVERE HIGH FORWARD FORCE POINT 1"
    Case LOWFWDFORCEPT1
        lstrMsg1 = "SEVERE LOW FORWARD FORCE POINT 1"
    Case HIGHFWDFORCEPT2
        lstrMsg1 = "SEVERE HIGH FORWARD FORCE POINT 2"
    Case LOWFWDFORCEPT2
        lstrMsg1 = "SEVERE LOW FORWARD FORCE POINT 2"
    Case HIGHFWDFORCEPT3
        lstrMsg1 = "SEVERE HIGH FORWARD FORCE POINT 3"
    Case LOWFWDFORCEPT3
        lstrMsg1 = "SEVERE LOW FORWARD FORCE POINT 3"
    Case HIGHREVFORCEPT1
        lstrMsg1 = "SEVERE HIGH REVERSE FORCE POINT 1"
    Case LOWREVFORCEPT1
        lstrMsg1 = "SEVERE LOW REVERSE FORCE POINT 1"
    Case HIGHREVFORCEPT2
        lstrMsg1 = "SEVERE HIGH REVERSE FORCE POINT 2"
    Case LOWREVFORCEPT2
        lstrMsg1 = "SEVERE LOW REVERSE FORCE POINT 2"
    Case HIGHREVFORCEPT3
        lstrMsg1 = "SEVERE HIGH REVERSE FORCE POINT 3"
    Case LOWREVFORCEPT3
        lstrMsg1 = "SEVERE LOW REVERSE FORCE POINT 3"
    Case HIGHPEAKFORCE
        lstrMsg1 = "SEVERE HIGH PEAK FORCE"
    Case LOWPEAKFORCE
        lstrMsg1 = "SEVERE LOW PEAK FORCE"
    Case HIGHMECHHYSTPT1
        lstrMsg1 = "SEVERE HIGH MECHANICAL HYSTERESIS POINT 1"
    Case LOWMECHHYSTPT1
        lstrMsg1 = "SEVERE LOW MECHANICAL HYSTERESIS POINT 1"
    Case HIGHMECHHYSTPT2
        lstrMsg1 = "SEVERE HIGH MECHANICAL HYSTERESIS POINT 2"
    Case LOWMECHHYSTPT2
        lstrMsg1 = "SEVERE LOW MECHANICAL HYSTERESIS POINT 2"
    Case HIGHMECHHYSTPT3
        lstrMsg1 = "SEVERE HIGH MECHANICAL HYSTERESIS POINT 3"
    Case LOWMECHHYSTPT3
        lstrMsg1 = "SEVERE LOW MECHANICAL HYSTERESIS POINT 3"
    Case HIGHMLXI          '2.8aANM '2.5ANM
        lstrMsg1 = "SEVERE HIGH MLX CURRENT!"
        lstrMsg2 = "Part should be removed and sent to Pedal Team."
    Case LOWMLXI           '2.8aANM '2.5ANM
        lstrMsg1 = "SEVERE LOW MLX CURRENT!"
        lstrMsg2 = "Part should be removed and sent to Pedal Team."
    Case HIGHABSLIN        '2.7ANM
        lstrMsg1 = "SEVERE HIGH ABSOLUTE LINEARITY DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWABSLIN
        lstrMsg1 = "SEVERE LOW ABSOLUTE LINEARITY DEVIATION on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHPEDALATREST   '2.8ANM
        lstrMsg1 = "SEVERE HIGH PEDAL AT REST LOC"
    Case LOWPEDALATREST
        lstrMsg1 = "SEVERE LOW PEDAL AT REST LOC"
    Case HIGHMLXI2 '2.8cANM
        lstrMsg1 = "SEVERE HIGH MLX WOT CURRENT!"
        lstrMsg2 = "Part should be removed and sent to Pedal Team."
    Case LOWMLXI2
        lstrMsg1 = "SEVERE LOW MLX WOT CURRENT!"
        lstrMsg2 = "Part should be removed and sent to Pedal Team."
    Case HIGHKDSTART  '2.8dANM
        lstrMsg1 = "SEVERE HIGH KICKDOWN START on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWKDSTART   '2.8dANM
        lstrMsg1 = "SEVERE LOW KICKDOWN START on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHKDSTOP   '2.8dANM
        lstrMsg1 = "SEVERE HIGH KICKDOWN STOP on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWKDSTOP    '2.8dANM
        lstrMsg1 = "SEVERE LOW KICKDOWN STOP on Output # " & Format(lintOutputID, "0") & ")"
    Case HIGHKDSPAN   '2.8dANM
        lstrMsg1 = "SEVERE HIGH KICKDOWN SPAN on Output # " & Format(lintOutputID, "0") & ")"
    Case LOWKDSPAN    '2.8dANM
        lstrMsg1 = "SEVERE LOW KICKDOWN SPAN on Output # " & Format(lintOutputID, "0") & ")"
End Select

'Display the message box
MsgBox lstrMsg1 & vbCrLf & vbCrLf & lstrMsg2, vbOKOnly + vbCritical + vbSystemModal, "Severe Unit Failure"

End Sub

Public Sub StatsClear()
'
'   PURPOSE:   To reset all production statistical variables
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer
Dim lintIndexNum As Integer
Dim lintLinType As Integer
Dim lintOutputCorType As Integer
Dim lintForcePt As Integer
Dim lintMechHysteresisPt As Integer
Dim lintProgrammerNum As Integer

'Clear the display control
Call frmMain.ctrResultsTabs1.ClearData(SCANSTATSGRID, 1, 8)
Call frmMain.ctrResultsTabs1.ClearData(SCANRESULTSGRID, 1, 4)

For lintChanNum = 0 To MAXCHANNUM
    'Index (FullClose Output, Midpoint Output, & FullOpen Output)
    For lintIndexNum = 1 To 3
        gudtScanStats(lintChanNum).Index(lintIndexNum).failCount.high = 0   'High Failure Count
        gudtScanStats(lintChanNum).Index(lintIndexNum).failCount.low = 0    'Low Failure Count
        gudtScanStats(lintChanNum).Index(lintIndexNum).max = -1000000       'Maximum Measured Value
        gudtScanStats(lintChanNum).Index(lintIndexNum).min = 1000000        'Minimum Measured Value
        gudtScanStats(lintChanNum).Index(lintIndexNum).sigma = 0            'Sum of Measured Values
        gudtScanStats(lintChanNum).Index(lintIndexNum).sigma2 = 0           'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).Index(lintIndexNum).n = 0                'Total Number of Values
    Next lintIndexNum

    'Output at Force Knee Location
    gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high = 0         'High Failure Count
    gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low = 0          'Low Failure Count
    gudtScanStats(lintChanNum).outputAtForceKnee.max = -1000000             'Maximum Measured Value
    gudtScanStats(lintChanNum).outputAtForceKnee.min = 1000000              'Minimum Measured Value
    gudtScanStats(lintChanNum).outputAtForceKnee.sigma = 0                  'Sum of Measured Values
    gudtScanStats(lintChanNum).outputAtForceKnee.sigma2 = 0                 'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).outputAtForceKnee.n = 0                      'Total Number of Values

    'Maximum Output
    gudtScanStats(lintChanNum).maxOutput.failCount.high = 0                 'High Failure Count
    gudtScanStats(lintChanNum).maxOutput.failCount.low = 0                  'Low Failure Count
    gudtScanStats(lintChanNum).maxOutput.max = -1000000                     'Maximum Measured Value
    gudtScanStats(lintChanNum).maxOutput.min = 1000000                      'Minimum Measured Value
    gudtScanStats(lintChanNum).maxOutput.sigma = 0                          'Sum of Measured Values
    gudtScanStats(lintChanNum).maxOutput.sigma2 = 0                         'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).maxOutput.n = 0                              'Total Number of Values

    'Linearity Deviation % of Tolerance
    'NOTE: LinType 1 = SinglePoint linearity
    '2.7ANM        2 = Absolute linearity
    For lintLinType = 1 To 2
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).failCount.high = 0     'High Failure Count
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).failCount.low = 0      'Low Failure Count
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).max = -1000000         'Maximum Measured Value
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).min = 1000000          'Minimum Measured Value
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).sigma = 0              'Sum of Measured Values
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).sigma2 = 0             'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).linDevPerTol(lintLinType).n = 0                  'Total Number of Values
    Next lintLinType

    'Slope Deviation Max
    gudtScanStats(lintChanNum).slopeMax.failCount.high = 0                  'High Failure Count
    gudtScanStats(lintChanNum).slopeMax.failCount.low = 0                   'Low Failure Count
    gudtScanStats(lintChanNum).slopeMax.max = -1000000                      'Maximum Measured Value
    gudtScanStats(lintChanNum).slopeMax.min = 1000000                       'Minimum Measured Value
    gudtScanStats(lintChanNum).slopeMax.sigma = 0                           'Sum of Measured Values
    gudtScanStats(lintChanNum).slopeMax.sigma2 = 0                          'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).slopeMax.n = 0                               'Total Number of Values

    'Slope Deviation Min
    gudtScanStats(lintChanNum).slopeMin.failCount.high = 0                  'High Failure Count
    gudtScanStats(lintChanNum).slopeMin.failCount.low = 0                   'Low Failure Count
    gudtScanStats(lintChanNum).slopeMin.max = -1000000                      'Maximum Measured Value
    gudtScanStats(lintChanNum).slopeMin.min = 1000000                       'Minimum Measured Value
    gudtScanStats(lintChanNum).slopeMin.sigma = 0                           'Sum of Measured Values
    gudtScanStats(lintChanNum).slopeMin.sigma2 = 0                          'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).slopeMin.n = 0                               'Total Number of Values

    'Full-Close Hysteresis
    gudtScanStats(lintChanNum).FullCloseHys.failCount.high = 0                  'High Failure Count
    gudtScanStats(lintChanNum).FullCloseHys.failCount.low = 0                   'Low Failure Count
    gudtScanStats(lintChanNum).FullCloseHys.max = -1000000                      'Maximum Measured Value
    gudtScanStats(lintChanNum).FullCloseHys.min = 1000000                       'Minimum Measured Value
    gudtScanStats(lintChanNum).FullCloseHys.sigma = 0                           'Sum of Measured Values
    gudtScanStats(lintChanNum).FullCloseHys.sigma2 = 0                          'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).FullCloseHys.n = 0                               'Total Number of Values

    'MLX Current '2.5ANM
    gudtScanStats(lintChanNum).mlxCurrent.failCount.high = 0                 'High Failure Count
    gudtScanStats(lintChanNum).mlxCurrent.failCount.low = 0                  'Low Failure Count
    gudtScanStats(lintChanNum).mlxCurrent.max = -1000000                     'Maximum Measured Value
    gudtScanStats(lintChanNum).mlxCurrent.min = 1000000                      'Minimum Measured Value
    gudtScanStats(lintChanNum).mlxCurrent.sigma = 0                          'Sum of Measured Values
    gudtScanStats(lintChanNum).mlxCurrent.sigma2 = 0                         'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).mlxCurrent.n = 0
    
    'MLX WOT Current '2.8cANM
    gudtScanStats(lintChanNum).mlxWCurrent.failCount.high = 0                 'High Failure Count
    gudtScanStats(lintChanNum).mlxWCurrent.failCount.low = 0                  'Low Failure Count
    gudtScanStats(lintChanNum).mlxWCurrent.max = -1000000                     'Maximum Measured Value
    gudtScanStats(lintChanNum).mlxWCurrent.min = 1000000                      'Minimum Measured Value
    gudtScanStats(lintChanNum).mlxWCurrent.sigma = 0                          'Sum of Measured Values
    gudtScanStats(lintChanNum).mlxWCurrent.sigma2 = 0                         'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).mlxWCurrent.n = 0                              'Total Number of Values
    
    'Output Correlation % of Tolerance
    'NOTE: OutputCorType 1 = forward output correlation
    '      OutputCorType 2 = reverse output correlation
    For lintOutputCorType = 1 To 2
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).failCount.high = 0     'High Failure Count
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).failCount.low = 0      'Low Failure Count
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).max = -1000000         'Maximum Measured Value
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).min = 1000000          'Minimum Measured Value
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).sigma = 0              'Sum of Measured Values
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).sigma2 = 0             'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).outputCorPerTol(lintOutputCorType).n = 0                  'Total Number of Values
    Next lintOutputCorType

    'Pedal-At-Rest Location
    gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.high = 0            'High Failure Count
    gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.low = 0             'Low Failure Count
    gudtScanStats(lintChanNum).pedalAtRestLoc.max = -1000000                'Maximum Measured Value
    gudtScanStats(lintChanNum).pedalAtRestLoc.min = 1000000                 'Minimum Measured Value
    gudtScanStats(lintChanNum).pedalAtRestLoc.sigma = 0                     'Sum of Measured Values
    gudtScanStats(lintChanNum).pedalAtRestLoc.sigma2 = 0                    'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).pedalAtRestLoc.n = 0                         'Total Number of Values

    'Force Knee Location
    gudtScanStats(lintChanNum).forceKneeLoc.failCount.high = 0              'High Failure Count
    gudtScanStats(lintChanNum).forceKneeLoc.failCount.low = 0               'Low Failure Count
    gudtScanStats(lintChanNum).forceKneeLoc.max = -1000000                  'Maximum Measured Value
    gudtScanStats(lintChanNum).forceKneeLoc.min = 1000000                   'Minimum Measured Value
    gudtScanStats(lintChanNum).forceKneeLoc.sigma = 0                       'Sum of Measured Values
    gudtScanStats(lintChanNum).forceKneeLoc.sigma2 = 0                      'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).forceKneeLoc.n = 0                           'Total Number of Values

    'Force Knee Force
    gudtScanStats(lintChanNum).forceKneeForce.failCount.high = 0            'High Failure Count
    gudtScanStats(lintChanNum).forceKneeForce.failCount.low = 0             'Low Failure Count
    gudtScanStats(lintChanNum).forceKneeForce.max = -1000000                'Maximum Measured Value
    gudtScanStats(lintChanNum).forceKneeForce.min = 1000000                 'Minimum Measured Value
    gudtScanStats(lintChanNum).forceKneeForce.sigma = 0                     'Sum of Measured Values
    gudtScanStats(lintChanNum).forceKneeForce.sigma2 = 0                    'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).forceKneeForce.n = 0                         'Total Number of Values

    'Pedal Force Points
    For lintForcePt = 1 To 3    'Three force points
        'Forward Points
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).failCount.high = 0               'High Failure Count
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).failCount.low = 0                'Low Failure Count
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).max = -1000000                   'Maximum Measured Value
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).min = 1000000                    'Minimum Measured Value
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).sigma = 0                        'Sum of Measured Values
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).sigma2 = 0                       'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).fwdForcePt(lintForcePt).n = 0                            'Total Number of Values
        'Reverse Points
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).failCount.high = 0               'High Failure Count
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).failCount.low = 0                'Low Failure Count
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).max = -1000000                   'Maximum Measured Value
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).min = 1000000                    'Minimum Measured Value
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).sigma = 0                        'Sum of Measured Values
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).sigma2 = 0                       'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).revForcePt(lintForcePt).n = 0                            'Total Number of Values
    Next lintForcePt

    'Peak Force
    gudtScanStats(lintChanNum).peakForce.failCount.high = 0                 'High Failure Count
    gudtScanStats(lintChanNum).peakForce.failCount.low = 0                  'Low Failure Count
    gudtScanStats(lintChanNum).peakForce.max = -1000000                     'Maximum Measured Value
    gudtScanStats(lintChanNum).peakForce.min = 1000000                      'Minimum Measured Value
    gudtScanStats(lintChanNum).peakForce.sigma = 0                          'Sum of Measured Values
    gudtScanStats(lintChanNum).peakForce.sigma2 = 0                         'Sum of Measured Squared Values
    gudtScanStats(lintChanNum).peakForce.n = 0                              'Total Number of Values

    'Mechanical Hysteresis
    For lintMechHysteresisPt = 1 To 3    'Three mechanical hysteresis points
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).failCount.high = 0      'High Failure Count
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).failCount.low = 0       'Low Failure Count
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).max = -1000000          'Maximum Measured Value
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).min = 1000000           'Minimum Measured Value
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).sigma = 0               'Sum of Measured Values
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).sigma2 = 0              'Sum of Measured Squared Values
        gudtScanStats(lintChanNum).mechHystPt(lintMechHysteresisPt).n = 0                   'Total Number of Values
    Next lintMechHysteresisPt
    
    '2.8dANM \/\/
    'Kickdown Start Location
    gudtScanStats(lintChanNum).KDStart.failCount.high = 0    'High Count
    gudtScanStats(lintChanNum).KDStart.failCount.low = 0     'Low Count
    gudtScanStats(lintChanNum).KDStart.max = -1000000        'Max Value
    gudtScanStats(lintChanNum).KDStart.min = 1000000         'Min Value
    gudtScanStats(lintChanNum).KDStart.sigma = 0             'Sum of Values
    gudtScanStats(lintChanNum).KDStart.sigma2 = 0            'Sum of Values Squared
    gudtScanStats(lintChanNum).KDStart.n = 0                 'Total Number of Values

    'Kickdown Force Span
    gudtScanStats(lintChanNum).KDSpan.failCount.high = 0     'High Count
    gudtScanStats(lintChanNum).KDSpan.failCount.low = 0      'Low Count
    gudtScanStats(lintChanNum).KDSpan.max = -1000000         'Max Value
    gudtScanStats(lintChanNum).KDSpan.min = 1000000          'Min Value
    gudtScanStats(lintChanNum).KDSpan.sigma = 0              'Sum of Values
    gudtScanStats(lintChanNum).KDSpan.sigma2 = 0             'Sum of Values Squared
    gudtScanStats(lintChanNum).KDSpan.n = 0                  'Total Number of Values

    'Kickdown Peak Location
    gudtScanStats(lintChanNum).KDPeak.failCount.high = 0     'High Count
    gudtScanStats(lintChanNum).KDPeak.failCount.low = 0      'Low Count
    gudtScanStats(lintChanNum).KDPeak.max = -1000000         'Max Value
    gudtScanStats(lintChanNum).KDPeak.min = 1000000          'Min Value
    gudtScanStats(lintChanNum).KDPeak.sigma = 0              'Sum of Values
    gudtScanStats(lintChanNum).KDPeak.sigma2 = 0             'Sum of Values Squared
    gudtScanStats(lintChanNum).KDPeak.n = 0                  'Total Number of Values

    'Kickdown Peak Force
    gudtScanStats(lintChanNum).KDPeakForce.failCount.high = 0   'High Count
    gudtScanStats(lintChanNum).KDPeakForce.failCount.low = 0    'Low Count
    gudtScanStats(lintChanNum).KDPeakForce.max = -1000000       'Max Value
    gudtScanStats(lintChanNum).KDPeakForce.min = 1000000        'Min Value
    gudtScanStats(lintChanNum).KDPeakForce.sigma = 0            'Sum of Values
    gudtScanStats(lintChanNum).KDPeakForce.sigma2 = 0           'Sum of Values Squared
    gudtScanStats(lintChanNum).KDPeakForce.n = 0                'Total Number of Values

    'Kickdown End Location
    gudtScanStats(lintChanNum).KDStop.failCount.high = 0     'High Count
    gudtScanStats(lintChanNum).KDStop.failCount.low = 0      'Low Count
    gudtScanStats(lintChanNum).KDStop.max = -1000000         'Max Value
    gudtScanStats(lintChanNum).KDStop.min = 1000000          'Min Value
    gudtScanStats(lintChanNum).KDStop.sigma = 0              'Sum of Values
    gudtScanStats(lintChanNum).KDStop.sigma2 = 0             'Sum of Values Squared
    gudtScanStats(lintChanNum).KDStop.n = 0                  'Total Number of Values
    '2.8dANM /\/\
    
Next lintChanNum

'Programming Stats
For lintProgrammerNum = 1 To 2

    'Index (FullClose & FullOpen)
    For lintIndexNum = 1 To 2
        'Value
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).failCount.high = 0      'High Failure Count
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).failCount.low = 0       'Low Failure Count
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).max = -1000000          'Maximum Measured Value
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).min = 1000000           'Minimum Measured Value
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).sigma = 0               'Sum of Measured Values
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).sigma2 = 0              'Sum of Measured Squared Values
        gudtProgStats(lintProgrammerNum).indexVal(lintIndexNum).n = 0                   'Total Number of Values
        'Location
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).failCount.high = 0      'High Failure Count
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).failCount.low = 0       'Low Failure Count
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).max = -1000000          'Maximum Measured Value
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).min = 1000000           'Minimum Measured Value
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).sigma = 0               'Sum of Measured Values
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).sigma2 = 0              'Sum of Measured Squared Values
        gudtProgStats(lintProgrammerNum).indexLoc(lintIndexNum).n = 0                   'Total Number of Values
    Next lintIndexNum

    'Clamp Low Value & Code
    gudtProgStats(lintProgrammerNum).clampLow.failCount.high = 0        'High Failure Count
    gudtProgStats(lintProgrammerNum).clampLow.failCount.low = 0         'Low Failure Count
    gudtProgStats(lintProgrammerNum).clampLow.max = -1000000            'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).clampLow.min = 1000000             'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).clampLow.sigma = 0                 'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).clampLow.sigma2 = 0                'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).clampLow.n = 0                     'Total Number of Values
    gudtProgStats(lintProgrammerNum).clampLowCode.max = -1000000        'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).clampLowCode.min = 1000000         'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).clampLowCode.sigma = 0             'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).clampLowCode.sigma2 = 0            'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).clampLowCode.n = 0                 'Total Number of Values

    'Clamp High Value & Code
    gudtProgStats(lintProgrammerNum).clampHigh.failCount.high = 0       'High Failure Count
    gudtProgStats(lintProgrammerNum).clampHigh.failCount.low = 0        'Low Failure Count
    gudtProgStats(lintProgrammerNum).clampHigh.max = -1000000           'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).clampHigh.min = 1000000            'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).clampHigh.sigma = 0                'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).clampHigh.sigma2 = 0               'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).clampHigh.n = 0                    'Total Number of Values
    gudtProgStats(lintProgrammerNum).clampHighCode.max = -1000000       'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).clampHighCode.min = 1000000        'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).clampHighCode.sigma = 0            'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).clampHighCode.sigma2 = 0           'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).clampHighCode.n = 0                'Total Number of Values

    'Offset Code
    gudtProgStats(lintProgrammerNum).offsetCode.max = -1000000          'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).offsetCode.min = 1000000           'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).offsetCode.sigma = 0               'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).offsetCode.sigma2 = 0              'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).offsetCode.n = 0                   'Total Number of Values

    'Rough Gain Code
    gudtProgStats(lintProgrammerNum).roughGainCode.max = -1000000       'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).roughGainCode.min = 1000000        'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).roughGainCode.sigma = 0            'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).roughGainCode.sigma2 = 0           'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).roughGainCode.n = 0                'Total Number of Values

    'Fine Gain Code
    gudtProgStats(lintProgrammerNum).fineGainCode.max = -1000000        'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).fineGainCode.min = 1000000         'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).fineGainCode.sigma = 0             'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).fineGainCode.sigma2 = 0            'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).fineGainCode.n = 0                 'Total Number of Values

    'Offset Seed Code
    gudtProgStats(lintProgrammerNum).OffsetSeedCode.max = -1000000      'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).OffsetSeedCode.min = 1000000       'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma = 0           'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2 = 0          'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).OffsetSeedCode.n = 0               'Total Number of Values

    'Rough Gain Seed Code
    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max = -1000000   'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min = 1000000    'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma = 0        'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2 = 0       'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n = 0            'Total Number of Values

    'Fine Gain Seed Code
    gudtProgStats(lintProgrammerNum).FineGainSeedCode.max = -1000000    'Maximum Measured Value
    gudtProgStats(lintProgrammerNum).FineGainSeedCode.min = 1000000     'Minimum Measured Value
    gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma = 0         'Sum of Measured Values
    gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2 = 0        'Sum of Measured Squared Values
    gudtProgStats(lintProgrammerNum).FineGainSeedCode.n = 0             'Total Number of Values

    'MLX Code Failure Counts
    gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high = 0
    gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high = 0
    gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high = 0
    gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high = 0
    gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high = 0
    gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high = 0

Next lintProgrammerNum

'Scan Summary information
gudtScanSummary.totalUnits = 0            'Total Units Tested
gudtScanSummary.totalGood = 0             'Total Good Units
gudtScanSummary.totalReject = 0           'Total Rejected Units
gudtScanSummary.totalSevere = 0           'Total Severe Units
gudtScanSummary.totalNoTest = 0           'Total Scan Errors
gudtScanSummary.currentTotal = 0          'Last xxx Parts
gudtScanSummary.currentGood = 0           'Last xxx Parts Good

'Program Summary information
gudtProgSummary.totalUnits = 0            'Total Units Tested
gudtProgSummary.totalGood = 0             'Total Good Units
gudtProgSummary.totalReject = 0           'Total Rejected Units
gudtProgSummary.totalSevere = 0           'Total Severe Units
gudtProgSummary.totalNoTest = 0           'Total Scan Errors
gudtProgSummary.currentTotal = 0          'Last xxx Parts
gudtProgSummary.currentGood = 0           'Last xxx Parts Good

End Sub

Public Sub StatsUpdateScanCounts()
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

Dim lintChanNum As Integer
Dim lintFaultNum As Integer

'If there is an anomaly, increment the scan error counter
If (gintAnomaly <> 0) Then              'Count No Test Units
    gudtScanSummary.totalNoTest = gudtScanSummary.totalNoTest + 1
Else
    'Check for Severe failures first; they have priority
    If gblnSevere Then                  'Count Failure if one occurred
        'Increment the severe counter if the part was a severe failure
        gudtScanSummary.totalSevere = gudtScanSummary.totalSevere + 1
        'Loop through the channels on current part
        For lintChanNum = 0 To MAXCHANNUM
            If gintSevere(lintChanNum, HIGHINDEXPT1) Then
                gudtScanStats(lintChanNum).Index(1).failCount.high = gudtScanStats(lintChanNum).Index(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWINDEXPT1) Then
                gudtScanStats(lintChanNum).Index(1).failCount.low = gudtScanStats(lintChanNum).Index(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHOUTPUTATFORCEKNEE) Then
                gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high = gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWOUTPUTATFORCEKNEE) Then
                gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low = gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHINDEXPT2) Then
                gudtScanStats(lintChanNum).Index(2).failCount.high = gudtScanStats(lintChanNum).Index(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWINDEXPT2) Then
                gudtScanStats(lintChanNum).Index(2).failCount.low = gudtScanStats(lintChanNum).Index(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHINDEXPT3) Then
                gudtScanStats(lintChanNum).Index(3).failCount.high = gudtScanStats(lintChanNum).Index(3).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWINDEXPT3) Then
                gudtScanStats(lintChanNum).Index(3).failCount.low = gudtScanStats(lintChanNum).Index(3).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMAXOUTPUT) Then
                gudtScanStats(lintChanNum).maxOutput.failCount.high = gudtScanStats(lintChanNum).maxOutput.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMAXOUTPUT) Then
                gudtScanStats(lintChanNum).maxOutput.failCount.low = gudtScanStats(lintChanNum).maxOutput.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHSINGLEPOINTLIN) Then
                gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWSINGLEPOINTLIN) Then
                gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHABSLIN) Then '2.7ANM
                gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWABSLIN) Then  '2.7ANM
                gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHSLOPE) Then
                gudtScanStats(lintChanNum).slopeMax.failCount.high = gudtScanStats(lintChanNum).slopeMax.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWSLOPE) Then
                gudtScanStats(lintChanNum).slopeMin.failCount.low = gudtScanStats(lintChanNum).slopeMin.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFWDOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.high = gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFWDOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.low = gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHREVOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.high = gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWREVOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.low = gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFORCEKNEELOC) Then
                gudtScanStats(lintChanNum).forceKneeLoc.failCount.high = gudtScanStats(lintChanNum).forceKneeLoc.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFORCEKNEELOC) Then
                gudtScanStats(lintChanNum).forceKneeLoc.failCount.low = gudtScanStats(lintChanNum).forceKneeLoc.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFORCEKNEEFWDFORCE) Then
                gudtScanStats(lintChanNum).forceKneeForce.failCount.high = gudtScanStats(lintChanNum).forceKneeForce.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFORCEKNEEFWDFORCE) Then
                gudtScanStats(lintChanNum).forceKneeForce.failCount.low = gudtScanStats(lintChanNum).forceKneeForce.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFWDFORCEPT1) Then
                gudtScanStats(lintChanNum).fwdForcePt(1).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFWDFORCEPT1) Then
                gudtScanStats(lintChanNum).fwdForcePt(1).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFWDFORCEPT2) Then
                gudtScanStats(lintChanNum).fwdForcePt(2).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFWDFORCEPT2) Then
                gudtScanStats(lintChanNum).fwdForcePt(2).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHFWDFORCEPT3) Then
                gudtScanStats(lintChanNum).fwdForcePt(3).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(3).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWFWDFORCEPT3) Then
                gudtScanStats(lintChanNum).fwdForcePt(3).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(3).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHREVFORCEPT1) Then
                gudtScanStats(lintChanNum).revForcePt(1).failCount.high = gudtScanStats(lintChanNum).revForcePt(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWREVFORCEPT1) Then
                gudtScanStats(lintChanNum).revForcePt(1).failCount.low = gudtScanStats(lintChanNum).revForcePt(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHREVFORCEPT2) Then
                gudtScanStats(lintChanNum).revForcePt(2).failCount.high = gudtScanStats(lintChanNum).revForcePt(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWREVFORCEPT2) Then
                gudtScanStats(lintChanNum).revForcePt(2).failCount.low = gudtScanStats(lintChanNum).revForcePt(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHREVFORCEPT3) Then
                gudtScanStats(lintChanNum).revForcePt(3).failCount.high = gudtScanStats(lintChanNum).revForcePt(3).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWREVFORCEPT3) Then
                gudtScanStats(lintChanNum).revForcePt(3).failCount.low = gudtScanStats(lintChanNum).revForcePt(3).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHPEAKFORCE) Then
                gudtScanStats(lintChanNum).peakForce.failCount.high = gudtScanStats(lintChanNum).peakForce.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWPEAKFORCE) Then
                gudtScanStats(lintChanNum).peakForce.failCount.low = gudtScanStats(lintChanNum).peakForce.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMECHHYSTPT1) Then
                gudtScanStats(lintChanNum).mechHystPt(1).failCount.high = gudtScanStats(lintChanNum).mechHystPt(1).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMECHHYSTPT1) Then
                gudtScanStats(lintChanNum).mechHystPt(1).failCount.low = gudtScanStats(lintChanNum).mechHystPt(1).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMECHHYSTPT2) Then
                gudtScanStats(lintChanNum).mechHystPt(2).failCount.high = gudtScanStats(lintChanNum).mechHystPt(2).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMECHHYSTPT2) Then
                gudtScanStats(lintChanNum).mechHystPt(2).failCount.low = gudtScanStats(lintChanNum).mechHystPt(2).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMECHHYSTPT3) Then
                gudtScanStats(lintChanNum).mechHystPt(3).failCount.high = gudtScanStats(lintChanNum).mechHystPt(3).failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMECHHYSTPT3) Then
                gudtScanStats(lintChanNum).mechHystPt(3).failCount.low = gudtScanStats(lintChanNum).mechHystPt(3).failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMLXI) Then '2.5ANM
                gudtScanStats(lintChanNum).mlxCurrent.failCount.high = gudtScanStats(lintChanNum).mlxCurrent.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMLXI) Then
                gudtScanStats(lintChanNum).mlxCurrent.failCount.low = gudtScanStats(lintChanNum).mlxCurrent.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHPEDALATREST) Then '2.8ANM
                gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.high = gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWPEDALATREST) Then
                gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.low = gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHMLXI2) Then '2.8cANM
                gudtScanStats(lintChanNum).mlxWCurrent.failCount.high = gudtScanStats(lintChanNum).mlxWCurrent.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWMLXI2) Then
                gudtScanStats(lintChanNum).mlxWCurrent.failCount.low = gudtScanStats(lintChanNum).mlxWCurrent.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHKDSTART) Then '2.8dANM
                gudtScanStats(lintChanNum).KDStart.failCount.high = gudtScanStats(lintChanNum).KDStart.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWKDSTART) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDStart.failCount.low = gudtScanStats(lintChanNum).KDStart.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHKDSTOP) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDStop.failCount.high = gudtScanStats(lintChanNum).KDStop.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWKDSTOP) Then   '2.8dANM
                gudtScanStats(lintChanNum).KDStop.failCount.low = gudtScanStats(lintChanNum).KDStop.failCount.low + 1
                Exit For
            ElseIf gintSevere(lintChanNum, HIGHKDSPAN) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDSpan.failCount.high = gudtScanStats(lintChanNum).KDSpan.failCount.high + 1
                Exit For
            ElseIf gintSevere(lintChanNum, LOWKDSPAN) Then   '2.8dANM
                gudtScanStats(lintChanNum).KDSpan.failCount.low = gudtScanStats(lintChanNum).KDSpan.failCount.low + 1
                Exit For
            End If
        Next lintChanNum
    ElseIf gblnScanFailure Then             'Count Failure if one occurred
        'Loop through the channels on current part
        For lintChanNum = 0 To MAXCHANNUM
            'Increment the failure count (prioritized)
            If gintFailure(lintChanNum, HIGHINDEXPT1) Then
                gudtScanStats(lintChanNum).Index(1).failCount.high = gudtScanStats(lintChanNum).Index(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWINDEXPT1) Then
                gudtScanStats(lintChanNum).Index(1).failCount.low = gudtScanStats(lintChanNum).Index(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHOUTPUTATFORCEKNEE) Then
                gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high = gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWOUTPUTATFORCEKNEE) Then
                gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low = gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHINDEXPT2) Then
                gudtScanStats(lintChanNum).Index(2).failCount.high = gudtScanStats(lintChanNum).Index(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWINDEXPT2) Then
                gudtScanStats(lintChanNum).Index(2).failCount.low = gudtScanStats(lintChanNum).Index(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHINDEXPT3) Then
                gudtScanStats(lintChanNum).Index(3).failCount.high = gudtScanStats(lintChanNum).Index(3).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWINDEXPT3) Then
                gudtScanStats(lintChanNum).Index(3).failCount.low = gudtScanStats(lintChanNum).Index(3).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMAXOUTPUT) Then
                gudtScanStats(lintChanNum).maxOutput.failCount.high = gudtScanStats(lintChanNum).maxOutput.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMAXOUTPUT) Then
                gudtScanStats(lintChanNum).maxOutput.failCount.low = gudtScanStats(lintChanNum).maxOutput.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHSINGLEPOINTLIN) Then
                gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWSINGLEPOINTLIN) Then
                gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low = gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHABSLIN) Then '2.7ANM
                gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWABSLIN) Then  '2.7ANM
                gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low = gudtScanStats(lintChanNum).linDevPerTol(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHSLOPE) Then
                gudtScanStats(lintChanNum).slopeMax.failCount.high = gudtScanStats(lintChanNum).slopeMax.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWSLOPE) Then
                gudtScanStats(lintChanNum).slopeMin.failCount.low = gudtScanStats(lintChanNum).slopeMin.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFCHYS) Then
                gudtScanStats(lintChanNum).FullCloseHys.failCount.high = gudtScanStats(lintChanNum).FullCloseHys.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFCHYS) Then
                gudtScanStats(lintChanNum).FullCloseHys.failCount.low = gudtScanStats(lintChanNum).FullCloseHys.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFWDOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.high = gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFWDOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.low = gudtScanStats(lintChanNum).outputCorPerTol(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHREVOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.high = gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWREVOUTPUTCOR) Then
                gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.low = gudtScanStats(lintChanNum).outputCorPerTol(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFORCEKNEELOC) Then
                gudtScanStats(lintChanNum).forceKneeLoc.failCount.high = gudtScanStats(lintChanNum).forceKneeLoc.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFORCEKNEELOC) Then
                gudtScanStats(lintChanNum).forceKneeLoc.failCount.low = gudtScanStats(lintChanNum).forceKneeLoc.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFORCEKNEEFWDFORCE) Then
                gudtScanStats(lintChanNum).forceKneeForce.failCount.high = gudtScanStats(lintChanNum).forceKneeForce.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFORCEKNEEFWDFORCE) Then
                gudtScanStats(lintChanNum).forceKneeForce.failCount.low = gudtScanStats(lintChanNum).forceKneeForce.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFWDFORCEPT1) Then
                gudtScanStats(lintChanNum).fwdForcePt(1).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFWDFORCEPT1) Then
                gudtScanStats(lintChanNum).fwdForcePt(1).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFWDFORCEPT2) Then
                gudtScanStats(lintChanNum).fwdForcePt(2).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFWDFORCEPT2) Then
                gudtScanStats(lintChanNum).fwdForcePt(2).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHFWDFORCEPT3) Then
                gudtScanStats(lintChanNum).fwdForcePt(3).failCount.high = gudtScanStats(lintChanNum).fwdForcePt(3).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWFWDFORCEPT3) Then
                gudtScanStats(lintChanNum).fwdForcePt(3).failCount.low = gudtScanStats(lintChanNum).fwdForcePt(3).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHREVFORCEPT1) Then
                gudtScanStats(lintChanNum).revForcePt(1).failCount.high = gudtScanStats(lintChanNum).revForcePt(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWREVFORCEPT1) Then
                gudtScanStats(lintChanNum).revForcePt(1).failCount.low = gudtScanStats(lintChanNum).revForcePt(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHREVFORCEPT2) Then
                gudtScanStats(lintChanNum).revForcePt(2).failCount.high = gudtScanStats(lintChanNum).revForcePt(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWREVFORCEPT2) Then
                gudtScanStats(lintChanNum).revForcePt(2).failCount.low = gudtScanStats(lintChanNum).revForcePt(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHREVFORCEPT3) Then
                gudtScanStats(lintChanNum).revForcePt(3).failCount.high = gudtScanStats(lintChanNum).revForcePt(3).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWREVFORCEPT3) Then
                gudtScanStats(lintChanNum).revForcePt(3).failCount.low = gudtScanStats(lintChanNum).revForcePt(3).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHPEAKFORCE) Then
                gudtScanStats(lintChanNum).peakForce.failCount.high = gudtScanStats(lintChanNum).peakForce.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWPEAKFORCE) Then
                gudtScanStats(lintChanNum).peakForce.failCount.low = gudtScanStats(lintChanNum).peakForce.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMECHHYSTPT1) Then
                gudtScanStats(lintChanNum).mechHystPt(1).failCount.high = gudtScanStats(lintChanNum).mechHystPt(1).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMECHHYSTPT1) Then
                gudtScanStats(lintChanNum).mechHystPt(1).failCount.low = gudtScanStats(lintChanNum).mechHystPt(1).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMECHHYSTPT2) Then
                gudtScanStats(lintChanNum).mechHystPt(2).failCount.high = gudtScanStats(lintChanNum).mechHystPt(2).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMECHHYSTPT2) Then
                gudtScanStats(lintChanNum).mechHystPt(2).failCount.low = gudtScanStats(lintChanNum).mechHystPt(2).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMECHHYSTPT3) Then
                gudtScanStats(lintChanNum).mechHystPt(3).failCount.high = gudtScanStats(lintChanNum).mechHystPt(3).failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMECHHYSTPT3) Then
                gudtScanStats(lintChanNum).mechHystPt(3).failCount.low = gudtScanStats(lintChanNum).mechHystPt(3).failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMLXI) Then '2.5ANM
                gudtScanStats(lintChanNum).mlxCurrent.failCount.high = gudtScanStats(lintChanNum).mlxCurrent.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMLXI) Then
                gudtScanStats(lintChanNum).mlxCurrent.failCount.low = gudtScanStats(lintChanNum).mlxCurrent.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHPEDALATREST) Then '2.8ANM
                gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.high = gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWPEDALATREST) Then
                gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.low = gudtScanStats(lintChanNum).pedalAtRestLoc.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHMLXI2) Then '2.8cANM
                gudtScanStats(lintChanNum).mlxWCurrent.failCount.high = gudtScanStats(lintChanNum).mlxWCurrent.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWMLXI2) Then
                gudtScanStats(lintChanNum).mlxWCurrent.failCount.low = gudtScanStats(lintChanNum).mlxWCurrent.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHKDSTART) Then '2.8dANM
                gudtScanStats(lintChanNum).KDStart.failCount.high = gudtScanStats(lintChanNum).KDStart.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWKDSTART) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDStart.failCount.low = gudtScanStats(lintChanNum).KDStart.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHKDSTOP) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDStop.failCount.high = gudtScanStats(lintChanNum).KDStop.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWKDSTOP) Then   '2.8dANM
                gudtScanStats(lintChanNum).KDStop.failCount.low = gudtScanStats(lintChanNum).KDStop.failCount.low + 1
                Exit For
            ElseIf gintFailure(lintChanNum, HIGHKDSPAN) Then  '2.8dANM
                gudtScanStats(lintChanNum).KDSpan.failCount.high = gudtScanStats(lintChanNum).KDSpan.failCount.high + 1
                Exit For
            ElseIf gintFailure(lintChanNum, LOWKDSPAN) Then   '2.8dANM
                gudtScanStats(lintChanNum).KDSpan.failCount.low = gudtScanStats(lintChanNum).KDSpan.failCount.low + 1
                Exit For
            End If
        Next lintChanNum
    Else
        gudtScanSummary.currentGood = gudtScanSummary.currentGood + 1   'XXX parts good
        gudtScanSummary.totalGood = gudtScanSummary.totalGood + 1       'Total parts good
    End If
End If
gudtScanSummary.currentTotal = gudtScanSummary.currentTotal + 1         'XXX part total
gudtScanSummary.totalUnits = gudtScanSummary.totalUnits + 1             'Total part total

End Sub

Public Sub StatsUpdateScanSums()
'
'   PURPOSE: To update the statistical sum information.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer

'Don't update stats sums for severe failures
If (gblnSevere = False) Then
    
    'Loop through the channels on current part
    For lintChanNum = 0 To MAXCHANNUM
        If Not gblnForceOnly Then              '1.6ANM added if block
            'Index 1 (FullClose Output)
            If gudtReading(lintChanNum).Index(1).Value > gudtScanStats(lintChanNum).Index(1).max Then
                gudtScanStats(lintChanNum).Index(1).max = gudtReading(lintChanNum).Index(1).Value
            End If
            If gudtReading(lintChanNum).Index(1).Value < gudtScanStats(lintChanNum).Index(1).min Then
                gudtScanStats(lintChanNum).Index(1).min = gudtReading(lintChanNum).Index(1).Value
            End If
            gudtScanStats(lintChanNum).Index(1).sigma = gudtScanStats(lintChanNum).Index(1).sigma + gudtReading(lintChanNum).Index(1).Value
            gudtScanStats(lintChanNum).Index(1).sigma2 = gudtScanStats(lintChanNum).Index(1).sigma2 + gudtReading(lintChanNum).Index(1).Value ^ 2
            gudtScanStats(lintChanNum).Index(1).n = gudtScanStats(lintChanNum).Index(1).n + 1
    
            'Output At Force Knee Location
            If gudtReading(lintChanNum).outputAtForceKnee > gudtScanStats(lintChanNum).outputAtForceKnee.max Then
                gudtScanStats(lintChanNum).outputAtForceKnee.max = gudtReading(lintChanNum).outputAtForceKnee
            End If
            If gudtReading(lintChanNum).outputAtForceKnee < gudtScanStats(lintChanNum).outputAtForceKnee.min Then
                gudtScanStats(lintChanNum).outputAtForceKnee.min = gudtReading(lintChanNum).outputAtForceKnee
            End If
            gudtScanStats(lintChanNum).outputAtForceKnee.sigma = gudtScanStats(lintChanNum).outputAtForceKnee.sigma + gudtReading(lintChanNum).outputAtForceKnee
            gudtScanStats(lintChanNum).outputAtForceKnee.sigma2 = gudtScanStats(lintChanNum).outputAtForceKnee.sigma2 + gudtReading(lintChanNum).outputAtForceKnee ^ 2
            gudtScanStats(lintChanNum).outputAtForceKnee.n = gudtScanStats(lintChanNum).outputAtForceKnee.n + 1
    
            'Index 2 (Midpoint Output)
            If gudtReading(lintChanNum).Index(2).Value > gudtScanStats(lintChanNum).Index(2).max Then
                gudtScanStats(lintChanNum).Index(2).max = gudtReading(lintChanNum).Index(2).Value
            End If
            If gudtReading(lintChanNum).Index(2).Value < gudtScanStats(lintChanNum).Index(2).min Then
                gudtScanStats(lintChanNum).Index(2).min = gudtReading(lintChanNum).Index(2).Value
            End If
            gudtScanStats(lintChanNum).Index(2).sigma = gudtScanStats(lintChanNum).Index(2).sigma + gudtReading(lintChanNum).Index(2).Value
            gudtScanStats(lintChanNum).Index(2).sigma2 = gudtScanStats(lintChanNum).Index(2).sigma2 + gudtReading(lintChanNum).Index(2).Value ^ 2
            gudtScanStats(lintChanNum).Index(2).n = gudtScanStats(lintChanNum).Index(2).n + 1
    
            'Index 3 (FullOpen Output)
            If gudtReading(lintChanNum).Index(3).Value > gudtScanStats(lintChanNum).Index(3).max Then
                gudtScanStats(lintChanNum).Index(3).max = gudtReading(lintChanNum).Index(3).Value
            End If
            If gudtReading(lintChanNum).Index(3).Value < gudtScanStats(lintChanNum).Index(3).min Then
                gudtScanStats(lintChanNum).Index(3).min = gudtReading(lintChanNum).Index(3).Value
            End If
            gudtScanStats(lintChanNum).Index(3).sigma = gudtScanStats(lintChanNum).Index(3).sigma + gudtReading(lintChanNum).Index(3).Value
            gudtScanStats(lintChanNum).Index(3).sigma2 = gudtScanStats(lintChanNum).Index(3).sigma2 + gudtReading(lintChanNum).Index(3).Value ^ 2
            gudtScanStats(lintChanNum).Index(3).n = gudtScanStats(lintChanNum).Index(3).n + 1
    
            'Index 3 (FullOpen Location)
            'If gudtReading(lintChanNum).Index(3).location > gudtScanStats(lintChanNum).Index(3).max Then
            '    gudtScanStats(lintChanNum).Index(3).max = gudtReading(lintChanNum).Index(3).location
            'End If
            'If gudtReading(lintChanNum).Index(3).location < gudtScanStats(lintChanNum).Index(3).min Then
            '    gudtScanStats(lintChanNum).Index(3).min = gudtReading(lintChanNum).Index(3).location
            'End If
            'gudtScanStats(lintChanNum).Index(3).sigma = gudtScanStats(lintChanNum).Index(3).sigma + gudtReading(lintChanNum).Index(3).location
            'gudtScanStats(lintChanNum).Index(3).sigma2 = gudtScanStats(lintChanNum).Index(3).sigma2 + gudtReading(lintChanNum).Index(3).location ^ 2
            'gudtScanStats(lintChanNum).Index(3).n = gudtScanStats(lintChanNum).Index(3).n + 1
    
            'Maximum Output
            If gudtReading(lintChanNum).maxOutput.Value > gudtScanStats(lintChanNum).maxOutput.max Then
                gudtScanStats(lintChanNum).maxOutput.max = gudtReading(lintChanNum).maxOutput.Value
            End If
            If gudtReading(lintChanNum).maxOutput.Value < gudtScanStats(lintChanNum).maxOutput.min Then
                gudtScanStats(lintChanNum).maxOutput.min = gudtReading(lintChanNum).maxOutput.Value
            End If
            gudtScanStats(lintChanNum).maxOutput.sigma = gudtScanStats(lintChanNum).maxOutput.sigma + gudtReading(lintChanNum).maxOutput.Value
            gudtScanStats(lintChanNum).maxOutput.sigma2 = gudtScanStats(lintChanNum).maxOutput.sigma2 + gudtReading(lintChanNum).maxOutput.Value ^ 2
            gudtScanStats(lintChanNum).maxOutput.n = gudtScanStats(lintChanNum).maxOutput.n + 1
    
            'SinglePoint Linearity Deviation % of Tolerance
            If gudtExtreme(lintChanNum).linDevPerTol(1).Value > gudtScanStats(lintChanNum).linDevPerTol(1).max Then
                gudtScanStats(lintChanNum).linDevPerTol(1).max = gudtExtreme(lintChanNum).linDevPerTol(1).Value
            End If
            If gudtExtreme(lintChanNum).linDevPerTol(1).Value < gudtScanStats(lintChanNum).linDevPerTol(1).min Then
                gudtScanStats(lintChanNum).linDevPerTol(1).min = gudtExtreme(lintChanNum).linDevPerTol(1).Value
            End If
            gudtScanStats(lintChanNum).linDevPerTol(1).sigma = gudtScanStats(lintChanNum).linDevPerTol(1).sigma + gudtExtreme(lintChanNum).linDevPerTol(1).Value
            gudtScanStats(lintChanNum).linDevPerTol(1).sigma2 = gudtScanStats(lintChanNum).linDevPerTol(1).sigma2 + gudtExtreme(lintChanNum).linDevPerTol(1).Value ^ 2
            gudtScanStats(lintChanNum).linDevPerTol(1).n = gudtScanStats(lintChanNum).linDevPerTol(1).n + 1
    
            '2.7ANM \/\/
            'Absolute Linearity Deviation % of Tolerance
            If gudtExtreme(lintChanNum).linDevPerTol(2).Value > gudtScanStats(lintChanNum).linDevPerTol(2).max Then
                gudtScanStats(lintChanNum).linDevPerTol(2).max = gudtExtreme(lintChanNum).linDevPerTol(2).Value
            End If
            If gudtExtreme(lintChanNum).linDevPerTol(2).Value < gudtScanStats(lintChanNum).linDevPerTol(2).min Then
                gudtScanStats(lintChanNum).linDevPerTol(2).min = gudtExtreme(lintChanNum).linDevPerTol(2).Value
            End If
            gudtScanStats(lintChanNum).linDevPerTol(2).sigma = gudtScanStats(lintChanNum).linDevPerTol(2).sigma + gudtExtreme(lintChanNum).linDevPerTol(2).Value
            gudtScanStats(lintChanNum).linDevPerTol(2).sigma2 = gudtScanStats(lintChanNum).linDevPerTol(2).sigma2 + gudtExtreme(lintChanNum).linDevPerTol(2).Value ^ 2
            gudtScanStats(lintChanNum).linDevPerTol(2).n = gudtScanStats(lintChanNum).linDevPerTol(2).n + 1
            '2.7ANM /\/\
    
            'Slope Deviation Max
            If gudtExtreme(lintChanNum).slope.high.Value > gudtScanStats(lintChanNum).slopeMax.max Then
                gudtScanStats(lintChanNum).slopeMax.max = gudtExtreme(lintChanNum).slope.high.Value
            End If
            If gudtExtreme(lintChanNum).slope.high.Value < gudtScanStats(lintChanNum).slopeMax.min Then
                gudtScanStats(lintChanNum).slopeMax.min = gudtExtreme(lintChanNum).slope.high.Value
            End If
            gudtScanStats(lintChanNum).slopeMax.sigma = gudtScanStats(lintChanNum).slopeMax.sigma + gudtExtreme(lintChanNum).slope.high.Value
            gudtScanStats(lintChanNum).slopeMax.sigma2 = gudtScanStats(lintChanNum).slopeMax.sigma2 + gudtExtreme(lintChanNum).slope.high.Value ^ 2
            gudtScanStats(lintChanNum).slopeMax.n = gudtScanStats(lintChanNum).slopeMax.n + 1
    
            'Slope Deviation Min
            If gudtExtreme(lintChanNum).slope.low.Value > gudtScanStats(lintChanNum).slopeMin.max Then
                gudtScanStats(lintChanNum).slopeMin.max = gudtExtreme(lintChanNum).slope.low.Value
            End If
            If gudtExtreme(lintChanNum).slope.low.Value < gudtScanStats(lintChanNum).slopeMin.min Then
                gudtScanStats(lintChanNum).slopeMin.min = gudtExtreme(lintChanNum).slope.low.Value
            End If
            gudtScanStats(lintChanNum).slopeMin.sigma = gudtScanStats(lintChanNum).slopeMin.sigma + gudtExtreme(lintChanNum).slope.low.Value
            gudtScanStats(lintChanNum).slopeMin.sigma2 = gudtScanStats(lintChanNum).slopeMin.sigma2 + gudtExtreme(lintChanNum).slope.low.Value ^ 2
            gudtScanStats(lintChanNum).slopeMin.n = gudtScanStats(lintChanNum).slopeMin.n + 1
    
            'Full-Close Hysteresis '2.5ANM fixed
            If gudtReading(lintChanNum).FullCloseHys.Value > gudtScanStats(lintChanNum).FullCloseHys.max Then
                gudtScanStats(lintChanNum).FullCloseHys.max = gudtReading(lintChanNum).FullCloseHys.Value
            End If
            If gudtReading(lintChanNum).FullCloseHys.Value < gudtScanStats(lintChanNum).FullCloseHys.min Then
                gudtScanStats(lintChanNum).FullCloseHys.min = gudtReading(lintChanNum).FullCloseHys.Value
            End If
            gudtScanStats(lintChanNum).FullCloseHys.sigma = gudtScanStats(lintChanNum).FullCloseHys.sigma + gudtReading(lintChanNum).FullCloseHys.Value
            gudtScanStats(lintChanNum).FullCloseHys.sigma2 = gudtScanStats(lintChanNum).FullCloseHys.sigma2 + gudtReading(lintChanNum).FullCloseHys.Value ^ 2
            gudtScanStats(lintChanNum).FullCloseHys.n = gudtScanStats(lintChanNum).FullCloseHys.n + 1
            
            'MLX Current '2.5ANM
            If gudtReading(lintChanNum).mlxCurrent > gudtScanStats(lintChanNum).mlxCurrent.max Then
                gudtScanStats(lintChanNum).mlxCurrent.max = gudtReading(lintChanNum).mlxCurrent
            End If
            If gudtReading(lintChanNum).mlxCurrent < gudtScanStats(lintChanNum).mlxCurrent.min Then
                gudtScanStats(lintChanNum).mlxCurrent.min = gudtReading(lintChanNum).mlxCurrent
            End If
            gudtScanStats(lintChanNum).mlxCurrent.sigma = gudtScanStats(lintChanNum).mlxCurrent.sigma + gudtReading(lintChanNum).mlxCurrent
            gudtScanStats(lintChanNum).mlxCurrent.sigma2 = gudtScanStats(lintChanNum).mlxCurrent.sigma2 + gudtReading(lintChanNum).mlxCurrent ^ 2
            gudtScanStats(lintChanNum).mlxCurrent.n = gudtScanStats(lintChanNum).mlxCurrent.n + 1
            
            'MLX WOT Current '2.8cANM
            If gudtReading(lintChanNum).mlxWCurrent > gudtScanStats(lintChanNum).mlxWCurrent.max Then
                gudtScanStats(lintChanNum).mlxWCurrent.max = gudtReading(lintChanNum).mlxWCurrent
            End If
            If gudtReading(lintChanNum).mlxWCurrent < gudtScanStats(lintChanNum).mlxWCurrent.min Then
                gudtScanStats(lintChanNum).mlxWCurrent.min = gudtReading(lintChanNum).mlxWCurrent
            End If
            gudtScanStats(lintChanNum).mlxWCurrent.sigma = gudtScanStats(lintChanNum).mlxWCurrent.sigma + gudtReading(lintChanNum).mlxWCurrent
            gudtScanStats(lintChanNum).mlxWCurrent.sigma2 = gudtScanStats(lintChanNum).mlxWCurrent.sigma2 + gudtReading(lintChanNum).mlxWCurrent ^ 2
            gudtScanStats(lintChanNum).mlxWCurrent.n = gudtScanStats(lintChanNum).mlxWCurrent.n + 1
        End If
        
        'Correlation, Force & Kickdown: only 1st channel of each part
        If (lintChanNum = CHAN0) Then
            If Not gblnForceOnly Then              '1.6ANM added if block
                'Forward Output Correlation % of Tolerance
                If gudtExtreme(lintChanNum).outputCorPerTol(1).Value > gudtScanStats(lintChanNum).outputCorPerTol(1).max Then
                    gudtScanStats(lintChanNum).outputCorPerTol(1).max = gudtExtreme(lintChanNum).outputCorPerTol(1).Value
                End If
                    If gudtExtreme(lintChanNum).outputCorPerTol(1).Value < gudtScanStats(lintChanNum).outputCorPerTol(1).min Then
                    gudtScanStats(lintChanNum).outputCorPerTol(1).min = gudtExtreme(lintChanNum).outputCorPerTol(1).Value
                End If
                gudtScanStats(lintChanNum).outputCorPerTol(1).sigma = gudtScanStats(lintChanNum).outputCorPerTol(1).sigma + gudtExtreme(lintChanNum).outputCorPerTol(1).Value
                gudtScanStats(lintChanNum).outputCorPerTol(1).sigma2 = gudtScanStats(lintChanNum).outputCorPerTol(1).sigma2 + gudtExtreme(lintChanNum).outputCorPerTol(1).Value ^ 2
                gudtScanStats(lintChanNum).outputCorPerTol(1).n = gudtScanStats(lintChanNum).outputCorPerTol(1).n + 1
    
                'Reverse Output Correlation % of Tolerance
                If gudtExtreme(lintChanNum).outputCorPerTol(2).Value > gudtScanStats(lintChanNum).outputCorPerTol(2).max Then
                    gudtScanStats(lintChanNum).outputCorPerTol(2).max = gudtExtreme(lintChanNum).outputCorPerTol(2).Value
                End If
                    If gudtExtreme(lintChanNum).outputCorPerTol(2).Value < gudtScanStats(lintChanNum).outputCorPerTol(2).min Then
                    gudtScanStats(lintChanNum).outputCorPerTol(2).min = gudtExtreme(lintChanNum).outputCorPerTol(2).Value
                End If
                gudtScanStats(lintChanNum).outputCorPerTol(2).sigma = gudtScanStats(lintChanNum).outputCorPerTol(2).sigma + gudtExtreme(lintChanNum).outputCorPerTol(2).Value
                gudtScanStats(lintChanNum).outputCorPerTol(2).sigma2 = gudtScanStats(lintChanNum).outputCorPerTol(2).sigma2 + gudtExtreme(lintChanNum).outputCorPerTol(2).Value ^ 2
                gudtScanStats(lintChanNum).outputCorPerTol(2).n = gudtScanStats(lintChanNum).outputCorPerTol(2).n + 1
            End If
            
            'Pedal-At-Rest Location
            If gudtReading(lintChanNum).pedalAtRestLoc > gudtScanStats(lintChanNum).pedalAtRestLoc.max Then
                gudtScanStats(lintChanNum).pedalAtRestLoc.max = gudtReading(lintChanNum).pedalAtRestLoc
            End If
            If gudtReading(lintChanNum).pedalAtRestLoc < gudtScanStats(lintChanNum).pedalAtRestLoc.min Then
                gudtScanStats(lintChanNum).pedalAtRestLoc.min = gudtReading(lintChanNum).pedalAtRestLoc
            End If
            gudtScanStats(lintChanNum).pedalAtRestLoc.sigma = gudtScanStats(lintChanNum).pedalAtRestLoc.sigma + gudtReading(lintChanNum).pedalAtRestLoc
            gudtScanStats(lintChanNum).pedalAtRestLoc.sigma2 = gudtScanStats(lintChanNum).pedalAtRestLoc.sigma2 + gudtReading(lintChanNum).pedalAtRestLoc ^ 2
            gudtScanStats(lintChanNum).pedalAtRestLoc.n = gudtScanStats(lintChanNum).pedalAtRestLoc.n + 1

            'Force Knee Location
            If gudtReading(lintChanNum).forceKnee.location > gudtScanStats(lintChanNum).forceKneeLoc.max Then
                gudtScanStats(lintChanNum).forceKneeLoc.max = gudtReading(lintChanNum).forceKnee.location
            End If
            If gudtReading(lintChanNum).forceKnee.location < gudtScanStats(lintChanNum).forceKneeLoc.min Then
                gudtScanStats(lintChanNum).forceKneeLoc.min = gudtReading(lintChanNum).forceKnee.location
            End If
            gudtScanStats(lintChanNum).forceKneeLoc.sigma = gudtScanStats(lintChanNum).forceKneeLoc.sigma + gudtReading(lintChanNum).forceKnee.location
            gudtScanStats(lintChanNum).forceKneeLoc.sigma2 = gudtScanStats(lintChanNum).forceKneeLoc.sigma2 + gudtReading(lintChanNum).forceKnee.location ^ 2
            gudtScanStats(lintChanNum).forceKneeLoc.n = gudtScanStats(lintChanNum).forceKneeLoc.n + 1

            'Forward Force at Force Knee Location
            If gudtReading(lintChanNum).forceKnee.Value > gudtScanStats(lintChanNum).forceKneeForce.max Then
                gudtScanStats(lintChanNum).forceKneeForce.max = gudtReading(lintChanNum).forceKnee.Value
            End If
            If gudtReading(lintChanNum).forceKnee.Value < gudtScanStats(lintChanNum).forceKneeForce.min Then
                gudtScanStats(lintChanNum).forceKneeForce.min = gudtReading(lintChanNum).forceKnee.Value
            End If
            gudtScanStats(lintChanNum).forceKneeForce.sigma = gudtScanStats(lintChanNum).forceKneeForce.sigma + gudtReading(lintChanNum).forceKnee.Value
            gudtScanStats(lintChanNum).forceKneeForce.sigma2 = gudtScanStats(lintChanNum).forceKneeForce.sigma2 + gudtReading(lintChanNum).forceKnee.Value ^ 2
            gudtScanStats(lintChanNum).forceKneeForce.n = gudtScanStats(lintChanNum).forceKneeForce.n + 1

            'Forward Force Point 1
            If gudtReading(lintChanNum).fwdForcePt(1).Value > gudtScanStats(lintChanNum).fwdForcePt(1).max Then
                gudtScanStats(lintChanNum).fwdForcePt(1).max = gudtReading(lintChanNum).fwdForcePt(1).Value
            End If
            If gudtReading(lintChanNum).fwdForcePt(1).Value < gudtScanStats(lintChanNum).fwdForcePt(1).min Then
                gudtScanStats(lintChanNum).fwdForcePt(1).min = gudtReading(lintChanNum).fwdForcePt(1).Value
            End If
            gudtScanStats(lintChanNum).fwdForcePt(1).sigma = gudtScanStats(lintChanNum).fwdForcePt(1).sigma + gudtReading(lintChanNum).fwdForcePt(1).Value
            gudtScanStats(lintChanNum).fwdForcePt(1).sigma2 = gudtScanStats(lintChanNum).fwdForcePt(1).sigma2 + gudtReading(lintChanNum).fwdForcePt(1).Value ^ 2
            gudtScanStats(lintChanNum).fwdForcePt(1).n = gudtScanStats(lintChanNum).fwdForcePt(1).n + 1

            'Forward Force Point 2
            If gudtReading(lintChanNum).fwdForcePt(2).Value > gudtScanStats(lintChanNum).fwdForcePt(2).max Then
                gudtScanStats(lintChanNum).fwdForcePt(2).max = gudtReading(lintChanNum).fwdForcePt(2).Value
            End If
            If gudtReading(lintChanNum).fwdForcePt(2).Value < gudtScanStats(lintChanNum).fwdForcePt(2).min Then
                gudtScanStats(lintChanNum).fwdForcePt(2).min = gudtReading(lintChanNum).fwdForcePt(2).Value
            End If
            gudtScanStats(lintChanNum).fwdForcePt(2).sigma = gudtScanStats(lintChanNum).fwdForcePt(2).sigma + gudtReading(lintChanNum).fwdForcePt(2).Value
            gudtScanStats(lintChanNum).fwdForcePt(2).sigma2 = gudtScanStats(lintChanNum).fwdForcePt(2).sigma2 + gudtReading(lintChanNum).fwdForcePt(2).Value ^ 2
            gudtScanStats(lintChanNum).fwdForcePt(2).n = gudtScanStats(lintChanNum).fwdForcePt(2).n + 1

            'Forward Force Point 3
            If gudtReading(lintChanNum).fwdForcePt(3).Value > gudtScanStats(lintChanNum).fwdForcePt(3).max Then
                gudtScanStats(lintChanNum).fwdForcePt(3).max = gudtReading(lintChanNum).fwdForcePt(3).Value
            End If
            If gudtReading(lintChanNum).fwdForcePt(3).Value < gudtScanStats(lintChanNum).fwdForcePt(3).min Then
                gudtScanStats(lintChanNum).fwdForcePt(3).min = gudtReading(lintChanNum).fwdForcePt(3).Value
            End If
            gudtScanStats(lintChanNum).fwdForcePt(3).sigma = gudtScanStats(lintChanNum).fwdForcePt(3).sigma + gudtReading(lintChanNum).fwdForcePt(3).Value
            gudtScanStats(lintChanNum).fwdForcePt(3).sigma2 = gudtScanStats(lintChanNum).fwdForcePt(3).sigma2 + gudtReading(lintChanNum).fwdForcePt(3).Value ^ 2
            gudtScanStats(lintChanNum).fwdForcePt(3).n = gudtScanStats(lintChanNum).fwdForcePt(3).n + 1

            'Reverse Force Point 1
            If gudtReading(lintChanNum).revForcePt(1).Value > gudtScanStats(lintChanNum).revForcePt(1).max Then
                gudtScanStats(lintChanNum).revForcePt(1).max = gudtReading(lintChanNum).revForcePt(1).Value
            End If
            If gudtReading(lintChanNum).revForcePt(1).Value < gudtScanStats(lintChanNum).revForcePt(1).min Then
                gudtScanStats(lintChanNum).revForcePt(1).min = gudtReading(lintChanNum).revForcePt(1).Value
            End If
            gudtScanStats(lintChanNum).revForcePt(1).sigma = gudtScanStats(lintChanNum).revForcePt(1).sigma + gudtReading(lintChanNum).revForcePt(1).Value
            gudtScanStats(lintChanNum).revForcePt(1).sigma2 = gudtScanStats(lintChanNum).revForcePt(1).sigma2 + gudtReading(lintChanNum).revForcePt(1).Value ^ 2
            gudtScanStats(lintChanNum).revForcePt(1).n = gudtScanStats(lintChanNum).revForcePt(1).n + 1

            'Reverse Force Point 2
            If gudtReading(lintChanNum).revForcePt(2).Value > gudtScanStats(lintChanNum).revForcePt(2).max Then
                gudtScanStats(lintChanNum).revForcePt(2).max = gudtReading(lintChanNum).revForcePt(2).Value
            End If
            If gudtReading(lintChanNum).revForcePt(2).Value < gudtScanStats(lintChanNum).revForcePt(2).min Then
                gudtScanStats(lintChanNum).revForcePt(2).min = gudtReading(lintChanNum).revForcePt(2).Value
            End If
            gudtScanStats(lintChanNum).revForcePt(2).sigma = gudtScanStats(lintChanNum).revForcePt(2).sigma + gudtReading(lintChanNum).revForcePt(2).Value
            gudtScanStats(lintChanNum).revForcePt(2).sigma2 = gudtScanStats(lintChanNum).revForcePt(2).sigma2 + gudtReading(lintChanNum).revForcePt(2).Value ^ 2
            gudtScanStats(lintChanNum).revForcePt(2).n = gudtScanStats(lintChanNum).revForcePt(2).n + 1

            'Reverse Force Point 3
            If gudtReading(lintChanNum).revForcePt(3).Value > gudtScanStats(lintChanNum).revForcePt(3).max Then
                gudtScanStats(lintChanNum).revForcePt(3).max = gudtReading(lintChanNum).revForcePt(3).Value
            End If
            If gudtReading(lintChanNum).revForcePt(3).Value < gudtScanStats(lintChanNum).revForcePt(3).min Then
                gudtScanStats(lintChanNum).revForcePt(3).min = gudtReading(lintChanNum).revForcePt(3).Value
            End If
            gudtScanStats(lintChanNum).revForcePt(3).sigma = gudtScanStats(lintChanNum).revForcePt(3).sigma + gudtReading(lintChanNum).revForcePt(3).Value
            gudtScanStats(lintChanNum).revForcePt(3).sigma2 = gudtScanStats(lintChanNum).revForcePt(3).sigma2 + gudtReading(lintChanNum).revForcePt(3).Value ^ 2
            gudtScanStats(lintChanNum).revForcePt(3).n = gudtScanStats(lintChanNum).revForcePt(3).n + 1

            'Peak Force
            If gudtReading(lintChanNum).peakForce > gudtScanStats(lintChanNum).peakForce.max Then
                gudtScanStats(lintChanNum).peakForce.max = gudtReading(lintChanNum).peakForce
            End If
            If gudtReading(lintChanNum).peakForce < gudtScanStats(lintChanNum).peakForce.min Then
                gudtScanStats(lintChanNum).peakForce.min = gudtReading(lintChanNum).peakForce
            End If
            gudtScanStats(lintChanNum).peakForce.sigma = gudtScanStats(lintChanNum).peakForce.sigma + gudtReading(lintChanNum).peakForce
            gudtScanStats(lintChanNum).peakForce.sigma2 = gudtScanStats(lintChanNum).peakForce.sigma2 + gudtReading(lintChanNum).peakForce ^ 2
            gudtScanStats(lintChanNum).peakForce.n = gudtScanStats(lintChanNum).peakForce.n + 1

            'Mechanical Hysteresis Point 1
            If gudtReading(lintChanNum).mechHystPt(1).Value > gudtScanStats(lintChanNum).mechHystPt(1).max Then
                gudtScanStats(lintChanNum).mechHystPt(1).max = gudtReading(lintChanNum).mechHystPt(1).Value
            End If
            If gudtReading(lintChanNum).mechHystPt(1).Value < gudtScanStats(lintChanNum).mechHystPt(1).min Then
                gudtScanStats(lintChanNum).mechHystPt(1).min = gudtReading(lintChanNum).mechHystPt(1).Value
            End If
            gudtScanStats(lintChanNum).mechHystPt(1).sigma = gudtScanStats(lintChanNum).mechHystPt(1).sigma + gudtReading(lintChanNum).mechHystPt(1).Value
            gudtScanStats(lintChanNum).mechHystPt(1).sigma2 = gudtScanStats(lintChanNum).mechHystPt(1).sigma2 + gudtReading(lintChanNum).mechHystPt(1).Value ^ 2
            gudtScanStats(lintChanNum).mechHystPt(1).n = gudtScanStats(lintChanNum).mechHystPt(1).n + 1

            'Mechanical Hysteresis Point 2
            If gudtReading(lintChanNum).mechHystPt(2).Value > gudtScanStats(lintChanNum).mechHystPt(2).max Then
                gudtScanStats(lintChanNum).mechHystPt(2).max = gudtReading(lintChanNum).mechHystPt(2).Value
            End If
            If gudtReading(lintChanNum).mechHystPt(2).Value < gudtScanStats(lintChanNum).mechHystPt(2).min Then
                gudtScanStats(lintChanNum).mechHystPt(2).min = gudtReading(lintChanNum).mechHystPt(2).Value
            End If
            gudtScanStats(lintChanNum).mechHystPt(2).sigma = gudtScanStats(lintChanNum).mechHystPt(2).sigma + gudtReading(lintChanNum).mechHystPt(2).Value
            gudtScanStats(lintChanNum).mechHystPt(2).sigma2 = gudtScanStats(lintChanNum).mechHystPt(2).sigma2 + gudtReading(lintChanNum).mechHystPt(2).Value ^ 2
            gudtScanStats(lintChanNum).mechHystPt(2).n = gudtScanStats(lintChanNum).mechHystPt(2).n + 1

            'Mechanical Hysteresis Point 3
            If gudtReading(lintChanNum).mechHystPt(3).Value > gudtScanStats(lintChanNum).mechHystPt(3).max Then
                gudtScanStats(lintChanNum).mechHystPt(3).max = gudtReading(lintChanNum).mechHystPt(3).Value
            End If
            If gudtReading(lintChanNum).mechHystPt(3).Value < gudtScanStats(lintChanNum).mechHystPt(3).min Then
                gudtScanStats(lintChanNum).mechHystPt(3).min = gudtReading(lintChanNum).mechHystPt(3).Value
            End If
            gudtScanStats(lintChanNum).mechHystPt(3).sigma = gudtScanStats(lintChanNum).mechHystPt(3).sigma + gudtReading(lintChanNum).mechHystPt(3).Value
            gudtScanStats(lintChanNum).mechHystPt(3).sigma2 = gudtScanStats(lintChanNum).mechHystPt(3).sigma2 + gudtReading(lintChanNum).mechHystPt(3).Value ^ 2
            gudtScanStats(lintChanNum).mechHystPt(3).n = gudtScanStats(lintChanNum).mechHystPt(3).n + 1

            '2.8dANM \/\/
            If gblnKD Then
                'Kickdown Start Location
                If gudtReading(lintChanNum).KDStart.location > gudtScanStats(lintChanNum).KDStart.max Then
                    gudtScanStats(lintChanNum).KDStart.max = gudtReading(lintChanNum).KDStart.location
                End If
                If gudtReading(lintChanNum).KDStart.location < gudtScanStats(lintChanNum).KDStart.min Then
                    gudtScanStats(lintChanNum).KDStart.min = gudtReading(lintChanNum).KDStart.location
                End If
                gudtScanStats(lintChanNum).KDStart.sigma = gudtScanStats(lintChanNum).KDStart.sigma + gudtReading(lintChanNum).KDStart.location
                gudtScanStats(lintChanNum).KDStart.sigma2 = gudtScanStats(lintChanNum).KDStart.sigma2 + gudtReading(lintChanNum).KDStart.location ^ 2
                gudtScanStats(lintChanNum).KDStart.n = gudtScanStats(lintChanNum).KDStart.n + 1

                'Kickdown Force Span
                If gudtReading(lintChanNum).KDSpan > gudtScanStats(lintChanNum).KDSpan.max Then
                    gudtScanStats(lintChanNum).KDSpan.max = gudtReading(lintChanNum).KDSpan
                End If
                If gudtReading(lintChanNum).KDSpan < gudtScanStats(lintChanNum).KDSpan.min Then
                    gudtScanStats(lintChanNum).KDSpan.min = gudtReading(lintChanNum).KDSpan
                End If
                gudtScanStats(lintChanNum).KDSpan.sigma = gudtScanStats(lintChanNum).KDSpan.sigma + gudtReading(lintChanNum).KDSpan
                gudtScanStats(lintChanNum).KDSpan.sigma2 = gudtScanStats(lintChanNum).KDSpan.sigma2 + gudtReading(lintChanNum).KDSpan ^ 2
                gudtScanStats(lintChanNum).KDSpan.n = gudtScanStats(lintChanNum).KDSpan.n + 1

                'Kickdown Peak Location
                If gudtReading(lintChanNum).KDPeak.location > gudtScanStats(lintChanNum).KDPeak.max Then
                    gudtScanStats(lintChanNum).KDPeak.max = gudtReading(lintChanNum).KDPeak.location
                End If
                If gudtReading(lintChanNum).KDPeak.location < gudtScanStats(lintChanNum).KDPeak.min Then
                    gudtScanStats(lintChanNum).KDPeak.min = gudtReading(lintChanNum).KDPeak.location
                End If
                gudtScanStats(lintChanNum).KDPeak.sigma = gudtScanStats(lintChanNum).KDPeak.sigma + gudtReading(lintChanNum).KDPeak.location
                gudtScanStats(lintChanNum).KDPeak.sigma2 = gudtScanStats(lintChanNum).KDPeak.sigma2 + gudtReading(lintChanNum).KDPeak.location ^ 2
                gudtScanStats(lintChanNum).KDPeak.n = gudtScanStats(lintChanNum).KDPeak.n + 1

                'Kickdown Peak Force
                If gudtReading(lintChanNum).KDPeak.Value > gudtScanStats(lintChanNum).KDPeakForce.max Then
                    gudtScanStats(lintChanNum).KDPeakForce.max = gudtReading(lintChanNum).KDPeak.Value
                End If
                If gudtReading(lintChanNum).KDPeak.Value < gudtScanStats(lintChanNum).KDPeakForce.min Then
                    gudtScanStats(lintChanNum).KDPeakForce.min = gudtReading(lintChanNum).KDPeak.Value
                End If
                gudtScanStats(lintChanNum).KDPeakForce.sigma = gudtScanStats(lintChanNum).KDPeakForce.sigma + gudtReading(lintChanNum).KDPeak.Value
                gudtScanStats(lintChanNum).KDPeakForce.sigma2 = gudtScanStats(lintChanNum).KDPeakForce.sigma2 + gudtReading(lintChanNum).KDPeak.Value ^ 2
                gudtScanStats(lintChanNum).KDPeakForce.n = gudtScanStats(lintChanNum).KDPeakForce.n + 1

                'Kickdown End Location
                If gudtReading(lintChanNum).KDStop.location > gudtScanStats(lintChanNum).KDStop.max Then
                    gudtScanStats(lintChanNum).KDStop.max = gudtReading(lintChanNum).KDStop.location
                End If
                If gudtReading(lintChanNum).KDStop.location < gudtScanStats(lintChanNum).KDStop.min Then
                    gudtScanStats(lintChanNum).KDStop.min = gudtReading(lintChanNum).KDStop.location
                End If
                gudtScanStats(lintChanNum).KDStop.sigma = gudtScanStats(lintChanNum).KDStop.sigma + gudtReading(lintChanNum).KDStop.location
                gudtScanStats(lintChanNum).KDStop.sigma2 = gudtScanStats(lintChanNum).KDStop.sigma2 + gudtReading(lintChanNum).KDStop.location ^ 2
                gudtScanStats(lintChanNum).KDStop.n = gudtScanStats(lintChanNum).KDStop.n + 1
            End If
            '2.8dANM /\/\

        End If
    Next lintChanNum
End If

End Sub

Public Sub StatsLoad()
'
'   PURPOSE:   To input production statistics into the program from
'              a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none
'2.0ANM new sub

Dim lintFileNum As Integer
Dim lintChanNum As Integer
Dim lintProgrammerNum As Integer
Dim lstrOperator As String
Dim lstrTemperature As String
Dim lstrComment As String
Dim lstrSeries As String

On Error GoTo StatsLoad_Err

'Clear statistics before starting a new lot or resuming an old lot
Call StatsClear

frmMain.MousePointer = vbHourglass

'Get a file number
lintFileNum = FreeFile

'Check to see if file exists if not exit sub
If gfsoFileSystemObject.FileExists(STATFILEPATH & gstrLotName & STATEXT) Then
    Open STATFILEPATH & gstrLotName & STATEXT For Input As #lintFileNum
Else
    frmMain.MousePointer = vbNormal
    Exit Sub
End If

'** General Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries
'Display to the form
frmMain.ctrSetupInfo1.Operator = lstrOperator
frmMain.ctrSetupInfo1.Temperature = lstrTemperature
frmMain.ctrSetupInfo1.Comment = lstrComment
frmMain.ctrSetupInfo1.Series = lstrSeries

'*** Scan Information ***
'Loop through all channels
For lintChanNum = CHAN0 To MAXCHANNUM
    'Index #1 (FullClose Output)
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
    'Output at Force Knee
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low, gudtScanStats(lintChanNum).outputAtForceKnee.max, gudtScanStats(lintChanNum).outputAtForceKnee.min, gudtScanStats(lintChanNum).outputAtForceKnee.sigma, gudtScanStats(lintChanNum).outputAtForceKnee.sigma2, gudtScanStats(lintChanNum).outputAtForceKnee.n
    'Index #2 (Midpoint Output)
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
    'Index #3 (FullOpen Output)
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(3).failCount.high, gudtScanStats(lintChanNum).Index(3).failCount.low, gudtScanStats(lintChanNum).Index(3).max, gudtScanStats(lintChanNum).Index(3).min, gudtScanStats(lintChanNum).Index(3).sigma, gudtScanStats(lintChanNum).Index(3).sigma2, gudtScanStats(lintChanNum).Index(3).n
    'Maximum Output
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
    'SinglePoint Linearity Deviation
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
    'Slope Max
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
    'Slope Min
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.max, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n '2.5ANM
    'Full-Close Hysteresis
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).FullCloseHys.failCount.high, gudtScanStats(lintChanNum).FullCloseHys.failCount.low, gudtScanStats(lintChanNum).FullCloseHys.max, gudtScanStats(lintChanNum).FullCloseHys.min, gudtScanStats(lintChanNum).FullCloseHys.sigma, gudtScanStats(lintChanNum).FullCloseHys.sigma2, gudtScanStats(lintChanNum).FullCloseHys.n '2.5ANM
Next lintChanNum
'Forward Output Correlation
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
'Reverse Output Correlation
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
'Pedal-At-Rest Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
'Force Knee Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).forceKneeLoc.failCount.high, gudtScanStats(CHAN0).forceKneeLoc.failCount.low, gudtScanStats(CHAN0).forceKneeLoc.max, gudtScanStats(CHAN0).forceKneeLoc.min, gudtScanStats(CHAN0).forceKneeLoc.sigma, gudtScanStats(CHAN0).forceKneeLoc.sigma2, gudtScanStats(CHAN0).forceKneeLoc.n
'Forward Force at Force Knee Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).forceKneeForce.failCount.high, gudtScanStats(CHAN0).forceKneeForce.failCount.low, gudtScanStats(CHAN0).forceKneeForce.max, gudtScanStats(CHAN0).forceKneeForce.min, gudtScanStats(CHAN0).forceKneeForce.sigma, gudtScanStats(CHAN0).forceKneeForce.sigma2, gudtScanStats(CHAN0).forceKneeForce.n
'Forward Force Point 1
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(1).failCount.high, gudtScanStats(CHAN0).fwdForcePt(1).failCount.low, gudtScanStats(CHAN0).fwdForcePt(1).max, gudtScanStats(CHAN0).fwdForcePt(1).min, gudtScanStats(CHAN0).fwdForcePt(1).sigma, gudtScanStats(CHAN0).fwdForcePt(1).sigma2, gudtScanStats(CHAN0).fwdForcePt(1).n
'Forward Force Point 2
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(2).failCount.high, gudtScanStats(CHAN0).fwdForcePt(2).failCount.low, gudtScanStats(CHAN0).fwdForcePt(2).max, gudtScanStats(CHAN0).fwdForcePt(2).min, gudtScanStats(CHAN0).fwdForcePt(2).sigma, gudtScanStats(CHAN0).fwdForcePt(2).sigma2, gudtScanStats(CHAN0).fwdForcePt(2).n
'Forward Force Point 3
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(3).failCount.high, gudtScanStats(CHAN0).fwdForcePt(3).failCount.low, gudtScanStats(CHAN0).fwdForcePt(3).max, gudtScanStats(CHAN0).fwdForcePt(3).min, gudtScanStats(CHAN0).fwdForcePt(3).sigma, gudtScanStats(CHAN0).fwdForcePt(3).sigma2, gudtScanStats(CHAN0).fwdForcePt(3).n
'Reverse Force Point 1
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(1).failCount.high, gudtScanStats(CHAN0).revForcePt(1).failCount.low, gudtScanStats(CHAN0).revForcePt(1).max, gudtScanStats(CHAN0).revForcePt(1).min, gudtScanStats(CHAN0).revForcePt(1).sigma, gudtScanStats(CHAN0).revForcePt(1).sigma2, gudtScanStats(CHAN0).revForcePt(1).n
'Reverse Force Point 2
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(2).failCount.high, gudtScanStats(CHAN0).revForcePt(2).failCount.low, gudtScanStats(CHAN0).revForcePt(2).max, gudtScanStats(CHAN0).revForcePt(2).min, gudtScanStats(CHAN0).revForcePt(2).sigma, gudtScanStats(CHAN0).revForcePt(2).sigma2, gudtScanStats(CHAN0).revForcePt(2).n
'Reverse Force Point 3
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(3).failCount.high, gudtScanStats(CHAN0).revForcePt(3).failCount.low, gudtScanStats(CHAN0).revForcePt(3).max, gudtScanStats(CHAN0).revForcePt(3).min, gudtScanStats(CHAN0).revForcePt(3).sigma, gudtScanStats(CHAN0).revForcePt(3).sigma2, gudtScanStats(CHAN0).revForcePt(3).n
'Peak Force
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
'Mechanical Hysteresis Point 1
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(1).failCount.high, gudtScanStats(CHAN0).mechHystPt(1).failCount.low, gudtScanStats(CHAN0).mechHystPt(1).max, gudtScanStats(CHAN0).mechHystPt(1).min, gudtScanStats(CHAN0).mechHystPt(1).sigma, gudtScanStats(CHAN0).mechHystPt(1).sigma2, gudtScanStats(CHAN0).mechHystPt(1).n
'Mechanical Hysteresis Point 2
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(2).failCount.high, gudtScanStats(CHAN0).mechHystPt(2).failCount.low, gudtScanStats(CHAN0).mechHystPt(2).max, gudtScanStats(CHAN0).mechHystPt(2).min, gudtScanStats(CHAN0).mechHystPt(2).sigma, gudtScanStats(CHAN0).mechHystPt(2).sigma2, gudtScanStats(CHAN0).mechHystPt(2).n
'Mechanical Hysteresis Point 3
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(3).failCount.high, gudtScanStats(CHAN0).mechHystPt(3).failCount.low, gudtScanStats(CHAN0).mechHystPt(3).max, gudtScanStats(CHAN0).mechHystPt(3).min, gudtScanStats(CHAN0).mechHystPt(3).sigma, gudtScanStats(CHAN0).mechHystPt(3).sigma2, gudtScanStats(CHAN0).mechHystPt(3).n

'*** Programming Information ***
'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Index #1 (FullClose) Values
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
    'Index #1 (FullClose) Locations
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
    'Index #2 (FullOpen) Values
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
    'Index #2 (FullOpen) Locations
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
    'Clamp Low
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
    'Clamp High
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
    'Offset Code
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
    'Rough Gain Code
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
    'Fine Gain Code
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
    'Clamp Low Code
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
    'Clamp High Code
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
    'MLX Code Failure Counts
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
Next lintProgrammerNum

'*** Programming Summary Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal

'*** Scanning Summary Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal

'MLX Current '2.5ANM
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mlxCurrent.failCount.high, gudtScanStats(CHAN0).mlxCurrent.failCount.low, gudtScanStats(CHAN0).mlxCurrent.max, gudtScanStats(CHAN0).mlxCurrent.min, gudtScanStats(CHAN0).mlxCurrent.sigma, gudtScanStats(CHAN0).mlxCurrent.sigma2, gudtScanStats(CHAN0).mlxCurrent.n
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN1).mlxCurrent.failCount.high, gudtScanStats(CHAN1).mlxCurrent.failCount.low, gudtScanStats(CHAN1).mlxCurrent.max, gudtScanStats(CHAN1).mlxCurrent.min, gudtScanStats(CHAN1).mlxCurrent.sigma, gudtScanStats(CHAN1).mlxCurrent.sigma2, gudtScanStats(CHAN1).mlxCurrent.n

'Abs Linearity Deviation '2.7ANM
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).linDevPerTol(2).failCount.high, gudtScanStats(CHAN0).linDevPerTol(2).failCount.low, gudtScanStats(CHAN0).linDevPerTol(2).max, gudtScanStats(CHAN0).linDevPerTol(2).min, gudtScanStats(CHAN0).linDevPerTol(2).sigma, gudtScanStats(CHAN0).linDevPerTol(2).sigma2, gudtScanStats(CHAN0).linDevPerTol(2).n
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN1).linDevPerTol(2).failCount.high, gudtScanStats(CHAN1).linDevPerTol(2).failCount.low, gudtScanStats(CHAN1).linDevPerTol(2).max, gudtScanStats(CHAN1).linDevPerTol(2).min, gudtScanStats(CHAN1).linDevPerTol(2).sigma, gudtScanStats(CHAN1).linDevPerTol(2).sigma2, gudtScanStats(CHAN1).linDevPerTol(2).n

'MLX WOT Current '2.8cANM
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mlxWCurrent.failCount.high, gudtScanStats(CHAN0).mlxWCurrent.failCount.low, gudtScanStats(CHAN0).mlxWCurrent.max, gudtScanStats(CHAN0).mlxWCurrent.min, gudtScanStats(CHAN0).mlxWCurrent.sigma, gudtScanStats(CHAN0).mlxWCurrent.sigma2, gudtScanStats(CHAN0).mlxWCurrent.n
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN1).mlxWCurrent.failCount.high, gudtScanStats(CHAN1).mlxWCurrent.failCount.low, gudtScanStats(CHAN1).mlxWCurrent.max, gudtScanStats(CHAN1).mlxWCurrent.min, gudtScanStats(CHAN1).mlxWCurrent.sigma, gudtScanStats(CHAN1).mlxWCurrent.sigma2, gudtScanStats(CHAN1).mlxWCurrent.n

'2.8dANM \/\/
'Kickdown Start Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).KDStart.failCount.high, gudtScanStats(CHAN0).KDStart.failCount.low, gudtScanStats(CHAN0).KDStart.max, gudtScanStats(CHAN0).KDStart.min, gudtScanStats(CHAN0).KDStart.sigma, gudtScanStats(CHAN0).KDStart.sigma2, gudtScanStats(CHAN0).KDStart.n
'Kickdown Force Span
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).KDSpan.failCount.high, gudtScanStats(CHAN0).KDSpan.failCount.low, gudtScanStats(CHAN0).KDSpan.max, gudtScanStats(CHAN0).KDSpan.min, gudtScanStats(CHAN0).KDSpan.sigma, gudtScanStats(CHAN0).KDSpan.sigma2, gudtScanStats(CHAN0).KDSpan.n
'Kickdown Peak Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).KDPeak.failCount.high, gudtScanStats(CHAN0).KDPeak.failCount.low, gudtScanStats(CHAN0).KDPeak.max, gudtScanStats(CHAN0).KDPeak.min, gudtScanStats(CHAN0).KDPeak.sigma, gudtScanStats(CHAN0).KDPeak.sigma2, gudtScanStats(CHAN0).KDPeak.n
'Kickdown Peak Force
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).KDPeakForce.failCount.high, gudtScanStats(CHAN0).KDPeakForce.failCount.low, gudtScanStats(CHAN0).KDPeakForce.max, gudtScanStats(CHAN0).KDPeakForce.min, gudtScanStats(CHAN0).KDPeakForce.sigma, gudtScanStats(CHAN0).KDPeakForce.sigma2, gudtScanStats(CHAN0).KDPeakForce.n
'Kickdown End Location
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).KDStop.failCount.high, gudtScanStats(CHAN0).KDStop.failCount.low, gudtScanStats(CHAN0).KDStop.max, gudtScanStats(CHAN0).KDStop.min, gudtScanStats(CHAN0).KDStop.sigma, gudtScanStats(CHAN0).KDStop.sigma2, gudtScanStats(CHAN0).KDStop.n
'2.8dANM /\/\

'Close the file
Close #lintFileNum
frmMain.MousePointer = vbNormal

Exit Sub
StatsLoad_Err:

    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Lot File!"

End Sub

Public Sub StatsSave()
'
'   PURPOSE:   To write production statistics to a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none
'2.0ANM new sub

Dim lintFileNum As Integer
Dim lintChanNum As Integer
Dim lintProgrammerNum As Integer
Dim lstrOperator As String
Dim lstrTemperature As String
Dim lstrComment As String
Dim lstrSeries As String

On Error GoTo StatsSave_Err

'Get a file number
lintFileNum = FreeFile
'Open the stats file
Open STATFILEPATH & gstrLotName & STATEXT For Output As #lintFileNum

'Take data from the form
lstrOperator = frmMain.ctrSetupInfo1.Operator
lstrTemperature = frmMain.ctrSetupInfo1.Temperature
lstrComment = frmMain.ctrSetupInfo1.Comment
lstrSeries = frmMain.ctrSetupInfo1.Series

'*** General Information ***
Write #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries

'*** Scan Information ***
'Loop through all channels
For lintChanNum = CHAN0 To MAXCHANNUM
    'Index #1 (FullClose Output)
    Write #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
    'Output at Force Knee
    Write #lintFileNum, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low, gudtScanStats(lintChanNum).outputAtForceKnee.max, gudtScanStats(lintChanNum).outputAtForceKnee.min, gudtScanStats(lintChanNum).outputAtForceKnee.sigma, gudtScanStats(lintChanNum).outputAtForceKnee.sigma2, gudtScanStats(lintChanNum).outputAtForceKnee.n
    'Index #2 (Midpoint Output)
    Write #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
    'Index #3 (FullOpen Output)
    Write #lintFileNum, gudtScanStats(lintChanNum).Index(3).failCount.high, gudtScanStats(lintChanNum).Index(3).failCount.low, gudtScanStats(lintChanNum).Index(3).max, gudtScanStats(lintChanNum).Index(3).min, gudtScanStats(lintChanNum).Index(3).sigma, gudtScanStats(lintChanNum).Index(3).sigma2, gudtScanStats(lintChanNum).Index(3).n
    'Maximum Output
    Write #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
    'SinglePoint Linearity Deviation
    Write #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
    'Slope Max
    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
    'Slope Min
    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.max, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n '2.5ANM
    'Full-Close Hysteresis
    Write #lintFileNum, gudtScanStats(lintChanNum).FullCloseHys.failCount.high, gudtScanStats(lintChanNum).FullCloseHys.failCount.low, gudtScanStats(lintChanNum).FullCloseHys.max, gudtScanStats(lintChanNum).FullCloseHys.min, gudtScanStats(lintChanNum).FullCloseHys.sigma, gudtScanStats(lintChanNum).FullCloseHys.sigma2, gudtScanStats(lintChanNum).FullCloseHys.n '2.5ANM
Next lintChanNum
'Forward Output Correlation
Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
'Reverse Output Correlation
Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
'Pedal-At-Rest Location
Write #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
'Force Knee Location
Write #lintFileNum, gudtScanStats(CHAN0).forceKneeLoc.failCount.high, gudtScanStats(CHAN0).forceKneeLoc.failCount.low, gudtScanStats(CHAN0).forceKneeLoc.max, gudtScanStats(CHAN0).forceKneeLoc.min, gudtScanStats(CHAN0).forceKneeLoc.sigma, gudtScanStats(CHAN0).forceKneeLoc.sigma2, gudtScanStats(CHAN0).forceKneeLoc.n
'Forward Force at Force Knee Location
Write #lintFileNum, gudtScanStats(CHAN0).forceKneeForce.failCount.high, gudtScanStats(CHAN0).forceKneeForce.failCount.low, gudtScanStats(CHAN0).forceKneeForce.max, gudtScanStats(CHAN0).forceKneeForce.min, gudtScanStats(CHAN0).forceKneeForce.sigma, gudtScanStats(CHAN0).forceKneeForce.sigma2, gudtScanStats(CHAN0).forceKneeForce.n
'Forward Force Point 1
Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(1).failCount.high, gudtScanStats(CHAN0).fwdForcePt(1).failCount.low, gudtScanStats(CHAN0).fwdForcePt(1).max, gudtScanStats(CHAN0).fwdForcePt(1).min, gudtScanStats(CHAN0).fwdForcePt(1).sigma, gudtScanStats(CHAN0).fwdForcePt(1).sigma2, gudtScanStats(CHAN0).fwdForcePt(1).n
'Forward Force Point 2
Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(2).failCount.high, gudtScanStats(CHAN0).fwdForcePt(2).failCount.low, gudtScanStats(CHAN0).fwdForcePt(2).max, gudtScanStats(CHAN0).fwdForcePt(2).min, gudtScanStats(CHAN0).fwdForcePt(2).sigma, gudtScanStats(CHAN0).fwdForcePt(2).sigma2, gudtScanStats(CHAN0).fwdForcePt(2).n
'Forward Force Point 3
Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(3).failCount.high, gudtScanStats(CHAN0).fwdForcePt(3).failCount.low, gudtScanStats(CHAN0).fwdForcePt(3).max, gudtScanStats(CHAN0).fwdForcePt(3).min, gudtScanStats(CHAN0).fwdForcePt(3).sigma, gudtScanStats(CHAN0).fwdForcePt(3).sigma2, gudtScanStats(CHAN0).fwdForcePt(3).n
'Reverse Force Point 1
Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(1).failCount.high, gudtScanStats(CHAN0).revForcePt(1).failCount.low, gudtScanStats(CHAN0).revForcePt(1).max, gudtScanStats(CHAN0).revForcePt(1).min, gudtScanStats(CHAN0).revForcePt(1).sigma, gudtScanStats(CHAN0).revForcePt(1).sigma2, gudtScanStats(CHAN0).revForcePt(1).n
'Reverse Force Point 2
Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(2).failCount.high, gudtScanStats(CHAN0).revForcePt(2).failCount.low, gudtScanStats(CHAN0).revForcePt(2).max, gudtScanStats(CHAN0).revForcePt(2).min, gudtScanStats(CHAN0).revForcePt(2).sigma, gudtScanStats(CHAN0).revForcePt(2).sigma2, gudtScanStats(CHAN0).revForcePt(2).n
'Reverse Force Point 3
Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(3).failCount.high, gudtScanStats(CHAN0).revForcePt(3).failCount.low, gudtScanStats(CHAN0).revForcePt(3).max, gudtScanStats(CHAN0).revForcePt(3).min, gudtScanStats(CHAN0).revForcePt(3).sigma, gudtScanStats(CHAN0).revForcePt(3).sigma2, gudtScanStats(CHAN0).revForcePt(3).n
'Peak Force
Write #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
'Mechanical Hysteresis Point 1
Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(1).failCount.high, gudtScanStats(CHAN0).mechHystPt(1).failCount.low, gudtScanStats(CHAN0).mechHystPt(1).max, gudtScanStats(CHAN0).mechHystPt(1).min, gudtScanStats(CHAN0).mechHystPt(1).sigma, gudtScanStats(CHAN0).mechHystPt(1).sigma2, gudtScanStats(CHAN0).mechHystPt(1).n
'Mechanical Hysteresis Point 2
Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(2).failCount.high, gudtScanStats(CHAN0).mechHystPt(2).failCount.low, gudtScanStats(CHAN0).mechHystPt(2).max, gudtScanStats(CHAN0).mechHystPt(2).min, gudtScanStats(CHAN0).mechHystPt(2).sigma, gudtScanStats(CHAN0).mechHystPt(2).sigma2, gudtScanStats(CHAN0).mechHystPt(2).n
'Mechanical Hysteresis Point 3
Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(3).failCount.high, gudtScanStats(CHAN0).mechHystPt(3).failCount.low, gudtScanStats(CHAN0).mechHystPt(3).max, gudtScanStats(CHAN0).mechHystPt(3).min, gudtScanStats(CHAN0).mechHystPt(3).sigma, gudtScanStats(CHAN0).mechHystPt(3).sigma2, gudtScanStats(CHAN0).mechHystPt(3).n

'*** Programming Information ***
'Loop through both programmers
For lintProgrammerNum = 1 To 2
    'Index #1 (FullClose) Values
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
    'Index #1 (FullClose) Locations
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
    'Index #2 (FullOpen) Values
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
    'Index #2 (FullOpen) Locations
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
    'Clamp Low
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
    'Clamp High
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
    'Offset Code
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
    'Rough Gain Code
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
    'Fine Gain Code
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
    'Clamp Low Code
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
    'Clamp High Code
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
    'MLX Code Failure Counts
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
Next lintProgrammerNum

'*** Programming Summary Information ***
Write #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal

'*** Scanning Summary Information ***
Write #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal

'MLX Current '2.5ANM
Write #lintFileNum, gudtScanStats(CHAN0).mlxCurrent.failCount.high, gudtScanStats(CHAN0).mlxCurrent.failCount.low, gudtScanStats(CHAN0).mlxCurrent.max, gudtScanStats(CHAN0).mlxCurrent.min, gudtScanStats(CHAN0).mlxCurrent.sigma, gudtScanStats(CHAN0).mlxCurrent.sigma2, gudtScanStats(CHAN0).mlxCurrent.n
Write #lintFileNum, gudtScanStats(CHAN1).mlxCurrent.failCount.high, gudtScanStats(CHAN1).mlxCurrent.failCount.low, gudtScanStats(CHAN1).mlxCurrent.max, gudtScanStats(CHAN1).mlxCurrent.min, gudtScanStats(CHAN1).mlxCurrent.sigma, gudtScanStats(CHAN1).mlxCurrent.sigma2, gudtScanStats(CHAN1).mlxCurrent.n

'Abs Linearity Deviation '2.7ANM
Write #lintFileNum, gudtScanStats(CHAN0).linDevPerTol(2).failCount.high, gudtScanStats(CHAN0).linDevPerTol(2).failCount.low, gudtScanStats(CHAN0).linDevPerTol(2).max, gudtScanStats(CHAN0).linDevPerTol(2).min, gudtScanStats(CHAN0).linDevPerTol(2).sigma, gudtScanStats(CHAN0).linDevPerTol(2).sigma2, gudtScanStats(CHAN0).linDevPerTol(2).n
Write #lintFileNum, gudtScanStats(CHAN1).linDevPerTol(2).failCount.high, gudtScanStats(CHAN1).linDevPerTol(2).failCount.low, gudtScanStats(CHAN1).linDevPerTol(2).max, gudtScanStats(CHAN1).linDevPerTol(2).min, gudtScanStats(CHAN1).linDevPerTol(2).sigma, gudtScanStats(CHAN1).linDevPerTol(2).sigma2, gudtScanStats(CHAN1).linDevPerTol(2).n

'MLX WOT Current '2.8cANM
Write #lintFileNum, gudtScanStats(CHAN0).mlxWCurrent.failCount.high, gudtScanStats(CHAN0).mlxWCurrent.failCount.low, gudtScanStats(CHAN0).mlxWCurrent.max, gudtScanStats(CHAN0).mlxWCurrent.min, gudtScanStats(CHAN0).mlxWCurrent.sigma, gudtScanStats(CHAN0).mlxWCurrent.sigma2, gudtScanStats(CHAN0).mlxWCurrent.n
Write #lintFileNum, gudtScanStats(CHAN1).mlxWCurrent.failCount.high, gudtScanStats(CHAN1).mlxWCurrent.failCount.low, gudtScanStats(CHAN1).mlxWCurrent.max, gudtScanStats(CHAN1).mlxWCurrent.min, gudtScanStats(CHAN1).mlxWCurrent.sigma, gudtScanStats(CHAN1).mlxWCurrent.sigma2, gudtScanStats(CHAN1).mlxWCurrent.n

'2.8dANM \/\/
'Kickdown Start Location
Write #lintFileNum, gudtScanStats(CHAN0).KDStart.failCount.high, gudtScanStats(CHAN0).KDStart.failCount.low, gudtScanStats(CHAN0).KDStart.max, gudtScanStats(CHAN0).KDStart.min, gudtScanStats(CHAN0).KDStart.sigma, gudtScanStats(CHAN0).KDStart.sigma2, gudtScanStats(CHAN0).KDStart.n
'Kickdown Force Span
Write #lintFileNum, gudtScanStats(CHAN0).KDSpan.failCount.high, gudtScanStats(CHAN0).KDSpan.failCount.low, gudtScanStats(CHAN0).KDSpan.max, gudtScanStats(CHAN0).KDSpan.min, gudtScanStats(CHAN0).KDSpan.sigma, gudtScanStats(CHAN0).KDSpan.sigma2, gudtScanStats(CHAN0).KDSpan.n
'Kickdown Peak Location
Write #lintFileNum, gudtScanStats(CHAN0).KDPeak.failCount.high, gudtScanStats(CHAN0).KDPeak.failCount.low, gudtScanStats(CHAN0).KDPeak.max, gudtScanStats(CHAN0).KDPeak.min, gudtScanStats(CHAN0).KDPeak.sigma, gudtScanStats(CHAN0).KDPeak.sigma2, gudtScanStats(CHAN0).KDPeak.n
'Kickdown Peak Force
Write #lintFileNum, gudtScanStats(CHAN0).KDPeakForce.failCount.high, gudtScanStats(CHAN0).KDPeakForce.failCount.low, gudtScanStats(CHAN0).KDPeakForce.max, gudtScanStats(CHAN0).KDPeakForce.min, gudtScanStats(CHAN0).KDPeakForce.sigma, gudtScanStats(CHAN0).KDPeakForce.sigma2, gudtScanStats(CHAN0).KDPeakForce.n
'Kickdown End Location
Write #lintFileNum, gudtScanStats(CHAN0).KDStop.failCount.high, gudtScanStats(CHAN0).KDStop.failCount.low, gudtScanStats(CHAN0).KDStop.max, gudtScanStats(CHAN0).KDStop.min, gudtScanStats(CHAN0).KDStop.sigma, gudtScanStats(CHAN0).KDStop.sigma2, gudtScanStats(CHAN0).KDStop.n
'2.8dANM /\/\

'Close the stats file
Close #lintFileNum
Call frmMain.RefreshLotFileList         'Add new files to lot file list

Exit Sub
StatsSave_Err:

    MsgBox Err.Description, vbOKOnly, "Error Saving Data to Lot File!"

End Sub

Public Sub SaveScanResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
'2.5ANM added MLX Idd
'2.7ANM added Abs. Lin.
'2.8cANM added MLX WOT

Dim lintFileNum As Integer
Dim lstrFileName As String

'Make the results file name
lstrFileName = gstrLotName + " Scan Results" & DATAEXT
'Get a file
lintFileNum = FreeFile

'If file does not exist then add a header
If Not gfsoFileSystemObject.FileExists(PARTSCANDATAPATH + lstrFileName) Then
    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
    'Part S/N, Date Code, Date/Time, Software Revision, Parameter File Name
    Print #lintFileNum, _
        "Part Number,"; _
        "Date Code,"; _
        "Date/Time,"; _
        "S/W Revision,"; _
        "Parameter File Name,"; _
    'AP1
    Print #lintFileNum, _
        "Full Close Value AP1 [%],"; _
        "Midpoint Value AP1 [%],"; _
        "Full Open Value AP1 [%],"; _
        "Maximum Output AP1 [%],"; _
        "Maximum SngPt Lin Deviation % of Tolerance AP1 [% Tol],"; _
        "Maximum SngPt Lin Deviation AP1 [%],"; _
        "Minimum SngPt Lin Deviation AP1 [%],"; _
        "Maximum Abs. Lin. Deviation % of Tolerance AP1 [% Tol],"; _
        "Maximum Abs. Lin. Deviation AP1 [%],"; _
        "Minimum Abs. Lin. Deviation AP1 [%],"; _
        "Maximum Slope Deviation AP1 [ratio to Ideal Slope],"; _
        "Minimum Slope Deviation AP1 [ratio to Ideal Slope],"; _
        "Full Close Hysteresis AP1 [%],"; _
        "Peak Hysteresis AP1 [%],";
    'AP2
    Print #lintFileNum, _
        "Full Close Value AP2 [%],"; _
        "Midpoint Value AP2 [%],"; _
        "Full Open Value AP2 [%],"; _
        "Maximum Output AP2 [%],"; _
        "Maximum SngPt Lin Deviation % of Tolerance AP2 [% Tol],"; _
        "Maximum SngPt Lin Deviation AP2 [%],"; _
        "Minimum SngPt Lin Deviation AP2 [%],"; _
        "Maximum Abs. Lin. Deviation % of Tolerance AP2 [% Tol],"; _
        "Maximum Abs. Lin. Deviation AP2 [%],"; _
        "Minimum Abs. Lin. Deviation AP2 [%],"; _
        "Maximum Slope Deviation AP2 [ratio to Ideal Slope],"; _
        "Minimum Slope Deviation AP2 [ratio to Ideal Slope],"; _
        "Full Close Hysteresis AP2 [%],"; _
        "Peak Hysteresis AP2 [%],";
    'Correlation
    Print #lintFileNum, _
        "Maximum Forward Output Correlation % of Tolerance [% Tol],"; _
        "Maximum Forward Output Correlation [%],"; _
        "Minimum Forward Output Correlation [%],"; _
        "Maximum Reverse Output Correlation % of Tolerance [% Tol],"; _
        "Maximum Reverse Output Correlation [%],"; _
        "Minimum Reverse Output Correlation [%],";
    'Force
    Print #lintFileNum, _
        "Pedal at Rest Location [°],"; _
        "Forward Force at " & Format(Round(gudtTest(CHAN0).fwdForcePt(1).location, 2), "#0.00") & "° [N],"; _
        "Forward Force at " & Format(Round(gudtTest(CHAN0).fwdForcePt(2).location, 2), "#0.00") & "° [N],"; _
        "Forward Force at " & Format(Round(gudtTest(CHAN0).fwdForcePt(3).location, 2), "#0.00") & "° [N],"; _
        "Reverse Force at " & Format(Round(gudtTest(CHAN0).revForcePt(1).location, 2), "#0.00") & "° [N],"; _
        "Reverse Force at " & Format(Round(gudtTest(CHAN0).revForcePt(2).location, 2), "#0.00") & "° [N],"; _
        "Reverse Force at " & Format(Round(gudtTest(CHAN0).revForcePt(3).location, 2), "#0.00") & "° [N],"; _
        "Peak Force [N],"; _
        "Mechanical Hysteresis at " & Format(Round(gudtTest(CHAN0).fwdForcePt(1).location, 2), "#0.00") & "° [% of Forward Force],"; _
        "Mechanical Hysteresis at " & Format(Round(gudtTest(CHAN0).fwdForcePt(2).location, 2), "#0.00") & "° [% of Forward Force],"; _
        "Mechanical Hysteresis at " & Format(Round(gudtTest(CHAN0).fwdForcePt(3).location, 2), "#0.00") & "° [% of Forward Force],"; _
        "Supply Idle Current AP1 [mA],"; _
        "Supply Idle Current AP2 [mA],"; _
        "Supply WOT Current AP1 [mA],"; _
        "Supply WOT Current AP2 [mA],"; _
        "Kickdown Start Loc [°],"; "Kickdown Force Span [N],"; "Kickdown Peak Loc [°],"; "Kickdown Peak Force [N],"; "Kickdown End Loc [°],";
    'Part Status, Comment, Operator Initials, Temperature, and Series
    Print #lintFileNum, _
        "Status,"; _
        "Comment,"; _
        "Operator,"; _
        "Temperature,"; _
        "Series,"
Else
    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
End If

'Part S/N, Date Code, Date/Time, Software Revision, Parameter File Name
Print #lintFileNum, _
    gstrSerialNumber; ","; _
    gstrDateCode; ","; _
    DateTime.Now; ","; _
    App.Major & "." & App.Minor & "." & App.Revision; ","; _
    gudtMachine.parameterName; ","; _
'Output #1
Print #lintFileNum, _
    Format(Round(gudtReading(CHAN0).Index(1).Value, 3), "##0.000"); ","; _
    Format(Round(gudtReading(CHAN0).Index(2).Value, 3), "##0.000"); ","; _
    Format(Round(gudtReading(CHAN0).Index(3).Value, 3), "##0.000", 2); ","; _
    Format(Round(gudtReading(CHAN0).maxOutput.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).linDevPerTol(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).SinglePointLin.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).SinglePointLin.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).linDevPerTol(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).AbsLin.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).AbsLin.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).slope.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).slope.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).FullCloseHys.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).hysteresis.Value, 2), "##0.00"); ",";
'Output #2
Print #lintFileNum, _
    Format(Round(gudtReading(CHAN1).Index(1).Value, 3), "##0.000"); ","; _
    Format(Round(gudtReading(CHAN1).Index(2).Value, 3), "##0.000"); ","; _
    Format(Round(gudtReading(CHAN1).Index(3).Value, 3), "##0.000", 2); ","; _
    Format(Round(gudtReading(CHAN1).maxOutput.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).linDevPerTol(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).SinglePointLin.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).SinglePointLin.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).linDevPerTol(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).AbsLin.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).AbsLin.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).slope.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).slope.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN1).FullCloseHys.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN1).hysteresis.Value, 2), "##0.00"); ",";
'Correlation
Print #lintFileNum, _
    Format(Round(gudtExtreme(CHAN0).outputCorPerTol(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.low.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).outputCorPerTol(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).revOutputCor.high.Value, 2), "##0.00"); ","; _
    Format(Round(gudtExtreme(CHAN0).revOutputCor.low.Value, 2), "##0.00"); ",";
'Force
Print #lintFileNum, _
    Format(Round(gudtReading(CHAN0).pedalAtRestLoc, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).fwdForcePt(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).fwdForcePt(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).fwdForcePt(3).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).revForcePt(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).revForcePt(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).revForcePt(3).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).peakForce, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).mechHystPt(1).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).mechHystPt(2).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).mechHystPt(3).Value, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).mlxCurrent, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN1).mlxCurrent, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).mlxWCurrent, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN1).mlxWCurrent, 2), "##0.00"); ","; _
    Format(Round(gudtReading(CHAN0).KDStart.location, 2), "##0.00"); ","; Format(Round(gudtReading(CHAN0).KDSpan, 2), "##0.00"); ","; Format(Round(gudtReading(CHAN0).KDPeak.location, 2), "##0.00"); ","; Format(Round(gudtReading(CHAN0).KDPeak.Value, 2), "##0.00"); ","; Format(Round(gudtReading(CHAN0).KDStop.location, 2), "##0.00"); ",";
'Part Status, Comment, Operator Initials, Temperature, and Series
If gblnScanFailure Then
    Print #lintFileNum, "REJECT,";
Else
    Print #lintFileNum, "PASS,";
End If
Print #lintFileNum, _
    frmMain.ctrSetupInfo1.Comment; ","; _
    frmMain.ctrSetupInfo1.Operator; ","; _
    frmMain.ctrSetupInfo1.Temperature; ","; _
    frmMain.ctrSetupInfo1.Series
'Close the file
Close #lintFileNum

End Sub
