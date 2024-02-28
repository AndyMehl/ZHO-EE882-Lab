Attribute VB_Name = "AMAD705"
'*************AMAD705.BAS - Analysis and Management of Acquired Data*************
'
'   705 Series Specific AMAD, supplemental to Pedal.Bas and Series705.BAS.
'   This module should handle all 705 series production programmer/scanners,
'   test lab programmer/scanners, and database recall software.
'   The software is to be kept in the pedal software library, EE947.
'
'VER    DATE      BY   PURPOSE OF MODIFICATION                          TAG
'1.0  03/27/2006  ANM  First release per PR 11801-K.                    1.0ANM
'1.1  08/17/2006  ANM  Updates per SCN# MISC-100 (3521).                1.1ANM
'1.2  10/30/2006  ANM  Updates per SCN# MISC-101 (3636).                1.2ANM
'1.3  01/18/2007  ANM  Updates for TR 8501-E (705 Prod.) and            1.3ANM
'                      SCN# MISC-102 (3702).
'1.4  04/12/2007  SRC  Added error trap to show error message when      1.4SRC
'                      laser database on process PC is not available.
'1.5  06/20/2007  SRC  Updated SNDATAPATH for new name.                 1.5SRC
'1.6  02/19/2008  ANM  Updates per SCN#s 4066 & 4067.                   1.6ANM
'1.7  05/19/2008  ANM  Updates per SCN# 4124.                           1.7ANM
'1.8  06/05/2008  ANM  Updates per SCN$ 4167.                           1.8ANM
'1.9  09/17/2009  ANM  Updates per SCN# 4397.                           1.9ANM
'2.0  06/22/2010  ANM  Updates per SCN# 4585.                           2.0ANM
'

Option Explicit

Public Const DATABASEPATH = "D:\Data\705\Database\"        'Path to database files
Public Const DATABASENAME = "705_Database"                 'Base name of the database
Public Const DATABASEEXTENSION = ".MDB"                    'Extension for database files
Public Const RAWDATAPATH = "D:\Data\705\Database\Rawdata\" 'Path to raw data files
Public Const RAWDATAEXTENSION = ".SDR"                     'Extension for raw data files (Stored Data Results)
Public Const MAXDATABASESIZE = 700000000                   'Maximum size of database file in bytes

'1.5SRC '1.3ANM 'Path to SN Database (Production use only)
Public Const SNDATAPATH = "\\cn7007PD0006\database$\705_Laser_Database.mdb"

'*** Type Definitions ***
Type AMADType
    DB_ID                   As Long     'Database ID
    ForceCalID              As Long     'Force Calibration ID
    LotID                   As Long     'Lot ID
    MachineParametersID     As Long     'Machine Parameters ID
    Output1Parameters       As Long     'Output #1 Scan Parameters
    Output2Parameters       As Long     'Output #2 Scan Parameters
    ProgParametersID        As Long     'Programming Parameters ID
    ScanParametersID        As Long     'Test Specifications ID
    SerialNumberID          As Long     'Serial Number ID
    ScanResultsID           As Long     'Scan Results ID
End Type

Public gatDataBaseKey As AMADType

Private mcnLocalDatabase As New Connection
Private mstrDatabaseName As String
Private mblnConnectionActive As Boolean
Private mcnLocalDatabase2 As New Connection
Private mstrDatabaseName2 As String
Private mblnConnectionActive2 As Boolean

Public Sub ActivateDatabase()
'
'   PURPOSE: To activate the database
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lrstDBState As New Recordset
Dim lstrSQLString As String

'Build a query to open the entire tblDbState table
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT *"
lstrSQLString = lstrSQLString & " From tblDbState"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstDBState.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Set the IsActive field to True
lrstDBState!IsActive = True

'Update the record now that it's changed
lrstDBState.Update

'Get the database ID from the new database
gatDataBaseKey.DB_ID = lrstDBState!DbId

'Close the Recordset
lrstDBState.Close

End Sub

Public Function AddForceCalRecord(rstForceCal As Recordset) As Long
'
'   PURPOSE: To add the current Force Calibration to tblForceCalibration.
'
'  INPUT(S): rstForceCal: Recordset that represents tblForceCalibration
' OUTPUT(S): key of new ForceCal Record

'Add the new Record
rstForceCal.AddNew

'Add the new Fields to the Record in tblForceCalibration
rstForceCal!ForceCalSensitivity = gsngNewtonsPerVolt
rstForceCal!ForceCalOffset = gsngForceAmplifierOffset

'Save the Record
rstForceCal.Update

'Assign the Function with the ID of the Record just created
AddForceCalRecord = rstForceCal!ForceCalID

End Function

Public Function AddLotRecord(rstLot As Recordset) As Long
'
'   PURPOSE: To add the current lot to tblLot.
'
'  INPUT(S): rstLot: Recordset that represents tblLot
' OUTPUT(S): key of new Lot Record

'Add the new Record
rstLot.AddNew

'Add the new Fields to the Record in tblLot
rstLot!LotName = gstrLotName

'Save the Record
rstLot.Update

'Assign the Function with the ID of the Record just created
AddLotRecord = rstLot!LotID

End Function

Public Function AddMachineParametersRecord(rstParameters As Recordset) As Long
'
'   PURPOSE: To add the current parameter file to tblMachineParameters
'
'  INPUT(S): rstParameters: Recordset that represents tblMachineParameters
' OUTPUT(S): key of new Parameters Record

'Add the new Record
rstParameters.AddNew

'***Add all the fields in tblMachineParameters***

'Key fields - Name, Revision, & Equipment Name
rstParameters!ParameterFileName = gudtMachine.parameterName
rstParameters!ParameterFileRevision = gudtMachine.parameterRev
rstParameters!EquipmentName = gstrSystemName
'Machine Parameters
rstParameters!seriesID = gudtMachine.seriesID
rstParameters!stationCode = gudtMachine.stationCode
rstParameters!PLC = gudtMachine.PLCCommType
rstParameters!SlopeDevInterval = gudtMachine.slopeInterval
rstParameters!SlopeDevIncrement = gudtMachine.slopeIncrement
rstParameters!FKStartTransitionSlope = gudtMachine.FKSlope
rstParameters!FKStartTransitionWindow = gudtMachine.FKWindow
rstParameters!FKStartTransitionPercentage = gudtMachine.FKPercentage
rstParameters!PedalZeroForce = gudtMachine.pedalAtRestLocForce
rstParameters!loadLocation = gudtMachine.loadLocation
rstParameters!HomeBlockOffset = gudtMachine.blockOffset
rstParameters!preScanStart = gudtMachine.preScanStart
rstParameters!preScanStop = gudtMachine.preScanStop
rstParameters!preScanVelocity = gudtMachine.preScanVelocity
rstParameters!preScanAcceleration = gudtMachine.preScanAcceleration
rstParameters!OffsetForStartScan = gudtMachine.offset4StartScan
rstParameters!scanLength = gudtMachine.scanLength
rstParameters!overTravel = gudtMachine.overTravel
rstParameters!scanVelocity = gudtMachine.scanVelocity
rstParameters!scanAcceleration = gudtMachine.scanAcceleration
rstParameters!progVelocity = gudtMachine.progVelocity
rstParameters!progAcceleration = gudtMachine.progAcceleration
rstParameters!EncCntPerDataPt = gudtMachine.countsPerTrigger
rstParameters!gearRatio = gudtMachine.gearRatio
rstParameters!EncoderResolution = gudtMachine.encReso
rstParameters!graphZeroOffset = gudtMachine.graphZeroOffset
rstParameters!xAxisLow = gudtMachine.xAxisLow
rstParameters!xAxisHigh = gudtMachine.xAxisHigh
rstParameters!Filter1Location = gudtMachine.filterLoc(CHAN0)
rstParameters!Filter2Location = gudtMachine.filterLoc(CHAN1)
rstParameters!Filter3Location = gudtMachine.filterLoc(CHAN2)
rstParameters!Filter4Location = gudtMachine.filterLoc(CHAN3)
rstParameters!VRefMode = gudtMachine.VRefMode
rstParameters!maxLBF = gudtMachine.maxLBF

'Save the Record
rstParameters.Update

'Assign the Function with the ID of the Record just created
AddMachineParametersRecord = rstParameters!MachineParametersID

End Function

Public Function AddOutput1ParametersRecord() As Long
'
'   PURPOSE: To add the current Output #1 parameters to the database
'
'  INPUT(S): none
' OUTPUT(S): key of new Parameters Record

Dim lrstParameters As New Recordset

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open "tblOutput1Parameters", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstParameters.AddNew

'***Add all the fields in tblPedalOutput1Parameters***

'Index 1 Parameters
lrstParameters!IdleIdeal = gudtTest(CHAN0).Index(1).ideal
lrstParameters!IdleHighLimit = gudtTest(CHAN0).Index(1).high
lrstParameters!IdleLowLimit = gudtTest(CHAN0).Index(1).low
lrstParameters!IdleLoc = gudtTest(CHAN0).Index(1).location
'Output At Force Knee Parameters
lrstParameters!OutputAtForceKneeIdeal = gudtTest(CHAN0).outputAtForceKnee.ideal
lrstParameters!OutputAtForceKneeRelativeHighLimit = gudtTest(CHAN0).outputAtForceKnee.high
lrstParameters!OutputAtForceKneeLowLimit = gudtTest(CHAN0).outputAtForceKnee.low
'Index 2 Parameters
lrstParameters!MidPointIdeal = gudtTest(CHAN0).Index(2).ideal
lrstParameters!MidPointHighLimit = gudtTest(CHAN0).Index(2).high
lrstParameters!MidPointLowLimit = gudtTest(CHAN0).Index(2).low
lrstParameters!MidPointLoc = gudtTest(CHAN0).Index(2).location
'Index 3 Parameters
lrstParameters!WOTIdealLoc = gudtTest(CHAN0).Index(3).ideal
lrstParameters!WOTLocHighLimit = gudtTest(CHAN0).Index(3).high
lrstParameters!WOTLocLowLimit = gudtTest(CHAN0).Index(3).low
lrstParameters!WOTValue = gudtTest(CHAN0).Index(3).location
'Maximum Output Parameters
lrstParameters!MaxOutputHighLimit = gudtTest(CHAN0).maxOutput.high
lrstParameters!MaxOutputLowLimit = gudtTest(CHAN0).maxOutput.low
'SinglePoint Linearity Parameters
lrstParameters!SingLinReg1StartLoc = gudtTest(CHAN0).SinglePointLin(1).start.location
lrstParameters!SingLinReg1StartHighLimit = gudtTest(CHAN0).SinglePointLin(1).start.high
lrstParameters!SingLinReg1StartLowLimit = gudtTest(CHAN0).SinglePointLin(1).start.low
lrstParameters!SingLinReg1StopLoc = gudtTest(CHAN0).SinglePointLin(1).stop.location
lrstParameters!SingLinReg1StopHighLimit = gudtTest(CHAN0).SinglePointLin(1).stop.high
lrstParameters!SingLinReg1StopLowLimit = gudtTest(CHAN0).SinglePointLin(1).stop.low
lrstParameters!SingLinReg2StartLoc = gudtTest(CHAN0).SinglePointLin(2).start.location
lrstParameters!SingLinReg2StartHighLimit = gudtTest(CHAN0).SinglePointLin(2).start.high
lrstParameters!SingLinReg2StartLowLimit = gudtTest(CHAN0).SinglePointLin(2).start.low
lrstParameters!SingLinReg2StopLoc = gudtTest(CHAN0).SinglePointLin(2).stop.location
lrstParameters!SingLinReg2StopHighLimit = gudtTest(CHAN0).SinglePointLin(2).stop.high
lrstParameters!SingLinReg2StopLowLimit = gudtTest(CHAN0).SinglePointLin(2).stop.low
lrstParameters!SingLinReg3StartLoc = gudtTest(CHAN0).SinglePointLin(3).start.location
lrstParameters!SingLinReg3StartHighLimit = gudtTest(CHAN0).SinglePointLin(3).start.high
lrstParameters!SingLinReg3StartLowLimit = gudtTest(CHAN0).SinglePointLin(3).start.low
lrstParameters!SingLinReg3StopLoc = gudtTest(CHAN0).SinglePointLin(3).stop.location
lrstParameters!SingLinReg3StopHighLimit = gudtTest(CHAN0).SinglePointLin(3).stop.high
lrstParameters!SingLinReg3StopLowLimit = gudtTest(CHAN0).SinglePointLin(3).stop.low
lrstParameters!SingLinReg4StartLoc = gudtTest(CHAN0).SinglePointLin(4).start.location
lrstParameters!SingLinReg4StartHighLimit = gudtTest(CHAN0).SinglePointLin(4).start.high
lrstParameters!SingLinReg4StartLowLimit = gudtTest(CHAN0).SinglePointLin(4).start.low
lrstParameters!SingLinReg4StopLoc = gudtTest(CHAN0).SinglePointLin(4).stop.location
lrstParameters!SingLinReg4StopHighLimit = gudtTest(CHAN0).SinglePointLin(4).stop.high
lrstParameters!SingLinReg4StopLowLimit = gudtTest(CHAN0).SinglePointLin(4).stop.low
lrstParameters!SingLinReg5StartLoc = gudtTest(CHAN0).SinglePointLin(5).start.location
lrstParameters!SingLinReg5StartHighLimit = gudtTest(CHAN0).SinglePointLin(5).start.high
lrstParameters!SingLinReg5StartLowLimit = gudtTest(CHAN0).SinglePointLin(5).start.low
lrstParameters!SingLinReg5StopLoc = gudtTest(CHAN0).SinglePointLin(5).stop.location
lrstParameters!SingLinReg5StopHighLimit = gudtTest(CHAN0).SinglePointLin(5).stop.high
lrstParameters!SingLinReg5StopLowLimit = gudtTest(CHAN0).SinglePointLin(5).stop.low
'1.8ANM \/\/
'Absolute Linearity Parameters
lrstParameters!AbsLinReg1StartLoc = gudtTest(CHAN0).AbsLin(1).start.location
lrstParameters!AbsLinReg1StartHighLimit = gudtTest(CHAN0).AbsLin(1).start.high
lrstParameters!AbsLinReg1StartLowLimit = gudtTest(CHAN0).AbsLin(1).start.low
lrstParameters!AbsLinReg1StopLoc = gudtTest(CHAN0).AbsLin(1).stop.location
lrstParameters!AbsLinReg1StopHighLimit = gudtTest(CHAN0).AbsLin(1).stop.high
lrstParameters!AbsLinReg1StopLowLimit = gudtTest(CHAN0).AbsLin(1).stop.low
lrstParameters!AbsLinReg2StartLoc = gudtTest(CHAN0).AbsLin(2).start.location
lrstParameters!AbsLinReg2StartHighLimit = gudtTest(CHAN0).AbsLin(2).start.high
lrstParameters!AbsLinReg2StartLowLimit = gudtTest(CHAN0).AbsLin(2).start.low
lrstParameters!AbsLinReg2StopLoc = gudtTest(CHAN0).AbsLin(2).stop.location
lrstParameters!AbsLinReg2StopHighLimit = gudtTest(CHAN0).AbsLin(2).stop.high
lrstParameters!AbsLinReg2StopLowLimit = gudtTest(CHAN0).AbsLin(2).stop.low
lrstParameters!AbsLinReg3StartLoc = gudtTest(CHAN0).AbsLin(3).start.location
lrstParameters!AbsLinReg3StartHighLimit = gudtTest(CHAN0).AbsLin(3).start.high
lrstParameters!AbsLinReg3StartLowLimit = gudtTest(CHAN0).AbsLin(3).start.low
lrstParameters!AbsLinReg3StopLoc = gudtTest(CHAN0).AbsLin(3).stop.location
lrstParameters!AbsLinReg3StopHighLimit = gudtTest(CHAN0).AbsLin(3).stop.high
lrstParameters!AbsLinReg3StopLowLimit = gudtTest(CHAN0).AbsLin(3).stop.low
lrstParameters!AbsLinReg4StartLoc = gudtTest(CHAN0).AbsLin(4).start.location
lrstParameters!AbsLinReg4StartHighLimit = gudtTest(CHAN0).AbsLin(4).start.high
lrstParameters!AbsLinReg4StartLowLimit = gudtTest(CHAN0).AbsLin(4).start.low
lrstParameters!AbsLinReg4StopLoc = gudtTest(CHAN0).AbsLin(4).stop.location
lrstParameters!AbsLinReg4StopHighLimit = gudtTest(CHAN0).AbsLin(4).stop.high
lrstParameters!AbsLinReg4StopLowLimit = gudtTest(CHAN0).AbsLin(4).stop.low
lrstParameters!AbsLinReg5StartLoc = gudtTest(CHAN0).AbsLin(5).start.location
lrstParameters!AbsLinReg5StartHighLimit = gudtTest(CHAN0).AbsLin(5).start.high
lrstParameters!AbsLinReg5StartLowLimit = gudtTest(CHAN0).AbsLin(5).start.low
lrstParameters!AbsLinReg5StopLoc = gudtTest(CHAN0).AbsLin(5).stop.location
lrstParameters!AbsLinReg5StopHighLimit = gudtTest(CHAN0).AbsLin(5).stop.high
lrstParameters!AbsLinReg5StopLowLimit = gudtTest(CHAN0).AbsLin(5).stop.low
'1.8ANM /\/\
'Slope Deviation Parameters
lrstParameters!IdealSlope = gudtTest(CHAN0).slope.ideal
lrstParameters!SlopeDevHighLimit = gudtTest(CHAN0).slope.high
lrstParameters!SlopeDevLowLimit = gudtTest(CHAN0).slope.low
lrstParameters!SlopeDevStartLoc = gudtTest(CHAN0).slope.start
lrstParameters!SlopeDevStopLoc = gudtTest(CHAN0).slope.stop
lrstParameters!FullCloseHysIdeal = gudtTest(CHAN0).FullCloseHys.ideal
lrstParameters!FullCloseHysHighLimit = gudtTest(CHAN0).FullCloseHys.high
lrstParameters!FullCloseHysLowLimit = gudtTest(CHAN0).FullCloseHys.low
'MLX Idd Paras '1.7ANM
lrstParameters!MLXIddIdeal = gudtTest(CHAN0).mlxCurrent.ideal
lrstParameters!MLXIddHighLimit = gudtTest(CHAN0).mlxCurrent.high
lrstParameters!MLXIddLowLimit = gudtTest(CHAN0).mlxCurrent.low
'MLX Idd Paras '2.0ANM
lrstParameters!MLXWIddIdeal = gudtTest(CHAN0).mlxWCurrent.ideal
lrstParameters!MLXWIddHighLimit = gudtTest(CHAN0).mlxWCurrent.high
lrstParameters!MLXWIddLowLimit = gudtTest(CHAN0).mlxWCurrent.low
'Eval
lrstParameters!EvaluationStartLoc = gudtTest(CHAN0).evaluate.start
lrstParameters!EvaluationStopLoc = gudtTest(CHAN0).evaluate.stop
'Save the Record
lrstParameters.Update

'Assign the Function with the ID of the Record just created
AddOutput1ParametersRecord = lrstParameters!Output1ParametersID

End Function

Public Function AddOutput2ParametersRecord() As Long
'
'   PURPOSE: To add the current Output #2 parameters to the database
'
'  INPUT(S): none
' OUTPUT(S): key of new Parameters Record

Dim lrstParameters As New Recordset

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open "tblOutput2Parameters", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstParameters.AddNew

'***Add all the fields in tblPedalOutput1Parameters***

'Index 1 Parameters
lrstParameters!IdleIdeal = gudtTest(CHAN1).Index(1).ideal
lrstParameters!IdleHighLimit = gudtTest(CHAN1).Index(1).high
lrstParameters!IdleLowLimit = gudtTest(CHAN1).Index(1).low
lrstParameters!IdleLoc = gudtTest(CHAN1).Index(1).location
'Output At Force Knee Parameters
lrstParameters!OutputAtForceKneeIdeal = gudtTest(CHAN1).outputAtForceKnee.ideal
lrstParameters!OutputAtForceKneeRelativeHighLimit = gudtTest(CHAN1).outputAtForceKnee.high
lrstParameters!OutputAtForceKneeLowLimit = gudtTest(CHAN1).outputAtForceKnee.low
'Index 2 Parameters
lrstParameters!MidPointIdeal = gudtTest(CHAN1).Index(2).ideal
lrstParameters!MidPointHighLimit = gudtTest(CHAN1).Index(2).high
lrstParameters!MidPointLowLimit = gudtTest(CHAN1).Index(2).low
lrstParameters!MidPointLoc = gudtTest(CHAN1).Index(2).location
'Index 3 Parameters
lrstParameters!WOTIdealLoc = gudtTest(CHAN1).Index(3).ideal
lrstParameters!WOTLocHighLimit = gudtTest(CHAN1).Index(3).high
lrstParameters!WOTLocLowLimit = gudtTest(CHAN1).Index(3).low
lrstParameters!WOTValue = gudtTest(CHAN1).Index(3).location
'Maximum Output Parameters
lrstParameters!MaxOutputHighLimit = gudtTest(CHAN1).maxOutput.high
lrstParameters!MaxOutputLowLimit = gudtTest(CHAN1).maxOutput.low
'SinglePoint Linearity Parameters
lrstParameters!SingLinReg1StartLoc = gudtTest(CHAN1).SinglePointLin(1).start.location
lrstParameters!SingLinReg1StartHighLimit = gudtTest(CHAN1).SinglePointLin(1).start.high
lrstParameters!SingLinReg1StartLowLimit = gudtTest(CHAN1).SinglePointLin(1).start.low
lrstParameters!SingLinReg1StopLoc = gudtTest(CHAN1).SinglePointLin(1).stop.location
lrstParameters!SingLinReg1StopHighLimit = gudtTest(CHAN1).SinglePointLin(1).stop.high
lrstParameters!SingLinReg1StopLowLimit = gudtTest(CHAN1).SinglePointLin(1).stop.low
lrstParameters!SingLinReg2StartLoc = gudtTest(CHAN1).SinglePointLin(2).start.location
lrstParameters!SingLinReg2StartHighLimit = gudtTest(CHAN1).SinglePointLin(2).start.high
lrstParameters!SingLinReg2StartLowLimit = gudtTest(CHAN1).SinglePointLin(2).start.low
lrstParameters!SingLinReg2StopLoc = gudtTest(CHAN1).SinglePointLin(2).stop.location
lrstParameters!SingLinReg2StopHighLimit = gudtTest(CHAN1).SinglePointLin(2).stop.high
lrstParameters!SingLinReg2StopLowLimit = gudtTest(CHAN1).SinglePointLin(2).stop.low
lrstParameters!SingLinReg3StartLoc = gudtTest(CHAN1).SinglePointLin(3).start.location
lrstParameters!SingLinReg3StartHighLimit = gudtTest(CHAN1).SinglePointLin(3).start.high
lrstParameters!SingLinReg3StartLowLimit = gudtTest(CHAN1).SinglePointLin(3).start.low
lrstParameters!SingLinReg3StopLoc = gudtTest(CHAN1).SinglePointLin(3).stop.location
lrstParameters!SingLinReg3StopHighLimit = gudtTest(CHAN1).SinglePointLin(3).stop.high
lrstParameters!SingLinReg3StopLowLimit = gudtTest(CHAN1).SinglePointLin(3).stop.low
lrstParameters!SingLinReg4StartLoc = gudtTest(CHAN1).SinglePointLin(4).start.location
lrstParameters!SingLinReg4StartHighLimit = gudtTest(CHAN1).SinglePointLin(4).start.high
lrstParameters!SingLinReg4StartLowLimit = gudtTest(CHAN1).SinglePointLin(4).start.low
lrstParameters!SingLinReg4StopLoc = gudtTest(CHAN1).SinglePointLin(4).stop.location
lrstParameters!SingLinReg4StopHighLimit = gudtTest(CHAN1).SinglePointLin(4).stop.high
lrstParameters!SingLinReg4StopLowLimit = gudtTest(CHAN1).SinglePointLin(4).stop.low
lrstParameters!SingLinReg5StartLoc = gudtTest(CHAN1).SinglePointLin(5).start.location
lrstParameters!SingLinReg5StartHighLimit = gudtTest(CHAN1).SinglePointLin(5).start.high
lrstParameters!SingLinReg5StartLowLimit = gudtTest(CHAN1).SinglePointLin(5).start.low
lrstParameters!SingLinReg5StopLoc = gudtTest(CHAN1).SinglePointLin(5).stop.location
lrstParameters!SingLinReg5StopHighLimit = gudtTest(CHAN1).SinglePointLin(5).stop.high
lrstParameters!SingLinReg5StopLowLimit = gudtTest(CHAN1).SinglePointLin(5).stop.low
'1.8ANM \/\/
'Absolute Linearity Parameters
lrstParameters!AbsLinReg1StartLoc = gudtTest(CHAN1).AbsLin(1).start.location
lrstParameters!AbsLinReg1StartHighLimit = gudtTest(CHAN1).AbsLin(1).start.high
lrstParameters!AbsLinReg1StartLowLimit = gudtTest(CHAN1).AbsLin(1).start.low
lrstParameters!AbsLinReg1StopLoc = gudtTest(CHAN1).AbsLin(1).stop.location
lrstParameters!AbsLinReg1StopHighLimit = gudtTest(CHAN1).AbsLin(1).stop.high
lrstParameters!AbsLinReg1StopLowLimit = gudtTest(CHAN1).AbsLin(1).stop.low
lrstParameters!AbsLinReg2StartLoc = gudtTest(CHAN1).AbsLin(2).start.location
lrstParameters!AbsLinReg2StartHighLimit = gudtTest(CHAN1).AbsLin(2).start.high
lrstParameters!AbsLinReg2StartLowLimit = gudtTest(CHAN1).AbsLin(2).start.low
lrstParameters!AbsLinReg2StopLoc = gudtTest(CHAN1).AbsLin(2).stop.location
lrstParameters!AbsLinReg2StopHighLimit = gudtTest(CHAN1).AbsLin(2).stop.high
lrstParameters!AbsLinReg2StopLowLimit = gudtTest(CHAN1).AbsLin(2).stop.low
lrstParameters!AbsLinReg3StartLoc = gudtTest(CHAN1).AbsLin(3).start.location
lrstParameters!AbsLinReg3StartHighLimit = gudtTest(CHAN1).AbsLin(3).start.high
lrstParameters!AbsLinReg3StartLowLimit = gudtTest(CHAN1).AbsLin(3).start.low
lrstParameters!AbsLinReg3StopLoc = gudtTest(CHAN1).AbsLin(3).stop.location
lrstParameters!AbsLinReg3StopHighLimit = gudtTest(CHAN1).AbsLin(3).stop.high
lrstParameters!AbsLinReg3StopLowLimit = gudtTest(CHAN1).AbsLin(3).stop.low
lrstParameters!AbsLinReg4StartLoc = gudtTest(CHAN1).AbsLin(4).start.location
lrstParameters!AbsLinReg4StartHighLimit = gudtTest(CHAN1).AbsLin(4).start.high
lrstParameters!AbsLinReg4StartLowLimit = gudtTest(CHAN1).AbsLin(4).start.low
lrstParameters!AbsLinReg4StopLoc = gudtTest(CHAN1).AbsLin(4).stop.location
lrstParameters!AbsLinReg4StopHighLimit = gudtTest(CHAN1).AbsLin(4).stop.high
lrstParameters!AbsLinReg4StopLowLimit = gudtTest(CHAN1).AbsLin(4).stop.low
lrstParameters!AbsLinReg5StartLoc = gudtTest(CHAN1).AbsLin(5).start.location
lrstParameters!AbsLinReg5StartHighLimit = gudtTest(CHAN1).AbsLin(5).start.high
lrstParameters!AbsLinReg5StartLowLimit = gudtTest(CHAN1).AbsLin(5).start.low
lrstParameters!AbsLinReg5StopLoc = gudtTest(CHAN1).AbsLin(5).stop.location
lrstParameters!AbsLinReg5StopHighLimit = gudtTest(CHAN1).AbsLin(5).stop.high
lrstParameters!AbsLinReg5StopLowLimit = gudtTest(CHAN1).AbsLin(5).stop.low
'1.8ANM /\/\
'Slope Deviation Parameters
lrstParameters!IdealSlope = gudtTest(CHAN1).slope.ideal
lrstParameters!SlopeDevHighLimit = gudtTest(CHAN1).slope.high
lrstParameters!SlopeDevLowLimit = gudtTest(CHAN1).slope.low
lrstParameters!SlopeDevStartLoc = gudtTest(CHAN1).slope.start
lrstParameters!SlopeDevStopLoc = gudtTest(CHAN1).slope.stop
lrstParameters!FullCloseHysIdeal = gudtTest(CHAN1).FullCloseHys.ideal
lrstParameters!FullCloseHysHighLimit = gudtTest(CHAN1).FullCloseHys.high
lrstParameters!FullCloseHysLowLimit = gudtTest(CHAN1).FullCloseHys.low
'MLX Idd Paras '1.7ANM
lrstParameters!MLXIddIdeal = gudtTest(CHAN1).mlxCurrent.ideal
lrstParameters!MLXIddHighLimit = gudtTest(CHAN1).mlxCurrent.high
lrstParameters!MLXIddLowLimit = gudtTest(CHAN1).mlxCurrent.low
'MLX Idd Paras '2.0ANM
lrstParameters!MLXWIddIdeal = gudtTest(CHAN1).mlxWCurrent.ideal
lrstParameters!MLXWIddHighLimit = gudtTest(CHAN1).mlxWCurrent.high
lrstParameters!MLXWIddLowLimit = gudtTest(CHAN1).mlxWCurrent.low
'Eval
lrstParameters!EvaluationStartLoc = gudtTest(CHAN1).evaluate.start
lrstParameters!EvaluationStopLoc = gudtTest(CHAN1).evaluate.stop
'Save the Record
lrstParameters.Update

'Assign the Function with the ID of the Record just created
AddOutput2ParametersRecord = lrstParameters!Output2ParametersID

End Function

Public Function AddProgrammingParametersRecord(rstParameters As Recordset) As Long
'
'   PURPOSE: To add the current parameter file to tblProgrammingParameters
'
'  INPUT(S): rstParameters: Recordset that represents tblProgrammingParameters
' OUTPUT(S): key of new Parameters Record

'Add the new Record
rstParameters.AddNew

'***Add all the fields in tblProgrammingParameters***

'Key fields - Name & Revision
rstParameters!ParameterFileName = gudtMachine.parameterName
rstParameters!ParameterFileRevision = gudtMachine.parameterRev
rstParameters!MLX90277RevLevel = gstrMLX90277Revision
'Programming Times Output #1
rstParameters!TporOutput1 = gudtPTC04(1).Tpor
rstParameters!TholdOutput1 = gudtPTC04(1).Thold
rstParameters!TprogOutput1 = gudtPTC04(1).Tprog
rstParameters!TpulsOutput1 = gudtPTC04(1).Tpuls
'Solver Parameters Output #1
rstParameters!Index1IdealOutput1 = gudtSolver(1).Index(1).IdealValue
rstParameters!Index1LocationOutput1 = gudtSolver(1).Index(1).IdealLocation
rstParameters!Index1TargetToleranceOutput1 = gudtSolver(1).Index(1).TargetTolerance
rstParameters!Index1PassFailToleranceOutput1 = gudtSolver(1).Index(1).PassFailTolerance
rstParameters!Index2IdealOutput1 = gudtSolver(1).Index(2).IdealValue
rstParameters!Index2LocationOutput1 = gudtSolver(1).Index(2).IdealLocation
rstParameters!Index2TargetToleranceOutput1 = gudtSolver(1).Index(2).TargetTolerance
rstParameters!Index2PassFailToleranceOutput1 = gudtSolver(1).Index(2).PassFailTolerance
rstParameters!FilterOutput1 = gudtSolver(1).Filter
rstParameters!InvertOutput1 = gudtSolver(1).InvertSlope
rstParameters!ModeOutput1 = gudtSolver(1).Mode
rstParameters!FaultLevelOutput1 = gudtSolver(1).FaultLevel
rstParameters!MaxOffsetDriftOutput1 = gudtSolver(1).MaxOffsetDrift
rstParameters!MaxAGNDSettingOutput1 = gudtSolver(1).MaxAGND
rstParameters!MinAGNDSettingOutput1 = gudtSolver(1).MinAGND
rstParameters!FCKADJSettingOutput1 = gudtSolver(1).FCKADJ
rstParameters!CKANACHSettingOutput1 = gudtSolver(1).CKANACH
rstParameters!CKDACCHSettingOutput1 = gudtSolver(1).CKDACCH
rstParameters!SlowModeSettingOutput1 = gudtSolver(1).SlowMode
rstParameters!InitialOffsetOutput1 = gudtSolver(1).InitialOffset
rstParameters!HighRGHighFG1 = gudtSolver(1).HighRGHighFG  '1.1ANM
rstParameters!HighRGLowFG1 = gudtSolver(1).HighRGLowFG    '1.1ANM
rstParameters!LowRGHighFG1 = gudtSolver(1).LowRGHighFG    '1.1ANM
rstParameters!LowRGLowFG1 = gudtSolver(1).LowRGLowFG      '1.1ANM
rstParameters!MinRoughGainOutput1 = gudtSolver(1).MinRG
rstParameters!MaxRoughGainOutput1 = gudtSolver(1).MaxRG
rstParameters!OffsetStepOutput1 = gudtSolver(1).OffsetStep
'1.6ANM rstParameters!RatioA1Output1 = gudtSolver(1).CodeRatio(1, 1)
'1.6ANM rstParameters!RatioA2Output1 = gudtSolver(1).CodeRatio(1, 2)
'1.6ANM rstParameters!RatioA3Output1 = gudtSolver(1).CodeRatio(1, 3)
rstParameters!RatioB1Output1 = gudtSolver(1).CodeRatio(2, 1)
rstParameters!RatioB2Output1 = gudtSolver(1).CodeRatio(2, 2)
rstParameters!RatioB3Output1 = gudtSolver(1).CodeRatio(2, 3)
rstParameters!ClampLowIdealOutput1 = gudtSolver(1).Clamp(1).IdealValue
rstParameters!ClampLowTargetToleranceOutput1 = gudtSolver(1).Clamp(1).TargetTolerance
rstParameters!ClampLowPassFailToleranceOutput1 = gudtSolver(1).Clamp(1).PassFailTolerance
rstParameters!ClampLowInitialCodeOutput1 = gudtSolver(1).Clamp(1).InitialCode
rstParameters!ClampHighIdealOutput1 = gudtSolver(1).Clamp(2).IdealValue
rstParameters!ClampHighTargetToleranceOutput1 = gudtSolver(1).Clamp(2).TargetTolerance
rstParameters!ClampHighPassFailToleranceOutput1 = gudtSolver(1).Clamp(2).PassFailTolerance
rstParameters!ClampHighInitialCodeOutput1 = gudtSolver(1).Clamp(2).InitialCode
rstParameters!ClampStepOutput1 = gudtSolver(1).ClampStep

'Programming Times Output #2
rstParameters!TporOutput2 = gudtPTC04(2).Tpor
rstParameters!TholdOutput2 = gudtPTC04(2).Thold
rstParameters!TprogOutput2 = gudtPTC04(2).Tprog
rstParameters!TpulsOutput2 = gudtPTC04(2).Tpuls
'Solver Parameters Output #2
rstParameters!Index1IdealOutput2 = gudtSolver(2).Index(1).IdealValue
rstParameters!Index1LocationOutput2 = gudtSolver(2).Index(1).IdealLocation
rstParameters!Index1TargetToleranceOutput2 = gudtSolver(2).Index(1).TargetTolerance
rstParameters!Index1PassFailToleranceOutput2 = gudtSolver(2).Index(1).PassFailTolerance
rstParameters!Index2IdealOutput2 = gudtSolver(2).Index(2).IdealValue
rstParameters!Index2LocationOutput2 = gudtSolver(2).Index(2).IdealLocation
rstParameters!Index2TargetToleranceOutput2 = gudtSolver(2).Index(2).TargetTolerance
rstParameters!Index2PassFailToleranceOutput2 = gudtSolver(2).Index(2).PassFailTolerance
rstParameters!FilterOutput2 = gudtSolver(2).Filter
rstParameters!InvertOutput2 = gudtSolver(2).InvertSlope
rstParameters!ModeOutput2 = gudtSolver(2).Mode
rstParameters!FaultLevelOutput2 = gudtSolver(2).FaultLevel
rstParameters!MaxOffsetDriftOutput2 = gudtSolver(2).MaxOffsetDrift
rstParameters!MaxAGNDSettingOutput2 = gudtSolver(2).MaxAGND
rstParameters!MinAGNDSettingOutput2 = gudtSolver(2).MinAGND
rstParameters!FCKADJSettingOutput2 = gudtSolver(2).FCKADJ
rstParameters!CKANACHSettingOutput2 = gudtSolver(2).CKANACH
rstParameters!CKDACCHSettingOutput2 = gudtSolver(2).CKDACCH
rstParameters!SlowModeSettingOutput2 = gudtSolver(2).SlowMode
rstParameters!InitialOffsetOutput2 = gudtSolver(2).InitialOffset
rstParameters!HighRGHighFG2 = gudtSolver(2).HighRGHighFG  '1.1ANM
rstParameters!HighRGLowFG2 = gudtSolver(2).HighRGLowFG    '1.1ANM
rstParameters!LowRGHighFG2 = gudtSolver(2).LowRGHighFG    '1.1ANM
rstParameters!LowRGLowFG2 = gudtSolver(2).LowRGLowFG      '1.1ANM
rstParameters!MinRoughGainOutput2 = gudtSolver(2).MinRG
rstParameters!MaxRoughGainOutput2 = gudtSolver(2).MaxRG
rstParameters!OffsetStepOutput2 = gudtSolver(2).OffsetStep
'1.6ANM rstParameters!RatioA1Output2 = gudtSolver(2).CodeRatio(1, 1)
'1.6ANM rstParameters!RatioA2Output2 = gudtSolver(2).CodeRatio(1, 2)
'1.6ANM rstParameters!RatioA3Output2 = gudtSolver(2).CodeRatio(1, 3)
rstParameters!RatioB1Output2 = gudtSolver(2).CodeRatio(2, 1)
rstParameters!RatioB2Output2 = gudtSolver(2).CodeRatio(2, 2)
rstParameters!RatioB3Output2 = gudtSolver(2).CodeRatio(2, 3)
rstParameters!ClampLowIdealOutput2 = gudtSolver(2).Clamp(1).IdealValue
rstParameters!ClampLowTargetToleranceOutput2 = gudtSolver(2).Clamp(1).TargetTolerance
rstParameters!ClampLowPassFailToleranceOutput2 = gudtSolver(2).Clamp(1).PassFailTolerance
rstParameters!ClampLowInitialCodeOutput2 = gudtSolver(2).Clamp(1).InitialCode
rstParameters!ClampHighIdealOutput2 = gudtSolver(2).Clamp(2).IdealValue
rstParameters!ClampHighTargetToleranceOutput2 = gudtSolver(2).Clamp(2).TargetTolerance
rstParameters!ClampHighPassFailToleranceOutput2 = gudtSolver(2).Clamp(2).PassFailTolerance
rstParameters!ClampHighInitialCodeOutput2 = gudtSolver(2).Clamp(2).InitialCode
rstParameters!ClampStepOutput2 = gudtSolver(2).ClampStep

'Save the Record
rstParameters.Update

'Assign the Function with the ID of the Record just created
AddProgrammingParametersRecord = rstParameters!ProgrammingParametersID

End Function

Public Sub AddProgrammingResultsRecord()
'
'   PURPOSE: To add the current Programming Results to tblProgrammingResults
'
'  INPUT(S): None
' OUTPUT(S): None

On Error GoTo ErrorSavingProgrammingData

Dim lrstProgResults As New Recordset

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstProgResults.Open "tblProgrammingResults", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstProgResults.AddNew

'Add the new Fields to the Record in tblDUTAnalysisResults
'Keys representing test setup & parameters
lrstProgResults!SerialNumberID = gatDataBaseKey.SerialNumberID
lrstProgResults!MachineParametersID = gatDataBaseKey.MachineParametersID
lrstProgResults!ProgrammingParametersID = gatDataBaseKey.ProgParametersID
lrstProgResults!LotID = gatDataBaseKey.LotID
'Software version & Date/Time Stamp
lrstProgResults!SoftwareRevision = App.Major & "." & App.Minor & "." & App.Revision
lrstProgResults!ProgDateTime = DateTime.Now
'Pallet Number
lrstProgResults!PalletNum = gintPalletNumber
'Error information
lrstProgResults!Anomaly = gintAnomaly

'If the scan was successful, go ahead and save the results data
If gintAnomaly = 0 Then
    'Programming Results
    lrstProgResults!PassFail = Not gblnProgFailure
    'Output #1
    'Indexes 1 & 2
    lrstProgResults!Index1ValOutput1 = gudtSolver(1).FinalIndexVal(1)
    lrstProgResults!Index1LocOutput1 = gudtSolver(1).FinalIndexLoc(1)
    lrstProgResults!Index2ValOutput1 = gudtSolver(1).FinalIndexVal(2)
    lrstProgResults!Index2LocOutput1 = gudtSolver(1).FinalIndexLoc(2)
    'Clamp Values
    lrstProgResults!ClampLowValueOutput1 = gudtSolver(1).FinalClampLowVal
    lrstProgResults!ClampHighValueOutput1 = gudtSolver(1).FinalClampHighVal
    'Final Codes
    lrstProgResults!OffsetCodeOutput1 = gudtSolver(1).FinalOffsetCode
    lrstProgResults!RoughGainCodeOutput1 = gudtSolver(1).FinalRGCode
    lrstProgResults!FineGainCodeOutput1 = gudtSolver(1).FinalFGCode
    lrstProgResults!ClampLowCodeOutput1 = gudtSolver(1).FinalClampLowCode
    lrstProgResults!ClampHighCodeOutput1 = gudtSolver(1).FinalClampHighCode
    'Output #2
    'Indexes 1 & 2
    lrstProgResults!Index1ValOutput2 = gudtSolver(2).FinalIndexVal(1)
    lrstProgResults!Index1LocOutput2 = gudtSolver(2).FinalIndexLoc(1)
    lrstProgResults!Index2ValOutput2 = gudtSolver(2).FinalIndexVal(2)
    lrstProgResults!Index2LocOutput2 = gudtSolver(2).FinalIndexLoc(2)
    'Clamp Values
    lrstProgResults!ClampLowValueOutput2 = gudtSolver(2).FinalClampLowVal
    lrstProgResults!ClampHighValueOutput2 = gudtSolver(2).FinalClampHighVal
    'Final Codes
    lrstProgResults!OffsetCodeOutput2 = gudtSolver(2).FinalOffsetCode
    lrstProgResults!RoughGainCodeOutput2 = gudtSolver(2).FinalRGCode
    lrstProgResults!FineGainCodeOutput2 = gudtSolver(2).FinalFGCode
    lrstProgResults!ClampLowCodeOutput2 = gudtSolver(2).FinalClampLowCode
    lrstProgResults!ClampHighCodeOutput2 = gudtSolver(2).FinalClampHighCode
End If

'User-entered test information
If frmMain.ctrSetupInfo1.Operator = "" Then
    lrstProgResults!Operator = Null 'Save Null if there is an emty string ("")
Else
    lrstProgResults!Operator = frmMain.ctrSetupInfo1.Operator
End If
If frmMain.ctrSetupInfo1.Temperature = "" Then
    lrstProgResults!Temperature = Null 'Save Null if there is an emty string ("")
Else
    lrstProgResults!Temperature = frmMain.ctrSetupInfo1.Temperature
End If
If frmMain.ctrSetupInfo1.Comment = "" Then
    lrstProgResults!Comment = Null 'Save Null if there is an emty string ("")
Else
    lrstProgResults!Comment = frmMain.ctrSetupInfo1.Comment
End If

'1.2ANM Record if part locked or not
lrstProgResults!Locked = gblnLockedPart

'Save the Record
lrstProgResults.Update

'Close the Recordset
lrstProgResults.Close

Exit Sub
ErrorSavingProgrammingData:

    gintAnomaly = 88
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Programming Results.", True, True)
End Sub

Public Function AddScanParametersRecord(rstParameters As Recordset) As Long
'
'   PURPOSE: To add the current parameter file to tblScanParameters
'
'  INPUT(S): rstParameters: Recordset that represents tblScanParameters
' OUTPUT(S): key of new Parameters Record

'Add the new Record
rstParameters.AddNew

'***Add all the fields in tblScanParameters***

'Key fields - Name & Revision
rstParameters!ParameterFileName = gudtMachine.parameterName
rstParameters!ParameterFileRevision = gudtMachine.parameterRev

'Add entries in the PedalOutputParameter tables for both outputs
rstParameters!Output1ID = AddOutput1ParametersRecord
rstParameters!Output2ID = AddOutput2ParametersRecord

'Forward Output Correlation Parameters
rstParameters!FwdOutputCorrelationReg1StartLoc = gudtTest(CHAN0).fwdOutputCor(1).start.location
rstParameters!FwdOutputCorrelationReg1StartHigh = gudtTest(CHAN0).fwdOutputCor(1).start.high
rstParameters!FwdOutputCorrelationReg1StartLow = gudtTest(CHAN0).fwdOutputCor(1).start.low
rstParameters!FwdOutputCorrelationReg1StopLoc = gudtTest(CHAN0).fwdOutputCor(1).stop.location
rstParameters!FwdOutputCorrelationReg1StopHigh = gudtTest(CHAN0).fwdOutputCor(1).stop.high
rstParameters!FwdOutputCorrelationReg1StopLow = gudtTest(CHAN0).fwdOutputCor(1).stop.low
rstParameters!FwdOutputCorrelationReg2StartLoc = gudtTest(CHAN0).fwdOutputCor(2).start.location
rstParameters!FwdOutputCorrelationReg2StartHigh = gudtTest(CHAN0).fwdOutputCor(2).start.high
rstParameters!FwdOutputCorrelationReg2StartLow = gudtTest(CHAN0).fwdOutputCor(2).start.low
rstParameters!FwdOutputCorrelationReg2StopLoc = gudtTest(CHAN0).fwdOutputCor(2).stop.location
rstParameters!FwdOutputCorrelationReg2StopHigh = gudtTest(CHAN0).fwdOutputCor(2).stop.high
rstParameters!FwdOutputCorrelationReg2StopLow = gudtTest(CHAN0).fwdOutputCor(2).stop.low
rstParameters!FwdOutputCorrelationReg3StartLoc = gudtTest(CHAN0).fwdOutputCor(3).start.location
rstParameters!FwdOutputCorrelationReg3StartHigh = gudtTest(CHAN0).fwdOutputCor(3).start.high
rstParameters!FwdOutputCorrelationReg3StartLow = gudtTest(CHAN0).fwdOutputCor(3).start.low
rstParameters!FwdOutputCorrelationReg3StopLoc = gudtTest(CHAN0).fwdOutputCor(3).stop.location
rstParameters!FwdOutputCorrelationReg3StopHigh = gudtTest(CHAN0).fwdOutputCor(3).stop.high
rstParameters!FwdOutputCorrelationReg3StopLow = gudtTest(CHAN0).fwdOutputCor(3).stop.low
rstParameters!FwdOutputCorrelationReg4StartLoc = gudtTest(CHAN0).fwdOutputCor(4).start.location
rstParameters!FwdOutputCorrelationReg4StartHigh = gudtTest(CHAN0).fwdOutputCor(4).start.high
rstParameters!FwdOutputCorrelationReg4StartLow = gudtTest(CHAN0).fwdOutputCor(4).start.low
rstParameters!FwdOutputCorrelationReg4StopLoc = gudtTest(CHAN0).fwdOutputCor(4).stop.location
rstParameters!FwdOutputCorrelationReg4StopHigh = gudtTest(CHAN0).fwdOutputCor(4).stop.high
rstParameters!FwdOutputCorrelationReg4StopLow = gudtTest(CHAN0).fwdOutputCor(4).stop.low
rstParameters!FwdOutputCorrelationReg5StartLoc = gudtTest(CHAN0).fwdOutputCor(5).start.location
rstParameters!FwdOutputCorrelationReg5StartHigh = gudtTest(CHAN0).fwdOutputCor(5).start.high
rstParameters!FwdOutputCorrelationReg5StartLow = gudtTest(CHAN0).fwdOutputCor(5).start.low
rstParameters!FwdOutputCorrelationReg5StopLoc = gudtTest(CHAN0).fwdOutputCor(5).stop.location
rstParameters!FwdOutputCorrelationReg5StopHigh = gudtTest(CHAN0).fwdOutputCor(5).stop.high
rstParameters!FwdOutputCorrelationReg5StopLow = gudtTest(CHAN0).fwdOutputCor(5).stop.low
'Reverse Output Correlation Parameters
rstParameters!RevOutputCorrelationReg1StartLoc = gudtTest(CHAN0).revOutputCor(1).start.location
rstParameters!RevOutputCorrelationReg1StartHigh = gudtTest(CHAN0).revOutputCor(1).start.high
rstParameters!RevOutputCorrelationReg1StartLow = gudtTest(CHAN0).revOutputCor(1).start.low
rstParameters!RevOutputCorrelationReg1StopLoc = gudtTest(CHAN0).revOutputCor(1).stop.location
rstParameters!RevOutputCorrelationReg1StopHigh = gudtTest(CHAN0).revOutputCor(1).stop.high
rstParameters!RevOutputCorrelationReg1StopLow = gudtTest(CHAN0).revOutputCor(1).stop.low
rstParameters!RevOutputCorrelationReg2StartLoc = gudtTest(CHAN0).revOutputCor(2).start.location
rstParameters!RevOutputCorrelationReg2StartHigh = gudtTest(CHAN0).revOutputCor(2).start.high
rstParameters!RevOutputCorrelationReg2StartLow = gudtTest(CHAN0).revOutputCor(2).start.low
rstParameters!RevOutputCorrelationReg2StopLoc = gudtTest(CHAN0).revOutputCor(2).stop.location
rstParameters!RevOutputCorrelationReg2StopHigh = gudtTest(CHAN0).revOutputCor(2).stop.high
rstParameters!RevOutputCorrelationReg2StopLow = gudtTest(CHAN0).revOutputCor(2).stop.low
rstParameters!RevOutputCorrelationReg3StartLoc = gudtTest(CHAN0).revOutputCor(3).start.location
rstParameters!RevOutputCorrelationReg3StartHigh = gudtTest(CHAN0).revOutputCor(3).start.high
rstParameters!RevOutputCorrelationReg3StartLow = gudtTest(CHAN0).revOutputCor(3).start.low
rstParameters!RevOutputCorrelationReg3StopLoc = gudtTest(CHAN0).revOutputCor(3).stop.location
rstParameters!RevOutputCorrelationReg3StopHigh = gudtTest(CHAN0).revOutputCor(3).stop.high
rstParameters!RevOutputCorrelationReg3StopLow = gudtTest(CHAN0).revOutputCor(3).stop.low
rstParameters!RevOutputCorrelationReg4StartLoc = gudtTest(CHAN0).revOutputCor(4).start.location
rstParameters!RevOutputCorrelationReg4StartHigh = gudtTest(CHAN0).revOutputCor(4).start.high
rstParameters!RevOutputCorrelationReg4StartLow = gudtTest(CHAN0).revOutputCor(4).start.low
rstParameters!RevOutputCorrelationReg4StopLoc = gudtTest(CHAN0).revOutputCor(4).stop.location
rstParameters!RevOutputCorrelationReg4StopHigh = gudtTest(CHAN0).revOutputCor(4).stop.high
rstParameters!RevOutputCorrelationReg4StopLow = gudtTest(CHAN0).revOutputCor(4).stop.low
rstParameters!RevOutputCorrelationReg5StartLoc = gudtTest(CHAN0).revOutputCor(5).start.location
rstParameters!RevOutputCorrelationReg5StartHigh = gudtTest(CHAN0).revOutputCor(5).start.high
rstParameters!RevOutputCorrelationReg5StartLow = gudtTest(CHAN0).revOutputCor(5).start.low
rstParameters!RevOutputCorrelationReg5StopLoc = gudtTest(CHAN0).revOutputCor(5).stop.location
rstParameters!RevOutputCorrelationReg5StopHigh = gudtTest(CHAN0).revOutputCor(5).stop.high
rstParameters!RevOutputCorrelationReg5StopLow = gudtTest(CHAN0).revOutputCor(5).stop.low
'Pedal-At-Rest Location Location
rstParameters!pedalAtRestLocationIdeal = gudtTest(CHAN0).pedalAtRestLoc.ideal
'Force Knee Location
rstParameters!ForceKneeLocationIdeal = gudtTest(CHAN0).forceKneeLoc.ideal
rstParameters!ForceKneeLocationHigh = gudtTest(CHAN0).forceKneeLoc.high
rstParameters!ForceKneeLocationLow = gudtTest(CHAN0).forceKneeLoc.low
rstParameters!FwdForceAtForceKneeLocationIdeal = gudtTest(CHAN0).forceKneeForce.ideal
rstParameters!FwdForceAtForceKneeLocationHigh = gudtTest(CHAN0).forceKneeForce.high
rstParameters!FwdForceAtForceKneeLocationLow = gudtTest(CHAN0).forceKneeForce.low
'Forward Force Points
rstParameters!FwdForcePt1Ideal = gudtTest(CHAN0).fwdForcePt(1).ideal
rstParameters!FwdForcePt1High = gudtTest(CHAN0).fwdForcePt(1).high
rstParameters!FwdForcePt1Low = gudtTest(CHAN0).fwdForcePt(1).low
rstParameters!FwdForcePt1Location = gudtTest(CHAN0).fwdForcePt(1).location
rstParameters!FwdForcePt2Ideal = gudtTest(CHAN0).fwdForcePt(2).ideal
rstParameters!FwdForcePt2High = gudtTest(CHAN0).fwdForcePt(2).high
rstParameters!FwdForcePt2Low = gudtTest(CHAN0).fwdForcePt(2).low
rstParameters!FwdForcePt2Location = gudtTest(CHAN0).fwdForcePt(2).location
rstParameters!FwdForcePt3Ideal = gudtTest(CHAN0).fwdForcePt(3).ideal
rstParameters!FwdForcePt3High = gudtTest(CHAN0).fwdForcePt(3).high
rstParameters!FwdForcePt3Low = gudtTest(CHAN0).fwdForcePt(3).low
rstParameters!FwdForcePt3Location = gudtTest(CHAN0).fwdForcePt(3).location
'Reverse Force Points
rstParameters!RevForcePt1Ideal = gudtTest(CHAN0).revForcePt(1).ideal
rstParameters!RevForcePt1High = gudtTest(CHAN0).revForcePt(1).high
rstParameters!RevForcePt1Low = gudtTest(CHAN0).revForcePt(1).low
rstParameters!RevForcePt1Location = gudtTest(CHAN0).revForcePt(1).location
rstParameters!RevForcePt2Ideal = gudtTest(CHAN0).revForcePt(2).ideal
rstParameters!RevForcePt2High = gudtTest(CHAN0).revForcePt(2).high
rstParameters!RevForcePt2Low = gudtTest(CHAN0).revForcePt(2).low
rstParameters!RevForcePt2Location = gudtTest(CHAN0).revForcePt(2).location
rstParameters!RevForcePt3Ideal = gudtTest(CHAN0).revForcePt(3).ideal
rstParameters!RevForcePt3High = gudtTest(CHAN0).revForcePt(3).high
rstParameters!RevForcePt3Low = gudtTest(CHAN0).revForcePt(3).low
rstParameters!RevForcePt3Location = gudtTest(CHAN0).revForcePt(3).location
rstParameters!PeakForceHigh = gudtTest(CHAN0).peakForce.high
rstParameters!PeakForceLow = gudtTest(CHAN0).peakForce.low
'Mechanical Hysteresis Points
rstParameters!MechHystPt1Ideal = gudtTest(CHAN0).mechHystPt(1).ideal
rstParameters!MechHystPt1High = gudtTest(CHAN0).mechHystPt(1).high
rstParameters!MechHystPt1Low = gudtTest(CHAN0).mechHystPt(1).low
rstParameters!MechHystPt1Location = gudtTest(CHAN0).mechHystPt(1).location
rstParameters!MechHystPt2Ideal = gudtTest(CHAN0).mechHystPt(2).ideal
rstParameters!MechHystPt2High = gudtTest(CHAN0).mechHystPt(2).high
rstParameters!MechHystPt2Low = gudtTest(CHAN0).mechHystPt(2).low
rstParameters!MechHystPt2Location = gudtTest(CHAN0).mechHystPt(2).location
rstParameters!MechHystPt3Ideal = gudtTest(CHAN0).mechHystPt(3).ideal
rstParameters!MechHystPt3High = gudtTest(CHAN0).mechHystPt(3).high
rstParameters!MechHystPt3Low = gudtTest(CHAN0).mechHystPt(3).low
rstParameters!MechHystPt3Location = gudtTest(CHAN0).mechHystPt(3).location

'Index Point Customer Specifications
rstParameters!IdleOutput1SpecHigh = gudtCustomerSpec(CHAN0).Index(1).high
rstParameters!IdleOutput1SpecLow = gudtCustomerSpec(CHAN0).Index(1).low
rstParameters!MidpointOutput1SpecHigh = gudtCustomerSpec(CHAN0).Index(2).high
rstParameters!MidpointOutput1SpecLow = gudtCustomerSpec(CHAN0).Index(2).low
rstParameters!WOTLocation1SpecHigh = gudtCustomerSpec(CHAN0).Index(3).high
rstParameters!WOTLocation1SpecLow = gudtCustomerSpec(CHAN0).Index(3).low
rstParameters!IdleOutput2SpecHigh = gudtCustomerSpec(CHAN1).Index(1).high
rstParameters!IdleOutput2SpecLow = gudtCustomerSpec(CHAN1).Index(1).low
rstParameters!MidpointOutput2SpecHigh = gudtCustomerSpec(CHAN1).Index(2).high
rstParameters!MidpointOutput2SpecLow = gudtCustomerSpec(CHAN1).Index(2).low
rstParameters!WOTLocation2SpecHigh = gudtCustomerSpec(CHAN1).Index(3).high
rstParameters!WOTLocation2SpecLow = gudtCustomerSpec(CHAN1).Index(3).low
'Force Knee Location Customer Specifications
rstParameters!ForceKneeLocationSpecHigh = gudtCustomerSpec(CHAN0).forceKneeLoc.high
rstParameters!ForceKneeLocationSpecLow = gudtCustomerSpec(CHAN0).forceKneeLoc.low
'Forward Force Point Customer Specifications
rstParameters!FwdForcePt1SpecHigh = gudtCustomerSpec(CHAN0).fwdForcePt(1).high
rstParameters!FwdForcePt1SpecLow = gudtCustomerSpec(CHAN0).fwdForcePt(1).low
rstParameters!FwdForcePt2SpecHigh = gudtCustomerSpec(CHAN0).fwdForcePt(2).high
rstParameters!FwdForcePt2SpecLow = gudtCustomerSpec(CHAN0).fwdForcePt(2).low
rstParameters!FwdForcePt3SpecHigh = gudtCustomerSpec(CHAN0).fwdForcePt(3).high
rstParameters!FwdForcePt3SpecLow = gudtCustomerSpec(CHAN0).fwdForcePt(3).low
'Reverse Force Point Customer Specifications
rstParameters!RevForcePt1SpecHigh = gudtCustomerSpec(CHAN0).revForcePt(1).high
rstParameters!RevForcePt1SpecLow = gudtCustomerSpec(CHAN0).revForcePt(1).low
rstParameters!RevForcePt2SpecHigh = gudtCustomerSpec(CHAN0).revForcePt(2).high
rstParameters!RevForcePt2SpecLow = gudtCustomerSpec(CHAN0).revForcePt(2).low
rstParameters!RevForcePt3SpecHigh = gudtCustomerSpec(CHAN0).revForcePt(3).high
rstParameters!RevForcePt3SpecLow = gudtCustomerSpec(CHAN0).revForcePt(3).low
'Mechanical Hysteresis Point Customer Specifications
rstParameters!MechHystPt1SpecHigh = gudtCustomerSpec(CHAN0).mechHystPt(1).high
rstParameters!MechHystPt1SpecLow = gudtCustomerSpec(CHAN0).mechHystPt(1).low
rstParameters!MechHystPt2SpecHigh = gudtCustomerSpec(CHAN0).mechHystPt(2).high
rstParameters!MechHystPt2SpecLow = gudtCustomerSpec(CHAN0).mechHystPt(2).low
rstParameters!MechHystPt3SpecHigh = gudtCustomerSpec(CHAN0).mechHystPt(3).high
rstParameters!MechHystPt3SpecLow = gudtCustomerSpec(CHAN0).mechHystPt(3).low

'Save the Record
rstParameters.Update

'Assign the Function with the ID of the Record just created
AddScanParametersRecord = rstParameters!ScanParametersID

End Function

Public Sub AddScanResultsRecord()
'
'   PURPOSE: To add the current Scan Results to tblScanResults
'
'  INPUT(S): None
' OUTPUT(S): None

Dim lrstScanResults As New Recordset

On Error GoTo ErrorSavingScanData

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstScanResults.Open "tblScanResults", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstScanResults.AddNew

'Add the new Fields to the Record in tblScanResults
'Keys representing test setup & parameters
lrstScanResults!SerialNumberID = gatDataBaseKey.SerialNumberID
lrstScanResults!MachineParametersID = gatDataBaseKey.MachineParametersID
lrstScanResults!ScanParametersID = gatDataBaseKey.ScanParametersID
lrstScanResults!LotID = gatDataBaseKey.LotID
lrstScanResults!ForceCalID = gatDataBaseKey.ForceCalID
'Software version & Date/Time Stamp
lrstScanResults!SoftwareRevision = App.Major & "." & App.Minor & "." & App.Revision
lrstScanResults!TestDateTime = DateTime.Now
'Pallet Number
lrstScanResults!PalletNum = gintPalletNumber
'Error information
lrstScanResults!Anomaly = gintAnomaly

'If the scan was successful, go ahead and save the results data
If gintAnomaly = 0 Then
    'Scanning Results
    lrstScanResults!PassFail = Not gblnScanFailure
    lrstScanResults!MLXOk = gblnMLXOk '1.9ANM
    'Output #1
    lrstScanResults!IdleValueOutput1 = gudtReading(CHAN0).Index(1).Value
    lrstScanResults!OutputAtForceKneeOutput1 = gudtReading(CHAN0).outputAtForceKnee
    lrstScanResults!MidpointValueOutput1 = gudtReading(CHAN0).Index(2).Value
    lrstScanResults!WOTLocOutput1 = gudtReading(CHAN0).Index(3).location
    lrstScanResults!MaxValueOutput1 = gudtReading(CHAN0).maxOutput.Value
    lrstScanResults!MaxLinDevPerTolValOutput1 = gudtExtreme(CHAN0).linDevPerTol(1).Value
    lrstScanResults!MaxLinDevPerTolLocOutput1 = gudtExtreme(CHAN0).linDevPerTol(1).location
    lrstScanResults!MaxSingLinDevValOutput1 = gudtExtreme(CHAN0).SinglePointLin.high.Value
    lrstScanResults!MaxSingLinDevLocOutput1 = gudtExtreme(CHAN0).SinglePointLin.high.location
    lrstScanResults!MinSingLinDevValOutput1 = gudtExtreme(CHAN0).SinglePointLin.low.Value
    lrstScanResults!MinSingLinDevLocOutput1 = gudtExtreme(CHAN0).SinglePointLin.low.location
    lrstScanResults!MaxLinDevPerTol2ValOutput1 = gudtExtreme(CHAN0).linDevPerTol(2).Value    '1.8ANM \/\/
    lrstScanResults!MaxLinDevPerTol2LocOutput1 = gudtExtreme(CHAN0).linDevPerTol(2).location
    lrstScanResults!MaxAbsLinDevValOutput1 = gudtExtreme(CHAN0).AbsLin.high.Value
    lrstScanResults!MaxAbsLinDevLocOutput1 = gudtExtreme(CHAN0).AbsLin.high.location
    lrstScanResults!MinAbsLinDevValOutput1 = gudtExtreme(CHAN0).AbsLin.low.Value
    lrstScanResults!MinAbsLinDevLocOutput1 = gudtExtreme(CHAN0).AbsLin.low.location          '1.8ANM /\/\
    lrstScanResults!MaxSlopeDevValOutput1 = gudtExtreme(CHAN0).slope.high.Value
    lrstScanResults!MaxSlopeDevLocOutput1 = gudtExtreme(CHAN0).slope.high.location
    lrstScanResults!MinSlopeDevValOutput1 = gudtExtreme(CHAN0).slope.low.Value
    lrstScanResults!MinSlopeDevLocOutput1 = gudtExtreme(CHAN0).slope.low.location
    lrstScanResults!PeakHysteresisValOutput1 = gudtExtreme(CHAN0).hysteresis.Value
    lrstScanResults!PeakHysteresisLocOutput1 = gudtExtreme(CHAN0).hysteresis.location
    lrstScanResults!FullCloseHysOutput1 = gudtReading(CHAN0).FullCloseHys.Value
    lrstScanResults!MLXIddOutput1 = gudtReading(CHAN0).mlxCurrent '1.7ANM
    lrstScanResults!MLXWIddOutput1 = gudtReading(CHAN0).mlxWCurrent '2.0ANM
    'Output #2
    lrstScanResults!IdleValueOutput2 = gudtReading(CHAN1).Index(1).Value
    lrstScanResults!OutputAtForceKneeOutput2 = gudtReading(CHAN1).outputAtForceKnee
    lrstScanResults!MidpointValueOutput2 = gudtReading(CHAN1).Index(2).Value
    lrstScanResults!WOTLocOutput2 = gudtReading(CHAN1).Index(3).location
    lrstScanResults!MaxValueOutput2 = gudtReading(CHAN1).maxOutput.Value
    lrstScanResults!MaxLinDevPerTolValOutput2 = gudtExtreme(CHAN1).linDevPerTol(1).Value
    lrstScanResults!MaxLinDevPerTolLocOutput2 = gudtExtreme(CHAN1).linDevPerTol(1).location
    lrstScanResults!MaxSingLinDevValOutput2 = gudtExtreme(CHAN1).SinglePointLin.high.Value
    lrstScanResults!MaxSingLinDevLocOutput2 = gudtExtreme(CHAN1).SinglePointLin.high.location
    lrstScanResults!MinSingLinDevValOutput2 = gudtExtreme(CHAN1).SinglePointLin.low.Value
    lrstScanResults!MinSingLinDevLocOutput2 = gudtExtreme(CHAN1).SinglePointLin.low.location
    lrstScanResults!MaxLinDevPerTol2ValOutput2 = gudtExtreme(CHAN1).linDevPerTol(2).Value    '1.8ANM \/\/
    lrstScanResults!MaxLinDevPerTol2LocOutput2 = gudtExtreme(CHAN1).linDevPerTol(2).location
    lrstScanResults!MaxAbsLinDevValOutput2 = gudtExtreme(CHAN1).AbsLin.high.Value
    lrstScanResults!MaxAbsLinDevLocOutput2 = gudtExtreme(CHAN1).AbsLin.high.location
    lrstScanResults!MinAbsLinDevValOutput2 = gudtExtreme(CHAN1).AbsLin.low.Value
    lrstScanResults!MinAbsLinDevLocOutput2 = gudtExtreme(CHAN1).AbsLin.low.location          '1.8ANM /\/\
    lrstScanResults!MaxSlopeDevValOutput2 = gudtExtreme(CHAN1).slope.high.Value
    lrstScanResults!MaxSlopeDevLocOutput2 = gudtExtreme(CHAN1).slope.high.location
    lrstScanResults!MinSlopeDevValOutput2 = gudtExtreme(CHAN1).slope.low.Value
    lrstScanResults!MinSlopeDevLocOutput2 = gudtExtreme(CHAN1).slope.low.location
    lrstScanResults!PeakHysteresisValOutput2 = gudtExtreme(CHAN1).hysteresis.Value
    lrstScanResults!PeakHysteresisLocOutput2 = gudtExtreme(CHAN1).hysteresis.location
    lrstScanResults!FullCloseHysOutput2 = gudtReading(CHAN1).FullCloseHys.Value
    lrstScanResults!MLXIddOutput2 = gudtReading(CHAN1).mlxCurrent '1.7ANM
    lrstScanResults!MLXWIddOutput2 = gudtReading(CHAN1).mlxWCurrent '2.0ANM
    'Output Correlation
    lrstScanResults!FwdOutputCorrPerTolVal = gudtExtreme(CHAN0).outputCorPerTol(1).Value
    lrstScanResults!FwdOutputCorrPerTolLoc = gudtExtreme(CHAN0).outputCorPerTol(1).location
    lrstScanResults!MaxFwdOutputCorrVal = gudtExtreme(CHAN0).fwdOutputCor.high.Value
    lrstScanResults!MaxFwdOutputCorrLoc = gudtExtreme(CHAN0).fwdOutputCor.high.location
    lrstScanResults!MinFwdOutputCorrVal = gudtExtreme(CHAN0).fwdOutputCor.low.Value
    lrstScanResults!MinFwdOutputCorrLoc = gudtExtreme(CHAN0).fwdOutputCor.low.location
    lrstScanResults!RevOutputCorrPerTolVal = gudtExtreme(CHAN0).outputCorPerTol(2).Value
    lrstScanResults!RevOutputCorrPerTolLoc = gudtExtreme(CHAN0).outputCorPerTol(2).location
    lrstScanResults!MaxRevOutputCorrVal = gudtExtreme(CHAN0).revOutputCor.high.Value
    lrstScanResults!MaxRevOutputCorrLoc = gudtExtreme(CHAN0).revOutputCor.high.location
    lrstScanResults!MinRevOutputCorrVal = gudtExtreme(CHAN0).revOutputCor.low.Value
    lrstScanResults!MinRevOutputCorrLoc = gudtExtreme(CHAN0).revOutputCor.low.location
    'Pedal-At-Rest Location
    lrstScanResults!pedalAtRestLoc = gudtReading(CHAN0).pedalAtRestLoc
    'Force Knee
    lrstScanResults!forceKneeLoc = gudtReading(CHAN0).forceKnee.location
    lrstScanResults!FwdForceAtForceKneeLoc = gudtReading(CHAN0).forceKnee.Value
    'Force Points
    lrstScanResults!FwdForcePt1 = gudtReading(CHAN0).fwdForcePt(1).Value
    lrstScanResults!FwdForcePt2 = gudtReading(CHAN0).fwdForcePt(2).Value
    lrstScanResults!FwdForcePt3 = gudtReading(CHAN0).fwdForcePt(3).Value
    lrstScanResults!RevForcePt1 = gudtReading(CHAN0).revForcePt(1).Value
    lrstScanResults!RevForcePt2 = gudtReading(CHAN0).revForcePt(2).Value
    lrstScanResults!RevForcePt3 = gudtReading(CHAN0).revForcePt(3).Value
    lrstScanResults!peakForce = gudtReading(CHAN0).peakForce
    lrstScanResults!MechHystPt1 = gudtReading(CHAN0).mechHystPt(1).Value
    lrstScanResults!MechHystPt2 = gudtReading(CHAN0).mechHystPt(2).Value
    lrstScanResults!MechHystPt3 = gudtReading(CHAN0).mechHystPt(3).Value
End If

'User-entered test information
If frmMain.ctrSetupInfo1.Operator = "" Then
    lrstScanResults!Operator = Null 'Save Null if there is an emty string ("")
Else
    lrstScanResults!Operator = frmMain.ctrSetupInfo1.Operator
End If
If frmMain.ctrSetupInfo1.Temperature = "" Then
    lrstScanResults!Temperature = Null 'Save Null if there is an emty string ("")
Else
    lrstScanResults!Temperature = frmMain.ctrSetupInfo1.Temperature
End If
If frmMain.ctrSetupInfo1.Comment = "" Then
    lrstScanResults!Comment = Null 'Save Null if there is an emty string ("")
Else
    lrstScanResults!Comment = frmMain.ctrSetupInfo1.Comment
End If

'Save the Record
lrstScanResults.Update

'Get the Scan Results ID
gatDataBaseKey.ScanResultsID = lrstScanResults!ScanResultsID

'Close the Recordset
lrstScanResults.Close

'If the scan was successful, go ahead and save the raw data
If gintAnomaly = 0 Then
    'Save the raw data associated with the DUT Test
    Call SaveRawData
End If

Exit Sub
ErrorSavingScanData:

    gintAnomaly = 89
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Scanning Results.", True, True)
End Sub

Public Function AddSerialNumberRecord(rstSerialNumber As Recordset) As Long
'
'   PURPOSE: To add the current Serial Number to tblSerialNumber
'
'  INPUT(S): rstSerialNumber: Recordset that represents tblSerialNumber
' OUTPUT(S): key of new Serial Number Record

Dim lintYear As Integer
Dim lintJulianDate As Integer
Dim lstrShift As String
Dim lintStation As Integer
Dim lstrPalletLoad As String

'Decode the Date Code
Call MLX90277.DecodeDateCode(gstrDateCode, lintYear, lintJulianDate, lstrShift, lintStation, lstrPalletLoad)

'Add the new Record
rstSerialNumber.AddNew

'Add the MLX information to the Serial Number table die A
rstSerialNumber!MLX_Y = gudtMLX90277(1).Read.Y
rstSerialNumber!MLX_X = gudtMLX90277(1).Read.X
rstSerialNumber!MLX_Wafer = gudtMLX90277(1).Read.Wafer
rstSerialNumber!MLX_Lot = gudtMLX90277(1).Read.Lot
'Add the MLX information to the Serial Number table die B '1.9ANM
rstSerialNumber!MLX_Y2 = gudtMLX90277(2).Read.Y
rstSerialNumber!MLX_X2 = gudtMLX90277(2).Read.X
rstSerialNumber!MLX_Wafer2 = gudtMLX90277(2).Read.Wafer
rstSerialNumber!MLX_Lot2 = gudtMLX90277(2).Read.Lot
'Add the Date Code information to the Serial Number table
rstSerialNumber!Year = lintYear
rstSerialNumber!JulianDate = lintJulianDate
rstSerialNumber!Shift = lstrShift
rstSerialNumber!Station = lintStation
'rstSerialNumber!PedalLoad = lstrPalletLoad

'Save the Record
rstSerialNumber.Update

'Assign the Function with the ID of the Record just created
AddSerialNumberRecord = rstSerialNumber!SerialNumberID

End Function

Public Sub AddUnserializedProgRecord()
'
'   PURPOSE: To add a record to the table containing information on unserialized
'            parts.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSavingUnserializedProg

Dim lrstUnserializedProg As New Recordset

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstUnserializedProg.Open "tblUnserializedProgram", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstUnserializedProg.AddNew

'Fill the new Record in tblUnserializedProgram
lrstUnserializedProg!MachineParametersID = gatDataBaseKey.MachineParametersID
lrstUnserializedProg!ProgrammingParametersID = gatDataBaseKey.ProgParametersID
lrstUnserializedProg!LotID = gatDataBaseKey.LotID
lrstUnserializedProg!SoftwareRevision = App.Major & "." & App.Minor & "." & App.Revision
lrstUnserializedProg!DateTime = DateTime.Now

'Save the Record
lrstUnserializedProg.Update

'Close the Recordset
lrstUnserializedProg.Close

Exit Sub
ErrorSavingUnserializedProg:
    
    gintAnomaly = 86
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Unserialized Programming Data.", True, True)
End Sub

Public Sub AddUnserializedScanRecord()
'
'   PURPOSE: To add a record to the table containing information on unserialized
'            parts.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSavingUnserializedScan

Dim lrstUnserializedScan As New Recordset

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstUnserializedScan.Open "tblUnserializedScan", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Add the new Record
lrstUnserializedScan.AddNew

'Fill the new Record in tblUnserializedScan
lrstUnserializedScan!MachineParametersID = gatDataBaseKey.MachineParametersID
lrstUnserializedScan!ScanParametersID = gatDataBaseKey.ScanParametersID
lrstUnserializedScan!LotID = gatDataBaseKey.LotID
lrstUnserializedScan!SoftwareRevision = App.Major & "." & App.Minor & "." & App.Revision
lrstUnserializedScan!DateTime = DateTime.Now

'Save the Record
lrstUnserializedScan.Update

'Close the Recordset
lrstUnserializedScan.Close

Exit Sub
ErrorSavingUnserializedScan:
    
    gintAnomaly = 87
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Unserialized Scanning Data.", True, True)
End Sub

Public Sub CloseDatabaseConnection()
'
'   PURPOSE: To close the current database connection.
'
'  INPUT(S): none
' OUTPUT(S): none

'Close the connection if it isn't already
If mcnLocalDatabase.State <> adStateClosed Then mcnLocalDatabase.Close

'Set the status variable to false
mblnConnectionActive = False

End Sub

Public Sub CloseDatabaseConnection2()
'
'   PURPOSE: To close the current database connection.
'
'  INPUT(S): none
' OUTPUT(S): none
'1.3ANM new sub

'Close the connection if it isn't already
If mcnLocalDatabase2.State <> adStateClosed Then mcnLocalDatabase2.Close

'Set the status variable to false
mblnConnectionActive2 = False

End Sub

Public Sub DeactivateDatabase()
'
'   PURPOSE: To deactivate the database
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lrstDBState As New Recordset
Dim lstrSQLString As String

'Build a query to open the entire tblDbState table
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT *"
lstrSQLString = lstrSQLString & " From tblDbState"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstDBState.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Set the IsActive field to False
lrstDBState!IsActive = False

'Update the record now that it's changed
lrstDBState.Update

'Close the Recordset
lrstDBState.Close

End Sub

Public Function FindSerialNumberID() As Boolean
'
'   PURPOSE: To identify if the current Serial Number exists in the
'            database, and if so, record the ID Number for that entry.
'
'  INPUT(S): none
' OUTPUT(S): Function returns whether or not the Serial Number exists

Dim lrstSerialNumber As New Recordset
Dim lstrSQLString As String

On Error GoTo ErrorFindingSerialNumberID

'Build a query to look for the current DUT information
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblSerialNumber.*"
lstrSQLString = lstrSQLString & " From tblSerialNumber"
lstrSQLString = lstrSQLString & " WHERE (((tblSerialNumber.MLX_Lot)=" & gudtMLX90277(1).Read.Lot & ") AND ((tblSerialNumber.MLX_Wafer)=" & gudtMLX90277(1).Read.Wafer & ")AND ((tblSerialNumber.MLX_X)=" & gudtMLX90277(1).Read.X & ")AND ((tblSerialNumber.MLX_Y)=" & gudtMLX90277(1).Read.Y & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstSerialNumber.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstSerialNumber.BOF And lrstSerialNumber.EOF Then
    'The function failed to find the serial number
    FindSerialNumberID = False
Else
    'Get the DUT ID number of the matching recordset
    gatDataBaseKey.SerialNumberID = lrstSerialNumber!SerialNumberID
    'The function found the serial number
    FindSerialNumberID = True
End If

'Close the Recordset
lrstSerialNumber.Close

Exit Function
ErrorFindingSerialNumberID:

    gintAnomaly = 74
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Finding Serial Number ID.", True, True)
End Function

Public Function GetConnectionStatus() As Boolean
'
'   PURPOSE: To return the status of the database connection
'
'  INPUT(S): none
' OUTPUT(S): returns the status of the database connection

GetConnectionStatus = mblnConnectionActive

End Function

Public Sub InitializeDatabaseConnection()
'
'   PURPOSE: To open a database connection to the appropriate database
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lblnConnectionOK As Boolean
Dim lintDatabaseNum As Integer

On Error GoTo DBInitializationErr

'Initialize the connection variable
lblnConnectionOK = False

For lintDatabaseNum = gintDatabaseStartNum To gintDatabaseStopNum
    'Display the current status
    frmMain.staMessage.Panels(1).Text = "Opening Connection to local database number " & CStr(lintDatabaseNum) & "."
    'Build the database name
    mstrDatabaseName = DATABASENAME & CStr(lintDatabaseNum) & DATABASEEXTENSION
    If OpenDatabaseConnection(mstrDatabaseName) Then
        'Now see if the database is active
        If IsDatabaseActive Then
            'Make sure it is not too big
            If Not IsDatabaseTooBig Then
                lblnConnectionOK = True
            Else
                'Display a message to inform the user that the database is beyond its size allowance
                MsgBox "One of the databases is beyond its allowable size.  Please contact Electronics!", vbOKOnly, "Database Problem"
            End If
        End If
    End If
    'If the connection was set, exit the for loop
    If lblnConnectionOK Then
        Exit For
    Else
        'Close the connection before trying to open another one
        Call CloseDatabaseConnection
    End If
Next lintDatabaseNum

If lblnConnectionOK Then
    'Let the user know the connection is set up
    frmMain.staMessage.Panels(1).Text = "Connection to local database number " & CStr(lintDatabaseNum) & " has been initialized."
Else
    gintAnomaly = 70
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Initialization Unsuccessful." & vbCrLf & _
                           "Verify Database is not being used by another program.", True, False)
    'Let the user know there is a problem, then close the program
    MsgBox "The program cannot run without a connection to the database!  Please contact Electronics!", vbCritical & vbOKOnly, "Database Problem"
    Unload frmMain
End If

Exit Sub
DBInitializationErr:

    gintAnomaly = 70
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Initialization Unsuccessful." & vbCrLf & _
                           "Verify Database is not being used by another program.", True, False)
    'Let the user know there is a problem, then close the program
    MsgBox "The program cannot run without a connection to the database!  Please contact Electronics!", vbCritical & vbOKOnly, "Database Problem"
    Unload frmMain

End Sub

Public Sub InitializeSNDatabaseConnection()
'
'   PURPOSE: To open a database connection to the appropriate database
'
'  INPUT(S): none
' OUTPUT(S): none
'1.3ANM new sub

Dim lblnConnectionOK As Boolean
Dim lintDatabaseNum As Integer

On Error GoTo DBInitializationErr

'Initialize the connection variable
lblnConnectionOK = False

'Display the current status
frmMain.staMessage.Panels(1).Text = "Opening Connection to SN database"

'If file does not exist then exit '1.4SRC \/\/
If Not gfsoFileSystemObject.FileExists(SNDATAPATH) Then
    gintAnomaly = 70
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Laser Database Error: Initialization Unsuccessful." & vbCrLf & _
                           "Verify Process PC is online and network connection is working.", True, False)
    'Let the user know there is a problem, then close the program
    MsgBox "The program cannot run without a connection to the laser database!", vbCritical & vbOKOnly, "Process PC Problem"
    Unload frmMain
End If
'If file does not exist then exit '1.4SRC /\/\

'Build the database name
mstrDatabaseName2 = SNDATAPATH
If OpenDatabaseConnection2(mstrDatabaseName2) Then
    lblnConnectionOK = True
End If

'If the connection was not set
If Not lblnConnectionOK Then
    'Close the connection before trying to open another one
    Call CloseDatabaseConnection2
End If

If lblnConnectionOK Then
    'Let the user know the connection is set up
    frmMain.staMessage.Panels(1).Text = "Connection to SN database has been initialized."
Else
    gintAnomaly = 70
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Initialization Unsuccessful." & vbCrLf & _
                           "Verify Database is not being used by another program.", True, False)
    'Let the user know there is a problem, then close the program
    MsgBox "The program cannot run without a connection to the database!  Please contact Electronics!", vbCritical & vbOKOnly, "Database Problem"
    Unload frmMain
End If

Exit Sub
DBInitializationErr:

    gintAnomaly = 70
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Initialization Unsuccessful." & vbCrLf & _
                           "Verify Database is not being used by another program.", True, False)
    'Let the user know there is a problem, then close the program
    MsgBox "The program cannot run without a connection to the database!  Please contact Electronics!", vbCritical & vbOKOnly, "Database Problem"
    Unload frmMain

End Sub

Public Function IsDatabaseActive() As Boolean
'
'   PURPOSE: To determine if the database is active.
'
'  INPUT(S): none
' OUTPUT(S): returns whether or not the database is active

Dim lrstDBState As New Recordset
Dim lstrSQLString As String

'Build a query to open the entire tblDbState table
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT *"
lstrSQLString = lstrSQLString & " From tblDbState"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstDBState.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstDBState.BOF And lrstDBState.EOF Then
    'If no records are returned, then the database is corrupt
    IsDatabaseActive = False
Else
    'Set the function equal to the IsActive field
    IsDatabaseActive = lrstDBState!IsActive
    'Also set the DB_ID
    gatDataBaseKey.DB_ID = lrstDBState!DbId
End If

'Close the Recordset
lrstDBState.Close

End Function

Public Function IsDatabaseTooBig() As Boolean
'
'   PURPOSE: To determine if the database is bigger than its maximum allowable
'            size.
'
'  INPUT(S): none
' OUTPUT(S): returns whether or not the database is too big

'Check the size using the file system object
If gfsoFileSystemObject.GetFile(DATABASEPATH & mstrDatabaseName).Size > MAXDATABASESIZE Then
    IsDatabaseTooBig = True
Else
    IsDatabaseTooBig = False
End If

End Function

Public Function IsOkToSwitch() As Boolean
'
'   PURPOSE: To check if it is ok to switch to the alternate database
'
'  INPUT(S): none
' OUTPUT(S): returns if it is ok to switch to the alternate database

Dim lrstDBState As New Recordset
Dim lstrSQLString As String

'Build a query to open the entire tblDbState table
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT *"
lstrSQLString = lstrSQLString & " From tblDbState"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstDBState.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstDBState.BOF And lrstDBState.EOF Then
    'If no records are returned, then the database is corrupt
    IsOkToSwitch = False
Else
    'Set the function equal to the OkToSwitch field
    IsOkToSwitch = lrstDBState!OkToSwitch
End If

'Close the Recordset
lrstDBState.Close

End Function

Public Function OpenDatabaseConnection(fileName As String) As Boolean
'
'   PURPOSE: To open a database connection to the requested database.
'
'  INPUT(S): none
' OUTPUT(S): returns a boolean based on whether or not there was an error opening

mblnConnectionActive = False

With mcnLocalDatabase

    'Set the modular database variable
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & DATABASEPATH & fileName & ";" & _
                        "Persist Security Info=False"
    'Open the database connection
    .Open

End With
    'The connection to the database is now active
    mblnConnectionActive = True
    OpenDatabaseConnection = mblnConnectionActive

End Function

Public Function OpenDatabaseConnection2(fileName As String) As Boolean
'
'   PURPOSE: To open a database connection to the requested database.
'
'  INPUT(S): none
' OUTPUT(S): returns a boolean based on whether or not there was an error opening

mblnConnectionActive2 = False

With mcnLocalDatabase2

    'Set the modular database variable
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & fileName & ";" & _
                        "Persist Security Info=False"
    'Open the database connection
    .Open

End With
    'The connection to the database is now active
    mblnConnectionActive2 = True
    OpenDatabaseConnection2 = mblnConnectionActive2

End Function

Public Sub ReadForceCalRecord()
'
'   PURPOSE: To read the appropriate Force Calibration data.
'
'  INPUT(S):
' OUTPUT(S):

On Error GoTo ErrorReadingForceCalData

Dim lrstForceCal As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current Force Calibration
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblForceCalibration.*"
lstrSQLString = lstrSQLString & " From tblForceCalibration"
lstrSQLString = lstrSQLString & " WHERE (((tblForceCalibration.ForceCalID)=" & gatDataBaseKey.ForceCalID & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstForceCal.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Read the new Fields to the Record in tblForceCalibration
gsngNewtonsPerVolt = lrstForceCal!ForceCalSensitivity
gsngForceAmplifierOffset = lrstForceCal!ForceCalOffset

'Close the Recordset
lrstForceCal.Close

Exit Sub
ErrorReadingForceCalData:

    gintAnomaly = 90
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Reading Force Calibration Data.", True, True)
End Sub

Public Sub ReadLotRecord()
'
'   PURPOSE: To read the appropriate lot name.
'
'  INPUT(S):
' OUTPUT(S):

On Error GoTo ErrorReadingLotName

Dim lrstLot As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current lot name
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblLot.*"
lstrSQLString = lstrSQLString & " From tblLot"
lstrSQLString = lstrSQLString & " WHERE (((tblLot.LotID)=" & """" & gatDataBaseKey.LotID & """" & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstLot.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Read the new Fields to the Record in tblLot
gstrLotName = lrstLot!LotName

Exit Sub
ErrorReadingLotName:

    gintAnomaly = 91
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Reading Lot Name.", True, True)
End Sub

Public Sub ReadMachineParametersRecord()
'
'   PURPOSE: To read the appropriate machine parameters from tblMachineParameters
'
'  INPUT(S):
' OUTPUT(S):

On Error GoTo ErrorReadingMachineParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblMachineParameters.*"
lstrSQLString = lstrSQLString & " From tblMachineParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblMachineParameters.MachineParametersID)=" & """" & gatDataBaseKey.MachineParametersID & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Key fields - Name, Revision, & Equipment Name
gudtMachine.parameterName = lrstParameters!ParameterFileName
gudtMachine.parameterRev = lrstParameters!ParameterFileRevision
gstrSystemName = lrstParameters!EquipmentName
'Machine Parameters
gudtMachine.seriesID = lrstParameters!seriesID
gudtMachine.stationCode = lrstParameters!stationCode
gudtMachine.PLCCommType = lrstParameters!PLC
gudtMachine.slopeInterval = lrstParameters!SlopeDevInterval
gudtMachine.slopeIncrement = lrstParameters!SlopeDevIncrement
gudtMachine.FKSlope = lrstParameters!FKStartTransitionSlope
gudtMachine.FKWindow = lrstParameters!FKStartTransitionWindow
gudtMachine.FKPercentage = lrstParameters!FKStartTransitionPercentage
gudtMachine.pedalAtRestLocForce = lrstParameters!PedalZeroForce
gudtMachine.loadLocation = lrstParameters!loadLocation
gudtMachine.blockOffset = lrstParameters!HomeBlockOffset
gudtMachine.preScanStart = lrstParameters!preScanStart
gudtMachine.preScanStop = lrstParameters!preScanStop
gudtMachine.preScanVelocity = lrstParameters!preScanVelocity
gudtMachine.preScanAcceleration = lrstParameters!preScanAcceleration
gudtMachine.offset4StartScan = lrstParameters!OffsetForStartScan
gudtMachine.scanLength = lrstParameters!scanLength
gudtMachine.overTravel = lrstParameters!overTravel
gudtMachine.scanVelocity = lrstParameters!scanVelocity
gudtMachine.scanAcceleration = lrstParameters!scanAcceleration
gudtMachine.progVelocity = lrstParameters!progVelocity
gudtMachine.progAcceleration = lrstParameters!progAcceleration
gudtMachine.countsPerTrigger = lrstParameters!EncCntPerDataPt
gudtMachine.gearRatio = lrstParameters!gearRatio
gudtMachine.encReso = lrstParameters!EncoderResolution
gudtMachine.graphZeroOffset = lrstParameters!graphZeroOffset
gudtMachine.xAxisLow = lrstParameters!xAxisLow
gudtMachine.xAxisHigh = lrstParameters!xAxisHigh
gudtMachine.filterLoc(CHAN0) = lrstParameters!Filter1Location
gudtMachine.filterLoc(CHAN1) = lrstParameters!Filter2Location
gudtMachine.filterLoc(CHAN2) = lrstParameters!Filter3Location
gudtMachine.filterLoc(CHAN3) = lrstParameters!Filter4Location
gudtMachine.VRefMode = lrstParameters!VRefMode
gudtMachine.maxLBF = lrstParameters!maxLBF

Exit Sub
ErrorReadingMachineParameters:

    gintAnomaly = 92
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Reading Machine Parameters.", True, True)
End Sub

Public Function ReadOutput1ParametersRecord() As Long
'
'   PURPOSE: To read the appropriate Output #2 parameters from
'            tblProgrammingParameters
'
'  INPUT(S):
' OUTPUT(S):

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblOutput1Parameters.*"
lstrSQLString = lstrSQLString & " From tblOutput1Parameters"
lstrSQLString = lstrSQLString & " WHERE (((tblOutput1Parameters.Output1ParametersID)=" & """" & gatDataBaseKey.Output1Parameters & """" & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open "tblOutput1Parameters", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Index 1 Parameters
gudtTest(CHAN0).Index(1).ideal = lrstParameters!IdleIdeal
gudtTest(CHAN0).Index(1).high = lrstParameters!IdleHighLimit
gudtTest(CHAN0).Index(1).low = lrstParameters!IdleLowLimit
gudtTest(CHAN0).Index(1).location = lrstParameters!IdleLoc
'Output At Force Knee Parameters
gudtTest(CHAN0).outputAtForceKnee.ideal = lrstParameters!OutputAtForceKneeIdeal
gudtTest(CHAN0).outputAtForceKnee.high = lrstParameters!OutputAtForceKneeRelativeHighLimit
gudtTest(CHAN0).outputAtForceKnee.low = lrstParameters!OutputAtForceKneeLowLimit
'Index 2 Parameters
gudtTest(CHAN0).Index(2).ideal = lrstParameters!MidPointIdeal
gudtTest(CHAN0).Index(2).high = lrstParameters!MidPointHighLimit
gudtTest(CHAN0).Index(2).low = lrstParameters!MidPointLowLimit
gudtTest(CHAN0).Index(2).location = lrstParameters!MidPointLoc
'Index 3 Parameters
gudtTest(CHAN0).Index(3).ideal = lrstParameters!WOTIdealLoc
gudtTest(CHAN0).Index(3).high = lrstParameters!WOTLocHighLimit
gudtTest(CHAN0).Index(3).low = lrstParameters!WOTLocLowLimit
gudtTest(CHAN0).Index(3).location = lrstParameters!WOTValue
'Maximum Output Parameters
gudtTest(CHAN0).maxOutput.high = lrstParameters!MaxOutputHighLimit
gudtTest(CHAN0).maxOutput.low = lrstParameters!MaxOutputLowLimit
'SinglePoint Linearity Parameters
gudtTest(CHAN0).SinglePointLin(1).start.location = lrstParameters!SingLinReg1StartLoc
gudtTest(CHAN0).SinglePointLin(1).start.high = lrstParameters!SingLinReg1StartHighLimit
gudtTest(CHAN0).SinglePointLin(1).start.low = lrstParameters!SingLinReg1StartLowLimit
gudtTest(CHAN0).SinglePointLin(1).stop.location = lrstParameters!SingLinReg1StopLoc
gudtTest(CHAN0).SinglePointLin(1).stop.high = lrstParameters!SingLinReg1StopHighLimit
gudtTest(CHAN0).SinglePointLin(1).stop.low = lrstParameters!SingLinReg1StopLowLimit
gudtTest(CHAN0).SinglePointLin(2).start.location = lrstParameters!SingLinReg2StartLoc
gudtTest(CHAN0).SinglePointLin(2).start.high = lrstParameters!SingLinReg2StartHighLimit
gudtTest(CHAN0).SinglePointLin(2).start.low = lrstParameters!SingLinReg2StartLowLimit
gudtTest(CHAN0).SinglePointLin(2).stop.location = lrstParameters!SingLinReg2StopLoc
gudtTest(CHAN0).SinglePointLin(2).stop.high = lrstParameters!SingLinReg2StopHighLimit
gudtTest(CHAN0).SinglePointLin(2).stop.low = lrstParameters!SingLinReg2StopLowLimit
gudtTest(CHAN0).SinglePointLin(3).start.location = lrstParameters!SingLinReg3StartLoc
gudtTest(CHAN0).SinglePointLin(3).start.high = lrstParameters!SingLinReg3StartHighLimit
gudtTest(CHAN0).SinglePointLin(3).start.low = lrstParameters!SingLinReg3StartLowLimit
gudtTest(CHAN0).SinglePointLin(3).stop.location = lrstParameters!SingLinReg3StopLoc
gudtTest(CHAN0).SinglePointLin(3).stop.high = lrstParameters!SingLinReg3StopHighLimit
gudtTest(CHAN0).SinglePointLin(3).stop.low = lrstParameters!SingLinReg3StopLowLimit
gudtTest(CHAN0).SinglePointLin(4).start.location = lrstParameters!SingLinReg4StartLoc
gudtTest(CHAN0).SinglePointLin(4).start.high = lrstParameters!SingLinReg4StartHighLimit
gudtTest(CHAN0).SinglePointLin(4).start.low = lrstParameters!SingLinReg4StartLowLimit
gudtTest(CHAN0).SinglePointLin(4).stop.location = lrstParameters!SingLinReg4StopLoc
gudtTest(CHAN0).SinglePointLin(4).stop.high = lrstParameters!SingLinReg4StopHighLimit
gudtTest(CHAN0).SinglePointLin(4).stop.low = lrstParameters!SingLinReg4StopLowLimit
gudtTest(CHAN0).SinglePointLin(5).start.location = lrstParameters!SingLinReg5StartLoc
gudtTest(CHAN0).SinglePointLin(5).start.high = lrstParameters!SingLinReg5StartHighLimit
gudtTest(CHAN0).SinglePointLin(5).start.low = lrstParameters!SingLinReg5StartLowLimit
gudtTest(CHAN0).SinglePointLin(5).stop.location = lrstParameters!SingLinReg5StopLoc
gudtTest(CHAN0).SinglePointLin(5).stop.high = lrstParameters!SingLinReg5StopHighLimit
gudtTest(CHAN0).SinglePointLin(5).stop.low = lrstParameters!SingLinReg5StopLowLimit
'Slope Deviation Parameters
gudtTest(CHAN0).slope.ideal = lrstParameters!IdealSlope
gudtTest(CHAN0).slope.high = lrstParameters!SlopeDevHighLimit
gudtTest(CHAN0).slope.low = lrstParameters!SlopeDevLowLimit
gudtTest(CHAN0).slope.start = lrstParameters!SlopeDevStartLoc
gudtTest(CHAN0).slope.stop = lrstParameters!SlopeDevStopLoc
gudtTest(CHAN0).FullCloseHys.ideal = lrstParameters!FullCloseHysIdeal
gudtTest(CHAN0).FullCloseHys.high = lrstParameters!FullCloseHighLimit
gudtTest(CHAN0).FullCloseHys.low = lrstParameters!FullCloseLowLimit
gudtTest(CHAN0).evaluate.start = lrstParameters!EvaluationStartLoc
gudtTest(CHAN0).evaluate.stop = lrstParameters!EvaluationStopLoc

'Close the Recordset
lrstParameters.Close

End Function

Public Function ReadOutput2ParametersRecord() As Long
'
'   PURPOSE: To read the appropriate Output #2 parameters from
'            tblProgrammingParameters
'
'  INPUT(S):
' OUTPUT(S):

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblOutput2Parameters.*"
lstrSQLString = lstrSQLString & " From tblOutput2Parameters"
lstrSQLString = lstrSQLString & " WHERE (((tblOutput2Parameters.Output2ParametersID)=" & """" & gatDataBaseKey.Output2Parameters & """" & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open "tblOutput2Parameters", mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Index 1 Parameters
gudtTest(CHAN1).Index(1).ideal = lrstParameters!IdleIdeal
gudtTest(CHAN1).Index(1).high = lrstParameters!IdleHighLimit
gudtTest(CHAN1).Index(1).low = lrstParameters!IdleLowLimit
gudtTest(CHAN1).Index(1).location = lrstParameters!IdleLoc
'Output At Force Knee Parameters
gudtTest(CHAN1).outputAtForceKnee.ideal = lrstParameters!OutputAtForceKneeIdeal
gudtTest(CHAN1).outputAtForceKnee.high = lrstParameters!OutputAtForceKneeRelativeHighLimit
gudtTest(CHAN1).outputAtForceKnee.low = lrstParameters!OutputAtForceKneeLowLimit
'Index 2 Parameters
gudtTest(CHAN1).Index(2).ideal = lrstParameters!MidPointIdeal
gudtTest(CHAN1).Index(2).high = lrstParameters!MidPointHighLimit
gudtTest(CHAN1).Index(2).low = lrstParameters!MidPointLowLimit
gudtTest(CHAN1).Index(2).location = lrstParameters!MidPointLoc
'Index 3 Parameters
gudtTest(CHAN1).Index(3).ideal = lrstParameters!WOTIdealLoc
gudtTest(CHAN1).Index(3).high = lrstParameters!WOTLocHighLimit
gudtTest(CHAN1).Index(3).low = lrstParameters!WOTLocLowLimit
gudtTest(CHAN1).Index(3).location = lrstParameters!WOTValue
'Maximum Output Parameters
gudtTest(CHAN1).maxOutput.high = lrstParameters!MaxOutputHighLimit
gudtTest(CHAN1).maxOutput.low = lrstParameters!MaxOutputLowLimit
'SinglePoint Linearity Parameters
gudtTest(CHAN1).SinglePointLin(1).start.location = lrstParameters!SingLinReg1StartLoc
gudtTest(CHAN1).SinglePointLin(1).start.high = lrstParameters!SingLinReg1StartHighLimit
gudtTest(CHAN1).SinglePointLin(1).start.low = lrstParameters!SingLinReg1StartLowLimit
gudtTest(CHAN1).SinglePointLin(1).stop.location = lrstParameters!SingLinReg1StopLoc
gudtTest(CHAN1).SinglePointLin(1).stop.high = lrstParameters!SingLinReg1StopHighLimit
gudtTest(CHAN1).SinglePointLin(1).stop.low = lrstParameters!SingLinReg1StopLowLimit
gudtTest(CHAN1).SinglePointLin(2).start.location = lrstParameters!SingLinReg2StartLoc
gudtTest(CHAN1).SinglePointLin(2).start.high = lrstParameters!SingLinReg2StartHighLimit
gudtTest(CHAN1).SinglePointLin(2).start.low = lrstParameters!SingLinReg2StartLowLimit
gudtTest(CHAN1).SinglePointLin(2).stop.location = lrstParameters!SingLinReg2StopLoc
gudtTest(CHAN1).SinglePointLin(2).stop.high = lrstParameters!SingLinReg2StopHighLimit
gudtTest(CHAN1).SinglePointLin(2).stop.low = lrstParameters!SingLinReg2StopLowLimit
gudtTest(CHAN1).SinglePointLin(3).start.location = lrstParameters!SingLinReg3StartLoc
gudtTest(CHAN1).SinglePointLin(3).start.high = lrstParameters!SingLinReg3StartHighLimit
gudtTest(CHAN1).SinglePointLin(3).start.low = lrstParameters!SingLinReg3StartLowLimit
gudtTest(CHAN1).SinglePointLin(3).stop.location = lrstParameters!SingLinReg3StopLoc
gudtTest(CHAN1).SinglePointLin(3).stop.high = lrstParameters!SingLinReg3StopHighLimit
gudtTest(CHAN1).SinglePointLin(3).stop.low = lrstParameters!SingLinReg3StopLowLimit
gudtTest(CHAN1).SinglePointLin(4).start.location = lrstParameters!SingLinReg4StartLoc
gudtTest(CHAN1).SinglePointLin(4).start.high = lrstParameters!SingLinReg4StartHighLimit
gudtTest(CHAN1).SinglePointLin(4).start.low = lrstParameters!SingLinReg4StartLowLimit
gudtTest(CHAN1).SinglePointLin(4).stop.location = lrstParameters!SingLinReg4StopLoc
gudtTest(CHAN1).SinglePointLin(4).stop.high = lrstParameters!SingLinReg4StopHighLimit
gudtTest(CHAN1).SinglePointLin(4).stop.low = lrstParameters!SingLinReg4StopLowLimit
gudtTest(CHAN1).SinglePointLin(5).start.location = lrstParameters!SingLinReg5StartLoc
gudtTest(CHAN1).SinglePointLin(5).start.high = lrstParameters!SingLinReg5StartHighLimit
gudtTest(CHAN1).SinglePointLin(5).start.low = lrstParameters!SingLinReg5StartLowLimit
gudtTest(CHAN1).SinglePointLin(5).stop.location = lrstParameters!SingLinReg5StopLoc
gudtTest(CHAN1).SinglePointLin(5).stop.high = lrstParameters!SingLinReg5StopHighLimit
gudtTest(CHAN1).SinglePointLin(5).stop.low = lrstParameters!SingLinReg5StopLowLimit
'Slope Deviation Parameters
gudtTest(CHAN1).slope.ideal = lrstParameters!IdealSlope
gudtTest(CHAN1).slope.high = lrstParameters!SlopeDevHighLimit
gudtTest(CHAN1).slope.low = lrstParameters!SlopeDevLowLimit
gudtTest(CHAN1).slope.start = lrstParameters!SlopeDevStartLoc
gudtTest(CHAN1).slope.stop = lrstParameters!SlopeDevStopLoc
gudtTest(CHAN1).FullCloseHys.ideal = lrstParameters!FullCloseHysIdeal
gudtTest(CHAN1).FullCloseHys.high = lrstParameters!FullCloseHighLimit
gudtTest(CHAN1).FullCloseHys.low = lrstParameters!FullCloseLowLimit
gudtTest(CHAN1).evaluate.start = lrstParameters!EvaluationStartLoc
gudtTest(CHAN1).evaluate.stop = lrstParameters!EvaluationStopLoc

'Close the Recordset
lrstParameters.Close

End Function

Public Sub ReadProgrammingParametersRecord()
'
'   PURPOSE: To read the appropriate programming parameters from
'            tblProgrammingParameters
'
'  INPUT(S):
' OUTPUT(S):

On Error GoTo ErrorReadingProgrammingParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblProgrammingParameters.*"
lstrSQLString = lstrSQLString & " From tblProgrammingParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblProgrammingParameters.ProgrammingParametersID)=" & """" & gatDataBaseKey.ProgParametersID & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Key fields - Name & Revision
gudtMachine.parameterName = lrstParameters!ParameterFileName
gudtMachine.parameterRev = lrstParameters!ParameterFileRevision
gstrMLX90277Revision = lrstParameters!MLX90277RevLevel
'Programming Times Output #1
gudtPTC04(1).Tpor = lrstParameters!TporOutput1
gudtPTC04(1).Thold = lrstParameters!TholdOutput1
gudtPTC04(1).Tprog = lrstParameters!TprogOutput1
gudtPTC04(1).Tpuls = lrstParameters!TpulsOutput1
'Solver Parameters Output #1
gudtSolver(1).Index(1).IdealValue = lrstParameters!Index1IdealOutput1
gudtSolver(1).Index(1).IdealLocation = lrstParameters!Index1LocationOutput1
gudtSolver(1).Index(1).TargetTolerance = lrstParameters!Index1TargetToleranceOutput1
gudtSolver(1).Index(1).PassFailTolerance = lrstParameters!Index1PassFailToleranceOutput1
gudtSolver(1).Index(2).IdealValue = lrstParameters!Index2IdealOutput1
gudtSolver(1).Index(2).IdealLocation = lrstParameters!Index2LocationOutput1
gudtSolver(1).Index(2).TargetTolerance = lrstParameters!Index2TargetToleranceOutput1
gudtSolver(1).Index(2).PassFailTolerance = lrstParameters!Index2PassFailToleranceOutput1
gudtSolver(1).Filter = lrstParameters!FilterOutput1
gudtSolver(1).InvertSlope = lrstParameters!InvertOutput1
gudtSolver(1).Mode = lrstParameters!ModeOutput1
gudtSolver(1).FaultLevel = lrstParameters!FaultLevelOutput1
gudtSolver(1).MaxOffsetDrift = lrstParameters!MaxOffsetDriftOutput1
gudtSolver(1).MaxAGND = lrstParameters!MaxAGNDSettingOutput1
gudtSolver(1).MinAGND = lrstParameters!MinAGNDSettingOutput1
gudtSolver(1).FCKADJ = lrstParameters!FCKADJSettingOutput1
gudtSolver(1).CKANACH = lrstParameters!CKANACHSettingOutput1
gudtSolver(1).CKDACCH = lrstParameters!CKDACCHSettingOutput1
gudtSolver(1).SlowMode = lrstParameters!SlowModeSettingOutput1
gudtSolver(1).InitialOffset = lrstParameters!InitialOffsetOutput1
gudtSolver(1).HighRGHighFG = lrstParameters!HighRGHighFG1  '1.1ANM
gudtSolver(1).HighRGLowFG = lrstParameters!HighRGLowFG1    '1.1ANM
gudtSolver(1).LowRGHighFG = lrstParameters!LowRGHighFG1    '1.1ANM
gudtSolver(1).LowRGLowFG = lrstParameters!LowRGLowFG1      '1.1ANM
gudtSolver(1).MinRG = lrstParameters!MinRoughGainOutput1
gudtSolver(1).MaxRG = lrstParameters!MaxRoughGainOutput1
gudtSolver(1).OffsetStep = lrstParameters!OffsetStepOutput1
gudtSolver(1).CodeRatio(1, 1) = lrstParameters!RatioA1Output1
gudtSolver(1).CodeRatio(1, 2) = lrstParameters!RatioA2Output1
gudtSolver(1).CodeRatio(1, 3) = lrstParameters!RatioA3Output1
gudtSolver(1).CodeRatio(2, 1) = lrstParameters!RatioB1Output1
gudtSolver(1).CodeRatio(2, 2) = lrstParameters!RatioB2Output1
gudtSolver(1).CodeRatio(2, 3) = lrstParameters!RatioB3Output1
gudtSolver(1).Clamp(1).IdealValue = lrstParameters!ClampLowIdealOutput1
gudtSolver(1).Clamp(1).TargetTolerance = lrstParameters!ClampLowTargetToleranceOutput1
gudtSolver(1).Clamp(1).PassFailTolerance = lrstParameters!ClampLowPassFailToleranceOutput1
gudtSolver(1).Clamp(1).InitialCode = lrstParameters!ClampLowInitialCodeOutput1
gudtSolver(1).Clamp(2).IdealValue = lrstParameters!ClampHighIdealOutput1
gudtSolver(1).Clamp(2).TargetTolerance = lrstParameters!ClampHighTargetToleranceOutput1
gudtSolver(1).Clamp(2).PassFailTolerance = lrstParameters!ClampHighPassFailToleranceOutput1
gudtSolver(1).Clamp(2).InitialCode = lrstParameters!ClampHighInitialCodeOutput1
gudtSolver(1).ClampStep = lrstParameters!ClampStepOutput1

'Programming Times Output #2
gudtPTC04(2).Tpor = lrstParameters!TporOutput2
gudtPTC04(2).Thold = lrstParameters!TholdOutput2
gudtPTC04(2).Tprog = lrstParameters!TprogOutput2
gudtPTC04(2).Tpuls = lrstParameters!TpulsOutput2
'Solver Parameters Output #2
gudtSolver(2).Index(1).IdealValue = lrstParameters!Index1IdealOutput2
gudtSolver(2).Index(1).IdealLocation = lrstParameters!Index1LocationOutput2
gudtSolver(2).Index(1).TargetTolerance = lrstParameters!Index1TargetToleranceOutput2
gudtSolver(2).Index(1).PassFailTolerance = lrstParameters!Index1PassFailToleranceOutput2
gudtSolver(2).Index(2).IdealValue = lrstParameters!Index2IdealOutput2
gudtSolver(2).Index(2).IdealLocation = lrstParameters!Index2LocationOutput2
gudtSolver(2).Index(2).TargetTolerance = lrstParameters!Index2TargetToleranceOutput2
gudtSolver(2).Index(2).PassFailTolerance = lrstParameters!Index2PassFailToleranceOutput2
gudtSolver(2).Filter = lrstParameters!FilterOutput2
gudtSolver(2).InvertSlope = lrstParameters!InvertOutput2
gudtSolver(2).Mode = lrstParameters!ModeOutput2
gudtSolver(2).FaultLevel = lrstParameters!FaultLevelOutput2
gudtSolver(2).MaxOffsetDrift = lrstParameters!MaxOffsetDriftOutput2
gudtSolver(2).MaxAGND = lrstParameters!MaxAGNDSettingOutput2
gudtSolver(2).MinAGND = lrstParameters!MinAGNDSettingOutput2
gudtSolver(2).FCKADJ = lrstParameters!FCKADJSettingOutput2
gudtSolver(2).CKANACH = lrstParameters!CKANACHSettingOutput2
gudtSolver(2).CKDACCH = lrstParameters!CKDACCHSettingOutput2
gudtSolver(2).SlowMode = lrstParameters!SlowModeSettingOutput2
gudtSolver(2).InitialOffset = lrstParameters!InitialOffsetOutput2
gudtSolver(2).HighRGHighFG = lrstParameters!HighRGHighFG2  '1.1ANM
gudtSolver(2).HighRGLowFG = lrstParameters!HighRGLowFG2    '1.1ANM
gudtSolver(2).LowRGHighFG = lrstParameters!LowRGHighFG2    '1.1ANM
gudtSolver(2).LowRGLowFG = lrstParameters!LowRGLowFG2      '1.1ANM
gudtSolver(2).MinRG = lrstParameters!MinRoughGainOutput2
gudtSolver(2).MaxRG = lrstParameters!MaxRoughGainOutput2
gudtSolver(2).OffsetStep = lrstParameters!OffsetStepOutput2
gudtSolver(2).CodeRatio(1, 1) = lrstParameters!RatioA1Output2
gudtSolver(2).CodeRatio(1, 2) = lrstParameters!RatioA2Output2
gudtSolver(2).CodeRatio(1, 3) = lrstParameters!RatioA3Output2
gudtSolver(2).CodeRatio(2, 1) = lrstParameters!RatioB1Output2
gudtSolver(2).CodeRatio(2, 2) = lrstParameters!RatioB2Output2
gudtSolver(2).CodeRatio(2, 3) = lrstParameters!RatioB3Output2
gudtSolver(2).Clamp(1).IdealValue = lrstParameters!ClampLowIdealOutput2
gudtSolver(2).Clamp(1).TargetTolerance = lrstParameters!ClampLowTargetToleranceOutput2
gudtSolver(2).Clamp(1).PassFailTolerance = lrstParameters!ClampLowPassFailToleranceOutput2
gudtSolver(2).Clamp(1).InitialCode = lrstParameters!ClampLowInitialCodeOutput2
gudtSolver(2).Clamp(2).IdealValue = lrstParameters!ClampHighIdealOutput2
gudtSolver(2).Clamp(2).TargetTolerance = lrstParameters!ClampHighTargetToleranceOutput2
gudtSolver(2).Clamp(2).PassFailTolerance = lrstParameters!ClampHighPassFailToleranceOutput2
gudtSolver(2).Clamp(2).InitialCode = lrstParameters!ClampHighInitialCodeOutput2
gudtSolver(2).ClampStep = lrstParameters!ClampStepOutput2

Exit Sub
ErrorReadingProgrammingParameters:

    gintAnomaly = 93
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Reading Programming Parameter.", True, True)
End Sub

Public Sub ReadScanParametersRecord()
'
'   PURPOSE: To read the appropriate programming parameters from
'            tblScanParameters, tblOutput1Parameters, & tblOutput2Parameters
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorReadingScanParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblScanParameters.*"
lstrSQLString = lstrSQLString & " From tblScanParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblScanParameters.ScanParametersID)=" & """" & gatDataBaseKey.ScanParametersID & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'Key fields - Name & Revision
gudtMachine.parameterName = lrstParameters!ParameterFileName
gudtMachine.parameterRev = lrstParameters!ParameterFileRevision

'Add entries in the PedalOutputParameter tables for both outputs
gatDataBaseKey.Output1Parameters = lrstParameters!Output1ID
gatDataBaseKey.Output2Parameters = lrstParameters!Output2ID

'Forward Output Correlation Parameters
gudtTest(CHAN0).fwdOutputCor(1).start.location = lrstParameters!FwdOutputCorrelationReg1StartLoc
gudtTest(CHAN0).fwdOutputCor(1).start.high = lrstParameters!FwdOutputCorrelationReg1StartHigh
gudtTest(CHAN0).fwdOutputCor(1).start.low = lrstParameters!FwdOutputCorrelationReg1StartLow
gudtTest(CHAN0).fwdOutputCor(1).stop.location = lrstParameters!FwdOutputCorrelationReg1StopLoc
gudtTest(CHAN0).fwdOutputCor(1).stop.high = lrstParameters!FwdOutputCorrelationReg1StopHigh
gudtTest(CHAN0).fwdOutputCor(1).stop.low = lrstParameters!FwdOutputCorrelationReg1StopLow
gudtTest(CHAN0).fwdOutputCor(2).start.location = lrstParameters!FwdOutputCorrelationReg2StartLoc
gudtTest(CHAN0).fwdOutputCor(2).start.high = lrstParameters!FwdOutputCorrelationReg2StartHigh
gudtTest(CHAN0).fwdOutputCor(2).start.low = lrstParameters!FwdOutputCorrelationReg2StartLow
gudtTest(CHAN0).fwdOutputCor(2).stop.location = lrstParameters!FwdOutputCorrelationReg2StopLoc
gudtTest(CHAN0).fwdOutputCor(2).stop.high = lrstParameters!FwdOutputCorrelationReg2StopHigh
gudtTest(CHAN0).fwdOutputCor(2).stop.low = lrstParameters!FwdOutputCorrelationReg2StopLow
gudtTest(CHAN0).fwdOutputCor(3).start.location = lrstParameters!FwdOutputCorrelationReg3StartLoc
gudtTest(CHAN0).fwdOutputCor(3).start.high = lrstParameters!FwdOutputCorrelationReg3StartHigh
gudtTest(CHAN0).fwdOutputCor(3).start.low = lrstParameters!FwdOutputCorrelationReg3StartLow
gudtTest(CHAN0).fwdOutputCor(3).stop.location = lrstParameters!FwdOutputCorrelationReg3StopLoc
gudtTest(CHAN0).fwdOutputCor(3).stop.high = lrstParameters!FwdOutputCorrelationReg3StopHigh
gudtTest(CHAN0).fwdOutputCor(3).stop.low = lrstParameters!FwdOutputCorrelationReg3StopLow
gudtTest(CHAN0).fwdOutputCor(4).start.location = lrstParameters!FwdOutputCorrelationReg4StartLoc
gudtTest(CHAN0).fwdOutputCor(4).start.high = lrstParameters!FwdOutputCorrelationReg4StartHigh
gudtTest(CHAN0).fwdOutputCor(4).start.low = lrstParameters!FwdOutputCorrelationReg4StartLow
gudtTest(CHAN0).fwdOutputCor(4).stop.location = lrstParameters!FwdOutputCorrelationReg4StopLoc
gudtTest(CHAN0).fwdOutputCor(4).stop.high = lrstParameters!FwdOutputCorrelationReg4StopHigh
gudtTest(CHAN0).fwdOutputCor(4).stop.low = lrstParameters!FwdOutputCorrelationReg4StopLow
gudtTest(CHAN0).fwdOutputCor(5).start.location = lrstParameters!FwdOutputCorrelationReg5StartLoc
gudtTest(CHAN0).fwdOutputCor(5).start.high = lrstParameters!FwdOutputCorrelationReg5StartHigh
gudtTest(CHAN0).fwdOutputCor(5).start.low = lrstParameters!FwdOutputCorrelationReg5StartLow
gudtTest(CHAN0).fwdOutputCor(5).stop.location = lrstParameters!FwdOutputCorrelationReg5StopLoc
gudtTest(CHAN0).fwdOutputCor(5).stop.high = lrstParameters!FwdOutputCorrelationReg5StopHigh
gudtTest(CHAN0).fwdOutputCor(5).stop.low = lrstParameters!FwdOutputCorrelationReg5StopLow
'Reverse Output Correlation Parameters
gudtTest(CHAN0).revOutputCor(1).start.location = lrstParameters!RevOutputCorrelationReg1StartLoc
gudtTest(CHAN0).revOutputCor(1).start.high = lrstParameters!RevOutputCorrelationReg1StartHigh
gudtTest(CHAN0).revOutputCor(1).start.low = lrstParameters!RevOutputCorrelationReg1StartLow
gudtTest(CHAN0).revOutputCor(1).stop.location = lrstParameters!RevOutputCorrelationReg1StopLoc
gudtTest(CHAN0).revOutputCor(1).stop.high = lrstParameters!RevOutputCorrelationReg1StopHigh
gudtTest(CHAN0).revOutputCor(1).stop.low = lrstParameters!RevOutputCorrelationReg1StopLow
gudtTest(CHAN0).revOutputCor(2).start.location = lrstParameters!RevOutputCorrelationReg2StartLoc
gudtTest(CHAN0).revOutputCor(2).start.high = lrstParameters!RevOutputCorrelationReg2StartHigh
gudtTest(CHAN0).revOutputCor(2).start.low = lrstParameters!RevOutputCorrelationReg2StartLow
gudtTest(CHAN0).revOutputCor(2).stop.location = lrstParameters!RevOutputCorrelationReg2StopLoc
gudtTest(CHAN0).revOutputCor(2).stop.high = lrstParameters!RevOutputCorrelationReg2StopHigh
gudtTest(CHAN0).revOutputCor(2).stop.low = lrstParameters!RevOutputCorrelationReg2StopLow
gudtTest(CHAN0).revOutputCor(3).start.location = lrstParameters!RevOutputCorrelationReg3StartLoc
gudtTest(CHAN0).revOutputCor(3).start.high = lrstParameters!RevOutputCorrelationReg3StartHigh
gudtTest(CHAN0).revOutputCor(3).start.low = lrstParameters!RevOutputCorrelationReg3StartLow
gudtTest(CHAN0).revOutputCor(3).stop.location = lrstParameters!RevOutputCorrelationReg3StopLoc
gudtTest(CHAN0).revOutputCor(3).stop.high = lrstParameters!RevOutputCorrelationReg3StopHigh
gudtTest(CHAN0).revOutputCor(3).stop.low = lrstParameters!RevOutputCorrelationReg3StopLow
gudtTest(CHAN0).revOutputCor(4).start.location = lrstParameters!RevOutputCorrelationReg4StartLoc
gudtTest(CHAN0).revOutputCor(4).start.high = lrstParameters!RevOutputCorrelationReg4StartHigh
gudtTest(CHAN0).revOutputCor(4).start.low = lrstParameters!RevOutputCorrelationReg4StartLow
gudtTest(CHAN0).revOutputCor(4).stop.location = lrstParameters!RevOutputCorrelationReg4StopLoc
gudtTest(CHAN0).revOutputCor(4).stop.high = lrstParameters!RevOutputCorrelationReg4StopHigh
gudtTest(CHAN0).revOutputCor(4).stop.low = lrstParameters!RevOutputCorrelationReg4StopLow
gudtTest(CHAN0).revOutputCor(5).start.location = lrstParameters!RevOutputCorrelationReg5StartLoc
gudtTest(CHAN0).revOutputCor(5).start.high = lrstParameters!RevOutputCorrelationReg5StartHigh
gudtTest(CHAN0).revOutputCor(5).start.low = lrstParameters!RevOutputCorrelationReg5StartLow
gudtTest(CHAN0).revOutputCor(5).stop.location = lrstParameters!RevOutputCorrelationReg5StopLoc
gudtTest(CHAN0).revOutputCor(5).stop.high = lrstParameters!RevOutputCorrelationReg5StopHigh
gudtTest(CHAN0).revOutputCor(5).stop.low = lrstParameters!RevOutputCorrelationReg5StopLow
'Pedal-At-Rest Location Location
gudtTest(CHAN0).pedalAtRestLoc.ideal = lrstParameters!pedalAtRestLocationIdeal
'Force Knee Location
gudtTest(CHAN0).forceKneeLoc.ideal = lrstParameters!ForceKneeLocationIdeal
gudtTest(CHAN0).forceKneeLoc.high = lrstParameters!ForceKneeLocationHigh
gudtTest(CHAN0).forceKneeLoc.low = lrstParameters!ForceKneeLocationLow
gudtTest(CHAN0).forceKneeForce.ideal = lrstParameters!FwdForceAtForceKneeLocationIdeal
gudtTest(CHAN0).forceKneeForce.high = lrstParameters!FwdForceAtForceKneeLocationHigh
gudtTest(CHAN0).forceKneeForce.low = lrstParameters!FwdForceAtForceKneeLocationLow
'Forward Force Points
gudtTest(CHAN0).fwdForcePt(1).ideal = lrstParameters!FwdForcePt1Ideal
gudtTest(CHAN0).fwdForcePt(1).high = lrstParameters!FwdForcePt1High
gudtTest(CHAN0).fwdForcePt(1).low = lrstParameters!FwdForcePt1Low
gudtTest(CHAN0).fwdForcePt(1).location = lrstParameters!FwdForcePt1Location
gudtTest(CHAN0).fwdForcePt(2).ideal = lrstParameters!FwdForcePt2Ideal
gudtTest(CHAN0).fwdForcePt(2).high = lrstParameters!FwdForcePt2High
gudtTest(CHAN0).fwdForcePt(2).low = lrstParameters!FwdForcePt2Low
gudtTest(CHAN0).fwdForcePt(2).location = lrstParameters!FwdForcePt2Location
gudtTest(CHAN0).fwdForcePt(3).ideal = lrstParameters!FwdForcePt3Ideal
gudtTest(CHAN0).fwdForcePt(3).high = lrstParameters!FwdForcePt3High
gudtTest(CHAN0).fwdForcePt(3).low = lrstParameters!FwdForcePt3Low
gudtTest(CHAN0).fwdForcePt(3).location = lrstParameters!FwdForcePt3Location
'Reverse Force Points
gudtTest(CHAN0).revForcePt(1).ideal = lrstParameters!RevForcePt1Ideal
gudtTest(CHAN0).revForcePt(1).high = lrstParameters!RevForcePt1High
gudtTest(CHAN0).revForcePt(1).low = lrstParameters!RevForcePt1Low
gudtTest(CHAN0).revForcePt(1).location = lrstParameters!RevForcePt1Location
gudtTest(CHAN0).revForcePt(2).ideal = lrstParameters!RevForcePt2Ideal
gudtTest(CHAN0).revForcePt(2).high = lrstParameters!RevForcePt2High
gudtTest(CHAN0).revForcePt(2).low = lrstParameters!RevForcePt2Low
gudtTest(CHAN0).revForcePt(2).location = lrstParameters!RevForcePt2Location
gudtTest(CHAN0).revForcePt(3).ideal = lrstParameters!RevForcePt3Ideal
gudtTest(CHAN0).revForcePt(3).high = lrstParameters!RevForcePt3High
gudtTest(CHAN0).revForcePt(3).low = lrstParameters!RevForcePt3Low
gudtTest(CHAN0).revForcePt(3).location = lrstParameters!RevForcePt3Location
gudtTest(CHAN0).peakForce.high = lrstParameters!PeakForceHigh
gudtTest(CHAN0).peakForce.low = lrstParameters!PeakForceLow
'Mechanical Hysteresis Points
gudtTest(CHAN0).mechHystPt(1).ideal = lrstParameters!MechHystPt1Ideal
gudtTest(CHAN0).mechHystPt(1).high = lrstParameters!MechHystPt1High
gudtTest(CHAN0).mechHystPt(1).low = lrstParameters!MechHystPt1Low
gudtTest(CHAN0).mechHystPt(1).location = lrstParameters!MechHystPt1Location
gudtTest(CHAN0).mechHystPt(2).ideal = lrstParameters!MechHystPt2Ideal
gudtTest(CHAN0).mechHystPt(2).high = lrstParameters!MechHystPt2High
gudtTest(CHAN0).mechHystPt(2).low = lrstParameters!MechHystPt2Low
gudtTest(CHAN0).mechHystPt(2).location = lrstParameters!MechHystPt2Location
gudtTest(CHAN0).mechHystPt(3).ideal = lrstParameters!MechHystPt3Ideal
gudtTest(CHAN0).mechHystPt(3).high = lrstParameters!MechHystPt3High
gudtTest(CHAN0).mechHystPt(3).low = lrstParameters!MechHystPt3Low
gudtTest(CHAN0).mechHystPt(3).location = lrstParameters!MechHystPt3Location

'Force Knee Location Customer Specifications
gudtCustomerSpec(CHAN0).forceKneeLoc.high = lrstParameters!ForceKneeLocationSpecHigh
gudtCustomerSpec(CHAN0).forceKneeLoc.low = lrstParameters!ForceKneeLocationSpecLow
'Index Point Customer Specifications
gudtCustomerSpec(CHAN0).Index(1).high = lrstParameters!IdleOutput1SpecHigh
gudtCustomerSpec(CHAN0).Index(1).low = lrstParameters!IdleOutput1SpecLow
gudtCustomerSpec(CHAN0).Index(2).high = lrstParameters!MidpointOutput1SpecHigh
gudtCustomerSpec(CHAN0).Index(2).low = lrstParameters!MidpointOutput1SpecLow
gudtCustomerSpec(CHAN0).Index(3).high = lrstParameters!WOTLocation1SpecHigh
gudtCustomerSpec(CHAN0).Index(3).low = lrstParameters!WOTLocation1SpecLow
gudtCustomerSpec(CHAN1).Index(1).high = lrstParameters!IdleOutput2SpecHigh
gudtCustomerSpec(CHAN1).Index(1).low = lrstParameters!IdleOutput2SpecLow
gudtCustomerSpec(CHAN1).Index(2).high = lrstParameters!MidpointOutput2SpecHigh
gudtCustomerSpec(CHAN1).Index(2).low = lrstParameters!MidpointOutput2SpecLow
gudtCustomerSpec(CHAN1).Index(3).high = lrstParameters!WOTLocation2SpecHigh
gudtCustomerSpec(CHAN1).Index(3).low = lrstParameters!WOTLocation2SpecLow
'Forward Force Point Customer Specifications
gudtCustomerSpec(CHAN0).fwdForcePt(1).high = lrstParameters!FwdForcePt1SpecHigh
gudtCustomerSpec(CHAN0).fwdForcePt(1).low = lrstParameters!FwdForcePt1SpecLow
gudtCustomerSpec(CHAN0).fwdForcePt(2).high = lrstParameters!FwdForcePt2SpecHigh
gudtCustomerSpec(CHAN0).fwdForcePt(2).low = lrstParameters!FwdForcePt2SpecLow
gudtCustomerSpec(CHAN0).fwdForcePt(3).high = lrstParameters!FwdForcePt3SpecHigh
gudtCustomerSpec(CHAN0).fwdForcePt(3).low = lrstParameters!FwdForcePt3SpecLow
'Reverse Force Point Customer Specifications
gudtCustomerSpec(CHAN0).revForcePt(1).high = lrstParameters!RevForcePt1SpecHigh
gudtCustomerSpec(CHAN0).revForcePt(1).low = lrstParameters!RevForcePt1SpecLow
gudtCustomerSpec(CHAN0).revForcePt(2).high = lrstParameters!RevForcePt2SpecHigh
gudtCustomerSpec(CHAN0).revForcePt(2).low = lrstParameters!RevForcePt2SpecLow
gudtCustomerSpec(CHAN0).revForcePt(3).high = lrstParameters!RevForcePt3SpecHigh
gudtCustomerSpec(CHAN0).revForcePt(3).low = lrstParameters!RevForcePt3SpecLow
'Mechanical Hysteresis Point Customer Specifications
gudtCustomerSpec(CHAN0).mechHystPt(1).high = lrstParameters!MechHystPt1SpecHigh
gudtCustomerSpec(CHAN0).mechHystPt(1).low = lrstParameters!MechHystPt1SpecLow
gudtCustomerSpec(CHAN0).mechHystPt(2).high = lrstParameters!MechHystPt2SpecHigh
gudtCustomerSpec(CHAN0).mechHystPt(2).low = lrstParameters!MechHystPt2SpecLow
gudtCustomerSpec(CHAN0).mechHystPt(3).high = lrstParameters!MechHystPt3SpecHigh
gudtCustomerSpec(CHAN0).mechHystPt(3).low = lrstParameters!MechHystPt3SpecLow

'Close the Recordset
lrstParameters.Close

Exit Sub
ErrorReadingScanParameters:

    gintAnomaly = 94
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Reading Scan Parameters.", True, True)
End Sub

Public Function ReadSerialNumberRecord() As Boolean
'
'   PURPOSE: To identify if the current Serial Number ID exists in the
'            database, and if so, get the Serial Number for that entry.
'
'  INPUT(S): none
' OUTPUT(S): Function returns whether or not the Serial Number exists

On Error GoTo ErrorReadingSerialNumberID

Dim lrstSerialNumber As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current DUT information
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblSerialNumber.*"
lstrSQLString = lstrSQLString & " From tblSerialNumber"
lstrSQLString = lstrSQLString & " WHERE (((tblSerialNumber.SerialNumberID)=" & gatDataBaseKey.SerialNumberID & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstSerialNumber.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstSerialNumber.BOF And lrstSerialNumber.EOF Then
    gatDataBaseKey.SerialNumberID = 0
Else
gudtMLX90277(1).Read.Lot = lrstSerialNumber!MLX_Lot
gudtMLX90277(1).Read.Wafer = lrstSerialNumber!MLX_Wafer
gudtMLX90277(1).Read.X = lrstSerialNumber!MLX_X
gudtMLX90277(1).Read.Y = lrstSerialNumber!MLX_Y
gstrSerialNumber = MLX90277.EncodePartID
End If

'Close the Recordset
lrstSerialNumber.Close

Exit Function
ErrorReadingSerialNumberID:

    gintAnomaly = 95
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Reading Serial Number.", True, True)
End Function

Public Sub RetrieveRawData()
'
'   PURPOSE: To Retrieve the raw data file for the current DUT
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintFileNum As Integer
Dim lstrPathName As String
Dim lstrFileName As String
Dim lintIndex As Integer

'Get the next available file number
lintFileNum = FreeFile
'Build the file path and name
lstrPathName = RAWDATAPATH & CStr(gatDataBaseKey.DB_ID) & "\"
lstrFileName = CStr(gatDataBaseKey.ScanResultsID) & RAWDATAEXTENSION

Open (lstrPathName & lstrFileName) For Binary As #lintFileNum

For lintIndex = 0 To gintMaxData - 1
    'Read the forward scan raw data from the file
    Get #lintFileNum, , gintForward(CHAN0, lintIndex)   'Vout#1
    Get #lintFileNum, , gintForward(CHAN1, lintIndex)   'Vout#2
    Get #lintFileNum, , gintForward(CHAN2, lintIndex)   'Force
    Get #lintFileNum, , gintForSupply(lintIndex)        'VRef
    'Read the reverse scan raw data from the file
    Get #lintFileNum, , gintReverse(CHAN0, lintIndex)   'Vout#1
    Get #lintFileNum, , gintReverse(CHAN1, lintIndex)   'Vout#2
    Get #lintFileNum, , gintReverse(CHAN2, lintIndex)   'Force
    Get #lintFileNum, , gintRevSupply(lintIndex)        'VRef
Next

'Close the file
Close lintFileNum

Exit Sub
ErrorRetrievingRawData:

    gintAnomaly = 73
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Retrieving Raw Data." & vbCrLf & _
                           "Verify folder permissions are correct.", True, True)
End Sub

Public Sub SaveRawData()
'
'   PURPOSE: To save the raw data file for the current DUT
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSavingRawData

Dim lintFileNum As Integer
Dim lstrPathName As String
Dim lstrFileName As String
Dim lstrDirPath As String
Dim lintIndex As Integer

'Get the next available file number
lintFileNum = FreeFile
'Build the file path and name
lstrDirPath = RAWDATAPATH & CStr(gatDataBaseKey.DB_ID)
lstrPathName = lstrDirPath & "\"

If Not gfsoFileSystemObject.FolderExists(lstrDirPath) Then
    gfsoFileSystemObject.CreateFolder (lstrDirPath)
End If

lstrFileName = CStr(gatDataBaseKey.ScanResultsID) & RAWDATAEXTENSION

'Open the file
Open (lstrPathName & lstrFileName) For Binary As #lintFileNum

For lintIndex = 0 To gintMaxData - 1
    'Write to the forward scan raw data to the file
    Put #lintFileNum, , gintForward(CHAN0, lintIndex)   'Vout#1
    Put #lintFileNum, , gintForward(CHAN1, lintIndex)   'Vout#2
    Put #lintFileNum, , gintForward(CHAN2, lintIndex)   'Force
    Put #lintFileNum, , gintForSupply(lintIndex)        'VRef
    'Write to the reverse scan raw data to the file
    Put #lintFileNum, , gintReverse(CHAN0, lintIndex)   'Vout#1
    Put #lintFileNum, , gintReverse(CHAN1, lintIndex)   'Vout#2
    Put #lintFileNum, , gintReverse(CHAN2, lintIndex)   'Force
    Put #lintFileNum, , gintRevSupply(lintIndex)        'VRef
Next

'Close the file
Close lintFileNum

Exit Sub
ErrorSavingRawData:

    gintAnomaly = 72
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Raw Data." & vbCrLf & _
                           "Verify folder permissions are correct.", True, True)
End Sub

Public Sub SetForceCal()
'
'   PURPOSE: To identify if the current Force Calibration exists in the
'            database, and if so, record the ID Number for that entry.  If not,
'            call the routine that saves the Force Calibration data.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSettingForceCalData

Dim lrstForceCal As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current Force Calibration
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblForceCalibration.*"
lstrSQLString = lstrSQLString & " From tblForceCalibration"
lstrSQLString = lstrSQLString & " WHERE (((tblForceCalibration.ForceCalOffset)=" & gsngForceAmplifierOffset & ") AND ((tblForceCalibration.ForceCalSensitivity)=" & gsngNewtonsPerVolt & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstForceCal.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstForceCal.BOF And lrstForceCal.EOF Then
    'If no match for the Force Calibration was found, add an entry to the table
    gatDataBaseKey.ForceCalID = AddForceCalRecord(lrstForceCal)
Else
    'Get the ForceCalibrationID number of the matching recordset
    gatDataBaseKey.ForceCalID = lrstForceCal!ForceCalID
End If

'Close the Recordset
lrstForceCal.Close

Exit Sub
ErrorSettingForceCalData:
    gintAnomaly = 80
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Force Calibration Data.", True, True)

End Sub

Public Sub SetSN()
'
'   PURPOSE: To send current SN to database on laser
'
'  INPUT(S): none
' OUTPUT(S): none
'1.3ANM new sub

On Error GoTo ErrorSettingSN

Dim lrstSN As New Recordset
Dim lstrSQLString As String
Dim llngKey As Long

'Build a query to look for the current Force Calibration
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblSN.*"
lstrSQLString = lstrSQLString & " From tblSN"
lstrSQLString = lstrSQLString & " WHERE ((tblSN.SerialNumber)=" & """" & gstrSerialNumber & """" & ");"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstSN.Open lstrSQLString, mcnLocalDatabase2, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstSN.BOF And lrstSN.EOF Then
    'If no match for the SN was found, add an entry to the table
    lrstSN.AddNew

    'Add the new Fields to the Record in tblSN
    lrstSN!SerialNumber = gstrSerialNumber
    
    'Save the Record
    lrstSN.Update
End If

'Close the Recordset
lrstSN.Close

Exit Sub
ErrorSettingSN:
    gintAnomaly = 80
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving SN.", True, True)

End Sub

Public Sub SetLot()
'
'   PURPOSE: To identify if the current lot exists in the database,
'            and if so, record the ID Number for that entry.  If not, call the
'            routine that saves the lot name.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSettingLot:

Dim lrstLot As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current lot name
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblLot.*"
lstrSQLString = lstrSQLString & " From tblLot"
lstrSQLString = lstrSQLString & " WHERE (((tblLot.LotName)=" & """" & gstrLotName & """" & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstLot.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstLot.BOF And lrstLot.EOF Then
    'If no match for the Lot was found, add an entry to the table
    gatDataBaseKey.LotID = AddLotRecord(lrstLot)
Else
    'Get the Lot ID number of the matching recordset
    gatDataBaseKey.LotID = lrstLot!LotID
End If

'Close the Recordset
lrstLot.Close

Exit Sub
ErrorSettingLot:
    
    gintAnomaly = 81
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Lot Name.", True, True)
End Sub

Public Sub SetMachineParameters()
'
'   PURPOSE: To identify if the current parameter file exists in the database,
'            and if so, record the ID Number for that entry.  If not, call the
'            routines that save the machine parameters.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSettingMachineParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblMachineParameters.*"
lstrSQLString = lstrSQLString & " From tblMachineParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblMachineParameters.ParameterFileName)=" & """" & gudtMachine.parameterName & """" & ") AND ((tblMachineParameters.ParameterFileRevision)=" & """" & gudtMachine.parameterRev & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstParameters.BOF And lrstParameters.EOF Then
    'If no match for the Parameters was found, add an entry to the table
    gatDataBaseKey.MachineParametersID = AddMachineParametersRecord(lrstParameters)
Else
    'Get the ParametersID number of the matching recordset
    gatDataBaseKey.MachineParametersID = lrstParameters!MachineParametersID
End If

'Close the Recordset
lrstParameters.Close

Exit Sub
ErrorSettingMachineParameters:

    gintAnomaly = 82
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Machine Parameters.", True, True)
End Sub

Public Sub SetProgrammingParameters()
'
'   PURPOSE: To identify if the current parameter file exists in the database,
'            and if so, record the ID Number for that entry.  If not, call the
'            routines that save the programming parameters.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSettingProgrammingParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblProgrammingParameters.*"
lstrSQLString = lstrSQLString & " From tblProgrammingParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblProgrammingParameters.ParameterFileName)=" & """" & gudtMachine.parameterName & """" & ") AND ((tblProgrammingParameters.ParameterFileRevision)=" & """" & gudtMachine.parameterRev & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstParameters.BOF And lrstParameters.EOF Then
    'If no match for the Parameters was found, add an entry to the table
    gatDataBaseKey.ProgParametersID = AddProgrammingParametersRecord(lrstParameters)
Else
    'Get the ParametersID number of the matching recordset
    gatDataBaseKey.ProgParametersID = lrstParameters!ProgrammingParametersID
End If

'Close the Recordset
lrstParameters.Close

Exit Sub
ErrorSettingProgrammingParameters:

    gintAnomaly = 83
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Programming Parameters.", True, True)
End Sub

Public Sub SetScanParameters()
'
'   PURPOSE: To identify if the current parameter file exists in the database,
'            and if so, record the ID Number for that entry.  If not, call the
'            routines that save the programming parameters.
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo ErrorSettingScanParameters

Dim lrstParameters As New Recordset
Dim lstrSQLString As String

'Build a query to look for the current parameter file name and revision
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblScanParameters.*"
lstrSQLString = lstrSQLString & " From tblScanParameters"
lstrSQLString = lstrSQLString & " WHERE (((tblScanParameters.ParameterFileName)=" & """" & gudtMachine.parameterName & """" & ") AND ((tblScanParameters.ParameterFileRevision)=" & """" & gudtMachine.parameterRev & """" & "));"

'Open the RecordSet (as a Dynamic Recordset, with Optimistic Locking)
lrstParameters.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstParameters.BOF And lrstParameters.EOF Then
    'If no match for the Parameters was found, add an entry to the table
    gatDataBaseKey.ScanParametersID = AddScanParametersRecord(lrstParameters)
Else
    'Get the ParametersID number of the matching recordset
    gatDataBaseKey.ScanParametersID = lrstParameters!ScanParametersID
End If

'Close the Recordset
lrstParameters.Close

Exit Sub
ErrorSettingScanParameters:

    gintAnomaly = 84
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Scanning Parameters.", True, True)
End Sub

Public Sub SetSerialNumber()
'
'   PURPOSE: To identify if the current Serial Number exists in the
'            database, and if so, record the ID Number for that entry.  If not,
'            call the routine that saves the Serial Number data.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lrstSerialNumber As New Recordset
Dim lstrSQLString As String

On Error GoTo ErrorSettingSerialNumber

'Build a query to look for the current DUT information
lstrSQLString = ""
lstrSQLString = lstrSQLString & "SELECT tblSerialNumber.*"
lstrSQLString = lstrSQLString & " From tblSerialNumber"
lstrSQLString = lstrSQLString & " WHERE (((tblSerialNumber.MLX_Lot)=" & gudtMLX90277(1).Read.Lot & ") AND ((tblSerialNumber.MLX_Wafer)=" & gudtMLX90277(1).Read.Wafer & ")AND ((tblSerialNumber.MLX_X)=" & gudtMLX90277(1).Read.X & ")AND ((tblSerialNumber.MLX_Y)=" & gudtMLX90277(1).Read.Y & "));"

'Open the Recordset (as a Dynamic Recordset, with Optimistic Locking)
lrstSerialNumber.Open lstrSQLString, mcnLocalDatabase, adOpenDynamic, adLockOptimistic

'If both Beginning Of File and End Of File then no records match the query
If lrstSerialNumber.BOF And lrstSerialNumber.EOF Then
    'If no match for the DUT was found, add an entry to the table
    gatDataBaseKey.SerialNumberID = AddSerialNumberRecord(lrstSerialNumber)
Else
    'Get the DUT ID number of the matching recordset
    gatDataBaseKey.SerialNumberID = lrstSerialNumber!SerialNumberID
End If

'Close the Recordset
lrstSerialNumber.Close

Exit Sub
ErrorSettingSerialNumber:

    gintAnomaly = 85
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Serial Number.", True, True)
End Sub

Public Function SwitchDatabase() As Boolean
'
'   PURPOSE: To switch the database connection to the alternate database
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo DBSwitchError

Dim lstrDatabaseName As String
Dim lintDatabaseNum As Integer
Dim lblnConnectionOK As Boolean

'First, check if it is ok to switch to the alternate database
If IsOkToSwitch Then

    If mblnConnectionActive = True Then
        'Make the current database inactive
        Call DeactivateDatabase
        'Close the current database connection
        Call CloseDatabaseConnection
    End If

    'Display the current status
    frmMain.staMessage.Panels(1).Text = "Testing Connection to alternate database"

    'Determine which database we are connected to, then switch to the other one
    For lintDatabaseNum = 1 To 2
        If mstrDatabaseName = DATABASENAME & CStr(lintDatabaseNum) & DATABASEEXTENSION Then
            'Switch based on which database we are already on
            If lintDatabaseNum = 1 Then
                lstrDatabaseName = DATABASENAME & "2" & DATABASEEXTENSION
            Else
                lstrDatabaseName = DATABASENAME & "1" & DATABASEEXTENSION
            End If
            If OpenDatabaseConnection(lstrDatabaseName) Then
                lblnConnectionOK = True
            Else
                'Display a message to inform the user that the database is beyond its size allowance
                MsgBox "Error opening alternate database!  Please contact Electronics!", vbOKOnly, "Database Problem"
            End If
            'Exit if we tried to connect
            Exit For
        End If
    Next lintDatabaseNum

    If Not lblnConnectionOK Then
        'Switch back to the first database
        If Not OpenDatabaseConnection(mstrDatabaseName) Then
            MsgBox "Error connecting to either database! Scanner cannot run until connected to database!", vbOKOnly, "Database Error!"
        Else
            MsgBox "Connected to original database.  Notify Electronics that it is over it's size limit!", vbOKOnly, "Database Error!"
            'Make the database active
            Call ActivateDatabase
        End If
    Else
        'Make the database active
        Call ActivateDatabase
        mstrDatabaseName = lstrDatabaseName
        frmMain.staMessage.Panels(1).Text = "Connection to alternate database made."
    End If
End If

Exit Function
DBSwitchError:

    gintAnomaly = 71
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Switching to New Database." & vbCrLf & _
                           "Verify Database is not being used by another program.", True, True)
End Function
