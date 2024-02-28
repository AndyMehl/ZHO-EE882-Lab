Attribute VB_Name = "AMAD705_2"
'*************AMAD705_2.BAS - Analysis and Management of Acquired Data*************
'
'   705 Series Specific AMAD, supplemental to Pedal.Bas and Series705.BAS.
'   This module should handle all 705 series production programmer/scanners,
'   test lab programmer/scanners, and database recall software.
'   The software is to be kept in the pedal software library, EE947.
'
'VER    DATE      BY   PURPOSE OF MODIFICATION                          TAG
'1.0  08/07/2007  ANM  First release.                                   1.0ANM
'1.1  08/31/2007  ANM  Removed force knee per SCN# 705F-007 (3973).     1.1ANM
'                      Removed original AMAD items in other module.
'1.2  09/24/2007  ANM  Update to fix cycle time issue with UDB.         1.2ANM
'1.3  11/02/2007  ANM  Update to remove/rename solver items per SCN#    1.3ANM
'                      705F-011 (4018).
'1.4  05/19/2008  ANM  Updates per SCN# 4124.                           1.4ANM
'

Option Explicit

'Public Const StationID = "A675E7C7-CFDE-47F9-9F3F-CB1E817D2A30"
Public gstrProcParamIDs() As String
Public gconnAmad As New ADODB.Connection
Public gstrStationName As String
Public gstrDatabaseName As String
Public gstrTsopName As String
Public gstrLotType As String
Public grsMpcPrecedence As New ADODB.Recordset
Public gstrSubProcess As String
Public grsTsopAnomaly As New ADODB.Recordset

Public Const UDBRAWDATAPATH = "D:\Data\705\Database\Rawdata\UDB" 'Path to raw data files

Public Const MPCTYPE_IDEAL = "Ideal"
Public Const MPCTYPE_HIGHLIMIT = "High Limit"
Public Const MPCTYPE_LOWLIMIT = "Low Limit"
Public Const MPCTYPE_HIGHTOLERANCE = "High Tolerance"
Public Const MPCTYPE_LOWTOLERANCE = "Low Tolerance"

Public Const MPCVALTYPE_ABSCISSA1 = "Abscissa1"
Public Const MPCVALTYPE_ORDINATE1 = "Ordinate1"
Public Const MPCVALTYPE_ABSCISSA2 = "Abscissa2"
Public Const MPCVALTYPE_ORDINATE2 = "Ordinate2"

Public Const INIT = "Initialization"
Public Const CONF = "Configuration"
Public Const PROG = "Programming"
Public Const SCAN = "FunctionalTest"

Public tmpFwdOutputCorrelation As Single
Public tmpRevOutputCorrelation As Single
Public tmpSingLinDevValOut1 As Single
Public tmpSingLinDevValOut2 As Single
Public tmpSlopeDevValOut1 As Single
Public tmpSlopeDevValOut2 As Single
Public tmpWOTValueOut1 As Single
Public tmpWOTValueOut2 As Single

'*** Type Definitions ***
Type DBKeys
    StationID               As String   'Primary Key for tblStation
    TSOP_ID                 As String   'Primary Key for tblTSOP
    ProcessParameterID      As String   'Primary Key for tblProcessParameters
    ProductID               As String   'Primary Key for tblProducts
    TsopStartupID           As String   'Primary Key for tblTSOP_Startup
    LotID                   As String   'Primary Key for tblLot
    DeviceInProcessID       As String   'Primary Key for tblDeviceInProcess
    TSOP_ModeID             As String   'Primary Key for tblTSOP_Modes
    ProgrammingID           As String   'Primary Key for tblProgramming
    TestID                  As String   'Primary Key for tblTests
    AnomalyID               As String   'Primary Key for tblAnomalies
End Type
Public gdbkDbKeys As DBKeys

'Old Items \/\/\/
'Public Const DATABASEPATH = "D:\Data\705\Database\"        'Path to database files
'Public Const DATABASENAME = "705_Database"                 'Base name of the database
'Public Const DATABASEEXTENSION = ".MDB"                    'Extension for database files
'Public Const RAWDATAPATH = "D:\Data\705\Database\Rawdata\" 'Path to raw data files
'Public Const RAWDATAEXTENSION = ".SDR"                     'Extension for raw data files (Stored Data Results)
'Public Const MAXDATABASESIZE = 700000000                   'Maximum size of database file in bytes
'
''Path to SN Database (Production use only)
'Public Const SNDATAPATH = "\\cn1806055\database$\705_Laser_Database.mdb"
'
''*** Type Definitions ***
'Type AMADType
'    DB_ID                   As Long     'Database ID
'    ForceCalID              As Long     'Force Calibration ID
'    LotID                   As Long     'Lot ID
'    MachineParametersID     As Long     'Machine Parameters ID
'    Output1Parameters       As Long     'Output #1 Scan Parameters
'    Output2Parameters       As Long     'Output #2 Scan Parameters
'    ProgParametersID        As Long     'Programming Parameters ID
'    ScanParametersID        As Long     'Test Specifications ID
'    SerialNumberID          As Long     'Serial Number ID
'    ScanResultsID           As Long     'Scan Results ID
'End Type
'
'Public gatDataBaseKey As AMADType
'
Private mcnLocalDatabase As New Connection
Private mgstrDatabaseName As String
Private mblnConnectionActive As Boolean
Private mcnLocalDatabase2 As New Connection
Private mgstrDatabaseName2 As String
Private mblnConnectionActive2 As Boolean
'Old Items /\/\/\

Public Sub StoreForceCalConstants()
'
'   PURPOSE: To write Force Calibration Constants to the database
'
On Error GoTo ErrorSettingForceCalData

    Call InsertDynamicStarutpParamValue("ForceCalOffset", gsngForceAmplifierOffset)
    Call InsertDynamicStarutpParamValue("ForceCalSensitivity", gsngNewtonsPerVolt)

Exit Sub
ErrorSettingForceCalData:
    gintAnomaly = 80
    'Log the error to the error log and display the error message
    Call Pedal.ErrorLogFile("Database Error: Error Saving Force Calibration Data.", True, True)

End Sub

Public Sub InsertDynamicStarutpParamValue(ParmName As String, ParmValue As Single)
On Error GoTo ERROR_InsertDynamicStarutpParamValue

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsertDynamicStartupParamValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = ParmName
    cmd.Parameters(2) = gdbkDbKeys.ProductID
    cmd.Parameters(3) = gdbkDbKeys.StationID
    cmd.Parameters(4) = gdbkDbKeys.TsopStartupID
    cmd.Parameters(5) = ParmValue
        
    cmd.Execute
    
    Set cmd = Nothing

    gconnAmad.Close
    
EXIT_InsertDynamicStarutpParamValue:
    Exit Sub
ERROR_InsertDynamicStarutpParamValue:
    MsgBox "Error in InsertDynamicStarutpParamValue:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDynamicStarutpParamValue
End Sub

Public Function GetLotIdWithInsert(LotName As String) As String
On Error GoTo ERROR_GetLotId

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetLotIDWithInsert"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = LotName
    cmd.Parameters(2) = gstrLotType
    
    cmd.Execute
    
    GetLotIdWithInsert = cmd.Parameters(3)
    
    Set cmd = Nothing

    gconnAmad.Close
    
EXIT_GetLotId:
    Exit Function
ERROR_GetLotId:
    MsgBox "Error in GetLotId:" & Err.number & "- " & Err.Description
    Resume EXIT_GetLotId
End Function

Public Sub DetermineFailureMode()
    
    Dim bFoundFailure As Boolean
    
    bFoundFailure = False
    If gblnSevere Then
        grsMpcPrecedence.MoveFirst
        Do Until grsMpcPrecedence.EOF Or bFoundFailure = True
            If gintSevere(grsMpcPrecedence!ChNum, grsMpcPrecedence!Precedence) = True Then
                bFoundFailure = True
                Call UpdateFailureCount(grsMpcPrecedence!ChNum, grsMpcPrecedence!Precedence)
            End If
            grsMpcPrecedence.MoveNext
        Loop
    ElseIf gblnScanFailure Then
        grsMpcPrecedence.MoveFirst
        Do Until grsMpcPrecedence.EOF Or bFoundFailure = True
            If gintFailure(grsMpcPrecedence!ChNum, grsMpcPrecedence!Precedence) = True Then
                bFoundFailure = True
                Call UpdateFailureCount(grsMpcPrecedence!ChNum, grsMpcPrecedence!Precedence)
            End If
            grsMpcPrecedence.MoveNext
        Loop
    End If

End Sub

Public Sub UpdateFailureCount(intChanNum, intPrecedence)

    Select Case intPrecedence
        Case HIGHINDEXPT1
            gudtScanStats(intChanNum).Index(1).failCount.high = gudtScanStats(intChanNum).Index(1).failCount.high + 1 '': Stop
        Case LOWINDEXPT1
            gudtScanStats(intChanNum).Index(1).failCount.low = gudtScanStats(intChanNum).Index(1).failCount.low + 1 '': Stop
        Case HIGHOUTPUTATFORCEKNEE
            gudtScanStats(intChanNum).outputAtForceKnee.failCount.high = gudtScanStats(intChanNum).outputAtForceKnee.failCount.high + 1 '': Stop
        Case LOWOUTPUTATFORCEKNEE
            gudtScanStats(intChanNum).outputAtForceKnee.failCount.low = gudtScanStats(intChanNum).outputAtForceKnee.failCount.low + 1 '': Stop
        Case HIGHINDEXPT2
            gudtScanStats(intChanNum).Index(2).failCount.high = gudtScanStats(intChanNum).Index(2).failCount.high + 1 '': Stop
        Case LOWINDEXPT2
            gudtScanStats(intChanNum).Index(2).failCount.low = gudtScanStats(intChanNum).Index(2).failCount.low + 1 '': Stop
        Case HIGHINDEXPT3
            gudtScanStats(intChanNum).Index(3).failCount.high = gudtScanStats(intChanNum).Index(3).failCount.high + 1 ': Stop
        Case LOWINDEXPT3
            gudtScanStats(intChanNum).Index(3).failCount.low = gudtScanStats(intChanNum).Index(3).failCount.low + 1 ': Stop
        Case HIGHMAXOUTPUT
            gudtScanStats(intChanNum).maxOutput.failCount.high = gudtScanStats(intChanNum).maxOutput.failCount.high + 1 ': Stop
        Case LOWMAXOUTPUT
            gudtScanStats(intChanNum).maxOutput.failCount.low = gudtScanStats(intChanNum).maxOutput.failCount.low + 1 ': Stop
        Case HIGHSINGLEPOINTLIN
            gudtScanStats(intChanNum).linDevPerTol(1).failCount.high = gudtScanStats(intChanNum).linDevPerTol(1).failCount.high + 1 ': Stop
        Case LOWSINGLEPOINTLIN
            gudtScanStats(intChanNum).linDevPerTol(1).failCount.low = gudtScanStats(intChanNum).linDevPerTol(1).failCount.low + 1 ': Stop
        Case HIGHSLOPE
            gudtScanStats(intChanNum).slopeMax.failCount.high = gudtScanStats(intChanNum).slopeMax.failCount.high + 1 ': Stop
        Case LOWSLOPE
            gudtScanStats(intChanNum).slopeMin.failCount.low = gudtScanStats(intChanNum).slopeMin.failCount.low + 1 ': Stop
        Case HIGHFCHYS
            gudtScanStats(intChanNum).FullCloseHys.failCount.high = gudtScanStats(intChanNum).FullCloseHys.failCount.high + 1 ': Stop
        Case LOWFCHYS
            gudtScanStats(intChanNum).FullCloseHys.failCount.low = gudtScanStats(intChanNum).FullCloseHys.failCount.low + 1 ': Stop
        Case HIGHFWDOUTPUTCOR
            gudtScanStats(intChanNum).outputCorPerTol(1).failCount.high = gudtScanStats(intChanNum).outputCorPerTol(1).failCount.high + 1 ': Stop
        Case LOWFWDOUTPUTCOR
            gudtScanStats(intChanNum).outputCorPerTol(1).failCount.low = gudtScanStats(intChanNum).outputCorPerTol(1).failCount.low + 1 ': Stop
        Case HIGHREVOUTPUTCOR
            gudtScanStats(intChanNum).outputCorPerTol(2).failCount.high = gudtScanStats(intChanNum).outputCorPerTol(2).failCount.high + 1 ': Stop
        Case LOWREVOUTPUTCOR
            gudtScanStats(intChanNum).outputCorPerTol(2).failCount.low = gudtScanStats(intChanNum).outputCorPerTol(2).failCount.low + 1 ': Stop
        Case HIGHFORCEKNEELOC
            gudtScanStats(intChanNum).forceKneeLoc.failCount.high = gudtScanStats(intChanNum).forceKneeLoc.failCount.high + 1 ': Stop
        Case LOWFORCEKNEELOC
            gudtScanStats(intChanNum).forceKneeLoc.failCount.low = gudtScanStats(intChanNum).forceKneeLoc.failCount.low + 1 ': Stop
        Case HIGHFORCEKNEEFWDFORCE
            gudtScanStats(intChanNum).forceKneeForce.failCount.high = gudtScanStats(intChanNum).forceKneeForce.failCount.high + 1 ': Stop
        Case LOWFORCEKNEEFWDFORCE
            gudtScanStats(intChanNum).forceKneeForce.failCount.low = gudtScanStats(intChanNum).forceKneeForce.failCount.low + 1 ': Stop
        Case HIGHFWDFORCEPT1
            gudtScanStats(intChanNum).fwdForcePt(1).failCount.high = gudtScanStats(intChanNum).fwdForcePt(1).failCount.high + 1 ': Stop
        Case LOWFWDFORCEPT1
            gudtScanStats(intChanNum).fwdForcePt(1).failCount.low = gudtScanStats(intChanNum).fwdForcePt(1).failCount.low + 1 ': Stop
        Case HIGHFWDFORCEPT2
            gudtScanStats(intChanNum).fwdForcePt(2).failCount.high = gudtScanStats(intChanNum).fwdForcePt(2).failCount.high + 1 ': Stop
        Case LOWFWDFORCEPT2
            gudtScanStats(intChanNum).fwdForcePt(2).failCount.low = gudtScanStats(intChanNum).fwdForcePt(2).failCount.low + 1 ': Stop
        Case HIGHFWDFORCEPT3
            gudtScanStats(intChanNum).fwdForcePt(3).failCount.high = gudtScanStats(intChanNum).fwdForcePt(3).failCount.high + 1 ': Stop
        Case LOWFWDFORCEPT3
            gudtScanStats(intChanNum).fwdForcePt(3).failCount.low = gudtScanStats(intChanNum).fwdForcePt(3).failCount.low + 1 ': Stop
        Case HIGHREVFORCEPT1
            gudtScanStats(intChanNum).revForcePt(1).failCount.high = gudtScanStats(intChanNum).revForcePt(1).failCount.high + 1 ': Stop
        Case LOWREVFORCEPT1
            gudtScanStats(intChanNum).revForcePt(1).failCount.low = gudtScanStats(intChanNum).revForcePt(1).failCount.low + 1 ': Stop
        Case HIGHREVFORCEPT2
            gudtScanStats(intChanNum).revForcePt(2).failCount.high = gudtScanStats(intChanNum).revForcePt(2).failCount.high + 1 ': Stop
        Case LOWREVFORCEPT2
            gudtScanStats(intChanNum).revForcePt(2).failCount.low = gudtScanStats(intChanNum).revForcePt(2).failCount.low + 1 ': Stop
        Case HIGHREVFORCEPT3
             gudtScanStats(intChanNum).revForcePt(3).failCount.high = gudtScanStats(intChanNum).revForcePt(3).failCount.high + 1 ': Stop
        Case LOWREVFORCEPT3
            gudtScanStats(intChanNum).revForcePt(3).failCount.low = gudtScanStats(intChanNum).revForcePt(3).failCount.low + 1 ': Stop
        Case HIGHPEAKFORCE
            gudtScanStats(intChanNum).peakForce.failCount.high = gudtScanStats(intChanNum).peakForce.failCount.high + 1 ': Stop
        Case LOWPEAKFORCE
            gudtScanStats(intChanNum).peakForce.failCount.low = gudtScanStats(intChanNum).peakForce.failCount.low + 1 ': Stop
        Case HIGHMECHHYSTPT1
            gudtScanStats(intChanNum).mechHystPt(1).failCount.high = gudtScanStats(intChanNum).mechHystPt(1).failCount.high + 1 ': Stop
        Case LOWMECHHYSTPT1
            gudtScanStats(intChanNum).mechHystPt(1).failCount.low = gudtScanStats(intChanNum).mechHystPt(1).failCount.low + 1 ': Stop
        Case HIGHMECHHYSTPT2
            gudtScanStats(intChanNum).mechHystPt(2).failCount.high = gudtScanStats(intChanNum).mechHystPt(2).failCount.high + 1 ': Stop
        Case LOWMECHHYSTPT2
            gudtScanStats(intChanNum).mechHystPt(2).failCount.low = gudtScanStats(intChanNum).mechHystPt(2).failCount.low + 1 ': Stop
        Case HIGHMECHHYSTPT3
            gudtScanStats(intChanNum).mechHystPt(3).failCount.high = gudtScanStats(intChanNum).mechHystPt(3).failCount.high + 1 ': Stop
        Case LOWMECHHYSTPT3
            gudtScanStats(intChanNum).mechHystPt(3).failCount.low = gudtScanStats(intChanNum).mechHystPt(3).failCount.low + 1 ': Stop
    End Select

End Sub

Public Sub TestMpcPrecedence()

Dim lintChanNum As Integer
Dim lintFailureNum As Integer

    'Initialize failure arrays to no failures
    For lintChanNum = CHAN0 To MAXCHANNUM
        For lintFailureNum = 0 To MAXFAULTCNT
            If lintFailureNum = 0 Then
                gintFailure(lintChanNum, lintFailureNum) = 0
                gintSevere(lintChanNum, lintFailureNum) = 0
            Else
                gintFailure(lintChanNum, lintFailureNum) = False
                gintSevere(lintChanNum, lintFailureNum) = False
            End If
        Next lintFailureNum
    Next lintChanNum
    
    Dim i As Integer
    
    For lintChanNum = 0 To MAXCHANNUM
        For i = 1 To 2
            If i = 1 Then
                gblnSevere = True
                gblnScanFailure = False
            Else
                gblnSevere = False
                gblnScanFailure = True
            End If
            For lintFailureNum = 1 To MAXFAULTCNT
                'Set Failure Mode
                If gblnSevere Then
                    gintSevere(lintChanNum, lintFailureNum) = True
                ElseIf gblnScanFailure Then
                    gintFailure(lintChanNum, lintFailureNum) = True
                End If
                'Determine Failure Mode
                Call DetermineFailureMode
                'Clear Failure Mode
                If gblnSevere Then
                    gintSevere(lintChanNum, lintFailureNum) = False
                ElseIf gblnScanFailure Then
                    gintFailure(lintChanNum, lintFailureNum) = False
                End If
            Next lintFailureNum
        Next i
    Next lintChanNum

End Sub

Public Sub GetMpcPrecedence()

    Dim cmd As New ADODB.command
    Dim i As Integer, j As Integer
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopMpcPrecedence"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    
    Set grsMpcPrecedence = cmd.Execute
        
    Set cmd = Nothing
    
    gconnAmad.Close
    
End Sub

Public Function GetProductID() As String
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetProductID"
        .CommandType = adCmdStoredProc
    End With

    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    
    cmd.Execute
        
    GetProductID = cmd.Parameters(2)

    Set cmd = Nothing
    
    gconnAmad.Close
    
End Function

'Public Function GetProgrammingIdWithInsert() As String
'On Error GoTo ERROR_GetProgrammingIdWithInsert
'
'    'Test Variables
'    Dim lstrOperator As String
'    Dim lstrTemperature As String
'    Dim lstrComment As String
'
'    'Test Values
'    lstrOperator = frmMain.ctrSetupInfo1.Operator
'    lstrTemperature = frmMain.ctrSetupInfo1.Temperature
'    lstrComment = frmMain.ctrSetupInfo1.Comment
'
'    Dim cmd As New ADODB.command
'    Dim par As New ADODB.Parameter
'
'    gconnAmad.Open
'
'    With cmd
'        .ActiveConnection = gconnAmad
'        .CommandText = "pspTsopInsProgrammingRecord"
'        .CommandType = adCmdStoredProc
'    End With
'
'    cmd.Parameters(1) = gdbkDbKeys.DeviceInProcessID
'    cmd.Parameters(2) = gdbkDbKeys.ProcessParameterID
'    cmd.Parameters(3) = gdbkDbKeys.LotID
'    cmd.Parameters(4) = gdbkDbKeys.TsopStartupID
'    cmd.Parameters(5) = gdbkDbKeys.TSOP_ModeID
'    cmd.Parameters(6) = DateTime.Now
'    cmd.Parameters(7) = App.Major & "." & App.Minor & "." & App.Revision
'    cmd.Parameters(8) = lstrOperator
'    cmd.Parameters(9) = lstrTemperature
'    cmd.Parameters(10) = lstrComment
'
'    cmd.Execute
'
'    GetProgrammingIdWithInsert = cmd.Parameters(11)
'
'    gconnAmad.Close
'    Set cmd = Nothing
'
'EXIT_GetProgrammingIdWithInsert:
'    Exit Function
'ERROR_GetProgrammingIdWithInsert:
'    MsgBox "Error in GetProgrammingIdWithInsert:" & Err.number & "- " & Err.Description
'    Resume EXIT_GetProgrammingIdWithInsert
'End Function

Public Sub StoreDynamicProgMpcValues()

    Call InsertDynamicProgMpcValue("Index1Val", 1, MPCTYPE_LOWLIMIT, gudtSolver(1).Index(1).low)
    Call InsertDynamicProgMpcValue("Index1Val", 1, MPCTYPE_HIGHLIMIT, gudtSolver(1).Index(1).high)
    Call InsertDynamicProgMpcValue("Index2Val", 1, MPCTYPE_LOWLIMIT, gudtSolver(1).Index(2).low)
    Call InsertDynamicProgMpcValue("Index2Val", 1, MPCTYPE_HIGHLIMIT, gudtSolver(1).Index(2).high)
    Call InsertDynamicProgMpcValue("Index1Val", 2, MPCTYPE_LOWLIMIT, gudtSolver(2).Index(1).low)
    Call InsertDynamicProgMpcValue("Index1Val", 2, MPCTYPE_HIGHLIMIT, gudtSolver(2).Index(1).high)
    Call InsertDynamicProgMpcValue("Index2Val", 2, MPCTYPE_LOWLIMIT, gudtSolver(2).Index(2).low)
    Call InsertDynamicProgMpcValue("Index2Val", 2, MPCTYPE_HIGHLIMIT, gudtSolver(2).Index(2).high)

End Sub

Public Sub InsertDynamicProgMpcValue(ParmName As String, SignalNumber As Integer, MpcType As String _
                                        , MpcValue As Single)
On Error GoTo ERROR_InsertDynamicProgMpcValue

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsDynamicProgMpcValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProgrammingID
    cmd.Parameters(2) = gdbkDbKeys.ProductID
    cmd.Parameters(3) = gdbkDbKeys.StationID
    cmd.Parameters(4) = ParmName
    cmd.Parameters(5) = SignalNumber
    cmd.Parameters(6) = MpcType
    cmd.Parameters(7) = MpcValue
        
    cmd.Execute

EXIT_InsertDynamicProgMpcValue:
    gconnAmad.Close
    Set cmd = Nothing
    Exit Sub
ERROR_InsertDynamicProgMpcValue:
    MsgBox "Error in InsertDynamicProgMpcValue:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDynamicProgMpcValue
End Sub

Public Sub StoreDynamicTestMpcValues()

'Test Only. No Dynamic Test Parameters exist now
'    Call InsertDynamicTestMpcValue("ParamName", 0, "MpcType", 0)

End Sub

Public Sub InsertDynamicTestMpcValue(ParmName As String, SignalNumber As Integer, MpcType As String _
                                        , MpcValue As Single)
On Error GoTo ERROR_InsertDynamicTestMpcValue

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsDynamicTestMpcValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.TestID
    cmd.Parameters(2) = gdbkDbKeys.ProductID
    cmd.Parameters(3) = gdbkDbKeys.StationID
    cmd.Parameters(4) = ParmName
    cmd.Parameters(5) = SignalNumber
    cmd.Parameters(6) = MpcType
    cmd.Parameters(7) = MpcValue
        
    cmd.Execute

EXIT_InsertDynamicTestMpcValue:
    gconnAmad.Close
    Set cmd = Nothing
    Exit Sub
ERROR_InsertDynamicTestMpcValue:
    MsgBox "Error in InsertDynamicTestMpcValue:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDynamicTestMpcValue
End Sub

Public Function InsertProgrammingResults() As Boolean 'TER_06/28/07
On Error GoTo ERROR_InsertProgrammingResults
    
    Dim cmd As New ADODB.command
    Dim sXml As String
       
    Call BuildXmlStringInsertProgResults(sXml)
    
    If gconnAmad.State = adStateClosed Then
        gconnAmad.Open
    End If
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsProgResults"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = gdbkDbKeys.ProgrammingID
    cmd.Parameters(3) = sXml
    cmd.Parameters(4) = Not gblnProgFailure
    cmd.Execute
    
    If cmd.Parameters(0) <> 0 Or IsNull(cmd.Parameters(0)) Or cmd.Parameters(0) = "" Then
        InsertProgrammingResults = False
    Else
        InsertProgrammingResults = True
    End If

EXIT_InsertProgrammingResults:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Function
ERROR_InsertProgrammingResults:
    MsgBox "Error in InsertProgrammingResults:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertProgrammingResults
End Function

Public Sub BuildXmlStringInsertProgResults(sXml As String)

    sXml = "<Root>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampHighCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalClampHighCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampHighCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalClampHighCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampHighValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalClampHighVal & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampHighValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalClampHighVal & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampLowCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalClampLowCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampLowCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalClampLowCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampLowValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalClampLowVal & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "ClampLowValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalClampLowVal & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FineGainCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalFGCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FineGainCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalFGCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index1Loc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalIndexLoc(1) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index1Loc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalIndexLoc(1) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index1Val" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalIndexVal(1) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index1Val" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalIndexVal(1) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index2Loc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalIndexLoc(2) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index2Loc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalIndexLoc(2) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index2Val" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalIndexVal(2) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "Index2Val" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalIndexVal(2) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "OffsetCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalOffsetCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "OffsetCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalOffsetCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RoughGainCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(1).FinalRGCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RoughGainCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtSolver(2).FinalRGCode & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "AGNDCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(1).Read.AGND & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "AGNDCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(2).Read.AGND & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FCKADJCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(1).Read.FCKADJ & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FCKADJCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(2).Read.FCKADJ & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "CKANACHCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(1).Read.CKANACH & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "CKANACHCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(2).Read.CKANACH & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "CKDACCHCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(1).Read.CKDACCH & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "CKDACCHCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(2).Read.CKDACCH & """"
    sXml = sXml & "></Result>"
'    sXml = sXml & "<Result"
'    sXml = sXml & " MetricName=" & """" & "SlowModeCode" & """"
'    sXml = sXml & " SignalNumber=" & """" & "1" & """"
'    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(1).Read.SlowMode & """"
'    sXml = sXml & "></Result>"
'    sXml = sXml & "<Result"
'    sXml = sXml & " MetricName=" & """" & "SlowModeCode" & """"
'    sXml = sXml & " SignalNumber=" & """" & "2" & """"
'    sXml = sXml & " MetricValue=" & """" & gudtMLX90277(2).Read.SlowMode & """"
'    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SlowModeCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & -1 * CInt(gudtMLX90277(1).Read.SlowMode) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SlowModeCode" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & -1 * CInt(gudtMLX90277(2).Read.SlowMode) & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "</Root>"

End Sub

Public Function GetTestIdWithInsert() As String
'On Error GoTo ERROR_GetTestIdWithInsert
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsTestRecord"
        .CommandType = adCmdStoredProc
    End With

    cmd.Parameters(1) = gdbkDbKeys.DeviceInProcessID
    cmd.Parameters(2) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(3) = gdbkDbKeys.LotID
    cmd.Parameters(4) = gdbkDbKeys.TsopStartupID
    cmd.Parameters(5) = gdbkDbKeys.TSOP_ModeID
    cmd.Parameters(6) = DateTime.Now
    cmd.Parameters(7) = App.Major & "." & App.Minor & "." & App.Revision
    cmd.Parameters(8) = frmMain.ctrSetupInfo1.Operator
    cmd.Parameters(9) = frmMain.ctrSetupInfo1.Temperature
    cmd.Parameters(10) = frmMain.ctrSetupInfo1.Comment

    cmd.Execute
        
    GetTestIdWithInsert = cmd.Parameters(11)

    gconnAmad.Close
    Set cmd = Nothing
    
EXIT_GetTestIdWithInsert:
    Exit Function
ERROR_GetTestIdWithInsert:
    MsgBox "Error in GetTestIdWithInsert:" & Err.number & "- " & Err.Description
    Resume EXIT_GetTestIdWithInsert
End Function

Public Function InsertTestResults() As Boolean 'TER_06/28/07
On Error GoTo ERROR_InsertTestResults
    
    Dim cmd As New ADODB.command
    Dim sXml As String
    
    'Test Values
    gblnProgFailure = False
    
    Call BuildXmlStringInsertTestResults(sXml)
    
    If gconnAmad.State = adStateClosed Then
        gconnAmad.Open
    End If
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsTestResults"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = gdbkDbKeys.TestID
    cmd.Parameters(3) = sXml
    cmd.Parameters(4) = Not gblnScanFailure
    
    cmd.Execute
    
    If cmd.Parameters(0) <> 0 Or IsNull(cmd.Parameters(0)) Or cmd.Parameters(0) = "" Then
        InsertTestResults = False
    Else
        InsertTestResults = True
    End If

    Call AMAD705_2.SaveRawData 'ANM
    
EXIT_InsertTestResults:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Function
ERROR_InsertTestResults:
    MsgBox "Error in InsertTestResults:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertTestResults
End Function

Public Sub BuildXmlStringInsertTestResults(sXml As String)

If Abs(gudtExtreme(CHAN0).fwdOutputCor.high.Value) > Abs(gudtExtreme(CHAN0).fwdOutputCor.low.Value) Then
    tmpFwdOutputCorrelation = gudtExtreme(CHAN0).fwdOutputCor.high.Value
Else
    tmpFwdOutputCorrelation = gudtExtreme(CHAN0).fwdOutputCor.low.Value
End If

If Abs(gudtExtreme(CHAN0).revOutputCor.high.Value) > Abs(gudtExtreme(CHAN0).revOutputCor.low.Value) Then
    tmpRevOutputCorrelation = gudtExtreme(CHAN0).revOutputCor.high.Value
Else
    tmpRevOutputCorrelation = gudtExtreme(CHAN0).revOutputCor.low.Value
End If

If Abs(gudtExtreme(CHAN0).SinglePointLin.high.Value) > Abs(gudtExtreme(CHAN0).SinglePointLin.low.Value) Then
    tmpSingLinDevValOut1 = gudtExtreme(CHAN0).SinglePointLin.high.Value
Else
    tmpSingLinDevValOut1 = gudtExtreme(CHAN0).SinglePointLin.low.Value
End If

If Abs(gudtExtreme(CHAN1).SinglePointLin.high.Value) > Abs(gudtExtreme(CHAN1).SinglePointLin.low.Value) Then
    tmpSingLinDevValOut2 = gudtExtreme(CHAN1).SinglePointLin.high.Value
Else
    tmpSingLinDevValOut2 = gudtExtreme(CHAN1).SinglePointLin.low.Value
End If

If Abs(gudtExtreme(CHAN0).slope.high.Value) > Abs(gudtExtreme(CHAN0).slope.low.Value) Then
    tmpSlopeDevValOut1 = gudtExtreme(CHAN0).slope.high.Value
Else
    tmpSlopeDevValOut1 = gudtExtreme(CHAN0).slope.low.Value
End If

If Abs(gudtExtreme(CHAN1).slope.high.Value) > Abs(gudtExtreme(CHAN1).slope.low.Value) Then
    tmpSlopeDevValOut2 = gudtExtreme(CHAN1).slope.high.Value
Else
    tmpSlopeDevValOut2 = gudtExtreme(CHAN1).slope.low.Value
End If

tmpWOTValueOut1 = gudtReading(CHAN0).Index(3).Value
tmpWOTValueOut2 = gudtReading(CHAN1).Index(3).Value

    sXml = "<Root>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "forceKneeLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).forceKnee.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FullCloseHys" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).FullCloseHys.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FullCloseHys" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).FullCloseHys.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdForceAtForceKneeLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).forceKnee.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdForcePt1" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).fwdForcePt(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdForcePt2" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).fwdForcePt(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdForcePt3" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).fwdForcePt(3).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdOutputCorrelation" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & tmpFwdOutputCorrelation & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdOutputCorrPerTolLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).outputCorPerTol(1).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "FwdOutputCorrPerTolVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).outputCorPerTol(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "IdleValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).Index(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "IdleValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).Index(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxFwdOutputCorrLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).fwdOutputCor.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxFwdOutputCorrVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).fwdOutputCor.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxLinDevPerTolLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).linDevPerTol(1).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxLinDevPerTolLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).linDevPerTol(1).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxLinDevPerTolVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).linDevPerTol(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxLinDevPerTolVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).linDevPerTol(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxRevOutputCorrLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).revOutputCor.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxRevOutputCorrVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).revOutputCor.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSingLinDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).SinglePointLin.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSingLinDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).SinglePointLin.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).SinglePointLin.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).SinglePointLin.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSlopeDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).slope.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSlopeDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).slope.high.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).slope.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxSlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).slope.high.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).maxOutput.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).maxOutput.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxValueLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).maxOutput.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MaxValueLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).maxOutput.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MechHystPt1" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).mechHystPt(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MechHystPt2" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).mechHystPt(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MechHystPt3" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).mechHystPt(3).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MidpointValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).Index(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MidpointValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).Index(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinFwdOutputCorrLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).fwdOutputCor.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinFwdOutputCorrVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).fwdOutputCor.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinRevOutputCorrLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).revOutputCor.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinRevOutputCorrVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).revOutputCor.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSingLinDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).SinglePointLin.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSingLinDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).SinglePointLin.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).SinglePointLin.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).SinglePointLin.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSlopeDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).slope.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSlopeDevLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).slope.low.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).slope.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "MinSlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).slope.low.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "OutputAtForceKnee" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).outputAtForceKnee & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "OutputAtForceKnee" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).outputAtForceKnee & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "peakForce" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).peakForce & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "PeakHysteresisLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).hysteresis.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "PeakHysteresisLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).hysteresis.location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "PeakHysteresisVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).hysteresis.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "PeakHysteresisVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN1).hysteresis.Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "pedalAtRestLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).pedalAtRestLoc & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevForcePt1" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).revForcePt(1).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevForcePt2" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).revForcePt(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevForcePt3" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).revForcePt(3).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevOutputCorrelation" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & tmpRevOutputCorrelation & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevOutputCorrPerTolLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).outputCorPerTol(2).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "RevOutputCorrPerTolVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtExtreme(CHAN0).outputCorPerTol(2).Value & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & tmpSingLinDevValOut1 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SingLinDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & tmpSingLinDevValOut2 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & tmpSlopeDevValOut1 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "SlopeDevVal" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & tmpSlopeDevValOut2 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "WOTLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).Index(3).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "WOTLoc" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).Index(3).location & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "WOTValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & tmpWOTValueOut1 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result"
    sXml = sXml & " MetricName=" & """" & "WOTValue" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & tmpWOTValueOut2 & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result" '1.4ANM
    sXml = sXml & " MetricName=" & """" & "MLXIdd" & """"
    sXml = sXml & " SignalNumber=" & """" & "1" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN0).mlxCurrent & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "<Result" '1.4ANM
    sXml = sXml & " MetricName=" & """" & "MLXIdd" & """"
    sXml = sXml & " SignalNumber=" & """" & "2" & """"
    sXml = sXml & " MetricValue=" & """" & gudtReading(CHAN1).mlxCurrent & """"
    sXml = sXml & "></Result>"
    sXml = sXml & "</Root>"

End Sub

Public Sub InitRecordsets()
    
    gconnAmad.Open
    
    With grsMpcPrecedence
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .ActiveConnection = gconnAmad
    End With
    
    With grsTsopAnomaly
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append "AnomalyMessage", adVarChar, 50
        .Fields.Append "AnomalyDateTime", adVarChar, 50
        .Fields.Append "Operator", adVarChar, 10
        .Fields.Append "UndefinedAnomalyNumber", adSmallInt, , adFldIsNullable
    End With
    grsTsopAnomaly.Open
    
    gconnAmad.Close

End Sub

Public Function NewTsopStartupID() As String
On Error GoTo ERROR_NewTsopStartupID 'Need error trap in case error writing to database
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopNewTsopStartupID"
        .CommandType = adCmdStoredProc
    End With

    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupDateTime", adVarChar, adParamInput, 50, CStr(Now()))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    cmd.Execute
        
    NewTsopStartupID = cmd.Parameters(2)

    Set par = Nothing
    Set cmd = Nothing

    gconnAmad.Close
    
EXIT_NewTsopStartupID:
    Exit Function
ERROR_NewTsopStartupID:
    MsgBox "Error in NewTsopStartupID:" & Err.number & "- " & Err.Description
    Resume EXIT_NewTsopStartupID
    
End Function

Public Function GetStationID(StationName As String) As String
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetStationID"
        .CommandType = adCmdStoredProc
    End With

    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationName", adVarChar, adParamInput, 50, StationName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    cmd.Execute
        
    GetStationID = cmd.Parameters(2)

    Set par = Nothing
    Set cmd = Nothing
    
    gconnAmad.Close
    
End Function

Public Function GetTsopID(TsopName As String) As String
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetTsopID"
        .CommandType = adCmdStoredProc
    End With

    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopName", adVarChar, adParamInput, 50, TsopName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    cmd.Execute
        
    GetTsopID = cmd.Parameters(2)

    Set par = Nothing
    Set cmd = Nothing
    
    gconnAmad.Close
    
End Function

Public Function GetTsopModeID() As String

On Error GoTo ERROR_GetTsopModeID
    
Dim cmd As New ADODB.command
Dim intTsopModeValue As Integer

intTsopModeValue = TsopModeValue()
    
gconnAmad.Open

With cmd
    .ActiveConnection = gconnAmad
    .CommandText = "pspTsopGetTsopModeID"
    .CommandType = adCmdStoredProc
End With

cmd.Parameters(1) = gdbkDbKeys.TSOP_ID
cmd.Parameters(2) = intTsopModeValue

cmd.Execute

If IsNull(cmd.Parameters(3)) Then
    GetTsopModeID = ""
Else
    GetTsopModeID = cmd.Parameters(3)
End If

EXIT_GetTsopModeID:
    gconnAmad.Close
    Exit Function
ERROR_GetTsopModeID:
    MsgBox "Error in GetTsopModeID:" & Err.number & "- " & Err.Description
    Resume EXIT_GetTsopModeID

End Function

Public Function GetTsopFunctionPositionValue(strFunctionName As String) As Integer

On Error GoTo ERROR_GetTsopFunctionPositionValue
    
Dim cmd As New ADODB.command

gconnAmad.Open

With cmd
    .ActiveConnection = gconnAmad
    .CommandText = "pspTsopGetTsopFunctionPositionValue"
    .CommandType = adCmdStoredProc
End With

cmd.Parameters(1) = gdbkDbKeys.TSOP_ID
cmd.Parameters(2) = strFunctionName

cmd.Execute

GetTsopFunctionPositionValue = cmd.Parameters(3)

EXIT_GetTsopFunctionPositionValue:
    gconnAmad.Close
    Exit Function
ERROR_GetTsopFunctionPositionValue:
    MsgBox "Error in GetTsopFunctionPositionValue:" & Err.number & "- " & Err.Description
    Resume EXIT_GetTsopFunctionPositionValue

End Function

Public Function TsopModeValue() As Integer
    
    Dim intProgramParameters As Integer
    Dim intLockAll As Integer
    Dim intLockNonRejects As Integer
    Dim intLockNone As Integer
    Dim intTestOutputVsPosition As Integer
    Dim intTestForceVsPosition As Integer
    
''Testing
'    Stop
'
'    gblnProgramStart = 1
'    gblnLockICs = 1
'    gblnLockRejects = 0
'    gblnScanStart = 1
'    gblnForceOnly = 1
        
    intProgramParameters = -1 * CInt(gblnProgramStart) * (GetTsopFunctionPositionValue("Program Parameters"))
    intLockAll = -1 * CInt(gblnProgramStart And gblnLockICs And gblnLockRejects) * (GetTsopFunctionPositionValue("Lock All"))
    intLockNonRejects = -1 * CInt(gblnProgramStart And gblnLockICs And Not gblnLockRejects) * (GetTsopFunctionPositionValue("Lock NonRejects"))
    intLockNone = -1 * CInt(gblnProgramStart And Not gblnLockICs) * (GetTsopFunctionPositionValue("Lock None"))
    intTestOutputVsPosition = -1 * CInt(gblnScanStart And Not gblnForceOnly) * (GetTsopFunctionPositionValue("Test Output Versus Position"))
    intTestForceVsPosition = -1 * CInt(gblnScanStart) * (GetTsopFunctionPositionValue("Test Force Versus Position"))

    TsopModeValue = intProgramParameters + intLockAll + intLockNonRejects + intLockNone _
                    + intTestOutputVsPosition + intTestForceVsPosition


End Function

Public Sub PopulateParameterSetList()
'
'   PURPOSE: To populate the parameter file list
'
'  INPUT(S): none
' OUTPUT(S): none

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.command
Dim par As New ADODB.Parameter
Dim i As Integer

'Remove all list items
If frmMain.cboParameterFileName.ListCount > 0 Then
    For i = frmMain.cboParameterFileName.ListCount - 1 To 0 Step -1
        frmMain.cboParameterFileName.RemoveItem (i)
    Next i
End If
    
gconnAmad.Open

'Setup call to stored procedure
With cmd
    .ActiveConnection = gconnAmad
    .CommandText = "pspTsopGetParameterSetList"
    .CommandType = adCmdStoredProc
End With

'Set command for list from DB
Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50, gdbkDbKeys.StationID)
cmd.Parameters.Append par
Set par = cmd.CreateParameter("TsopID", adVarChar, adParamInput, 50, gdbkDbKeys.TSOP_ID)
cmd.Parameters.Append par
Set par = cmd.CreateParameter("LotType", adVarChar, adParamInput, 50, gstrLotType)
cmd.Parameters.Append par

'Finish record set
With rs
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .ActiveConnection = gconnAmad
End With

'Get the record set
Set rs = cmd.Execute

'Dimension array for list
ReDim gstrProcParamIDs(rs.RecordCount)

'Place list in parameter file drop down
i = 0
Do Until rs.EOF
    'Debug.Print i & ". " & rs!ParameterSet, rs!ProcessParameterID
    frmMain.cboParameterFileName.AddItem rs!ParameterSet
    gstrProcParamIDs(i) = rs!ProcessParameterID
    i = i + 1
    rs.MoveNext
Loop
        
'Close the record set
rs.Close

'Clear variables
Set rs = Nothing
Set par = Nothing
Set cmd = Nothing
    
gconnAmad.Close

End Sub

Public Sub InitStoredProcParams()
    
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
        
    gconnAmad.Open
    
    cmd.ActiveConnection = gconnAmad
    cmd.CommandType = adCmdStoredProc
        
    'Stored Procedure to return ProductID from selected Parameter Set
    cmd.CommandText = "pspTsopGetProductID"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to Return MPC Precedence recordset
    cmd.CommandText = "pspTsopMpcPrecedence"
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to return Metric Performance Criteria (MPC) Values
    cmd.CommandText = "pspTsopGetMpcValues"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MeasurandSigNumber", adInteger, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcType", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("Region", adInteger, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcValType", adInteger, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcValue", adDouble, adParamOutput)
    cmd.Parameters.Append par

    'Stored Procedure to return Non-Metric Parameter Values
    cmd.CommandText = "pspTsopGetParameterValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("Value", adDouble, adParamOutput)
    cmd.Parameters.Append par

    'Stored Procedure to return Non-Metric Parameter Enumerated Values
    cmd.CommandText = "pspTsopGetParameterEnumValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("Value", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par

    'Stored Procedure to return Max Number of Regions for a MPC
    cmd.CommandText = "pspTsopGetMpcRegionCount"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MeasurandSigNumber", adInteger, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("RegionCount", adDouble, adParamOutput)
    cmd.Parameters.Append par
    
    'Stored Procedure to insert Dynamic Startup Parameter Values
    cmd.CommandText = "pspTsopInsertDynamicStartupParamValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParamValue", adDouble, adParamInput)
    cmd.Parameters.Append par
    
    'Stored Procedure to return LotID for a LotName
    cmd.CommandText = "pspTsopGetLotIDWithInsert"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("LotName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("LotType", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("LotID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to search for a DeviceInProcess based on input parameters
    ' and return DeviceInProcessID if it exists
    cmd.CommandText = "pspTsopGetDipID"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("EncodedSerialNumber", adVarChar, adParamInput, 50) '1.2ANM
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DeviceInProcessID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to insert a DeviceInProcess and its associated Attribute Values and return its ID
    cmd.CommandText = "pspTsopInsDeviceInProcess"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("XmlString", adVarChar, adParamInput, 8000)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DeviceInProcessID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to Update DIP Attribute Values for existing DeviceInProcess
    cmd.CommandText = "pspTsopUpdateDipAttributeValues"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DeviceInProcessID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("XmlString", adVarChar, adParamInput, 8000)
    cmd.Parameters.Append par
    
    'Stored Procedure to Determine TSOP Function position value
    cmd.CommandText = "pspTsopGetTsopFunctionPositionValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopFunctionName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopFunctionValue", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    'Stored Procedure to Determine TSOP Mode and return Mode ID
    cmd.CommandText = "pspTsopGetTsopModeID"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopModeValue", adInteger, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopModeID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to Insert Programming Records
    cmd.CommandText = "pspTsopInsProgrammingRecord"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DeviceInProcessID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("LotID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopModeID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgDateTime", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopVersion", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgOperator", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgTemperature", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgComment", adVarChar, adParamInput, 255)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgrammingID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
   
    'Stored Procedure to insert Programming Result records
    cmd.CommandText = "pspTsopInsProgResults"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgrammingID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("XmlString", adVarChar, adParamInput, 8000)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgPassed", adBoolean, adParamInput)
    cmd.Parameters.Append par
  
    'Stored Procedure to Insert Test Records
    cmd.CommandText = "pspTsopInsTestRecord"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DeviceInProcessID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("LotID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopModeID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestDateTime", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopVersion", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestOperator", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestTemperature", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestComment", adVarChar, adParamInput, 255)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
 
    'Stored Procedure to insert Test Result records
    cmd.CommandText = "pspTsopInsTestResults"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProcessParameterID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("XmlString", adVarChar, adParamInput, 8000)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestPassed", adBoolean, adParamInput)
    cmd.Parameters.Append par

    'Stored Procedure to return Anomaly Information from Database
    cmd.CommandText = "pspTsopGetAnomalyInfo"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("AnomalyNumber", adInteger, adParamInput)
    cmd.Parameters.Append par
   
    'Stored Procedure to return AnomalyID for undefined anomalies
    cmd.CommandText = "pspTsopGetUndefinedAnomalyID"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("AnomalyID", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    
    'Stored Procedure to Insert TSOP Anomaly Records
    cmd.CommandText = "pspTsopInsTsopAnomaly"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("AnomalyID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopSubProcessID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TsopStartupID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("AnomalyMessage", adVarChar, adParamInput, 8000)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("AnomalyDateTime", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("Operator", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("UndefinedAnomalyNumber", adSmallInt, adParamInput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgrammingID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
       
    'Stored Procedure to insert Dynamic Programming Parameter Values 'TER_05/04/07
    cmd.CommandText = "pspTsopInsertDynamicProgParamValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgrammingID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParamValue", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
       
    'Stored Procedure to insert Dynamic Test Parameter Values 'TER_05/04/07
    cmd.CommandText = "pspTsopInsertDynamicTestParamValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParamValue", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    
    'Stored Procedure to insert Dynamic Programming MPC Values 'TER_05/04/07
    cmd.CommandText = "pspTsopInsDynamicProgMpcValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProgrammingID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MeasurandSigNumber", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcType", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcValue", adDouble, adParamInput)
    cmd.Parameters.Append par
              
    'Stored Procedure to insert Dynamic Test MPC Values 'TER_05/04/07
    cmd.CommandText = "pspTsopInsDynamicTestMpcValue"
    Set par = cmd.CreateParameter("ReturnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("TestID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ProductID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("StationID", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("ParameterName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MeasurandSigNumber", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcType", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("MpcValue", adDouble, adParamInput)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set cmd = Nothing

    gconnAmad.Close

End Sub

Public Function MpcValue(ParamName As String, MeasurandSigNumber As Integer, MpcType As String, Region As Integer, MpcValType As String) As Double
    
    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetMpcValues"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = ParamName
    cmd.Parameters(3) = MeasurandSigNumber
    cmd.Parameters(4) = MpcType
    cmd.Parameters(5) = Region
    cmd.Parameters(6) = MpcValType
    
    cmd.Execute
    
    MpcValue = cmd.Parameters(7)
    
'    Select Case Coordinate
'        Case "Abscissa1"
'            MpcValue = cmd.Parameters(5)
'        Case "Abscissa2"
'            MpcValue = cmd.Parameters(6)
'        Case "Ordinate1"
'            MpcValue = cmd.Parameters(7)
'        Case "Ordinate2"
'            MpcValue = cmd.Parameters(8)
'    End Select
    
    Set cmd = Nothing

    gconnAmad.Close

End Function

Public Function ParamValue(ParamName As String) As Double
    
    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetParameterValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = ParamName
    
    cmd.Execute
    
    ParamValue = cmd.Parameters(3)
    
    Set cmd = Nothing

    gconnAmad.Close
    
End Function

Public Function ParamEnumValue(ParamName As String) As String
    
    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetParameterEnumValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = ParamName
    
    cmd.Execute
    
    ParamEnumValue = cmd.Parameters(3)
    
    Set cmd = Nothing

    gconnAmad.Close

End Function

Public Function MPCRegionCount(ParamName As String, MeasurandSigNumber As Integer, MpcType As String) As Double
    
    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetMpcRegionCount"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProcessParameterID
    cmd.Parameters(2) = ParamName
    cmd.Parameters(3) = MeasurandSigNumber
    cmd.Parameters(4) = MpcType
    
    cmd.Execute
    
    MPCRegionCount = cmd.Parameters(5)
    
    Set cmd = Nothing

    gconnAmad.Close

End Function

Public Sub LoadParametersFromDb()
'
'   PURPOSE:   To input operational parameters into the program from
'              Database.
'
'  INPUT(S):
' OUTPUT(S):

'Not sure if these variables are needed yet
Dim lintRegionNum As Integer                'Region number
Dim lintRegionCount As Integer              'Count number
Dim lintRow As Integer                      'Row of table
Dim lintColumn As Integer                   'Column of table
Dim lstrParameterName As String             'Parameter or Metric Name
Dim lsngRead As Single                      'Read variable

'*** Test Parameters Output #1 ***
'Index 1 - FullClose By Location
' Metric Performance Criteria (MPC) Values
lstrParameterName = "IdleValue" 'Metric Parameter
gudtTest(CHAN0).Index(1).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(1).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(1).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
' There are two ways to load this variable
' This way loads it from an MPC record.
gudtTest(CHAN0).Index(1).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)
' This way loads it from a Parameter record
'lstrParameterName = "IdleLocOutput1" 'Non-metric Parameter
'gudtTest(CHAN0).Index(1).location = ParamValue(lstrParameterName)

'1.1ANM 'Output at Force Knee Location
'1.1ANM lstrParameterName = "OutputAtForceKnee" 'Metric Parameter
'1.1ANM gudtTest(CHAN0).outputAtForceKnee.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).outputAtForceKnee.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).outputAtForceKnee.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Index 2 - Midpoint By Location
lstrParameterName = "MidpointValue" 'Metric Parameter
gudtTest(CHAN0).Index(2).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(2).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(2).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(2).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Index 3 - FullOpen By Location
lstrParameterName = "WOTValue" 'Metric Parameter
gudtTest(CHAN0).Index(3).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(3).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(3).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).Index(3).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Maximum Output
lstrParameterName = "MaxValue" 'Metric Parameter
gudtTest(CHAN0).maxOutput.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).maxOutput.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).maxOutput.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'SinglePoint Linearity Deviation
lstrParameterName = "SingLinDevVal"
'Ideal
' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
For lintRegionNum = 1 To lintRegionCount
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, lintRegionNum, MPCVALTYPE_ORDINATE1)
Next lintRegionNum
'High Limit
' Start and stop location values (Abscissa1 and Abscissa2) are stored
'   for all MPC Types (Ideal, High Limit, and Low Limit) and could be loaded in those loops as well.
'   High was chosen because it explicitly loads start and stop value pairs
'   Start valuepairs are (Abscissa1, Ordinate1) and Stop value pairs are (Abscissa2 and Ordinate2)
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_HIGHLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA1)
    'Region Stop
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA2)
Next lintRegionNum
'Low Limit
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_LOWLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    'Region Stop
    gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
Next lintRegionNum

'Slope Deviation Start
lstrParameterName = "IdealSlope1Output1" 'Non-metric Parameter
' Ideal is a non-metric parameter, Ideal2 is used as a metric performance criteria
gudtTest(CHAN0).slope.ideal = ParamValue(lstrParameterName)
lstrParameterName = "SlopeDevVal" 'Metric Parameter
gudtTest(CHAN0).slope.ideal2 = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
lsngRead = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).slope.high = lsngRead / gudtTest(CHAN0).slope.ideal2
lsngRead = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).slope.low = lsngRead / gudtTest(CHAN0).slope.ideal2
' Start value could be loaded from Abscissa1 of any MPC Type (Ideal, High Limit, or Low Limit) for this metric
gudtTest(CHAN0).slope.start = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ABSCISSA1)

'Slope Deviation Stop
lstrParameterName = "SlopeDevVal" 'Metric Parameter
gudtTest(CHAN0).slope.stop = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ABSCISSA2)

'Full-Close Hysteresis
lstrParameterName = "FullCloseHys" 'Metric Parameter
gudtTest(CHAN0).FullCloseHys.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).FullCloseHys.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).FullCloseHys.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).FullCloseHys.location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'MLX Current '1.4ANM
lstrParameterName = "MLXIdd" 'Metric Parameter
gudtTest(CHAN0).mlxCurrent.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mlxCurrent.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mlxCurrent.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Evaluation Start
lstrParameterName = "EvaluationStartLocOutput1" 'Non-Metric Parameter
gudtTest(CHAN0).evaluate.start = ParamValue(lstrParameterName)

'Evaluation Stop
lstrParameterName = "EvaluationStopLocOutput1" 'Non-Metric Parameter
gudtTest(CHAN0).evaluate.stop = ParamValue(lstrParameterName)

'*** Test Parameters Output #2 ***

'Index 1 - FullClose By Location
' Metric Performance Criteria (MPC) Values
lstrParameterName = "IdleValue" 'Metric Parameter
gudtTest(CHAN1).Index(1).ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(1).high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(1).low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(1).location = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'1.1ANM 'Output at Force Knee Location
'1.1ANM lstrParameterName = "OutputAtForceKnee" 'Metric Parameter
'1.1ANM gudtTest(CHAN1).outputAtForceKnee.ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN1).outputAtForceKnee.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN1).outputAtForceKnee.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Index 2 - Midpoint By Location
lstrParameterName = "MidpointValue" 'Metric Parameter
gudtTest(CHAN1).Index(2).ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(2).high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(2).low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(2).location = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Index 3 - FullOpen By Location
lstrParameterName = "WOTValue" 'Metric Parameter
gudtTest(CHAN1).Index(3).ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(3).high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(3).low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).Index(3).location = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Maximum Output
lstrParameterName = "MaxValue" 'Metric Parameter
gudtTest(CHAN1).maxOutput.ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).maxOutput.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).maxOutput.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'SinglePoint Linearity Deviation
lstrParameterName = "SingLinDevVal"
'Ideal
' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount(lstrParameterName, 2, MPCTYPE_IDEAL)
For lintRegionNum = 1 To lintRegionCount
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, lintRegionNum, MPCVALTYPE_ORDINATE1)
Next lintRegionNum
'High Limit
' Start and stop location values (Abscissa1 and Abscissa2) are stored
'   for all MPC Types (Ideal, High Limit, and Low Limit) and could be loaded in those loops as well.
'   High was chosen because it explicitly loads start and stop value pairs
'   Start valuepairs are (Abscissa1, Ordinate1) and Stop value pairs are (Abscissa2 and Ordinate2)
lintRegionCount = MPCRegionCount(lstrParameterName, 2, MPCTYPE_HIGHLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.location = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA1)
    'Region Stop
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.location = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA2)
Next lintRegionNum
'Low Limit
lintRegionCount = MPCRegionCount(lstrParameterName, 2, MPCTYPE_LOWLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    'Region Stop
    gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
Next lintRegionNum

'Slope Deviation Start
lstrParameterName = "IdealSlope1Output2" 'Non-metric Parameter
' Ideal is a non-metric parameter, Ideal2 is used as a metric performance criteria
gudtTest(CHAN1).slope.ideal = ParamValue(lstrParameterName)
lstrParameterName = "SlopeDevVal" 'Metric Parameter
gudtTest(CHAN1).slope.ideal2 = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
lsngRead = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).slope.high = lsngRead / gudtTest(CHAN1).slope.ideal2
lsngRead = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).slope.low = lsngRead / gudtTest(CHAN1).slope.ideal2
' Start value could be loaded from Abscissa1 of any MPC Type (Ideal, High Limit, or Low Limit) for this metric
gudtTest(CHAN1).slope.start = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ABSCISSA1)

'Slope Deviation Stop
lstrParameterName = "SlopeDevVal" 'Metric Parameter
gudtTest(CHAN1).slope.stop = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ABSCISSA2)

'Full-Close Hysteresis
lstrParameterName = "FullCloseHys" 'Metric Parameter
gudtTest(CHAN1).FullCloseHys.ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).FullCloseHys.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).FullCloseHys.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).FullCloseHys.location = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'MLX Current '1.4ANM
lstrParameterName = "MLXIdd" 'Metric Parameter
gudtTest(CHAN1).mlxCurrent.ideal = MpcValue(lstrParameterName, 2, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).mlxCurrent.high = MpcValue(lstrParameterName, 2, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN1).mlxCurrent.low = MpcValue(lstrParameterName, 2, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Evaluation Start
lstrParameterName = "EvaluationStartLocOutput2" 'Non-Metric Parameter
gudtTest(CHAN1).evaluate.start = ParamValue(lstrParameterName)

'Evaluation Stop
lstrParameterName = "EvaluationStopLocOutput2" 'Non-Metric Parameter
gudtTest(CHAN1).evaluate.stop = ParamValue(lstrParameterName)

'*** Correlation Parameters ***

'Forward Output Correlation
lstrParameterName = "FwdOutputCorrelation"
'Ideal
' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
For lintRegionNum = 1 To lintRegionCount
    lstrParameterName = "OutputCorrelationRatioIdeal"
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).ideal = ParamValue(lstrParameterName)
Next lintRegionNum
lstrParameterName = "FwdOutputCorrelation"
'High Limit
' Start and stop location values (Abscissa1 and Abscissa2) are stored
'   for all MPC Types (Ideal, High Limit, and Low Limit) and could be loaded in those loops as well.
'   High was chosen because it explicitly loads start and stop value pairs
'   Start valuepairs are (Abscissa1, Ordinate1) and Stop value pairs are (Abscissa2 and Ordinate2)
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_HIGHLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA1)
    'Region Stop
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA2)
Next lintRegionNum
'Low Limit
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_LOWLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    'Region Stop
    gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
Next lintRegionNum

'Reverse Output Correlation
lstrParameterName = "RevOutputCorrelation"
'Ideal
' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_IDEAL)
For lintRegionNum = 1 To lintRegionCount
    lstrParameterName = "OutputCorrelationRatioIdeal"
    gudtTest(CHAN0).revOutputCor(lintRegionNum).ideal = ParamValue(lstrParameterName)
Next lintRegionNum
lstrParameterName = "RevOutputCorrelation"
'High Limit
' Start and stop location values (Abscissa1 and Abscissa2) are stored
'   for all MPC Types (Ideal, High Limit, and Low Limit) and could be loaded in those loops as well.
'   High was chosen because it explicitly loads start and stop value pairs
'   Start valuepairs are (Abscissa1, Ordinate1) and Stop value pairs are (Abscissa2 and Ordinate2)
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_HIGHLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA1)
    'Region Stop
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.location = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, lintRegionNum, MPCVALTYPE_ABSCISSA2)
Next lintRegionNum
'Low Limit
lintRegionCount = MPCRegionCount(lstrParameterName, 1, MPCTYPE_LOWLIMIT)
For lintRegionNum = 1 To lintRegionCount
    'Region Start
    gudtTest(CHAN0).revOutputCor(lintRegionNum).start.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE1)
    'Region Stop
    gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, lintRegionNum, MPCVALTYPE_ORDINATE2)
Next lintRegionNum

'*** Force Parameters ***

'Pedal at Rest Location
lstrParameterName = "pedalAtRestLoc" 'Metric Parameter
gudtTest(CHAN0).pedalAtRestLoc.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)

'1.1ANM 'Force Knee Location
'1.1ANM lstrParameterName = "forceKneeLoc" 'Metric Parameter
'1.1ANM gudtTest(CHAN0).forceKneeLoc.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).forceKneeLoc.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).forceKneeLoc.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'1.1ANM 'Forward Force at Force Knee Location
'1.1ANM lstrParameterName = "FwdForceAtForceKneeLoc" 'Metric Parameter
'1.1ANM gudtTest(CHAN0).forceKneeForce.ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).forceKneeForce.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
'1.1ANM gudtTest(CHAN0).forceKneeForce.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Forward Force Point #1
lstrParameterName = "FwdForcePt1" 'Metric Parameter
gudtTest(CHAN0).fwdForcePt(1).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(1).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(1).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(1).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Forward Force Point #2
lstrParameterName = "FwdForcePt2" 'Metric Parameter
gudtTest(CHAN0).fwdForcePt(2).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(2).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(2).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(2).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Forward Force Point #3
lstrParameterName = "FwdForcePt3" 'Metric Parameter
gudtTest(CHAN0).fwdForcePt(3).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(3).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(3).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).fwdForcePt(3).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Reverse Force Point #1
lstrParameterName = "RevForcePt1" 'Metric Parameter
gudtTest(CHAN0).revForcePt(1).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(1).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(1).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(1).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Reverse Force Point #2
lstrParameterName = "RevForcePt2" 'Metric Parameter
gudtTest(CHAN0).revForcePt(2).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(2).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(2).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(2).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Reverse Force Point #3
lstrParameterName = "RevForcePt3" 'Metric Parameter
gudtTest(CHAN0).revForcePt(3).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(3).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(3).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).revForcePt(3).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Peak Force
lstrParameterName = "peakForce" 'Metric Parameter
gudtTest(CHAN0).peakForce.high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).peakForce.low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)

'Mechanical Hysteresis Point #1
lstrParameterName = "MechHystPt1" 'Metric Parameter
gudtTest(CHAN0).mechHystPt(1).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(1).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(1).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(1).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Mechanical Hysteresis Point #2
lstrParameterName = "MechHystPt2" 'Metric Parameter
gudtTest(CHAN0).mechHystPt(2).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(2).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(2).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(2).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'Mechanical Hysteresis Point #3
lstrParameterName = "MechHystPt3" 'Metric Parameter
gudtTest(CHAN0).mechHystPt(3).ideal = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(3).high = MpcValue(lstrParameterName, 1, MPCTYPE_HIGHLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(3).low = MpcValue(lstrParameterName, 1, MPCTYPE_LOWLIMIT, 1, MPCVALTYPE_ORDINATE1)
gudtTest(CHAN0).mechHystPt(3).location = MpcValue(lstrParameterName, 1, MPCTYPE_IDEAL, 1, MPCVALTYPE_ABSCISSA1)

'*** STATS Parameters for CP & CPK calculations ***

'STATS: Index 1 (FullClose) Output #1
lstrParameterName = "IdleOutput1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(1).high = ParamValue(lstrParameterName)
lstrParameterName = "IdleOutput1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(1).low = ParamValue(lstrParameterName)

'STATS: Index 2 (Midpoint) Output #1
lstrParameterName = "MidpointOutput1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(2).high = ParamValue(lstrParameterName)
lstrParameterName = "MidpointOutput1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(2).low = ParamValue(lstrParameterName)

'STATS: Index 3 (FullOpen) Output #1
lstrParameterName = "WOTLocation1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(3).high = ParamValue(lstrParameterName)
lstrParameterName = "WOTLocation1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).Index(3).low = ParamValue(lstrParameterName)

'STATS: Index 1 (FullClose) Output #2
lstrParameterName = "IdleOutput2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(1).high = ParamValue(lstrParameterName)
lstrParameterName = "IdleOutput2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(1).low = ParamValue(lstrParameterName)

'STATS: Index 2 (Midpoint) Output #2
lstrParameterName = "MidpointOutput2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(2).high = ParamValue(lstrParameterName)
lstrParameterName = "MidpointOutput2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(2).low = ParamValue(lstrParameterName)

'STATS: Index 3 (FullOpen) Output #2
lstrParameterName = "WOTLocation2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(3).high = ParamValue(lstrParameterName)
lstrParameterName = "WOTLocation2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN1).Index(3).low = ParamValue(lstrParameterName)

'1.1ANM 'STATS: Force Knee Location
'1.1ANM lstrParameterName = "ForceKneeLocationSpecHigh" 'Non-Metric Parameter
'1.1ANM gudtCustomerSpec(CHAN0).forceKneeLoc.high = ParamValue(lstrParameterName)
'1.1ANM lstrParameterName = "ForceKneeLocationSpecLow" 'Non-Metric Parameter
'1.1ANM gudtCustomerSpec(CHAN0).forceKneeLoc.low = ParamValue(lstrParameterName)

'1.1ANM 'STATS: Forward Force at Force Knee Location
'1.1ANM lstrParameterName = "ForceKneeForceSpecHigh" 'Non-Metric Parameter
'1.1ANM gudtCustomerSpec(CHAN0).forceKneeForce.high = ParamValue(lstrParameterName)
'1.1ANM lstrParameterName = "ForceKneeForceSpecLow" 'Non-Metric Parameter
'1.1ANM gudtCustomerSpec(CHAN0).forceKneeForce.low = ParamValue(lstrParameterName)

'STATS: Forward Force Point #1
lstrParameterName = "FwdForcePt1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(1).high = ParamValue(lstrParameterName)
lstrParameterName = "FwdForcePt1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(1).low = ParamValue(lstrParameterName)

'STATS: Forward Force Point #2
lstrParameterName = "FwdForcePt2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(2).high = ParamValue(lstrParameterName)
lstrParameterName = "FwdForcePt2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(2).low = ParamValue(lstrParameterName)

'STATS: Forward Force Point #3
lstrParameterName = "FwdForcePt3SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(3).high = ParamValue(lstrParameterName)
lstrParameterName = "FwdForcePt3SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).fwdForcePt(3).low = ParamValue(lstrParameterName)

'STATS: Reverse Force Point #1
lstrParameterName = "RevForcePt1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(1).high = ParamValue(lstrParameterName)
lstrParameterName = "RevForcePt1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(1).low = ParamValue(lstrParameterName)

'STATS: Reverse Force Point #2
lstrParameterName = "RevForcePt2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(2).high = ParamValue(lstrParameterName)
lstrParameterName = "RevForcePt2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(2).low = ParamValue(lstrParameterName)

'STATS: Reverse Force Point #3
lstrParameterName = "RevForcePt3SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(3).high = ParamValue(lstrParameterName)
lstrParameterName = "RevForcePt3SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).revForcePt(3).low = ParamValue(lstrParameterName)

'STATS: Mechanical Hysteresis Point #1
lstrParameterName = "MechHystPt1SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(1).high = ParamValue(lstrParameterName)
lstrParameterName = "MechHystPt1SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(1).low = ParamValue(lstrParameterName)

'STATS: Mechanical Hysteresis Point #2
lstrParameterName = "MechHystPt2SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(2).high = ParamValue(lstrParameterName)
lstrParameterName = "MechHystPt2SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(2).low = ParamValue(lstrParameterName)

'STATS: Mechanical Hysteresis Point #3
lstrParameterName = "MechHystPt3SpecHigh" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(3).high = ParamValue(lstrParameterName)
lstrParameterName = "MechHystPt3SpecLow" 'Non-Metric Parameter
gudtCustomerSpec(CHAN0).mechHystPt(3).low = ParamValue(lstrParameterName)

'*** PTC-04 Parameters ***
'Output #1
lstrParameterName = "TpulsOutput1" 'Non-Metric Parameter
gudtPTC04(1).Tpuls = ParamValue(lstrParameterName)
lstrParameterName = "TporOutput1" 'Non-Metric Parameter
gudtPTC04(1).Tpor = ParamValue(lstrParameterName)
lstrParameterName = "TprogOutput1" 'Non-Metric Parameter
gudtPTC04(1).Tprog = ParamValue(lstrParameterName)
lstrParameterName = "TholdOutput1" 'Non-Metric Parameter
gudtPTC04(1).Thold = ParamValue(lstrParameterName)
'Output #2
lstrParameterName = "TpulsOutput2" 'Non-Metric Parameter
gudtPTC04(2).Tpuls = ParamValue(lstrParameterName)
lstrParameterName = "TporOutput2" 'Non-Metric Parameter
gudtPTC04(2).Tpor = ParamValue(lstrParameterName)
lstrParameterName = "TprogOutput2" 'Non-Metric Parameter
gudtPTC04(2).Tprog = ParamValue(lstrParameterName)
lstrParameterName = "TholdOutput2" 'Non-Metric Parameter
gudtPTC04(2).Thold = ParamValue(lstrParameterName)

'MLX 90277 Chip Revision
lstrParameterName = "MLX90277Revision" 'Non-metric Enumerated Parameter
gstrMLX90277Revision = ParamEnumValue(lstrParameterName)

'*** Solver Parameters ***
'Output #1
lstrParameterName = "Index1IdealOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(1).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "Index1LocationOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(1).IdealLocation = ParamValue(lstrParameterName)
lstrParameterName = "Index1TargetToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(1).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index1PassFailToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(1).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index2IdealOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(2).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "Index2LocationOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(2).IdealLocation = ParamValue(lstrParameterName)
lstrParameterName = "Index2TargetToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(2).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index2PassFailToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Index(2).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "FilterOutput1" 'Non-Metric Parameter
gudtSolver(1).Filter = ParamValue(lstrParameterName)
lstrParameterName = "InvertOutput1" 'Non-Metric Parameter
gudtSolver(1).InvertSlope = ParamValue(lstrParameterName)
lstrParameterName = "ModeOutput1" 'Non-Metric Parameter
gudtSolver(1).Mode = ParamValue(lstrParameterName)
lstrParameterName = "FaultLevelOutput1" 'Non-Metric Parameter
gudtSolver(1).FaultLevel = ParamValue(lstrParameterName)
lstrParameterName = "MaxOffsetDriftOutput1" 'Non-Metric Parameter
gudtSolver(1).MaxOffsetDrift = ParamValue(lstrParameterName)
lstrParameterName = "MaxAGNDSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).MaxAGND = ParamValue(lstrParameterName)
lstrParameterName = "MinAGNDSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).MinAGND = ParamValue(lstrParameterName)
lstrParameterName = "FCKADJSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).FCKADJ = ParamValue(lstrParameterName)
lstrParameterName = "CKANACHSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).CKANACH = ParamValue(lstrParameterName)
lstrParameterName = "CKDACCHSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).CKDACCH = ParamValue(lstrParameterName)
lstrParameterName = "SlowModeSettingOutput1" 'Non-Metric Parameter
gudtSolver(1).SlowMode = ParamValue(lstrParameterName)
lstrParameterName = "InitialOffsetOutput1" 'Non-Metric Parameter
gudtSolver(1).InitialOffset = ParamValue(lstrParameterName)
lstrParameterName = "HighRGHighFG1" 'Non-Metric Parameter
gudtSolver(1).HighRGHighFG = ParamValue(lstrParameterName)
lstrParameterName = "HighRGLowFG1" 'Non-Metric Parameter
gudtSolver(1).HighRGLowFG = ParamValue(lstrParameterName)
lstrParameterName = "LowRGHighFG1" 'Non-Metric Parameter
gudtSolver(1).LowRGHighFG = ParamValue(lstrParameterName)
lstrParameterName = "LowRGLowFG1" 'Non-Metric Parameter
gudtSolver(1).LowRGLowFG = ParamValue(lstrParameterName)
lstrParameterName = "MinRoughGainOutput1" 'Non-Metric Parameter
gudtSolver(1).MinRG = ParamValue(lstrParameterName)
lstrParameterName = "MaxRoughGainOutput1" 'Non-Metric Parameter
gudtSolver(1).MaxRG = ParamValue(lstrParameterName)
lstrParameterName = "OffsetStepOutput1" 'Non-Metric Parameter
gudtSolver(1).OffsetStep = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA1Output1" 'Non-Metric Parameter
'1.3ANM gudtSolver(1).CodeRatio(1, 1) = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA2Output1" 'Non-Metric Parameter
'1.3ANM gudtSolver(1).CodeRatio(1, 2) = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA3Output1" 'Non-Metric Parameter
'1.3ANM gudtSolver(1).CodeRatio(1, 3) = ParamValue(lstrParameterName)
lstrParameterName = "Output1ProgrammingTestRatio1" 'Non-Metric Parameter '1.3ANM
gudtSolver(1).CodeRatio(2, 1) = ParamValue(lstrParameterName)
lstrParameterName = "Output1ProgrammingTestRatio2" 'Non-Metric Parameter '1.3ANM
gudtSolver(1).CodeRatio(2, 2) = ParamValue(lstrParameterName)
lstrParameterName = "Output1ProgrammingTestRatio3" 'Non-Metric Parameter '1.3ANM
gudtSolver(1).CodeRatio(2, 3) = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowIdealOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(1).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowTargetToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(1).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowPassFailToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(1).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowInitialCodeOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(1).InitialCode = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighIdealOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(2).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighTargetToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(2).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighPassFailToleranceOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(2).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighInitialCodeOutput1" 'Non-Metric Parameter
gudtSolver(1).Clamp(2).InitialCode = ParamValue(lstrParameterName)
lstrParameterName = "ClampStepOutput1" 'Non-Metric Parameter
gudtSolver(1).ClampStep = ParamValue(lstrParameterName)

'Output #2
lstrParameterName = "Index1IdealOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(1).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "Index1LocationOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(1).IdealLocation = ParamValue(lstrParameterName)
lstrParameterName = "Index1TargetToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(1).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index1PassFailToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(1).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index2IdealOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(2).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "Index2LocationOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(2).IdealLocation = ParamValue(lstrParameterName)
lstrParameterName = "Index2TargetToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(2).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "Index2PassFailToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Index(2).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "FilterOutput2" 'Non-Metric Parameter
gudtSolver(2).Filter = ParamValue(lstrParameterName)
lstrParameterName = "InvertOutput2" 'Non-Metric Parameter
gudtSolver(2).InvertSlope = ParamValue(lstrParameterName)
lstrParameterName = "ModeOutput2" 'Non-Metric Parameter
gudtSolver(2).Mode = ParamValue(lstrParameterName)
lstrParameterName = "FaultLevelOutput2" 'Non-Metric Parameter
gudtSolver(2).FaultLevel = ParamValue(lstrParameterName)
lstrParameterName = "MaxOffsetDriftOutput2" 'Non-Metric Parameter
gudtSolver(2).MaxOffsetDrift = ParamValue(lstrParameterName)
lstrParameterName = "MaxAGNDSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).MaxAGND = ParamValue(lstrParameterName)
lstrParameterName = "MinAGNDSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).MinAGND = ParamValue(lstrParameterName)
lstrParameterName = "FCKADJSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).FCKADJ = ParamValue(lstrParameterName)
lstrParameterName = "CKANACHSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).CKANACH = ParamValue(lstrParameterName)
lstrParameterName = "CKDACCHSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).CKDACCH = ParamValue(lstrParameterName)
lstrParameterName = "SlowModeSettingOutput2" 'Non-Metric Parameter
gudtSolver(2).SlowMode = ParamValue(lstrParameterName)
lstrParameterName = "InitialOffsetOutput2" 'Non-Metric Parameter
gudtSolver(2).InitialOffset = ParamValue(lstrParameterName)
lstrParameterName = "HighRGHighFG2" 'Non-Metric Parameter
gudtSolver(2).HighRGHighFG = ParamValue(lstrParameterName)
lstrParameterName = "HighRGLowFG2" 'Non-Metric Parameter
gudtSolver(2).HighRGLowFG = ParamValue(lstrParameterName)
lstrParameterName = "LowRGHighFG2" 'Non-Metric Parameter
gudtSolver(2).LowRGHighFG = ParamValue(lstrParameterName)
lstrParameterName = "LowRGLowFG2" 'Non-Metric Parameter
gudtSolver(2).LowRGLowFG = ParamValue(lstrParameterName)
lstrParameterName = "MinRoughGainOutput2" 'Non-Metric Parameter
gudtSolver(2).MinRG = ParamValue(lstrParameterName)
lstrParameterName = "MaxRoughGainOutput2" 'Non-Metric Parameter
gudtSolver(2).MaxRG = ParamValue(lstrParameterName)
lstrParameterName = "OffsetStepOutput2" 'Non-Metric Parameter
gudtSolver(2).OffsetStep = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA1Output2" 'Non-Metric Parameter
'1.3ANM gudtSolver(2).CodeRatio(1, 1) = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA2Output2" 'Non-Metric Parameter
'1.3ANM gudtSolver(2).CodeRatio(1, 2) = ParamValue(lstrParameterName)
'1.3ANM lstrParameterName = "RatioA3Output2" 'Non-Metric Parameter
'1.3ANM gudtSolver(2).CodeRatio(1, 3) = ParamValue(lstrParameterName)
lstrParameterName = "Output2ProgrammingTestRatio1" 'Non-Metric Parameter '1.3ANM
gudtSolver(2).CodeRatio(2, 1) = ParamValue(lstrParameterName)
lstrParameterName = "Output2ProgrammingTestRatio2" 'Non-Metric Parameter '1.3ANM
gudtSolver(2).CodeRatio(2, 2) = ParamValue(lstrParameterName)
lstrParameterName = "Output2ProgrammingTestRatio3" 'Non-Metric Parameter '1.3ANM
gudtSolver(2).CodeRatio(2, 3) = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowIdealOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(1).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowTargetToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(1).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowPassFailToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(1).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampLowInitialCodeOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(1).InitialCode = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighIdealOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(2).IdealValue = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighTargetToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(2).TargetTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighPassFailToleranceOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(2).PassFailTolerance = ParamValue(lstrParameterName)
lstrParameterName = "ClampHighInitialCodeOutput2" 'Non-Metric Parameter
gudtSolver(2).Clamp(2).InitialCode = ParamValue(lstrParameterName)
lstrParameterName = "ClampStepOutput2" 'Non-Metric Parameter
gudtSolver(2).ClampStep = ParamValue(lstrParameterName)

'*** Machine Parameters ***
lstrParameterName = "" 'Non-Metric Parameter
gudtMachine.parameterName = frmMain.cboParameterFileName.Text
gudtMachine.parameterRev = "3.0.0"

lstrParameterName = "DriveArmSetupCode" 'Non-Metric Parameter
gudtMachine.BOMNumber = ParamValue(lstrParameterName)
lstrParameterName = "stationCode" 'Non-Metric Parameter
gudtMachine.stationCode = ParamValue(lstrParameterName)
lstrParameterName = "seriesID" 'Non-Metric Parameter
gudtMachine.seriesID = ParamValue(lstrParameterName)
lstrParameterName = "PLCComType" 'Non-Metric Parameter
gudtMachine.PLCCommType = ParamValue(lstrParameterName)
lstrParameterName = "RiseTarget" 'Non-Metric Parameter
gudtTest(CHAN0).riseTarget = ParamValue(lstrParameterName)
lstrParameterName = "loadLocation" 'Non-Metric Parameter
gudtMachine.loadLocation = ParamValue(lstrParameterName)
lstrParameterName = "preScanStart" 'Non-Metric Parameter
gudtMachine.preScanStart = ParamValue(lstrParameterName)
lstrParameterName = "preScanStop" 'Non-Metric Parameter
gudtMachine.preScanStop = ParamValue(lstrParameterName)
lstrParameterName = "OffsetForStartScan" 'Non-Metric Parameter
gudtMachine.offset4StartScan = ParamValue(lstrParameterName)
lstrParameterName = "scanLength" 'Non-Metric Parameter
gudtMachine.scanLength = ParamValue(lstrParameterName)
lstrParameterName = "overTravel" 'Non-Metric Parameter
gudtMachine.overTravel = ParamValue(lstrParameterName)
lstrParameterName = "EncCntPerDataPt" 'Non-Metric Parameter
gudtMachine.countsPerTrigger = ParamValue(lstrParameterName)
lstrParameterName = "gearRatio" 'Non-Metric Parameter
gudtMachine.gearRatio = ParamValue(lstrParameterName)
lstrParameterName = "EncoderResolution" 'Non-Metric Parameter
gudtMachine.encReso = ParamValue(lstrParameterName)
lstrParameterName = "PedalZeroForce" 'Non-Metric Parameter
gudtMachine.pedalAtRestLocForce = ParamValue(lstrParameterName)
lstrParameterName = "FKStartTransitionSlope" 'Non-Metric Parameter
gudtMachine.FKSlope = ParamValue(lstrParameterName)
lstrParameterName = "FKStartTransitionWindow" 'Non-Metric Parameter
gudtMachine.FKWindow = ParamValue(lstrParameterName)
lstrParameterName = "FKStartTransitionPercentage" 'Non-Metric Parameter
gudtMachine.FKPercentage = ParamValue(lstrParameterName)
lstrParameterName = "SlopeDevInterval" 'Non-Metric Parameter
gudtMachine.slopeInterval = ParamValue(lstrParameterName)
lstrParameterName = "SlopeDevIncrement" 'Non-Metric Parameter
gudtMachine.slopeIncrement = ParamValue(lstrParameterName)
lstrParameterName = "preScanVelocity" 'Non-Metric Parameter
gudtMachine.preScanVelocity = ParamValue(lstrParameterName)
lstrParameterName = "preScanAcceleration" 'Non-Metric Parameter
gudtMachine.preScanAcceleration = ParamValue(lstrParameterName)
lstrParameterName = "scanVelocity" 'Non-Metric Parameter
gudtMachine.scanVelocity = ParamValue(lstrParameterName)
lstrParameterName = "scanAcceleration" 'Non-Metric Parameter
gudtMachine.scanAcceleration = ParamValue(lstrParameterName)
lstrParameterName = "progVelocity" 'Non-Metric Parameter
gudtMachine.progVelocity = ParamValue(lstrParameterName)
lstrParameterName = "progAcceleration" 'Non-Metric Parameter
gudtMachine.progAcceleration = ParamValue(lstrParameterName)
lstrParameterName = "graphZeroOffset" 'Non-Metric Parameter
gudtMachine.graphZeroOffset = ParamValue(lstrParameterName)
lstrParameterName = "HourlyYieldPartCount" 'Non-Metric Parameter
gudtMachine.currentPartCount = ParamValue(lstrParameterName)
lstrParameterName = "YieldGreen" 'Non-Metric Parameter
gudtMachine.yieldGreen = ParamValue(lstrParameterName)
lstrParameterName = "YieldYellow" 'Non-Metric Parameter
gudtMachine.yieldYellow = ParamValue(lstrParameterName)
lstrParameterName = "xAxisLow" 'Non-Metric Parameter
gudtMachine.xAxisLow = ParamValue(lstrParameterName)
lstrParameterName = "xAxisHigh" 'Non-Metric Parameter
gudtMachine.xAxisHigh = ParamValue(lstrParameterName)
lstrParameterName = "Filter1Location" 'Non-Metric Parameter
gudtMachine.filterLoc(CHAN0) = ParamValue(lstrParameterName)
lstrParameterName = "Filter2Location" 'Non-Metric Parameter
gudtMachine.filterLoc(CHAN1) = ParamValue(lstrParameterName)
lstrParameterName = "Filter3Location" 'Non-Metric Parameter
gudtMachine.filterLoc(CHAN2) = ParamValue(lstrParameterName)
lstrParameterName = "Filter4Location" 'Non-Metric Parameter
gudtMachine.filterLoc(CHAN3) = ParamValue(lstrParameterName)
lstrParameterName = "HomeBlockOffset" 'Non-Metric Parameter
gudtMachine.blockOffset = ParamValue(lstrParameterName)
lstrParameterName = "VRefMode" 'Non-Metric Parameter
gudtMachine.VRefMode = ParamValue(lstrParameterName)
lstrParameterName = "maxLBF" 'Non-Metric Parameter
gudtMachine.maxLBF = ParamValue(lstrParameterName)

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

Public Sub ParameterViewer()
'
'   PURPOSE:   To load the parameters on the viewing page.
'
'  INPUT(S):
' OUTPUT(S):

'Not sure if these variables are needed yet
Dim lintRegionNum As Integer                'Region number
Dim lintRegionCount As Integer              'Count number
Dim lintRow As Integer                      'Row of table
Dim lintColumn As Integer                   'Column of table

frmParamViewer.MSHFlexGrid1.Cols = 5                    'Set up 5 columns
frmParamViewer.MSHFlexGrid1.Rows = 208                  'Set up number of rows
frmParamViewer.MSHFlexGrid1.ColWidth(0, 0) = 4500
frmParamViewer.MSHFlexGrid1.ColWidth(1, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(2, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(3, 0) = 1300
frmParamViewer.MSHFlexGrid1.ColWidth(4, 0) = 1300

lintColumn = 0
lintRow = 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Name": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Ideal": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "High": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Low": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Location": lintColumn = 0

lintRow = lintRow + 1

'*** Test Parameters Output #1 ***
'Index 1 - FullClose By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "IdleValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(1).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(1).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(1).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(1).location: lintColumn = 0

lintRow = lintRow + 1

'Output at Force Knee Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "OutputAtForceKnee": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).outputAtForceKnee.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).outputAtForceKnee.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).outputAtForceKnee.low: lintColumn = 0

lintRow = lintRow + 1

'Index 2 - Midpoint By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MidpointValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(2).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(2).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(2).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(2).location: lintColumn = 0

lintRow = lintRow + 1

'Index 3 - FullOpen By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "WOTValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(3).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(3).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(3).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).Index(3).location: lintColumn = 0

lintRow = lintRow + 1

'Maximum Output
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).maxOutput.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).maxOutput.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).maxOutput.low: lintColumn = 0

lintRow = lintRow + 1

'SinglePoint Linearity Deviation
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SingLinDevVal": lintColumn = lintColumn + 1

' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount("SingLinDevVal", 1, MPCTYPE_IDEAL)

For lintRegionNum = 1 To lintRegionCount
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).start.location: lintColumn = 1
    
    lintRow = lintRow + 1
    
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).SinglePointLin(lintRegionNum).stop.location: lintColumn = 1
    
    lintRow = lintRow + 1
Next lintRegionNum

lintColumn = 0

'Slope Deviation Start
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlopeDevVal": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.start: lintColumn = 1

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.ideal2: lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).slope.stop: lintColumn = 0

lintRow = lintRow + 1

'Full-Close Hysteresis
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FullCloseHys": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).FullCloseHys.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).FullCloseHys.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).FullCloseHys.low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).FullCloseHys.location: lintColumn = 0

lintRow = lintRow + 1

'Evaluation Start
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EvalStartLoc1": lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).evaluate.start: lintColumn = 0

lintRow = lintRow + 1

'Evaluation Stop
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EvalStopLoc1": lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).evaluate.stop: lintColumn = 0

lintRow = lintRow + 1

'*** Test Parameters Output #2 ***
'Index 1 - FullClose By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "IdleValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(1).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(1).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(1).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(1).location: lintColumn = 0

lintRow = lintRow + 1

'Output at Force Knee Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "OutputAtForceKnee": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).outputAtForceKnee.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).outputAtForceKnee.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).outputAtForceKnee.low: lintColumn = 0

lintRow = lintRow + 1

'Index 2 - Midpoint By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MidpointValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(2).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(2).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(2).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(2).location: lintColumn = 0

lintRow = lintRow + 1

'Index 3 - FullOpen By Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "WOTValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(3).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(3).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(3).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).Index(3).location: lintColumn = 0

lintRow = lintRow + 1

'Maximum Output
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxValue": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).maxOutput.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).maxOutput.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).maxOutput.low: lintColumn = 0

lintRow = lintRow + 1

'SinglePoint Linearity Deviation
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SingLinDevVal": lintColumn = lintColumn + 1

' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount("SingLinDevVal", 2, MPCTYPE_IDEAL)

For lintRegionNum = 1 To lintRegionCount
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).start.location: lintColumn = 1
    
    lintRow = lintRow + 1
    
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).SinglePointLin(lintRegionNum).stop.location: lintColumn = 1
    
    lintRow = lintRow + 1
Next lintRegionNum

lintColumn = 0

'Slope Deviation Start
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlopeDevVal": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.start: lintColumn = 1

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.ideal2: lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).slope.stop: lintColumn = 0

lintRow = lintRow + 1

'Full-Close Hysteresis
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FullCloseHys": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).FullCloseHys.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).FullCloseHys.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).FullCloseHys.low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).FullCloseHys.location: lintColumn = 0

lintRow = lintRow + 1

'Evaluation Start
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EvalStartLoc2": lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).evaluate.start: lintColumn = 0

lintRow = lintRow + 1

'Evaluation Stop
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EvalStopLoc2": lintColumn = 4
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN1).evaluate.stop: lintColumn = 0

lintRow = lintRow + 1

'*** Correlation Parameters ***
'Forward Output Correlation
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FwdOutputCorr": lintColumn = lintColumn + 1

' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount("FwdOutputCorrelation", 1, MPCTYPE_IDEAL)

For lintRegionNum = 1 To lintRegionCount
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).start.location: lintColumn = 1
    
    lintRow = lintRow + 1
    
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdOutputCor(lintRegionNum).stop.location: lintColumn = 1

    lintRow = lintRow + 1
Next lintRegionNum

lintColumn = 0

'Reverse Output Correlation
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RevOutputCorr": lintColumn = lintColumn + 1

' Call this stored procedure to learn the number of regions a particular MPC has
lintRegionCount = MPCRegionCount("RevOutputCorrelation", 1, MPCTYPE_IDEAL)

For lintRegionNum = 1 To lintRegionCount
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).start.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).start.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).start.location: lintColumn = 1
    
    lintRow = lintRow + 1
    
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).ideal: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.high: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.low: lintColumn = lintColumn + 1
    frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revOutputCor(lintRegionNum).stop.location: lintColumn = 1

    lintRow = lintRow + 1
Next lintRegionNum

lintColumn = 0

'*** Force Parameters ***
'Pedal at Rest Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "pedalAtRestLoc": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).pedalAtRestLoc.ideal: lintColumn = 0

lintRow = lintRow + 1

'Force Knee Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "forceKneeLoc": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeLoc.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeLoc.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeLoc.low: lintColumn = 0

lintRow = lintRow + 1

'Forward Force at Force Knee Location
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FwdForceAtForceKneeLoc": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeForce.ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeForce.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).forceKneeForce.low: lintColumn = 0

lintRow = lintRow + 1

'Forward Force Point #1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FwdForcePt1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(1).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(1).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(1).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(1).location: lintColumn = 0

lintRow = lintRow + 1

'Forward Force Point #2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FwdForcePt2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(2).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(2).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(2).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(2).location: lintColumn = 0

lintRow = lintRow + 1

'Forward Force Point #3
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FwdForcePt3": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(3).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(3).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(3).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).fwdForcePt(3).location: lintColumn = 0

lintRow = lintRow + 1

'Reverse Force Point #1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RevForcePt1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(1).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(1).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(1).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(1).location: lintColumn = 0

lintRow = lintRow + 1

'Reverse Force Point #2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RevForcePt2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(2).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(2).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(2).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(2).location: lintColumn = 0

lintRow = lintRow + 1

'Reverse Force Point #3
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RevForcePt3": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(3).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(3).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(3).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).revForcePt(3).location: lintColumn = 0

lintRow = lintRow + 1

'Peak Force
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "peakForce": lintColumn = 2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).peakForce.high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).peakForce.low: lintColumn = 0

lintRow = lintRow + 1

'Mechanical Hysteresis Point #1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MechHystPt1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(1).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(1).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(1).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(1).location: lintColumn = 0

lintRow = lintRow + 1

'Mechanical Hysteresis Point #2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MechHystPt2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(2).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(2).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(2).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(2).location: lintColumn = 0

lintRow = lintRow + 1

'Mechanical Hysteresis Point #3
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MechHystPt3": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(3).ideal: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(3).high: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(3).low: lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).mechHystPt(3).location: lintColumn = 0

lintRow = lintRow + 1

'*** PTC-04 Parameters ***
'Output #1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TpulsOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(1).Tpuls: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TporOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(1).Tpor: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TprogOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(1).Tprog: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TholdOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(1).Thold: lintColumn = 0

lintRow = lintRow + 1

'Output #2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TpulsOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(2).Tpuls: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TporOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(2).Tpor: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TprogOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(2).Tprog: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "TholdOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtPTC04(2).Thold: lintColumn = 0

lintRow = lintRow + 1

'MLX 90277 Chip Revision
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MLX90277Revision": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gstrMLX90277Revision: lintColumn = 0

lintRow = lintRow + 1

'*** Solver Parameters ***
'Output #1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1IdealOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(1).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1LocationOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(1).IdealLocation: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1TargetToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(1).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1PassFailToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(1).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2IdealOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(2).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2LocationOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(2).IdealLocation: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2TargetToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(2).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2PassFailToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Index(2).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FilterOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Filter: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "InvertOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).InvertSlope: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ModeOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Mode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FaultLevelOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).FaultLevel: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxOffsetDriftOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).MaxOffsetDrift: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxAGNDSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).MaxAGND: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MinAGNDSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).MinAGND: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FCKADJSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).FCKADJ: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "CKANACHSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CKANACH: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "CKDACCHSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CKDACCH: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlowModeSettingOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).SlowMode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "InitialOffsetOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).InitialOffset: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HighRGHighFG1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).HighRGHighFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HighRGLowFG1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).HighRGLowFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "LowRGHighFG1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).LowRGHighFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "LowRGLowFG1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).LowRGLowFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MinRoughGainOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).MinRG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxRoughGainOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).MaxRG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "OffsetStepOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).OffsetStep: lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA1Output1": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(1, 1): lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA2Output1": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(1, 2): lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA3Output1": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(1, 3): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output1ProgTestRatio1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(2, 1): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output1ProgTestRatio2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(2, 2): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output1ProgTestRatio3": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).CodeRatio(2, 3): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowIdealOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(1).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowTargetToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(1).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowPassFailToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(1).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowInitialCodeOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(1).InitialCode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighIdealOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(2).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighTargetToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(2).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighPassFailToleranceOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(2).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighInitialCodeOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).Clamp(2).InitialCode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampStepOutput1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(1).ClampStep: lintColumn = 0

lintRow = lintRow + 1

'Output #2
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1IdealOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(1).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1LocationOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(1).IdealLocation: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1TargetToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(1).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index1PassFailToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(1).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2IdealOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(2).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2LocationOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(2).IdealLocation: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2TargetToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(2).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Index2PassFailToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Index(2).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FilterOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Filter: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "InvertOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).InvertSlope: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ModeOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Mode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FaultLevelOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).FaultLevel: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxOffsetDriftOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).MaxOffsetDrift: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxAGNDSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).MaxAGND: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MinAGNDSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).MinAGND: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FCKADJSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).FCKADJ: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "CKANACHSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CKANACH: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "CKDACCHSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CKDACCH: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlowModeSettingOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).SlowMode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "InitialOffsetOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).InitialOffset: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HighRGHighFG2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).HighRGHighFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HighRGLowFG2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).HighRGLowFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "LowRGHighFG2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).LowRGHighFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "LowRGLowFG2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).LowRGLowFG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MinRoughGainOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).MinRG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "MaxRoughGainOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).MaxRG: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "OffsetStepOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).OffsetStep: lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA1Output2": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(1, 1): lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA2Output2": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(1, 2): lintColumn = 0

'1.3ANM lintRow = lintRow + 1

'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RatioA3Output2": lintColumn = lintColumn + 1
'1.3ANM frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(1, 3): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output2ProgTestRatio1": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(2, 1): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output2ProgTestRatio2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(2, 2): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Output2ProgTestRatio3": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).CodeRatio(2, 3): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowIdealOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(1).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowTargetToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(1).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowPassFailToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(1).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampLowInitialCodeOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(1).InitialCode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighIdealOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(2).IdealValue: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighTargetToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(2).TargetTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighPassFailToleranceOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(2).PassFailTolerance: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampHighInitialCodeOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).Clamp(2).InitialCode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ClampStepOutput2": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtSolver(2).ClampStep: lintColumn = 0

lintRow = lintRow + 1

'*** Machine Parameters ***
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "ParameterFile": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = frmMain.cboParameterFileName.Text: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "DriveArmSetupCode": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.BOMNumber: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "stationCode": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.stationCode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "seriesID": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.seriesID: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "PLCComType": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.PLCCommType: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "RiseTarget": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtTest(CHAN0).riseTarget: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "loadLocation": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.loadLocation: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "preScanStart": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.preScanStart: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "preScanStop": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.preScanStop: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "OffsetForStartScan": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.offset4StartScan: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "scanLength": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.scanLength: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "overTravel": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.overTravel: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EncCntPerDataPt": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.countsPerTrigger: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "gearRatio": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.gearRatio: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "EncoderResolution": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.encReso: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "PedalZeroForce": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.pedalAtRestLocForce: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FKStartTransitionSlope": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.FKSlope: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FKStartTransitionWindow": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.FKWindow: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "FKStartTransitionPercentage": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.FKPercentage: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlopeDevInterval": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.slopeInterval: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "SlopeDevIncrement": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.slopeIncrement: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "preScanVelocity": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.preScanVelocity: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "preScanAcceleration": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.preScanAcceleration: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "scanVelocity": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.scanVelocity: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "scanAcceleration": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.scanAcceleration: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "progVelocity": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.progVelocity: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "progAcceleration": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.progAcceleration: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "graphZeroOffset": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.graphZeroOffset: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HourlyYieldPartCount": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.currentPartCount: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "YieldGreen": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.yieldGreen: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "YieldYellow": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.yieldYellow: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "xAxisLow": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.xAxisLow: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "xAxisHigh": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.xAxisHigh: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Filter1Location": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.filterLoc(CHAN0): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Filter2Location": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.filterLoc(CHAN1): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Filter3Location": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.filterLoc(CHAN2): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "Filter4Location": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.filterLoc(CHAN3): lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "HomeBlockOffset": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.blockOffset: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "VRefMode": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.VRefMode: lintColumn = 0

lintRow = lintRow + 1

frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = "maxLBF": lintColumn = lintColumn + 1
frmParamViewer.MSHFlexGrid1.TextMatrix(lintRow, lintColumn) = gudtMachine.maxLBF: lintColumn = 0

End Sub

Public Sub CloseDatabaseConnection()
'
'   PURPOSE: To close the current database connection.
'
'  INPUT(S): none
' OUTPUT(S): none

'Close the connection if it isn't already
'If mcnLocalDatabase.State <> adStateClosed Then mcnLocalDatabase.Close

'Set the status variable to false
mblnConnectionActive = False

End Sub

'Public Sub CheckForProgrammingFaultsTestDynamicMPC()
''
''     PURPOSE:  To check for programming faults and set the pass/fail boolean
''
''    INPUT(S):  None.
''   OUTPUT(S):  None.
'
'Dim lintProgrammerNum As Integer
'Dim lintFaultNum As Integer
'Dim lsngIdealSlope As Single
'
''Check the Solver outputs for pass/fail
'For lintProgrammerNum = 1 To 2
'    'NOTE: The Index checks are based on the actual position WOT was programmed at:
'    'Calculate the ideal slope to use in calculating Index limits based on actual locations
'    lsngIdealSlope = (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(1).IdealValue) / (gudtSolver(lintProgrammerNum).Index(2).IdealLocation - gudtSolver(lintProgrammerNum).Index(1).IdealLocation)
'
'    'Check Index 1 (Idle)
'    'old
'    'Call Calc.CheckFault(intProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope, LOWPROGINDEX1, HIGHPROGINDEX1, gintProgFailure())
'
'    'New variables to store High and Low Limits (Dynamic MPCs)
'    gudtSolver(lintProgrammerNum).Index(1).low = gudtSolver(lintProgrammerNum).Index(1).IdealValue - gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope
'    gudtSolver(lintProgrammerNum).Index(1).high = gudtSolver(lintProgrammerNum).Index(1).IdealValue + gudtSolver(lintProgrammerNum).Index(1).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(1) - gudtSolver(lintProgrammerNum).Index(1).IdealLocation) * lsngIdealSlope
'    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).FinalIndexVal(1), gudtSolver(lintProgrammerNum).Index(1).low, gudtSolver(lintProgrammerNum).Index(1).high, LOWPROGINDEX1, HIGHPROGINDEX1, gintProgFailure())
'
'    'Check Index 2 (WOT)
'    'old
'    'Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope, LOWPROGINDEX2, HIGHPROGINDEX2, gintProgFailure())
'
'    'New variables to store High and Low Limits (Dynamic MPCs)
'    gudtSolver(lintProgrammerNum).Index(2).low = gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope
'    gudtSolver(lintProgrammerNum).Index(2).high = gudtSolver(lintProgrammerNum).Index(2).IdealValue + gudtSolver(lintProgrammerNum).Index(2).PassFailTolerance + (gudtSolver(lintProgrammerNum).FinalIndexLoc(2) - gudtSolver(lintProgrammerNum).Index(2).IdealLocation) * lsngIdealSlope
'    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).FinalIndexVal(2), gudtSolver(lintProgrammerNum).Index(2).low, gudtSolver(lintProgrammerNum).Index(2).high, LOWPROGINDEX2, HIGHPROGINDEX2, gintProgFailure())
'
'    'Check the Low Clamp
'    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).FinalClampLowVal, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue - gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(1).IdealValue + gudtSolver(lintProgrammerNum).Clamp(1).PassFailTolerance, LOWCLAMPLOW, HIGHCLAMPLOW, gintProgFailure())
'    'Check the High Clamp
'    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).FinalClampHighVal, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue - gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, gudtSolver(lintProgrammerNum).Clamp(2).IdealValue + gudtSolver(lintProgrammerNum).Clamp(2).PassFailTolerance, LOWCLAMPHIGH, HIGHCLAMPHIGH, gintProgFailure())
'    'Check Offset Drift Code
'    '2.1ANM gintProgFailure(lintProgrammerNum, HIGHOFFSETDRIFT) = (gudtMLX90277(lintProgrammerNum).Read.Drift > gudtSolver(lintProgrammerNum).MaxOffsetDrift)
'    'Check AGND Code
'    Call Calc.CheckFault(lintProgrammerNum, gudtSolver(lintProgrammerNum).MinAGND, gudtSolver(lintProgrammerNum).MaxAGND, gudtSolver(lintProgrammerNum).MinAGND, gudtSolver(lintProgrammerNum).MaxAGND, AGNDFAILURE, AGNDFAILURE, gintProgFailure())
'    'Check Oscillator Adjust Code
'    gintProgFailure(lintProgrammerNum, FCKADJFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.FCKADJ <> gudtSolver(lintProgrammerNum).FCKADJ)
'    'Check Capacitor Frequency Adjust Code
'    gintProgFailure(lintProgrammerNum, CKANACHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKANACH <> gudtSolver(lintProgrammerNum).CKANACH)
'    'Check DAC Code Frequency Adjust Code
'    gintProgFailure(lintProgrammerNum, CKDACCHFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.CKDACCH <> gudtSolver(lintProgrammerNum).CKDACCH)
'    'Check Slow Code
'    gintProgFailure(lintProgrammerNum, SLOWMODEFAILURE) = (gudtMLX90277(lintProgrammerNum).Read.SlowMode <> gudtSolver(lintProgrammerNum).SlowMode)
'Next lintProgrammerNum
'
''Check each output
'For lintProgrammerNum = 1 To 2
'    'Check every fault on each output
'    For lintFaultNum = 1 To PROGFAULTCNT
'        If gintProgFailure(lintProgrammerNum, lintFaultNum) Then
'            gblnProgFailure = True         'Failure occured
'        End If
'    Next lintFaultNum
'Next lintProgrammerNum
'
''Set the part status control
'If gblnProgFailure Then
'    frmMain.ctrStatus1.StatusOnText(1) = "REJECT"
'    frmMain.ctrStatus1.StatusOnColor(1) = vbRed
'Else
'    frmMain.ctrStatus1.StatusOnText(1) = "GOOD"
'    frmMain.ctrStatus1.StatusOnColor(1) = vbGreen
'End If
'
''Turn the status control on
'frmMain.ctrStatus1.StatusValue(1) = True
'
'End Sub

'Public Function GetDipID() As String
''
''   PURPOSE: To return the DIP ID
''
''  INPUT(S): none
'' OUTPUT(S): returns the the DIP ID
'
'On Error GoTo ERROR_GetDipID
'
'    Dim cmd As New ADODB.command
'    Dim sXml As String
'
'    Call BuildXmlStringFindDip(sXml)
'
'    gconnAmad.Open
'
'    With cmd
'        .ActiveConnection = gconnAmad
'        .CommandText = "pspTsopGetDipID"
'        .CommandType = adCmdStoredProc
'    End With
'
'    cmd.Parameters(1) = gdbkDbKeys.ProductID
'    cmd.Parameters(2) = sXml
'
'    cmd.Execute
'
'    If IsNull(cmd.Parameters(3)) Then
'        gconnAmad.Close
'        GetDipID = InsertDeviceInProcess
'    Else
'        GetDipID = UpdateDipAttributeValues(cmd.Parameters(3))
'    End If
'
'EXIT_GetDipID:
'    If gconnAmad.State = adStateOpen Then
'        gconnAmad.Close
'    End If
'    Set cmd = Nothing
'    Exit Function
'ERROR_GetDipID:
'    MsgBox "Error in GetDipID:" & Err.number & "- " & Err.Description
'    Resume EXIT_GetDipID
'
'End Function

Public Sub BuildXmlStringFindDip(sXml As String)

    'Build XML String of DIP Attribute Values
    sXml = "<Root>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " RowID=" & """" & "1" & """"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Lot" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Lot & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " RowID=" & """" & "2" & """"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Wafer" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Wafer & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " RowID=" & """" & "3" & """"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_X" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.X & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " RowID=" & """" & "4" & """"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Y" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Y & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "</Root>"
    
End Sub

Public Sub PopulateLotTypesList()

    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.command
    Dim par As New ADODB.Parameter
    Dim i As Integer
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetLotTypes"
        .CommandType = adCmdStoredProc
    End With
    
    With rs
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .ActiveConnection = gconnAmad
    End With
    
    Set rs = cmd.Execute
    
    i = 0
    Do Until rs.EOF
        frmLotType.cboLotFileType.AddItem rs!LotTypeName
        rs.MoveNext
    Loop
        
    rs.Close
    gconnAmad.Close
    Set rs = Nothing
    Set par = Nothing
    Set cmd = Nothing

End Sub

Public Function InsertDeviceInProcess() As String

On Error GoTo ERROR_InsertDeviceInProcess
    
    Dim cmd As New ADODB.command
    Dim sXml As String
    
    Call BuildXmlStringDipAttributes(sXml)
    
    If gconnAmad.State = adStateClosed Then
        gconnAmad.Open
    End If
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsDeviceInProcess"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProductID
    cmd.Parameters(2) = sXml
    
    cmd.Execute
    
    If IsNull(cmd.Parameters(3)) Then
        InsertDeviceInProcess = ""
    Else
        InsertDeviceInProcess = cmd.Parameters(3)
    End If

EXIT_InsertDeviceInProcess:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Function
ERROR_InsertDeviceInProcess:
    MsgBox "Error in InsertDeviceInProcess:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDeviceInProcess

End Function

Public Sub BuildXmlStringDipAttributes(sXml As String)

    Dim lintYear As Integer
    Dim lintJulianDate As Integer
    Dim lstrShift As String
    Dim lintStation As Integer
    Dim lstrPalletLoad As String

    'Disabled for testing
    'Decode the Date Code
    Call MLX90277.DecodeDateCode(gstrDateCode, lintYear, lintJulianDate, lstrShift, lintStation, lstrPalletLoad)
   
    'Build XML String of DIP Attribute Values
    sXml = "<Root>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Lot" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Lot & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Wafer" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Wafer & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_X" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.X & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "MLX_Y" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gudtMLX90277(1).Read.Y & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "ProdYear" & """"
    sXml = sXml & " DipAttributeValue=" & """" & lintYear & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "ProdDayOfYear" & """"
    sXml = sXml & " DipAttributeValue=" & """" & lintJulianDate & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "ProdShift" & """"
    sXml = sXml & " DipAttributeValue=" & """" & lstrShift & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "ProgrammingStation" & """"
    sXml = sXml & " DipAttributeValue=" & """" & lintStation & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "PedalLoadStation" & """"
    sXml = sXml & " DipAttributeValue=" & """" & lstrPalletLoad & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "EncodedSerialNumber" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gstrSerialNumber & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "<DipAttribute"
    sXml = sXml & " DipAttributeName=" & """" & "DateCode" & """"
    sXml = sXml & " DipAttributeValue=" & """" & gstrDateCode & """"
    sXml = sXml & "></DipAttribute>"
    sXml = sXml & "</Root>"

End Sub

Public Sub StoreDynamicProgParamValues()

    Call InsertDynamicProgParamValue("PalletNumber", gintPalletNumber)
    Call InsertDynamicProgParamValue("LockedPart", gblnLockedPart)

End Sub

Public Sub InsertDynamicProgParamValue(ParmName As String, ParmValue As Variant)
On Error GoTo ERROR_InsertDynamicProgParamValue

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsertDynamicProgParamValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.ProgrammingID
    cmd.Parameters(2) = gdbkDbKeys.ProductID
    cmd.Parameters(3) = gdbkDbKeys.StationID
    cmd.Parameters(4) = ParmName
    cmd.Parameters(5) = ParmValue
        
    cmd.Execute

EXIT_InsertDynamicProgParamValue:
    gconnAmad.Close
    Set cmd = Nothing
    Exit Sub
ERROR_InsertDynamicProgParamValue:
    MsgBox "Error in InsertDynamicProgParamValue:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDynamicProgParamValue
End Sub

Public Sub StoreDynamicTestParamValues()

    Call InsertDynamicTestParamValue("PalletNumber", gintPalletNumber)

End Sub

Public Sub InsertDynamicTestParamValue(ParmName As String, ParmValue As Variant)
On Error GoTo ERROR_InsertDynamicTestParamValue

    Dim cmd As New ADODB.command
    
    gconnAmad.Open
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsertDynamicTestParamValue"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.TestID
    cmd.Parameters(2) = gdbkDbKeys.ProductID
    cmd.Parameters(3) = gdbkDbKeys.StationID
    cmd.Parameters(4) = ParmName
    cmd.Parameters(5) = ParmValue
        
    cmd.Execute

EXIT_InsertDynamicTestParamValue:
    gconnAmad.Close
    Set cmd = Nothing
    Exit Sub
ERROR_InsertDynamicTestParamValue:
    MsgBox "Error in InsertDynamicTestParamValue:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDynamicTestParamValue
End Sub

Public Sub GetAnomalyInfo(rsAnomaly As ADODB.Recordset)
On Error GoTo ERROR_GetAnomalyInfo

    Dim cmd As New ADODB.command
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetAnomalyInfo"
        .CommandType = adCmdStoredProc
    End With
            
    With rsAnomaly
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .ActiveConnection = gconnAmad
    End With

    cmd.Parameters(1) = gdbkDbKeys.TSOP_ID
    cmd.Parameters(2) = gintAnomaly
    
    Set rsAnomaly = cmd.Execute

EXIT_GetAnomalyInfo:
    Set cmd = Nothing
    Exit Sub
ERROR_GetAnomalyInfo:
    MsgBox "Error in GetAnomalyInfo:" & Err.number & "- " & Err.Description
    Resume EXIT_GetAnomalyInfo

End Sub

Public Function GetUndefinedAnomalyID() As String
On Error GoTo ERROR_GetUndefinedAnomalyID

    Dim cmd As New ADODB.command
    Dim intTsopModeValue As Integer
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopGetUndefinedAnomalyID"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.TSOP_ID
    
    cmd.Execute
    
    If IsNull(cmd.Parameters(2)) Then
        GetUndefinedAnomalyID = ""
    Else
        GetUndefinedAnomalyID = cmd.Parameters(2)
    End If
    
EXIT_GetUndefinedAnomalyID:
    Set cmd = Nothing
    Exit Function
ERROR_GetUndefinedAnomalyID:
    MsgBox "Error in GetUndefinedAnomalyID:" & Err.number & "- " & Err.Description
    Resume EXIT_GetUndefinedAnomalyID

End Function

Public Sub InsertTsopAnomaly()
On Error GoTo ERROR_InsertTsopAnomaly
    
    Dim cmd As New ADODB.command
    
    If gconnAmad.State = adStateClosed Then
        gconnAmad.Open
    End If
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopInsTsopAnomaly"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = gdbkDbKeys.AnomalyID
    cmd.Parameters(2) = gstrSubProcess
    cmd.Parameters(3) = gdbkDbKeys.TsopStartupID
    cmd.Parameters(4) = grsTsopAnomaly!AnomalyMessage
    cmd.Parameters(5) = grsTsopAnomaly!AnomalyDateTime
    cmd.Parameters(6) = grsTsopAnomaly!Operator
    cmd.Parameters(7) = grsTsopAnomaly!UndefinedAnomalyNumber
    If gstrSubProcess = "Programming" Then
        cmd.Parameters(8) = gdbkDbKeys.ProgrammingID
        cmd.Parameters(9) = Null
    ElseIf gstrSubProcess = "FunctionalTest" Then
        cmd.Parameters(8) = Null
        cmd.Parameters(9) = gdbkDbKeys.TestID
    Else
        cmd.Parameters(8) = Null
        cmd.Parameters(9) = Null
    End If
        
    cmd.Execute
    
EXIT_InsertTsopAnomaly:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Sub
ERROR_InsertTsopAnomaly:
    MsgBox "Error in InsertTsopAnomaly:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertTsopAnomaly
End Sub

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

'Display the current status
frmMain.staMessage.Panels(1).Text = "Opening Connection to local database."
If OpenDatabaseConnection Then
    lblnConnectionOK = True
End If

'If the connection was set, exit the for loop
If Not lblnConnectionOK Then
    'Close the connection before trying to open another one
    Call CloseDatabaseConnection
End If

If lblnConnectionOK Then
    'Let the user know the connection is set up
    frmMain.staMessage.Panels(1).Text = "Connection to local database has been initialized."
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

Public Function OpenDatabaseConnection() As Boolean
'
'   PURPOSE: To open a database connection to the requested database.
'
'  INPUT(S): none
' OUTPUT(S): returns a boolean based on whether or not there was an error opening

mblnConnectionActive = False

Dim strConnection As String

strConnection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Initial Catalog=CN7007PD0004_AutoProductMfgTest;Data Source=CN7007PD0004"

With gconnAmad
    .ConnectionString = strConnection
    .CursorLocation = adUseClient        'This will allow getting record counts
'    .Open
End With

'The connection to the database is now active
mblnConnectionActive = True
OpenDatabaseConnection = mblnConnectionActive

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
lstrDirPath = UDBRAWDATAPATH
lstrPathName = lstrDirPath & "\"

If Not gfsoFileSystemObject.FolderExists(lstrDirPath) Then
    gfsoFileSystemObject.CreateFolder (lstrDirPath)
End If

lstrFileName = gdbkDbKeys.TestID & RAWDATAEXTENSION

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

'Public Function NewErrorLogFile(MessageText As String) As String
''
''   PURPOSE:   To output error and message to UDB
''
''  INPUT(S):   MessageText = Error message text
''
'' OUTPUT(S):   NewErrorLogFile = Error Description from UDB
''
'
'    Dim lstrType As String               'Error Type
'    Dim lstrErrorDescription As String   'Error Description
'
'    Dim rsAnomaly As New ADODB.Recordset 'TER_7/9/07 need recordset variable for anomaly info
'    Dim bUndefinedAnomaly As Boolean     'TER_7/9/07 flag to indicate undefined anomaly
'
'    gconnAmad.Open
'    Call GetAnomalyInfo(rsAnomaly)
'    If rsAnomaly.BOF And rsAnomaly.EOF Then
'        gdbkDbKeys.AnomalyID = GetUndefinedAnomalyID
'        bUndefinedAnomaly = True
'        lstrType = "Undefined Anomaly Type"
'        lstrErrorDescription = "Undefined Anomaly"
'    Else
'        bUndefinedAnomaly = False
'        gdbkDbKeys.AnomalyID = rsAnomaly!AnomalyID
'        lstrType = rsAnomaly!AnomalyType
'        lstrErrorDescription = rsAnomaly!AnomalyDescription
'    End If
'    gconnAmad.Close
'
'    If grsTsopAnomaly.RecordCount > 0 Then
'        grsTsopAnomaly.MoveFirst
'        Do Until grsTsopAnomaly.EOF
'            grsTsopAnomaly.Delete
'            grsTsopAnomaly.MoveNext
'        Loop
'    End If
'    grsTsopAnomaly.AddNew
'    grsTsopAnomaly!AnomalyMessage = lstrErrorDescription
'    grsTsopAnomaly!AnomalyDateTime = Now()
'    grsTsopAnomaly!Operator = frmMain.ctrSetupInfo1.Operator
'    If bUndefinedAnomaly Then
'        grsTsopAnomaly!UndefinedAnomalyNumber = gintAnomaly
'    Else
'        grsTsopAnomaly!UndefinedAnomalyNumber = Null
'    End If
'    grsTsopAnomaly.Update
'
'    Call InsertTsopAnomaly
'
'    grsTsopAnomaly.Delete
'
'    NewErrorLogFile = lstrErrorDescription
'
'End Function

Public Function UpdateDipAttributeValues(DeviceInProcessID As String) As String
On Error GoTo ERROR_InsertDeviceInProcess
    
    Dim cmd As New ADODB.command
    Dim sXml As String
    
    Call BuildXmlStringDipAttributes(sXml)
    
    If gconnAmad.State = adStateClosed Then
        gconnAmad.Open
    End If
    
    With cmd
        .ActiveConnection = gconnAmad
        .CommandText = "pspTsopUpdateDipAttributeValues"
        .CommandType = adCmdStoredProc
    End With
    
    cmd.Parameters(1) = DeviceInProcessID
    cmd.Parameters(2) = sXml
    
    cmd.Execute
    
    If cmd.Parameters(0) <> 0 Or IsNull(cmd.Parameters(0)) Then
        UpdateDipAttributeValues = ""
    Else
        UpdateDipAttributeValues = DeviceInProcessID
    End If

EXIT_InsertDeviceInProcess:
    If gconnAmad.State = adStateOpen Then
        gconnAmad.Close
    End If
    Set cmd = Nothing
    Exit Function
ERROR_InsertDeviceInProcess:
    MsgBox "Error in InsertDeviceInProcess:" & Err.number & "- " & Err.Description
    Resume EXIT_InsertDeviceInProcess
    
End Function
