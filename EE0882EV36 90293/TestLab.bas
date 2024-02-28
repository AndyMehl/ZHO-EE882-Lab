Attribute VB_Name = "TestLab"
'**********************************TestLab.BAS**********************************
'
'   Pedal Programming & Scanning Software, supplemented by SeriesXXX.Bas.
'   This module should handle all 700/702/703/705 series testlab programmer/
'   scanners and database recall software.
'   The software is to be kept in the pedal software library, EE947.
'
'VER    DATE      BY   PURPOSE OF MODIFICATION                          TAG
'1.0  03/28/2006  ANM  First release per SCN# MISC-092 (3365).       1.0ANM
'1.1  05/04/2006  ANM  Updates per SCN# MISC-094 (3423).             1.1ANM
'1.2  08/17/2006  ANM  Updates per SCN# MISC-100 (3521).             1.2ANM
'1.3  12/05/2006  ANM  Updates per SCN# MISC-101 (3636).             1.3ANM
'1.4  08/30/2007  ANM  Removed force knee per SCN# 705T-007 (3973).  1.4ANM
'1.5  02/06/2008  ANM  Updated to save ZG to file per SCN# 4087.     1.5ANM
'1.6  05/02/2008  ANM  Fixed stats and added MLX I.                  1.6ANM
'1.7  06/05/2008  ANM  Added Abs Lin per SCN# 4167.                  1.7ANM
'1.8  01/19/2009  ANM  Updated for PDF prints per SCN# 4258.         1.8ANM
'1.9  06/22/2010  ANM  Added MLX WOT I per SCN# 4585.                1.9ANM
'1.9a 01/05/2011  ANM  Update force per SCN# 4698.                   1.9aANM
'

Type ExposureType
    TypeofDust              As String               'Exposure type for dust type
    AmountofDust            As Single               'Exposure type for dust amount
    StirTime                As Single               'Exposure type for stir time
    SettleTime              As Single               'Exposure type for settle time
    Duration                As String               'Exposure type for duration
    NumberofCycles          As String               'Exposure type for number of cycles
    NewNumberofCycles       As String               'Exposure type for new number of cycles
    TotalNumberofCycles     As String               'Exposure type for total number of cycles
    Condition               As String               'Exposure type for condition
    Profile                 As String               'Exposure type for profile
    Temperature             As String               'Exposure type for temperature
    Planes                  As String               'Exposure type for planes
    Frequency               As String               'Exposure type for frequency
    LowTemp                 As Single               'Exposure type for low temp
    HighTemp                As Single               'Exposure type for high temp
    LowTempTime             As Single               'Exposure type for low temp time
    HighTempTime            As Single               'Exposure type for high temp time
    RelativeHumidity        As Single               'Exposure type for relative humididty
    Substance               As String               'Exposure type for Substance
End Type

Type ExposureData
    Dust                    As ExposureType         'Exposure variable for dust data
    Vibration               As ExposureType         'Exposure variable for vibration data
    ThermalShock            As ExposureType         'Exposure variable for thermal shock data
    SaltSpray               As ExposureType         'Exposure variable for salt spray data
    OperationalEndurance    As ExposureType         'Exposure variable for operational endurance data
    SnapBack                As ExposureType         'Exposure variable for snapback data
    HighTempSoak            As ExposureType         'Exposure variable for high temp soak data
    HTempHHumiditySoak      As ExposureType         'Exposure variable for high temp high humidity soak data
    LowTempSoak             As ExposureType         'Exposure variable for low temp soak data
    WaterSpray              As ExposureType         'Exposure variable for water spray data
    ChemResistance          As ExposureType         'Exposure variable for chemical resistance data
    Exposure                As ExposureType         'Exposure variable for exposure data
End Type
Public gudtExposure As ExposureData
Public gstrCustomerName As String
Public gstrCTSPartNum As String
Public gstrPartName As String

Public gblnDustExp As Boolean
Public gblnVibrationExp As Boolean
Public gblnDitherExp As Boolean
Public gblnOzoneExp As Boolean
Public gblnThermalShockExp As Boolean
Public gblnSaltSprayExp As Boolean
Public gblnInitialExp As Boolean
Public gblnOperStrnExp As Boolean
Public gblnLateralStrnExp As Boolean
Public gblnOpStrnStopExp As Boolean
Public gblnImpactStrnExp As Boolean
Public gblnOperEndurExp As Boolean
Public gblnSnapbackExp As Boolean
Public gblnHighTempExp As Boolean
Public gblnHighTempHighHumidExp As Boolean
Public gblnLowTempExp As Boolean
Public gblnWaterSprayExp As Boolean
Public gblnChemResExp As Boolean
Public gblnCondenExp As Boolean
Public gblnESDElecExp As Boolean
Public gblnEMWaveResElecExp As Boolean
Public gblnBilkCInjElecExp As Boolean
Public gblnIgnitionNoiseElecExp As Boolean
Public gblnNarRadEMEElecExp As Boolean
Public gblnExposure As Boolean
Public gblnGraphsLoaded As Boolean
Public gblnCustPrintType As Boolean
Public gblnTLPrintType As Boolean

'1.8ANM \/\/
Public gstrType As String
Public gblnSkipPDF As Boolean
Public Const PDFPATH = "D:\Data\705\PDF\"
Public Const PDFWINDOW = "Microsoft*"
Public Declare Function BringWindowToTop Lib "user32" (ByVal HWND As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Integer
Private Const GW_HWNDNEXT = 2
Private Declare Function GetWindow Lib "user32" (ByVal HWND As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWND As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWND As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const FORCEEXT = ".txt" '1.9aANM
Private Const FORCEFILEPATH = "C:\CTSLIBVB\Force\" '1.9aANM

Public Sub ForceLoad()
'
'   PURPOSE:   To input force gain/offset for system
'
'  INPUT(S): none
' OUTPUT(S): none
'1.9aANM new sub

Dim lintFileNum As Integer
Dim lsngFG As Single
Dim lsngFO As Single

On Error GoTo ForceLoad_Err

frmMain.MousePointer = vbHourglass

'Get a file number
lintFileNum = FreeFile

'Check to see if file exists if not exit sub
If gfsoFileSystemObject.FileExists(FORCEFILEPATH & "Force" & FORCEEXT) Then
    Open FORCEFILEPATH & "Force" & FORCEEXT For Input As #lintFileNum
Else
    frmMain.MousePointer = vbNormal
    Exit Sub
End If

'** General Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, lsngFG, lsngFO

'Set values
If (lsngFG <> 0) And (lsngFO <> 0) Then
    gsngForceGain = lsngFG
    gsngForceOffset = lsngFO
End If

'Close the file
Close #lintFileNum
frmMain.MousePointer = vbNormal

Exit Sub
ForceLoad_Err:

    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Force File!"

End Sub

Public Function FindWindowPartial(ByVal Title As String, ByVal Class As String) As Long
    Dim hWndThis As Long
    hWndThis = FindWindow(vbNullString, vbNullString)
    While hWndThis
        Dim sTitle As String, sClass As String
        sTitle = Space$(255)
        sTitle = left$(sTitle, GetWindowText(hWndThis, sTitle, Len(sTitle)))
        sClass = Space$(255)
        sClass = left$(sClass, GetClassName(hWndThis, sClass, Len(sClass)))
        If sTitle Like Title And sClass Like Class Then
            FindWindowPartial = hWndThis
            Exit Function
        End If
        hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
    Wend
End Function

Public Sub SaveTLProgResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
'1.5ANM added ZG x-position

Dim lintFileNum As Integer
Dim lstrFileName As String

'Make the results file name
lstrFileName = gstrLotName + " Programming Results" & DATAEXT
'Get a file
lintFileNum = FreeFile

'If file does not exist then add a header
If Not gfsoFileSystemObject.FileExists(PARTPROGDATAPATH + lstrFileName) Then
    Open PARTPROGDATAPATH + lstrFileName For Append As #lintFileNum
    'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, & Pallet Number
    Print #lintFileNum, _
        "Part Number,"; _
        "Sample Number,"; _
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
    'Programming Process Variables Vout #1
    Print #lintFileNum, _
        "Vout #1 Cycle 1 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #1 Cycle 1 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #1 Cycle 1 Step 1 Measured Slope 3 [%/°],"; _
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
    'Programming Process Variables Vout #1
    Print #lintFileNum, _
        "Vout #2 Cycle 1 Step 1 Measured Slope 1 [%/°],"; _
        "Vout #2 Cycle 1 Step 1 Measured Slope 2 [%/°],"; _
        "Vout #2 Cycle 1 Step 1 Measured Slope 3 [%/°],"; _
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
        "Series,"; _
        "TestLog #,"; _
        "Sample #"
Else
    Open PARTPROGDATAPATH + lstrFileName For Append As #lintFileNum
End If
'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, & Pallet Number
Print #lintFileNum, _
    gstrSerialNumber; ","; _
    gstrSampleNum; ","; _
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
'Programming Process Variables Vout#1
Print #lintFileNum, _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(1).Cycle(1).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
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
'Programming Process Variables Vout#2
Print #lintFileNum, _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(1).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(2).CalculatedSlope, 5), "0.00000"); ","; _
    Format(Round(gudtSolver(2).Cycle(1).Step(1).Test(3).CalculatedSlope, 5), "0.00000"); ","; _
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
    frmMain.ctrSetupInfo1.Temperature; ","; _
    frmMain.ctrSetupInfo1.Series; ","; _
    frmMain.ctrSetupInfo1.TLNum; ","; _
    frmMain.ctrSetupInfo1.Sample
'Close the file
Close #lintFileNum

End Sub

Public Sub Save700TLScanResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lstrFileName As String
'Dim lstrExp As String
'Dim lstrExp2 As String
'
'lstrExp = ""
'lstrExp2 = ""
'If gblnDustExp Then
'    lstrExp = "Dust "
'    lstrExp2 = lstrExp2 & "Dust(" & gudtExposure.Dust.TypeofDust & " " & gudtExposure.Dust.AmountofDust & " " & gudtExposure.Dust.StirTime & " " & gudtExposure.Dust.SettleTime & " " & gudtExposure.Dust.Frequency & " " & gudtExposure.Dust.NumberofCycles & " " & gudtExposure.Dust.Condition & ") "
'End If
'If gblnVibrationExp Then
'    lstrExp = lstrExp & "Vibration "
'    lstrExp2 = lstrExp2 & "Vibration(" & gudtExposure.Vibration.Profile & " " & gudtExposure.Vibration.Temperature & " " & gudtExposure.Vibration.Duration & " " & gudtExposure.Vibration.Planes & " " & gudtExposure.Vibration.NumberofCycles & " " & gudtExposure.Vibration.Frequency & ") "
'End If
'If gblnDitherExp Then lstrExp = lstrExp & "Dither "
'If gblnThermalShockExp Then
'    lstrExp = lstrExp & "Thermal Shock "
'    lstrExp2 = lstrExp2 & "Thermal Shock(" & CStr(gudtExposure.ThermalShock.LowTemp) & " " & gudtExposure.ThermalShock.LowTempTime & " " & CStr(gudtExposure.ThermalShock.HighTemp) & " " & gudtExposure.ThermalShock.HighTempTime & " " & gudtExposure.ThermalShock.NumberofCycles & " " & gudtExposure.ThermalShock.Condition & ") "
'End If
'If gblnSaltSprayExp Then
'    lstrExp = lstrExp & "Salt Spray "
'    lstrExp2 = lstrExp2 & "Salt Spray(" & CStr(gudtExposure.SaltSpray.Duration) & " " & gudtExposure.SaltSpray.Condition & ") "
'End If
'If gblnInitialExp Then lstrExp = lstrExp & "Initial "
'If gblnExposure Then
'    lstrExp = lstrExp & "Exposure "
'    lstrExp2 = lstrExp2 & "Exposure (" & gudtExposure.Exposure.Condition & ") "
'End If
'If gblnOperStrnExp Then lstrExp = lstrExp & "Operational Strength "
'If gblnLateralStrnExp Then lstrExp = lstrExp & "Lateral Strength "
'If gblnOpStrnStopExp Then lstrExp = lstrExp & "Operational Strength with Stopper "
'If gblnImpactStrnExp Then lstrExp = lstrExp & "Impact Strength "
'If gblnOperEndurExp Then
'    lstrExp = lstrExp & "Operational Endurance "
'    lstrExp2 = lstrExp2 & "Operational Endurance(" & gudtExposure.OperationalEndurance.Temperature & " " & gudtExposure.OperationalEndurance.NewNumberofCycles & " " & gudtExposure.OperationalEndurance.TotalNumberofCycles & " " & gudtExposure.OperationalEndurance.Condition & ") "
'End If
'If gblnSnapbackExp Then
'    lstrExp = lstrExp & "Snapback "
'    lstrExp2 = lstrExp2 & "Snapback(" & gudtExposure.SnapBack.Temperature & " " & gudtExposure.SnapBack.NumberofCycles & " " & gudtExposure.SnapBack.Condition & ") "
'End If
'If gblnHighTempExp Then
'    lstrExp = lstrExp & "High Temp Soak "
'    lstrExp2 = lstrExp2 & "High Temp Soak(" & CStr(gudtExposure.HighTempSoak.Temperature) & " " & gudtExposure.HighTempSoak.Duration & " " & gudtExposure.HighTempSoak.Condition & ") "
'End If
'If gblnHighTempHighHumidExp Then
'    lstrExp = lstrExp & "High Temp - High Humidity Soak "
'    lstrExp2 = lstrExp2 & "High Temp - High Humidity Soak(" & CStr(gudtExposure.HTempHHumiditySoak.Temperature) & " " & CStr(gudtExposure.HTempHHumiditySoak.RelativeHumidity) & " " & gudtExposure.HTempHHumiditySoak.Duration & " " & gudtExposure.HTempHHumiditySoak.Condition & ") "
'End If
'If gblnLowTempExp Then
'    lstrExp = lstrExp & "Low Temp Soak "
'    lstrExp2 = lstrExp2 & "Low Temp Soak(" & CStr(gudtExposure.LowTempSoak.Temperature) & " " & gudtExposure.LowTempSoak.Duration & " " & gudtExposure.LowTempSoak.Condition & ") "
'End If
'If gblnWaterSprayExp Then
'    lstrExp = lstrExp & "Water Spray "
'    lstrExp2 = lstrExp2 & "Water Spray(" & gudtExposure.WaterSpray.Duration & " " & gudtExposure.WaterSpray.Condition & ") "
'End If
'If gblnChemResExp Then
'    lstrExp = lstrExp & "Chemical Resistance "
'    lstrExp2 = lstrExp2 & "Chemical Resistance(" & CStr(gudtExposure.ChemResistance.Temperature) & " " & gudtExposure.ChemResistance.Duration & " " & gudtExposure.ChemResistance.Substance & ") "
'End If
'If gblnCondenExp Then lstrExp = lstrExp & "Condensation "
'If gblnESDElecExp Then lstrExp = lstrExp & "ElectroStatic Discharge "
'If gblnEMWaveResElecExp Then lstrExp = lstrExp & "Electromagnetic Wave Resistance "
'If gblnBilkCInjElecExp Then lstrExp = lstrExp & "Bilk Current Injection "
'If gblnIgnitionNoiseElecExp Then lstrExp = lstrExp & "Ignition Noise "
'If gblnNarRadEMEElecExp Then lstrExp = lstrExp & "Narrowband Radiated Electromagnetic Energy "
'
''Make the results file name
'lstrFileName = gstrLotName + " Scan Results" & DATAEXT
''Get a file
'lintFileNum = FreeFile
'
''If file does not exist then add a header
'If Not gfsoFileSystemObject.FileExists(PARTSCANDATAPATH + lstrFileName) Then
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'    'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'    Print #lintFileNum, _
'        "Part Number,"; _
'        "Sample Number,"; _
'        "TestLog #,"; _
'        "Date Code,"; _
'        "Date/Time,"; _
'        "S/W Revision,"; _
'        "Parameter File Name,"; _
'        "Pallet Number,"; _
'        "Exposures,"; _
'        "Exposure Data,";
'    'Output #1
'    Print #lintFileNum, _
'        "Idle Value Output #1 [%],"; _
'        "Force Knee Output #1 [%],"; _
'        "Midpoint Value Output #1 [%],"; _
'        "WOT Location Output #1 [°],"; _
'        "Full-Pedal-Travel Output #1 [%],"; _
'        "Max Output #1 [%],"; _
'        "Max Absolute Linearity Deviation % of Tol Output #1 [% Tol],"; _
'        "Max Absolute Linearity Deviation Output #1 [%],"; _
'        "Min Absolute Linearity Deviation Output #1 [%],"; _
'        "Max Slope Deviation Output #1 [% of Ideal Slope],"; _
'        "Min Slope Deviation Output #1 [% of Ideal Slope],"; _
'        "Peak Hysteresis Output #1 [%],";
'    'Output #2
'    Print #lintFileNum, _
'        "Idle Value Output #2 [%],"; _
'        "Force Knee Output #2 [%],"; _
'        "Midpoint Value Output #2 [%],"; _
'        "WOT Location Output #2 [°],"; _
'        "Full-Pedal-Travel Output #2 [%],"; _
'        "Max Output #2 [%],"; _
'        "Max Absolute Linearity Deviation % of Tol Output #2 [% Tol],"; _
'        "Max Absolute Linearity Deviation Output #2 [%],"; _
'        "Min Absolute Linearity Deviation Output #2 [%],"; _
'        "Max Slope Deviation Output #2 [% of Ideal Slope],"; _
'        "Min Slope Deviation Output #2 [% of Ideal Slope],"; _
'        "Peak Hysteresis Output #2 [%],";
'    'Correlation
'    Print #lintFileNum, _
'        "Max Forward Output Corr % of Tolerance [% Tol],"; _
'        "Max Forward Output Corr [%],"; _
'        "Min Forward Output Corr [%],"; _
'        "Max Reverse Output Corr % of Tolerance [% Tol],"; _
'        "Max Reverse Output Corr [%],"; _
'        "Min Reverse Output Corr [%],";
'    'Force
'    Print #lintFileNum, _
'        "Pedal at Rest Location [°],"; _
'        "Force Knee Location [°],"; _
'        "Forward Force at Force Knee Location [N],"; _
'        "Full-Pedal-Travel Location [°],"; _
'        "Full-Pedal-Travel Force [N],"; _
'        "Forward Force Point 1 [N],"; _
'        "Forward Force Point 2 [N],"; _
'        "Forward Force Point 3 [N],"; _
'        "Reverse Force Point 1 [N],"; _
'        "Reverse Force Point 2 [N],"; _
'        "Reverse Force Point 3 [N],"; _
'        "Peak Force [N],"; _
'        "Mechanical Hysteresis Point 1 [% of Forward Force],"; _
'        "Mechanical Hysteresis Point 2 [% of Forward Force],"; _
'        "Mechanical Hysteresis Point 3 [% of Forward Force],"; _
'    'Kickdown
'    Print #lintFileNum, _
'        "Kickdown Start Location [°],"; _
'        "Kickdown Start Force [N],"; _
'        "Output #1 at Kickdown Start Location [%],"; _
'        "Kickdown On Location [°],"; _
'        "Kickdown On Span [°],"; _
'        "Kickdown Peak Location [°],"; _
'        "Kickdown Peak - FPT Span [°],"; _
'        "Kickdown Peak Force [N],"; _
'        "Kickdown Force Span [N],"; _
'    'Part Status, Comment, Operator Initials, Temperature, and Series
'    Print #lintFileNum, _
'        "Status,"; _
'        "Comment,"; _
'        "Operator,"; _
'        "Temperature,"; _
'        "Series,"
'Else
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'End If
''Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'Print #lintFileNum, _
'    gstrSerialNumber; ","; _
'    gstrSampleNum; ","; _
'    frmMain.ctrSetupInfo1.TLNum; ","; _
'    gstrDateCode; ","; _
'    DateTime.Now; ","; _
'    App.Major & "." & App.Minor & "." & App.Revision; ","; _
'    gudtMachine.parameterName; ","; _
'    gintPalletNumber; ","; _
'    lstrExp; ","; _
'    lstrExp2; ",";
''Output #1
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).outputAtForceKnee, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(2).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(3).location, 3), "##0.000", 2); ","; _
'    Format(Round(gudtReading(CHAN0).outputAtFPTLoc, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).linDevPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).absoluteLin.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).absoluteLin.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).hysteresis.Value, 2), "##0.00"); ","; _
''Output #2
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN1).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).outputAtForceKnee, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(2).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(3).location, 3), "##0.000", 2); ","; _
'    Format(Round(gudtReading(CHAN1).outputAtFPTLoc, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).linDevPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).absoluteLin.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).absoluteLin.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).hysteresis.Value, 2), "##0.00"); ",";
''Correlation
'Print #lintFileNum, _
'    Format(Round(gudtExtreme(CHAN0).outputCorPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).outputCorPerTol(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).revOutputCor.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).revOutputCor.low.Value, 2), "##0.00"); ","; _
''Force
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).pedalAtRestLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).forceKnee.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).forceKnee.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fullPedalTravel.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fullPedalTravel.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fwdForcePt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fwdForcePt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fwdForcePt(3).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).revForcePt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).revForcePt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).revForcePt(3).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).peakForce, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).mechHystPt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).mechHystPt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).mechHystPt(3).Value, 2), "##0.00"); ","; _
''Kickdown
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).kickdownStart.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownStart.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).outputAtKDStartLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnSpan, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownFPTSpan, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownForceSpan, 2), "##0.00"); ","; _
''Part Status, Comment, Operator Initials, Temperature, and Series
'If gblnScanFailure Then
'    Print #lintFileNum, "REJECT,";
'Else
'    Print #lintFileNum, "PASS,";
'End If
'Print #lintFileNum, _
'    frmMain.ctrSetupInfo1.Comment; ","; _
'    frmMain.ctrSetupInfo1.Operator; ","; _
'    frmMain.ctrSetupInfo1.Temperature; ","; _
'    frmMain.ctrSetupInfo1.Series
''Close the file
'Close #lintFileNum

End Sub

Public Sub Stats700TLLoad()
'
'   PURPOSE:   To input production statistics into the program from
'              a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsLoad_Err
'
''Clear statistics before starting a new lot or resuming an old lot
'Call StatsClear
'
'frmMain.MousePointer = vbHourglass
'
''Get a file number
'lintFileNum = FreeFile
'
''Check to see if file exists if not exit sub
'If gfsoFileSystemObject.FileExists(STATFILEPATH & gstrLotName & STATEXT) Then
'    Open STATFILEPATH & gstrLotName & STATEXT For Input As #lintFileNum
'Else
'    frmMain.MousePointer = vbNormal
'    Exit Sub
'End If
'
''** General Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
''Display to the form
'frmMain.ctrSetupInfo1.Operator = lstrOperator
'frmMain.ctrSetupInfo1.Temperature = lstrTemperature
'frmMain.ctrSetupInfo1.Comment = lstrComment
'frmMain.ctrSetupInfo1.Series = lstrSeries
'frmMain.ctrSetupInfo1.TLNum = lstrTLNum
'frmMain.ctrSetupInfo1.Sample = lstrSample
'
''*** Scan Information ***
''Loop through all channels
'For lintChanNum = CHAN0 To MAXCHANNUM
'    'Index #1 (Idle Output)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
'    'Output at Force Knee
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low, gudtScanStats(lintChanNum).outputAtForceKnee.max, gudtScanStats(lintChanNum).outputAtForceKnee.min, gudtScanStats(lintChanNum).outputAtForceKnee.sigma, gudtScanStats(lintChanNum).outputAtForceKnee.sigma2, gudtScanStats(lintChanNum).outputAtForceKnee.n
'    'Index #2 (Midpoint Output)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
'    'Index #3 (WOT Location)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(3).failCount.high, gudtScanStats(lintChanNum).Index(3).failCount.low, gudtScanStats(lintChanNum).Index(3).max, gudtScanStats(lintChanNum).Index(3).min, gudtScanStats(lintChanNum).Index(3).sigma, gudtScanStats(lintChanNum).Index(3).sigma2, gudtScanStats(lintChanNum).Index(3).n
'    'Output at Full-Pedal-Travel
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).outputAtFPTLoc.failCount.high, gudtScanStats(lintChanNum).outputAtFPTLoc.failCount.low, gudtScanStats(lintChanNum).outputAtFPTLoc.max, gudtScanStats(lintChanNum).outputAtFPTLoc.min, gudtScanStats(lintChanNum).outputAtFPTLoc.sigma, gudtScanStats(lintChanNum).outputAtFPTLoc.sigma2, gudtScanStats(lintChanNum).outputAtFPTLoc.n
'    'Maximum Output
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
'    'Absolute Linearity Deviation
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
'    'Slope Max
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
'    'Slope Min
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n
'Next lintChanNum
''Forward Output Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
''Reverse Output Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
''Pedal-At-Rest Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
''Force Knee Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).forceKneeLoc.failCount.high, gudtScanStats(CHAN0).forceKneeLoc.failCount.low, gudtScanStats(CHAN0).forceKneeLoc.max, gudtScanStats(CHAN0).forceKneeLoc.min, gudtScanStats(CHAN0).forceKneeLoc.sigma, gudtScanStats(CHAN0).forceKneeLoc.sigma2, gudtScanStats(CHAN0).forceKneeLoc.n
''Forward Force at Force Knee Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).forceKneeForce.failCount.high, gudtScanStats(CHAN0).forceKneeForce.failCount.low, gudtScanStats(CHAN0).forceKneeForce.max, gudtScanStats(CHAN0).forceKneeForce.min, gudtScanStats(CHAN0).forceKneeForce.sigma, gudtScanStats(CHAN0).forceKneeForce.sigma2, gudtScanStats(CHAN0).forceKneeForce.n
''Full-Pedal-Travel Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fullPedalTravel.failCount.high, gudtScanStats(CHAN0).fullPedalTravel.failCount.low, gudtScanStats(CHAN0).fullPedalTravel.max, gudtScanStats(CHAN0).fullPedalTravel.min, gudtScanStats(CHAN0).fullPedalTravel.sigma, gudtScanStats(CHAN0).fullPedalTravel.sigma2, gudtScanStats(CHAN0).fullPedalTravel.n
''Forward Force Point 1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(1).failCount.high, gudtScanStats(CHAN0).fwdForcePt(1).failCount.low, gudtScanStats(CHAN0).fwdForcePt(1).max, gudtScanStats(CHAN0).fwdForcePt(1).min, gudtScanStats(CHAN0).fwdForcePt(1).sigma, gudtScanStats(CHAN0).fwdForcePt(1).sigma2, gudtScanStats(CHAN0).fwdForcePt(1).n
''Forward Force Point 2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(2).failCount.high, gudtScanStats(CHAN0).fwdForcePt(2).failCount.low, gudtScanStats(CHAN0).fwdForcePt(2).max, gudtScanStats(CHAN0).fwdForcePt(2).min, gudtScanStats(CHAN0).fwdForcePt(2).sigma, gudtScanStats(CHAN0).fwdForcePt(2).sigma2, gudtScanStats(CHAN0).fwdForcePt(2).n
''Forward Force Point 3
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(3).failCount.high, gudtScanStats(CHAN0).fwdForcePt(3).failCount.low, gudtScanStats(CHAN0).fwdForcePt(3).max, gudtScanStats(CHAN0).fwdForcePt(3).min, gudtScanStats(CHAN0).fwdForcePt(3).sigma, gudtScanStats(CHAN0).fwdForcePt(3).sigma2, gudtScanStats(CHAN0).fwdForcePt(3).n
''Reverse Force Point 1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(1).failCount.high, gudtScanStats(CHAN0).revForcePt(1).failCount.low, gudtScanStats(CHAN0).revForcePt(1).max, gudtScanStats(CHAN0).revForcePt(1).min, gudtScanStats(CHAN0).revForcePt(1).sigma, gudtScanStats(CHAN0).revForcePt(1).sigma2, gudtScanStats(CHAN0).revForcePt(1).n
''Reverse Force Point 2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(2).failCount.high, gudtScanStats(CHAN0).revForcePt(2).failCount.low, gudtScanStats(CHAN0).revForcePt(2).max, gudtScanStats(CHAN0).revForcePt(2).min, gudtScanStats(CHAN0).revForcePt(2).sigma, gudtScanStats(CHAN0).revForcePt(2).sigma2, gudtScanStats(CHAN0).revForcePt(2).n
''Reverse Force Point 3
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).revForcePt(3).failCount.high, gudtScanStats(CHAN0).revForcePt(3).failCount.low, gudtScanStats(CHAN0).revForcePt(3).max, gudtScanStats(CHAN0).revForcePt(3).min, gudtScanStats(CHAN0).revForcePt(3).sigma, gudtScanStats(CHAN0).revForcePt(3).sigma2, gudtScanStats(CHAN0).revForcePt(3).n
''Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
''Mechanical Hysteresis Point 1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(1).failCount.high, gudtScanStats(CHAN0).mechHystPt(1).failCount.low, gudtScanStats(CHAN0).mechHystPt(1).max, gudtScanStats(CHAN0).mechHystPt(1).min, gudtScanStats(CHAN0).mechHystPt(1).sigma, gudtScanStats(CHAN0).mechHystPt(1).sigma2, gudtScanStats(CHAN0).mechHystPt(1).n
''Mechanical Hysteresis Point 2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(2).failCount.high, gudtScanStats(CHAN0).mechHystPt(2).failCount.low, gudtScanStats(CHAN0).mechHystPt(2).max, gudtScanStats(CHAN0).mechHystPt(2).min, gudtScanStats(CHAN0).mechHystPt(2).sigma, gudtScanStats(CHAN0).mechHystPt(2).sigma2, gudtScanStats(CHAN0).mechHystPt(2).n
''Mechanical Hysteresis Point 3
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHystPt(3).failCount.high, gudtScanStats(CHAN0).mechHystPt(3).failCount.low, gudtScanStats(CHAN0).mechHystPt(3).max, gudtScanStats(CHAN0).mechHystPt(3).min, gudtScanStats(CHAN0).mechHystPt(3).sigma, gudtScanStats(CHAN0).mechHystPt(3).sigma2, gudtScanStats(CHAN0).mechHystPt(3).n
''Kickdown Start Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownStartLoc.failCount.high, gudtScanStats(CHAN0).kickdownStartLoc.failCount.low, gudtScanStats(CHAN0).kickdownStartLoc.max, gudtScanStats(CHAN0).kickdownStartLoc.min, gudtScanStats(CHAN0).kickdownStartLoc.sigma, gudtScanStats(CHAN0).kickdownStartLoc.sigma2, gudtScanStats(CHAN0).kickdownStartLoc.n
''Kickdown Start Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownStartForce.failCount.high, gudtScanStats(CHAN0).kickdownStartForce.failCount.low, gudtScanStats(CHAN0).kickdownStartForce.max, gudtScanStats(CHAN0).kickdownStartForce.min, gudtScanStats(CHAN0).kickdownStartForce.sigma, gudtScanStats(CHAN0).kickdownStartForce.sigma2, gudtScanStats(CHAN0).kickdownStartForce.n
''Output #1 at Kickdown Start Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputAtKDStartLoc.failCount.high, gudtScanStats(CHAN0).outputAtKDStartLoc.failCount.low, gudtScanStats(CHAN0).outputAtKDStartLoc.max, gudtScanStats(CHAN0).outputAtKDStartLoc.min, gudtScanStats(CHAN0).outputAtKDStartLoc.sigma, gudtScanStats(CHAN0).outputAtKDStartLoc.sigma2, gudtScanStats(CHAN0).outputAtKDStartLoc.n
''Kickdown On Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownOnLoc.failCount.high, gudtScanStats(CHAN0).kickdownOnLoc.failCount.low, gudtScanStats(CHAN0).kickdownOnLoc.max, gudtScanStats(CHAN0).kickdownOnLoc.min, gudtScanStats(CHAN0).kickdownOnLoc.sigma, gudtScanStats(CHAN0).kickdownOnLoc.sigma2, gudtScanStats(CHAN0).kickdownOnLoc.n
''Kickdown On Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownOnSpan.failCount.high, gudtScanStats(CHAN0).kickdownOnSpan.failCount.low, gudtScanStats(CHAN0).kickdownOnSpan.max, gudtScanStats(CHAN0).kickdownOnSpan.min, gudtScanStats(CHAN0).kickdownOnSpan.sigma, gudtScanStats(CHAN0).kickdownOnSpan.sigma2, gudtScanStats(CHAN0).kickdownOnSpan.n
''Kickdown Peak Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.high, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.low, gudtScanStats(CHAN0).kickdownPeakLoc.max, gudtScanStats(CHAN0).kickdownPeakLoc.min, gudtScanStats(CHAN0).kickdownPeakLoc.sigma, gudtScanStats(CHAN0).kickdownPeakLoc.sigma2, gudtScanStats(CHAN0).kickdownPeakLoc.n
''Kickdown FPT Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownFPTSpan.failCount.high, gudtScanStats(CHAN0).kickdownFPTSpan.failCount.low, gudtScanStats(CHAN0).kickdownFPTSpan.max, gudtScanStats(CHAN0).kickdownFPTSpan.min, gudtScanStats(CHAN0).kickdownFPTSpan.sigma, gudtScanStats(CHAN0).kickdownFPTSpan.sigma2, gudtScanStats(CHAN0).kickdownFPTSpan.n
''Kickdown Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownPeakForce.failCount.high, gudtScanStats(CHAN0).kickdownPeakForce.failCount.low, gudtScanStats(CHAN0).kickdownPeakForce.max, gudtScanStats(CHAN0).kickdownPeakForce.min, gudtScanStats(CHAN0).kickdownPeakForce.sigma, gudtScanStats(CHAN0).kickdownPeakForce.sigma2, gudtScanStats(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownForceSpan.failCount.high, gudtScanStats(CHAN0).kickdownForceSpan.failCount.low, gudtScanStats(CHAN0).kickdownForceSpan.max, gudtScanStats(CHAN0).kickdownForceSpan.min, gudtScanStats(CHAN0).kickdownForceSpan.sigma, gudtScanStats(CHAN0).kickdownForceSpan.sigma2, gudtScanStats(CHAN0).kickdownForceSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Idle) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Idle) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (WOT) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (WOT) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Exposure data
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the file
'Close #lintFileNum
'frmMain.MousePointer = vbNormal
'
'Exit Sub
'StatsLoad_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Lot File!"

End Sub

Public Sub Stats700TLSave()
'
'   PURPOSE:   To write production statistics to a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsSave_Err
'
''Get a file number
'lintFileNum = FreeFile
''Open the stats file
'Open STATFILEPATH & gstrLotName & STATEXT For Output As #lintFileNum
'
''Take data from the form
'lstrOperator = frmMain.ctrSetupInfo1.Operator
'lstrTemperature = frmMain.ctrSetupInfo1.Temperature
'lstrComment = frmMain.ctrSetupInfo1.Comment
'lstrSeries = frmMain.ctrSetupInfo1.Series
'lstrTLNum = frmMain.ctrSetupInfo1.TLNum
'lstrSample = frmMain.ctrSetupInfo1.Sample
'
''*** General Information ***
'Write #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
'
''*** Scan Information ***
''Loop through all channels
'For lintChanNum = CHAN0 To MAXCHANNUM
'    'Index #1 (Idle Output)
'    Write #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
'    'Output at Force Knee
'    Write #lintFileNum, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.high, gudtScanStats(lintChanNum).outputAtForceKnee.failCount.low, gudtScanStats(lintChanNum).outputAtForceKnee.max, gudtScanStats(lintChanNum).outputAtForceKnee.min, gudtScanStats(lintChanNum).outputAtForceKnee.sigma, gudtScanStats(lintChanNum).outputAtForceKnee.sigma2, gudtScanStats(lintChanNum).outputAtForceKnee.n
'    'Index #2 (Midpoint Output)
'    Write #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
'    'Index #3 (WOT Location)
'    Write #lintFileNum, gudtScanStats(lintChanNum).Index(3).failCount.high, gudtScanStats(lintChanNum).Index(3).failCount.low, gudtScanStats(lintChanNum).Index(3).max, gudtScanStats(lintChanNum).Index(3).min, gudtScanStats(lintChanNum).Index(3).sigma, gudtScanStats(lintChanNum).Index(3).sigma2, gudtScanStats(lintChanNum).Index(3).n
'    'Output at Full-Pedal-Travel
'    Write #lintFileNum, gudtScanStats(lintChanNum).outputAtFPTLoc.failCount.high, gudtScanStats(lintChanNum).outputAtFPTLoc.failCount.low, gudtScanStats(lintChanNum).outputAtFPTLoc.max, gudtScanStats(lintChanNum).outputAtFPTLoc.min, gudtScanStats(lintChanNum).outputAtFPTLoc.sigma, gudtScanStats(lintChanNum).outputAtFPTLoc.sigma2, gudtScanStats(lintChanNum).outputAtFPTLoc.n
'    'Maximum Output
'    Write #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
'    'Absolute Linearity Deviation
'    Write #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
'    'Slope Max
'    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
'    'Slope Min
'    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n
'Next lintChanNum
''Forward Output Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
''Reverse Output Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
''Pedal-At-Rest Location
'Write #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
''Force Knee Location
'Write #lintFileNum, gudtScanStats(CHAN0).forceKneeLoc.failCount.high, gudtScanStats(CHAN0).forceKneeLoc.failCount.low, gudtScanStats(CHAN0).forceKneeLoc.max, gudtScanStats(CHAN0).forceKneeLoc.min, gudtScanStats(CHAN0).forceKneeLoc.sigma, gudtScanStats(CHAN0).forceKneeLoc.sigma2, gudtScanStats(CHAN0).forceKneeLoc.n
''Forward Force at Force Knee Location
'Write #lintFileNum, gudtScanStats(CHAN0).forceKneeForce.failCount.high, gudtScanStats(CHAN0).forceKneeForce.failCount.low, gudtScanStats(CHAN0).forceKneeForce.max, gudtScanStats(CHAN0).forceKneeForce.min, gudtScanStats(CHAN0).forceKneeForce.sigma, gudtScanStats(CHAN0).forceKneeForce.sigma2, gudtScanStats(CHAN0).forceKneeForce.n
''Full-Pedal-Travel Location
'Write #lintFileNum, gudtScanStats(CHAN0).fullPedalTravel.failCount.high, gudtScanStats(CHAN0).fullPedalTravel.failCount.low, gudtScanStats(CHAN0).fullPedalTravel.max, gudtScanStats(CHAN0).fullPedalTravel.min, gudtScanStats(CHAN0).fullPedalTravel.sigma, gudtScanStats(CHAN0).fullPedalTravel.sigma2, gudtScanStats(CHAN0).fullPedalTravel.n
''Forward Force Point 1
'Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(1).failCount.high, gudtScanStats(CHAN0).fwdForcePt(1).failCount.low, gudtScanStats(CHAN0).fwdForcePt(1).max, gudtScanStats(CHAN0).fwdForcePt(1).min, gudtScanStats(CHAN0).fwdForcePt(1).sigma, gudtScanStats(CHAN0).fwdForcePt(1).sigma2, gudtScanStats(CHAN0).fwdForcePt(1).n
''Forward Force Point 2
'Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(2).failCount.high, gudtScanStats(CHAN0).fwdForcePt(2).failCount.low, gudtScanStats(CHAN0).fwdForcePt(2).max, gudtScanStats(CHAN0).fwdForcePt(2).min, gudtScanStats(CHAN0).fwdForcePt(2).sigma, gudtScanStats(CHAN0).fwdForcePt(2).sigma2, gudtScanStats(CHAN0).fwdForcePt(2).n
''Forward Force Point 3
'Write #lintFileNum, gudtScanStats(CHAN0).fwdForcePt(3).failCount.high, gudtScanStats(CHAN0).fwdForcePt(3).failCount.low, gudtScanStats(CHAN0).fwdForcePt(3).max, gudtScanStats(CHAN0).fwdForcePt(3).min, gudtScanStats(CHAN0).fwdForcePt(3).sigma, gudtScanStats(CHAN0).fwdForcePt(3).sigma2, gudtScanStats(CHAN0).fwdForcePt(3).n
''Reverse Force Point 1
'Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(1).failCount.high, gudtScanStats(CHAN0).revForcePt(1).failCount.low, gudtScanStats(CHAN0).revForcePt(1).max, gudtScanStats(CHAN0).revForcePt(1).min, gudtScanStats(CHAN0).revForcePt(1).sigma, gudtScanStats(CHAN0).revForcePt(1).sigma2, gudtScanStats(CHAN0).revForcePt(1).n
''Reverse Force Point 2
'Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(2).failCount.high, gudtScanStats(CHAN0).revForcePt(2).failCount.low, gudtScanStats(CHAN0).revForcePt(2).max, gudtScanStats(CHAN0).revForcePt(2).min, gudtScanStats(CHAN0).revForcePt(2).sigma, gudtScanStats(CHAN0).revForcePt(2).sigma2, gudtScanStats(CHAN0).revForcePt(2).n
''Reverse Force Point 3
'Write #lintFileNum, gudtScanStats(CHAN0).revForcePt(3).failCount.high, gudtScanStats(CHAN0).revForcePt(3).failCount.low, gudtScanStats(CHAN0).revForcePt(3).max, gudtScanStats(CHAN0).revForcePt(3).min, gudtScanStats(CHAN0).revForcePt(3).sigma, gudtScanStats(CHAN0).revForcePt(3).sigma2, gudtScanStats(CHAN0).revForcePt(3).n
''Peak Force
'Write #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
''Mechanical Hysteresis Point 1
'Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(1).failCount.high, gudtScanStats(CHAN0).mechHystPt(1).failCount.low, gudtScanStats(CHAN0).mechHystPt(1).max, gudtScanStats(CHAN0).mechHystPt(1).min, gudtScanStats(CHAN0).mechHystPt(1).sigma, gudtScanStats(CHAN0).mechHystPt(1).sigma2, gudtScanStats(CHAN0).mechHystPt(1).n
''Mechanical Hysteresis Point 2
'Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(2).failCount.high, gudtScanStats(CHAN0).mechHystPt(2).failCount.low, gudtScanStats(CHAN0).mechHystPt(2).max, gudtScanStats(CHAN0).mechHystPt(2).min, gudtScanStats(CHAN0).mechHystPt(2).sigma, gudtScanStats(CHAN0).mechHystPt(2).sigma2, gudtScanStats(CHAN0).mechHystPt(2).n
''Mechanical Hysteresis Point 3
'Write #lintFileNum, gudtScanStats(CHAN0).mechHystPt(3).failCount.high, gudtScanStats(CHAN0).mechHystPt(3).failCount.low, gudtScanStats(CHAN0).mechHystPt(3).max, gudtScanStats(CHAN0).mechHystPt(3).min, gudtScanStats(CHAN0).mechHystPt(3).sigma, gudtScanStats(CHAN0).mechHystPt(3).sigma2, gudtScanStats(CHAN0).mechHystPt(3).n
''Kickdown Start Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownStartLoc.failCount.high, gudtScanStats(CHAN0).kickdownStartLoc.failCount.low, gudtScanStats(CHAN0).kickdownStartLoc.max, gudtScanStats(CHAN0).kickdownStartLoc.min, gudtScanStats(CHAN0).kickdownStartLoc.sigma, gudtScanStats(CHAN0).kickdownStartLoc.sigma2, gudtScanStats(CHAN0).kickdownStartLoc.n
''Kickdown Start Force
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownStartForce.failCount.high, gudtScanStats(CHAN0).kickdownStartForce.failCount.low, gudtScanStats(CHAN0).kickdownStartForce.max, gudtScanStats(CHAN0).kickdownStartForce.min, gudtScanStats(CHAN0).kickdownStartForce.sigma, gudtScanStats(CHAN0).kickdownStartForce.sigma2, gudtScanStats(CHAN0).kickdownStartForce.n
''Output #1 at Kickdown Start Location
'Write #lintFileNum, gudtScanStats(CHAN0).outputAtKDStartLoc.failCount.high, gudtScanStats(CHAN0).outputAtKDStartLoc.failCount.low, gudtScanStats(CHAN0).outputAtKDStartLoc.max, gudtScanStats(CHAN0).outputAtKDStartLoc.min, gudtScanStats(CHAN0).outputAtKDStartLoc.sigma, gudtScanStats(CHAN0).outputAtKDStartLoc.sigma2, gudtScanStats(CHAN0).outputAtKDStartLoc.n
''Kickdown On Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownOnLoc.failCount.high, gudtScanStats(CHAN0).kickdownOnLoc.failCount.low, gudtScanStats(CHAN0).kickdownOnLoc.max, gudtScanStats(CHAN0).kickdownOnLoc.min, gudtScanStats(CHAN0).kickdownOnLoc.sigma, gudtScanStats(CHAN0).kickdownOnLoc.sigma2, gudtScanStats(CHAN0).kickdownOnLoc.n
''Kickdown On Span
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownOnSpan.failCount.high, gudtScanStats(CHAN0).kickdownOnSpan.failCount.low, gudtScanStats(CHAN0).kickdownOnSpan.max, gudtScanStats(CHAN0).kickdownOnSpan.min, gudtScanStats(CHAN0).kickdownOnSpan.sigma, gudtScanStats(CHAN0).kickdownOnSpan.sigma2, gudtScanStats(CHAN0).kickdownOnSpan.n
''Kickdown Peak Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.high, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.low, gudtScanStats(CHAN0).kickdownPeakLoc.max, gudtScanStats(CHAN0).kickdownPeakLoc.min, gudtScanStats(CHAN0).kickdownPeakLoc.sigma, gudtScanStats(CHAN0).kickdownPeakLoc.sigma2, gudtScanStats(CHAN0).kickdownPeakLoc.n
''Kickdown FPT Span
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownFPTSpan.failCount.high, gudtScanStats(CHAN0).kickdownFPTSpan.failCount.low, gudtScanStats(CHAN0).kickdownFPTSpan.max, gudtScanStats(CHAN0).kickdownFPTSpan.min, gudtScanStats(CHAN0).kickdownFPTSpan.sigma, gudtScanStats(CHAN0).kickdownFPTSpan.sigma2, gudtScanStats(CHAN0).kickdownFPTSpan.n
''Kickdown Peak Force
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownPeakForce.failCount.high, gudtScanStats(CHAN0).kickdownPeakForce.failCount.low, gudtScanStats(CHAN0).kickdownPeakForce.max, gudtScanStats(CHAN0).kickdownPeakForce.min, gudtScanStats(CHAN0).kickdownPeakForce.sigma, gudtScanStats(CHAN0).kickdownPeakForce.sigma2, gudtScanStats(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownForceSpan.failCount.high, gudtScanStats(CHAN0).kickdownForceSpan.failCount.low, gudtScanStats(CHAN0).kickdownForceSpan.max, gudtScanStats(CHAN0).kickdownForceSpan.min, gudtScanStats(CHAN0).kickdownForceSpan.sigma, gudtScanStats(CHAN0).kickdownForceSpan.sigma2, gudtScanStats(CHAN0).kickdownForceSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Idle) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Idle) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (WOT) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (WOT) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'Write #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'Write #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Exposure data
'Write #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'Write #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'Write #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'Write #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'Write #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'Write #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'Write #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'Write #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'Write #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'Write #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'Write #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the stats file
'Close #lintFileNum
'Call frmMain.RefreshLotFileList         'Add new files to lot file list
'
'Exit Sub
'StatsSave_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Saving Data to Lot File!"

End Sub

Public Sub Save702TLScanResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lstrFileName As String
'Dim lstrExp As String
'Dim lstrExp2 As String
'
'lstrExp = ""
'lstrExp2 = ""
'If gblnDustExp Then
'    lstrExp = "Dust "
'    lstrExp2 = lstrExp2 & "Dust(" & gudtExposure.Dust.TypeofDust & " " & gudtExposure.Dust.AmountofDust & " " & gudtExposure.Dust.StirTime & " " & gudtExposure.Dust.SettleTime & " " & gudtExposure.Dust.Frequency & " " & gudtExposure.Dust.NumberofCycles & " " & gudtExposure.Dust.Condition & ") "
'End If
'If gblnVibrationExp Then
'    lstrExp = lstrExp & "Vibration "
'    lstrExp2 = lstrExp2 & "Vibration(" & gudtExposure.Vibration.Profile & " " & gudtExposure.Vibration.Temperature & " " & gudtExposure.Vibration.Duration & " " & gudtExposure.Vibration.Planes & " " & gudtExposure.Vibration.NumberofCycles & " " & gudtExposure.Vibration.Frequency & ") "
'End If
'If gblnDitherExp Then lstrExp = lstrExp & "Dither "
'If gblnThermalShockExp Then
'    lstrExp = lstrExp & "Thermal Shock "
'    lstrExp2 = lstrExp2 & "Thermal Shock(" & CStr(gudtExposure.ThermalShock.LowTemp) & " " & gudtExposure.ThermalShock.LowTempTime & " " & CStr(gudtExposure.ThermalShock.HighTemp) & " " & gudtExposure.ThermalShock.HighTempTime & " " & gudtExposure.ThermalShock.NumberofCycles & " " & gudtExposure.ThermalShock.Condition & ") "
'End If
'If gblnSaltSprayExp Then
'    lstrExp = lstrExp & "Salt Spray "
'    lstrExp2 = lstrExp2 & "Salt Spray(" & CStr(gudtExposure.SaltSpray.Duration) & " " & gudtExposure.SaltSpray.Condition & ") "
'End If
'If gblnInitialExp Then lstrExp = lstrExp & "Initial "
'If gblnExposure Then
'    lstrExp = lstrExp & "Exposure "
'    lstrExp2 = lstrExp2 & "Exposure (" & gudtExposure.Exposure.Condition & ") "
'End If
'If gblnOperStrnExp Then lstrExp = lstrExp & "Operational Strength "
'If gblnLateralStrnExp Then lstrExp = lstrExp & "Lateral Strength "
'If gblnOpStrnStopExp Then lstrExp = lstrExp & "Operational Strength with Stopper "
'If gblnImpactStrnExp Then lstrExp = lstrExp & "Impact Strength "
'If gblnOperEndurExp Then
'    lstrExp = lstrExp & "Operational Endurance "
'    lstrExp2 = lstrExp2 & "Operational Endurance(" & gudtExposure.OperationalEndurance.Temperature & " " & gudtExposure.OperationalEndurance.NewNumberofCycles & " " & gudtExposure.OperationalEndurance.TotalNumberofCycles & " " & gudtExposure.OperationalEndurance.Condition & ") "
'End If
'If gblnSnapbackExp Then
'    lstrExp = lstrExp & "Snapback "
'    lstrExp2 = lstrExp2 & "Snapback(" & gudtExposure.SnapBack.Temperature & " " & gudtExposure.SnapBack.NumberofCycles & " " & gudtExposure.SnapBack.Condition & ") "
'End If
'If gblnHighTempExp Then
'    lstrExp = lstrExp & "High Temp Soak "
'    lstrExp2 = lstrExp2 & "High Temp Soak(" & CStr(gudtExposure.HighTempSoak.Temperature) & " " & gudtExposure.HighTempSoak.Duration & " " & gudtExposure.HighTempSoak.Condition & ") "
'End If
'If gblnHighTempHighHumidExp Then
'    lstrExp = lstrExp & "High Temp - High Humidity Soak "
'    lstrExp2 = lstrExp2 & "High Temp - High Humidity Soak(" & CStr(gudtExposure.HTempHHumiditySoak.Temperature) & " " & CStr(gudtExposure.HTempHHumiditySoak.RelativeHumidity) & " " & gudtExposure.HTempHHumiditySoak.Duration & " " & gudtExposure.HTempHHumiditySoak.Condition & ") "
'End If
'If gblnLowTempExp Then
'    lstrExp = lstrExp & "Low Temp Soak "
'    lstrExp2 = lstrExp2 & "Low Temp Soak(" & CStr(gudtExposure.LowTempSoak.Temperature) & " " & gudtExposure.LowTempSoak.Duration & " " & gudtExposure.LowTempSoak.Condition & ") "
'End If
'If gblnWaterSprayExp Then
'    lstrExp = lstrExp & "Water Spray "
'    lstrExp2 = lstrExp2 & "Water Spray(" & gudtExposure.WaterSpray.Duration & " " & gudtExposure.WaterSpray.Condition & ") "
'End If
'If gblnChemResExp Then
'    lstrExp = lstrExp & "Chemical Resistance "
'    lstrExp2 = lstrExp2 & "Chemical Resistance(" & CStr(gudtExposure.ChemResistance.Temperature) & " " & gudtExposure.ChemResistance.Duration & " " & gudtExposure.ChemResistance.Substance & ") "
'End If
'If gblnCondenExp Then lstrExp = lstrExp & "Condensation "
'If gblnESDElecExp Then lstrExp = lstrExp & "ElectroStatic Discharge "
'If gblnEMWaveResElecExp Then lstrExp = lstrExp & "Electromagnetic Wave Resistance "
'If gblnBilkCInjElecExp Then lstrExp = lstrExp & "Bilk Current Injection "
'If gblnIgnitionNoiseElecExp Then lstrExp = lstrExp & "Ignition Noise "
'If gblnNarRadEMEElecExp Then lstrExp = lstrExp & "Narrowband Radiated Electromagnetic Energy "
'
''Make the results file name
'lstrFileName = gstrLotName + " Scan Results" & DATAEXT
''Get a file
'lintFileNum = FreeFile
'
''If file does not exist then add a header
'If Not gfsoFileSystemObject.FileExists(PARTSCANDATAPATH + lstrFileName) Then
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'    'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'    Print #lintFileNum, _
'        "Part Number,"; _
'        "Sample Number,"; _
'        "TestLog #,"; _
'        "Date Code,"; _
'        "Date/Time,"; _
'        "S/W Revision,"; _
'        "Parameter File Name,"; _
'        "Pallet Number,"; _
'        "Exposures,"; _
'        "Exposure Data,";
'    'Pedal Location, Rising Point, Output #1, & Output #2
'    Print #lintFileNum, _
'        "Pedal at Rest Location [°],"; _
'        "Rising Point Location [°],"; _
'        "Full-Close Value Output #1 [%],"; _
'        "Full-Open Location Output #1 [°],"; _
'        "Max Output Output #1 [%],"; _
'        "Max Linearity Deviation % of Tol Output #1 [% Tol],"; _
'        "Max Linearity Deviation Output #1 [%],"; _
'        "Min Linearity Deviation Output #1 [%],"; _
'        "Max Slope Deviation Output #1 [% Ideal Slope],"; _
'        "Min Slope Deviation Output #1 [% Ideal Slope],"; _
'        "Peak Hysteresis Output #1 [%],"; _
'        "Full-Close Value Output #2 [%],"; _
'        "Full-Open Location Output #2 [°],"; _
'        "Max Output Output #2 [%],"; _
'        "Max Linearity Deviation % of Tol Output #2 [% Tol],"; _
'        "Max Linearity Deviation Output #2 [%],"; _
'        "Min Linearity Deviation Output #2 [%],"; _
'        "Max Slope Deviation Output #2 [% Ideal Slope],"; _
'        "Min Slope Deviation Output #2 [% Ideal Slope],"; _
'        "Peak Hysteresis Output #2 [%],";
'    'Correlation, Pedal Effort, & Kickdown
'    Print #lintFileNum, _
'        "Max Output Corr % of Tolerance [% Tol],"; _
'        "Max Output Corr Output #1/Output #2 [%],"; _
'        "Min Output Corr Output #1/Output #2 [%],"; _
'        "Pedal Effort Pressing Point 1 [N],"; _
'        "Pedal Effort Releasing Point 1 [N],"; _
'        "Pedal Effort Pressing Point 2 [N],"; _
'        "Pedal Effort Releasing Point 2 [N],"; _
'        "Peak Force [N],"; _
'        "Mechanical Hysteresis Pt 1 [N],"; _
'        "Kickdown Start Location [°],"; _
'        "Kickdown Peak Location [°],"; _
'        "Kickdown Peak Force [N],"; _
'        "Kickdown Force Span [N],"; _
'        "Kickdown On Location [°],"; _
'        "Kickdown On Span [°],";
'    'Part Status, Comment, Operator Initials, Temperature, and Series
'    Print #lintFileNum, _
'        "Status,"; _
'        "Comment,"; _
'        "Operator,"; _
'        "Temperature,"; _
'        "Series,"
'Else
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'End If
''Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'Print #lintFileNum, _
'    gstrSerialNumber; ","; _
'    gstrSampleNum; ","; _
'    frmMain.ctrSetupInfo1.TLNum; ","; _
'    gstrDateCode; ","; _
'    DateTime.Now; ","; _
'    App.Major & "." & App.Minor & "." & App.Revision; ","; _
'    gudtMachine.parameterName; ","; _
'    gintPalletNumber; ","; _
'    lstrExp; ","; _
'    lstrExp2; ",";
''Pedal Location, Rising Point, Output #1, & Output #2
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).pedalAtRestLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).risingPoint.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(2).location, 3), "##0.000", 2); ","; _
'    Format(Round(gudtReading(CHAN0).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).linDevPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).linDev.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).linDev.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).hysteresis.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(2).location, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).linDevPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).linDev.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).linDev.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).hysteresis.Value, 2), "##0.00"); ",";
''Correlation, Pedal Effort, & Kickdown
'Print #lintFileNum, _
'    Format(Round(gudtExtreme(CHAN0).outputCorPerTol(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fwdForcePt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).revForcePt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).fwdForcePt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).revForcePt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).peakForce, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).mechHysteresis(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownStart.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownForceSpan, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnSpan, 2), "##0.00"); ",";
''Part Status, Comment, Operator Initials, Temperature, and Series
'If gblnScanFailure Then
'    Print #lintFileNum, "REJECT,";
'Else
'    Print #lintFileNum, "PASS,";
'End If
'Print #lintFileNum, _
'    frmMain.ctrSetupInfo1.Comment; ","; _
'    frmMain.ctrSetupInfo1.Operator; ","; _
'    frmMain.ctrSetupInfo1.Temperature; ","; _
'    frmMain.ctrSetupInfo1.Series
''Close the file
'Close #lintFileNum

End Sub

Public Sub Stats702TLLoad()
'
'   PURPOSE:   To input production statistics into the program from
'              a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsLoad_Err
'
''Clear statistics before starting a new lot or resuming an old lot
'Call StatsClear
'
'frmMain.MousePointer = vbHourglass
'
''Get a file number
'lintFileNum = FreeFile
'
''Check to see if file exists if not exit sub
'If gfsoFileSystemObject.FileExists(STATFILEPATH & gstrLotName & STATEXT) Then
'    Open STATFILEPATH & gstrLotName & STATEXT For Input As #lintFileNum
'Else
'    frmMain.MousePointer = vbNormal
'    Exit Sub
'End If
'
''** General Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
''Display to the form
'frmMain.ctrSetupInfo1.Operator = lstrOperator
'frmMain.ctrSetupInfo1.Temperature = lstrTemperature
'frmMain.ctrSetupInfo1.Comment = lstrComment
'frmMain.ctrSetupInfo1.Series = lstrSeries
'frmMain.ctrSetupInfo1.TLNum = lstrTLNum
'frmMain.ctrSetupInfo1.Sample = lstrSample
'
''*** Scan Information ***
''Pedal at Rest Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).pedalAtRestLoc.failCount.high, gudtScanSums(CHAN0).pedalAtRestLoc.failCount.low, gudtScanSums(CHAN0).pedalAtRestLoc.max, gudtScanSums(CHAN0).pedalAtRestLoc.min, gudtScanSums(CHAN0).pedalAtRestLoc.sigma, gudtScanSums(CHAN0).pedalAtRestLoc.sigma2, gudtScanSums(CHAN0).pedalAtRestLoc.n
''Rising Point
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).risingPoint.failCount.high, gudtScanSums(CHAN0).risingPoint.failCount.low, gudtScanSums(CHAN0).risingPoint.max, gudtScanSums(CHAN0).risingPoint.min, gudtScanSums(CHAN0).risingPoint.sigma, gudtScanSums(CHAN0).risingPoint.sigma2, gudtScanSums(CHAN0).risingPoint.n
''Loop through all channels
'For lintChanNum = CHAN0 To MAXCHANNUM
'    'Index #1 (Full-Close)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(lintChanNum).Index(1).failCount.high, gudtScanSums(lintChanNum).Index(1).failCount.low, gudtScanSums(lintChanNum).Index(1).max, gudtScanSums(lintChanNum).Index(1).min, gudtScanSums(lintChanNum).Index(1).sigma, gudtScanSums(lintChanNum).Index(1).sigma2, gudtScanSums(lintChanNum).Index(1).n
'    'Index #2 (Full-Open)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(lintChanNum).Index(2).failCount.high, gudtScanSums(lintChanNum).Index(2).failCount.low, gudtScanSums(lintChanNum).Index(2).max, gudtScanSums(lintChanNum).Index(2).min, gudtScanSums(lintChanNum).Index(2).sigma, gudtScanSums(lintChanNum).Index(2).sigma2, gudtScanSums(lintChanNum).Index(2).n
'    'Maximum Output
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(lintChanNum).maxOutput.failCount.high, gudtScanSums(lintChanNum).maxOutput.failCount.low, gudtScanSums(lintChanNum).maxOutput.max, gudtScanSums(lintChanNum).maxOutput.min, gudtScanSums(lintChanNum).maxOutput.sigma, gudtScanSums(lintChanNum).maxOutput.sigma2, gudtScanSums(lintChanNum).maxOutput.n
'    'Linearity
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(lintChanNum).linDevPerTol(1).failCount.high, gudtScanSums(lintChanNum).linDevPerTol(1).failCount.low, gudtScanSums(lintChanNum).linDevPerTol(1).max, gudtScanSums(lintChanNum).linDevPerTol(1).min, gudtScanSums(lintChanNum).linDevPerTol(1).sigma, gudtScanSums(lintChanNum).linDevPerTol(1).sigma2, gudtScanSums(lintChanNum).linDevPerTol(1).n
'    'Slope Deviation
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(lintChanNum).slopeMax.failCount.high, gudtScanSums(lintChanNum).slopeMax.failCount.low, gudtScanSums(lintChanNum).slopeMax.max, gudtScanSums(lintChanNum).slopeMax.min, gudtScanSums(lintChanNum).slopeMax.sigma, gudtScanSums(lintChanNum).slopeMax.sigma2, gudtScanSums(lintChanNum).slopeMax.n, gudtScanSums(lintChanNum).slopeMin.max, gudtScanSums(lintChanNum).slopeMin.min, gudtScanSums(lintChanNum).slopeMin.sigma, gudtScanSums(lintChanNum).slopeMin.sigma2, gudtScanSums(lintChanNum).slopeMin.n
'Next lintChanNum
''Output Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).outputCorPerTol(1).failCount.high, gudtScanSums(CHAN0).outputCorPerTol(1).failCount.low, gudtScanSums(CHAN0).outputCorPerTol(1).max, gudtScanSums(CHAN0).outputCorPerTol(1).min, gudtScanSums(CHAN0).outputCorPerTol(1).sigma, gudtScanSums(CHAN0).outputCorPerTol(1).sigma2, gudtScanSums(CHAN0).outputCorPerTol(1).n
''Pedal Effort Pressing Pt1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).fwdForcePt(1).failCount.high, gudtScanSums(CHAN0).fwdForcePt(1).failCount.low, gudtScanSums(CHAN0).fwdForcePt(1).max, gudtScanSums(CHAN0).fwdForcePt(1).min, gudtScanSums(CHAN0).fwdForcePt(1).sigma, gudtScanSums(CHAN0).fwdForcePt(1).sigma2, gudtScanSums(CHAN0).fwdForcePt(1).n
''Pedal Effort Releasing Pt1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).revForcePt(1).failCount.high, gudtScanSums(CHAN0).revForcePt(1).failCount.low, gudtScanSums(CHAN0).revForcePt(1).max, gudtScanSums(CHAN0).revForcePt(1).min, gudtScanSums(CHAN0).revForcePt(1).sigma, gudtScanSums(CHAN0).revForcePt(1).sigma2, gudtScanSums(CHAN0).revForcePt(1).n
''Pedal Effort Pressing Pt2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).fwdForcePt(2).failCount.high, gudtScanSums(CHAN0).fwdForcePt(2).failCount.low, gudtScanSums(CHAN0).fwdForcePt(2).max, gudtScanSums(CHAN0).fwdForcePt(2).min, gudtScanSums(CHAN0).fwdForcePt(2).sigma, gudtScanSums(CHAN0).fwdForcePt(2).sigma2, gudtScanSums(CHAN0).fwdForcePt(2).n
''Pedal Effort Releasing Pt2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).revForcePt(2).failCount.high, gudtScanSums(CHAN0).revForcePt(2).failCount.low, gudtScanSums(CHAN0).revForcePt(2).max, gudtScanSums(CHAN0).revForcePt(2).min, gudtScanSums(CHAN0).revForcePt(2).sigma, gudtScanSums(CHAN0).revForcePt(2).sigma2, gudtScanSums(CHAN0).revForcePt(2).n
''Mechanical Hysteresis Pt1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).mechHysteresis(1).failCount.high, gudtScanSums(CHAN0).mechHysteresis(1).failCount.low, gudtScanSums(CHAN0).mechHysteresis(1).max, gudtScanSums(CHAN0).mechHysteresis(1).min, gudtScanSums(CHAN0).mechHysteresis(1).sigma, gudtScanSums(CHAN0).mechHysteresis(1).sigma2, gudtScanSums(CHAN0).mechHysteresis(1).n
''Kickdown Start Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownStartLoc.failCount.high, gudtScanSums(CHAN0).kickdownStartLoc.failCount.low, gudtScanSums(CHAN0).kickdownStartLoc.max, gudtScanSums(CHAN0).kickdownStartLoc.min, gudtScanSums(CHAN0).kickdownStartLoc.sigma, gudtScanSums(CHAN0).kickdownStartLoc.sigma2, gudtScanSums(CHAN0).kickdownStartLoc.n
''Kickdown Peak Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownPeakLoc.failCount.high, gudtScanSums(CHAN0).kickdownPeakLoc.failCount.low, gudtScanSums(CHAN0).kickdownPeakLoc.max, gudtScanSums(CHAN0).kickdownPeakLoc.min, gudtScanSums(CHAN0).kickdownPeakLoc.sigma, gudtScanSums(CHAN0).kickdownPeakLoc.sigma2, gudtScanSums(CHAN0).kickdownPeakLoc.n
''Kickdown Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownPeakForce.failCount.high, gudtScanSums(CHAN0).kickdownPeakForce.failCount.low, gudtScanSums(CHAN0).kickdownPeakForce.max, gudtScanSums(CHAN0).kickdownPeakForce.min, gudtScanSums(CHAN0).kickdownPeakForce.sigma, gudtScanSums(CHAN0).kickdownPeakForce.sigma2, gudtScanSums(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownForceSpan.failCount.high, gudtScanSums(CHAN0).kickdownForceSpan.failCount.low, gudtScanSums(CHAN0).kickdownForceSpan.max, gudtScanSums(CHAN0).kickdownForceSpan.min, gudtScanSums(CHAN0).kickdownForceSpan.sigma, gudtScanSums(CHAN0).kickdownForceSpan.sigma2, gudtScanSums(CHAN0).kickdownForceSpan.n
''Kickdown On Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownOnLoc.failCount.high, gudtScanSums(CHAN0).kickdownOnLoc.failCount.low, gudtScanSums(CHAN0).kickdownOnLoc.max, gudtScanSums(CHAN0).kickdownOnLoc.min, gudtScanSums(CHAN0).kickdownOnLoc.sigma, gudtScanSums(CHAN0).kickdownOnLoc.sigma2, gudtScanSums(CHAN0).kickdownOnLoc.n
''Kickdown On Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).kickdownOnSpan.failCount.high, gudtScanSums(CHAN0).kickdownOnSpan.failCount.low, gudtScanSums(CHAN0).kickdownOnSpan.max, gudtScanSums(CHAN0).kickdownOnSpan.min, gudtScanSums(CHAN0).kickdownOnSpan.sigma, gudtScanSums(CHAN0).kickdownOnSpan.sigma2, gudtScanSums(CHAN0).kickdownOnSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Full-Close) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Full-Close) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (Full-Open) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (Full-Open) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSums(CHAN0).peakForce.failCount.high, gudtScanSums(CHAN0).peakForce.failCount.low, gudtScanSums(CHAN0).peakForce.max, gudtScanSums(CHAN0).peakForce.min, gudtScanSums(CHAN0).peakForce.sigma, gudtScanSums(CHAN0).peakForce.sigma2, gudtScanSums(CHAN0).peakForce.n
'
''Exposure data
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the file
'Close #lintFileNum
'frmMain.MousePointer = vbNormal
'
'Exit Sub
'StatsLoad_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Lot File!"

End Sub

Public Sub Stats702TLSave()
'
'   PURPOSE:   To write production statistics to a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsSave_Err
'
''Get a file number
'lintFileNum = FreeFile
''Open the stats file
'Open STATFILEPATH & gstrLotName & STATEXT For Output As #lintFileNum
'
''Take data from the form
'lstrOperator = frmMain.ctrSetupInfo1.Operator
'lstrTemperature = frmMain.ctrSetupInfo1.Temperature
'lstrComment = frmMain.ctrSetupInfo1.Comment
'lstrSeries = frmMain.ctrSetupInfo1.Series
'lstrTLNum = frmMain.ctrSetupInfo1.TLNum
'lstrSample = frmMain.ctrSetupInfo1.Sample
'
''*** General Information ***
'Write #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
'
''*** Scan Information ***
''Pedal at Rest Location
'Write #lintFileNum, gudtScanSums(CHAN0).pedalAtRestLoc.failCount.high, gudtScanSums(CHAN0).pedalAtRestLoc.failCount.low, gudtScanSums(CHAN0).pedalAtRestLoc.max, gudtScanSums(CHAN0).pedalAtRestLoc.min, gudtScanSums(CHAN0).pedalAtRestLoc.sigma, gudtScanSums(CHAN0).pedalAtRestLoc.sigma2, gudtScanSums(CHAN0).pedalAtRestLoc.n
''Rising Point
'Write #lintFileNum, gudtScanSums(CHAN0).risingPoint.failCount.high, gudtScanSums(CHAN0).risingPoint.failCount.low, gudtScanSums(CHAN0).risingPoint.max, gudtScanSums(CHAN0).risingPoint.min, gudtScanSums(CHAN0).risingPoint.sigma, gudtScanSums(CHAN0).risingPoint.sigma2, gudtScanSums(CHAN0).risingPoint.n
''Loop through all channels
'For lintChanNum = 0 To MAXCHANNUM
'    'Index 1 (Full-Close)
'    Write #lintFileNum, gudtScanSums(lintChanNum).Index(1).failCount.high, gudtScanSums(lintChanNum).Index(1).failCount.low, gudtScanSums(lintChanNum).Index(1).max, gudtScanSums(lintChanNum).Index(1).min, gudtScanSums(lintChanNum).Index(1).sigma, gudtScanSums(lintChanNum).Index(1).sigma2, gudtScanSums(lintChanNum).Index(1).n
'    'Index 2 (Full-Open)
'    Write #lintFileNum, gudtScanSums(lintChanNum).Index(2).failCount.high, gudtScanSums(lintChanNum).Index(2).failCount.low, gudtScanSums(lintChanNum).Index(2).max, gudtScanSums(lintChanNum).Index(2).min, gudtScanSums(lintChanNum).Index(2).sigma, gudtScanSums(lintChanNum).Index(2).sigma2, gudtScanSums(lintChanNum).Index(2).n
'    'Maximum Output
'    Write #lintFileNum, gudtScanSums(lintChanNum).maxOutput.failCount.high, gudtScanSums(lintChanNum).maxOutput.failCount.low, gudtScanSums(lintChanNum).maxOutput.max, gudtScanSums(lintChanNum).maxOutput.min, gudtScanSums(lintChanNum).maxOutput.sigma, gudtScanSums(lintChanNum).maxOutput.sigma2, gudtScanSums(lintChanNum).maxOutput.n
'    'Linearity
'    Write #lintFileNum, gudtScanSums(lintChanNum).linDevPerTol(1).failCount.high, gudtScanSums(lintChanNum).linDevPerTol(1).failCount.low, gudtScanSums(lintChanNum).linDevPerTol(1).max, gudtScanSums(lintChanNum).linDevPerTol(1).min, gudtScanSums(lintChanNum).linDevPerTol(1).sigma, gudtScanSums(lintChanNum).linDevPerTol(1).sigma2, gudtScanSums(lintChanNum).linDevPerTol(1).n
'    'Slope Deviation
'    Write #lintFileNum, gudtScanSums(lintChanNum).slopeMax.failCount.high, gudtScanSums(lintChanNum).slopeMax.failCount.low, gudtScanSums(lintChanNum).slopeMax.max, gudtScanSums(lintChanNum).slopeMax.min, gudtScanSums(lintChanNum).slopeMax.sigma, gudtScanSums(lintChanNum).slopeMax.sigma2, gudtScanSums(lintChanNum).slopeMax.n, gudtScanSums(lintChanNum).slopeMin.max, gudtScanSums(lintChanNum).slopeMin.min, gudtScanSums(lintChanNum).slopeMin.sigma, gudtScanSums(lintChanNum).slopeMin.sigma2, gudtScanSums(lintChanNum).slopeMin.n
'Next lintChanNum
''Output Correlation
'Write #lintFileNum, gudtScanSums(CHAN0).outputCorPerTol(1).failCount.high, gudtScanSums(CHAN0).outputCorPerTol(1).failCount.low, gudtScanSums(CHAN0).outputCorPerTol(1).max, gudtScanSums(CHAN0).outputCorPerTol(1).min, gudtScanSums(CHAN0).outputCorPerTol(1).sigma, gudtScanSums(CHAN0).outputCorPerTol(1).sigma2, gudtScanSums(CHAN0).outputCorPerTol(1).n
''Pedal Effort Pressing Pt1
'Write #lintFileNum, gudtScanSums(CHAN0).fwdForcePt(1).failCount.high, gudtScanSums(CHAN0).fwdForcePt(1).failCount.low, gudtScanSums(CHAN0).fwdForcePt(1).max, gudtScanSums(CHAN0).fwdForcePt(1).min, gudtScanSums(CHAN0).fwdForcePt(1).sigma, gudtScanSums(CHAN0).fwdForcePt(1).sigma2, gudtScanSums(CHAN0).fwdForcePt(1).n
''Pedal Effort Releasing Pt1
'Write #lintFileNum, gudtScanSums(CHAN0).revForcePt(1).failCount.high, gudtScanSums(CHAN0).revForcePt(1).failCount.low, gudtScanSums(CHAN0).revForcePt(1).max, gudtScanSums(CHAN0).revForcePt(1).min, gudtScanSums(CHAN0).revForcePt(1).sigma, gudtScanSums(CHAN0).revForcePt(1).sigma2, gudtScanSums(CHAN0).revForcePt(1).n
''Pedal Effort Pressing Pt2
'Write #lintFileNum, gudtScanSums(CHAN0).fwdForcePt(2).failCount.high, gudtScanSums(CHAN0).fwdForcePt(2).failCount.low, gudtScanSums(CHAN0).fwdForcePt(2).max, gudtScanSums(CHAN0).fwdForcePt(2).min, gudtScanSums(CHAN0).fwdForcePt(2).sigma, gudtScanSums(CHAN0).fwdForcePt(2).sigma2, gudtScanSums(CHAN0).fwdForcePt(2).n
''Pedal Effort Releasing Pt2
'Write #lintFileNum, gudtScanSums(CHAN0).revForcePt(2).failCount.high, gudtScanSums(CHAN0).revForcePt(2).failCount.low, gudtScanSums(CHAN0).revForcePt(2).max, gudtScanSums(CHAN0).revForcePt(2).min, gudtScanSums(CHAN0).revForcePt(2).sigma, gudtScanSums(CHAN0).revForcePt(2).sigma2, gudtScanSums(CHAN0).revForcePt(2).n
''Mechanical Hysteresis Pt1
'Write #lintFileNum, gudtScanSums(CHAN0).mechHysteresis(1).failCount.high, gudtScanSums(CHAN0).mechHysteresis(1).failCount.low, gudtScanSums(CHAN0).mechHysteresis(1).max, gudtScanSums(CHAN0).mechHysteresis(1).min, gudtScanSums(CHAN0).mechHysteresis(1).sigma, gudtScanSums(CHAN0).mechHysteresis(1).sigma2, gudtScanSums(CHAN0).mechHysteresis(1).n
''Kickdown Start Location
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownStartLoc.failCount.high, gudtScanSums(CHAN0).kickdownStartLoc.failCount.low, gudtScanSums(CHAN0).kickdownStartLoc.max, gudtScanSums(CHAN0).kickdownStartLoc.min, gudtScanSums(CHAN0).kickdownStartLoc.sigma, gudtScanSums(CHAN0).kickdownStartLoc.sigma2, gudtScanSums(CHAN0).kickdownStartLoc.n
''Kickdown Peak Location
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownPeakLoc.failCount.high, gudtScanSums(CHAN0).kickdownPeakLoc.failCount.low, gudtScanSums(CHAN0).kickdownPeakLoc.max, gudtScanSums(CHAN0).kickdownPeakLoc.min, gudtScanSums(CHAN0).kickdownPeakLoc.sigma, gudtScanSums(CHAN0).kickdownPeakLoc.sigma2, gudtScanSums(CHAN0).kickdownPeakLoc.n
''Kickdown Peak Force
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownPeakForce.failCount.high, gudtScanSums(CHAN0).kickdownPeakForce.failCount.low, gudtScanSums(CHAN0).kickdownPeakForce.max, gudtScanSums(CHAN0).kickdownPeakForce.min, gudtScanSums(CHAN0).kickdownPeakForce.sigma, gudtScanSums(CHAN0).kickdownPeakForce.sigma2, gudtScanSums(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownForceSpan.failCount.high, gudtScanSums(CHAN0).kickdownForceSpan.failCount.low, gudtScanSums(CHAN0).kickdownForceSpan.max, gudtScanSums(CHAN0).kickdownForceSpan.min, gudtScanSums(CHAN0).kickdownForceSpan.sigma, gudtScanSums(CHAN0).kickdownForceSpan.sigma2, gudtScanSums(CHAN0).kickdownForceSpan.n
''Kickdown On Location
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownOnLoc.failCount.high, gudtScanSums(CHAN0).kickdownOnLoc.failCount.low, gudtScanSums(CHAN0).kickdownOnLoc.max, gudtScanSums(CHAN0).kickdownOnLoc.min, gudtScanSums(CHAN0).kickdownOnLoc.sigma, gudtScanSums(CHAN0).kickdownOnLoc.sigma2, gudtScanSums(CHAN0).kickdownOnLoc.n
''Kickdown On Span
'Write #lintFileNum, gudtScanSums(CHAN0).kickdownOnSpan.failCount.high, gudtScanSums(CHAN0).kickdownOnSpan.failCount.low, gudtScanSums(CHAN0).kickdownOnSpan.max, gudtScanSums(CHAN0).kickdownOnSpan.min, gudtScanSums(CHAN0).kickdownOnSpan.sigma, gudtScanSums(CHAN0).kickdownOnSpan.sigma2, gudtScanSums(CHAN0).kickdownOnSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Full-Close) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Full-Close) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (Full-Open) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (Full-Open) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'Write #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'Write #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Peak Force
'Write #lintFileNum, gudtScanSums(CHAN0).peakForce.failCount.high, gudtScanSums(CHAN0).peakForce.failCount.low, gudtScanSums(CHAN0).peakForce.max, gudtScanSums(CHAN0).peakForce.min, gudtScanSums(CHAN0).peakForce.sigma, gudtScanSums(CHAN0).peakForce.sigma2, gudtScanSums(CHAN0).peakForce.n
'
''Exposure data
'Write #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'Write #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'Write #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'Write #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'Write #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'Write #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'Write #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'Write #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'Write #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'Write #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'Write #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the stats file
'Close #lintFileNum
'Call frmMain.RefreshLotFileList         'Add new files to lot file list
'
'Exit Sub
'StatsSave_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Saving Data to Lot File!"

End Sub

Public Sub Save703TLScanResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lstrFileName As String
'Dim lstrExp As String
'Dim lstrExp2 As String
'
'lstrExp = ""
'lstrExp2 = ""
'If gblnDustExp Then
'    lstrExp = "Dust "
'    lstrExp2 = lstrExp2 & "Dust(" & gudtExposure.Dust.TypeofDust & " " & gudtExposure.Dust.AmountofDust & " " & gudtExposure.Dust.StirTime & " " & gudtExposure.Dust.SettleTime & " " & gudtExposure.Dust.Frequency & " " & gudtExposure.Dust.NumberofCycles & " " & gudtExposure.Dust.Condition & ") "
'End If
'If gblnVibrationExp Then
'    lstrExp = lstrExp & "Vibration "
'    lstrExp2 = lstrExp2 & "Vibration(" & gudtExposure.Vibration.Profile & " " & gudtExposure.Vibration.Temperature & " " & gudtExposure.Vibration.Duration & " " & gudtExposure.Vibration.Planes & " " & gudtExposure.Vibration.NumberofCycles & " " & gudtExposure.Vibration.Frequency & ") "
'End If
'If gblnDitherExp Then lstrExp = lstrExp & "Dither "
'If gblnThermalShockExp Then
'    lstrExp = lstrExp & "Thermal Shock "
'    lstrExp2 = lstrExp2 & "Thermal Shock(" & CStr(gudtExposure.ThermalShock.LowTemp) & " " & gudtExposure.ThermalShock.LowTempTime & " " & CStr(gudtExposure.ThermalShock.HighTemp) & " " & gudtExposure.ThermalShock.HighTempTime & " " & gudtExposure.ThermalShock.NumberofCycles & " " & gudtExposure.ThermalShock.Condition & ") "
'End If
'If gblnSaltSprayExp Then
'    lstrExp = lstrExp & "Salt Spray "
'    lstrExp2 = lstrExp2 & "Salt Spray(" & CStr(gudtExposure.SaltSpray.Duration) & " " & gudtExposure.SaltSpray.Condition & ") "
'End If
'If gblnInitialExp Then lstrExp = lstrExp & "Initial "
'If gblnExposure Then
'    lstrExp = lstrExp & "Exposure "
'    lstrExp2 = lstrExp2 & "Exposure (" & gudtExposure.Exposure.Condition & ") "
'End If
'If gblnOperStrnExp Then lstrExp = lstrExp & "Operational Strength "
'If gblnLateralStrnExp Then lstrExp = lstrExp & "Lateral Strength "
'If gblnOpStrnStopExp Then lstrExp = lstrExp & "Operational Strength with Stopper "
'If gblnImpactStrnExp Then lstrExp = lstrExp & "Impact Strength "
'If gblnOperEndurExp Then
'    lstrExp = lstrExp & "Operational Endurance "
'    lstrExp2 = lstrExp2 & "Operational Endurance(" & gudtExposure.OperationalEndurance.Temperature & " " & gudtExposure.OperationalEndurance.NewNumberofCycles & " " & gudtExposure.OperationalEndurance.TotalNumberofCycles & " " & gudtExposure.OperationalEndurance.Condition & ") "
'End If
'If gblnSnapbackExp Then
'    lstrExp = lstrExp & "Snapback "
'    lstrExp2 = lstrExp2 & "Snapback(" & gudtExposure.SnapBack.Temperature & " " & gudtExposure.SnapBack.NumberofCycles & " " & gudtExposure.SnapBack.Condition & ") "
'End If
'If gblnHighTempExp Then
'    lstrExp = lstrExp & "High Temp Soak "
'    lstrExp2 = lstrExp2 & "High Temp Soak(" & CStr(gudtExposure.HighTempSoak.Temperature) & " " & gudtExposure.HighTempSoak.Duration & " " & gudtExposure.HighTempSoak.Condition & ") "
'End If
'If gblnHighTempHighHumidExp Then
'    lstrExp = lstrExp & "High Temp - High Humidity Soak "
'    lstrExp2 = lstrExp2 & "High Temp - High Humidity Soak(" & CStr(gudtExposure.HTempHHumiditySoak.Temperature) & " " & CStr(gudtExposure.HTempHHumiditySoak.RelativeHumidity) & " " & gudtExposure.HTempHHumiditySoak.Duration & " " & gudtExposure.HTempHHumiditySoak.Condition & ") "
'End If
'If gblnLowTempExp Then
'    lstrExp = lstrExp & "Low Temp Soak "
'    lstrExp2 = lstrExp2 & "Low Temp Soak(" & CStr(gudtExposure.LowTempSoak.Temperature) & " " & gudtExposure.LowTempSoak.Duration & " " & gudtExposure.LowTempSoak.Condition & ") "
'End If
'If gblnWaterSprayExp Then
'    lstrExp = lstrExp & "Water Spray "
'    lstrExp2 = lstrExp2 & "Water Spray(" & gudtExposure.WaterSpray.Duration & " " & gudtExposure.WaterSpray.Condition & ") "
'End If
'If gblnChemResExp Then
'    lstrExp = lstrExp & "Chemical Resistance "
'    lstrExp2 = lstrExp2 & "Chemical Resistance(" & CStr(gudtExposure.ChemResistance.Temperature) & " " & gudtExposure.ChemResistance.Duration & " " & gudtExposure.ChemResistance.Substance & ") "
'End If
'If gblnCondenExp Then lstrExp = lstrExp & "Condensation "
'If gblnESDElecExp Then lstrExp = lstrExp & "ElectroStatic Discharge "
'If gblnEMWaveResElecExp Then lstrExp = lstrExp & "Electromagnetic Wave Resistance "
'If gblnBilkCInjElecExp Then lstrExp = lstrExp & "Bilk Current Injection "
'If gblnIgnitionNoiseElecExp Then lstrExp = lstrExp & "Ignition Noise "
'If gblnNarRadEMEElecExp Then lstrExp = lstrExp & "Narrowband Radiated Electromagnetic Energy "
'
''Make the results file name
'lstrFileName = gstrLotName + " Scan Results" & DATAEXT
''Get a file
'lintFileNum = FreeFile
'
''If file does not exist then add a header
'If Not gfsoFileSystemObject.FileExists(PARTSCANDATAPATH + lstrFileName) Then
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'    'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'    Print #lintFileNum, _
'        "Part Number,"; _
'        "Sample Number,"; _
'        "TestLog #,"; _
'        "Date Code,"; _
'        "Date/Time,"; _
'        "S/W Revision,"; _
'        "Parameter File Name,"; _
'        "Pallet Number,"; _
'        "Exposures,"; _
'        "Exposure Data,";
'    'Rising Point, Output #1, & Output #2
'    Print #lintFileNum, _
'        "Rising Point Location [°],"; _
'        "Idle Output Vout #1 [%],"; _
'        "WOT Output Vout #1 [%],"; _
'        "Max Output Vout #1 [%],"; _
'        "Max Absolute Linearity Deviation Vout #1 [%],"; _
'        "Min Absolute Linearity Deviation Vout #1 [%],"; _
'        "Max Slope Deviation Vout #1 [% of Ideal],"; _
'        "Min Slope Deviation Vout #1 [% of Ideal],"; _
'        "Peak Hysteresis Vout #1 [%],"; _
'        "Idle Output Vout #2 [%],"; _
'        "WOT Output Vout #2 [%],"; _
'        "Max Output Vout #2 [%],"; _
'        "Max Absolute Linearity Deviation Vout #2 [%],"; _
'        "Min Absolute Linearity Deviation Vout #2 [%],"; _
'        "Max Slope Deviation Vout #2 [% of Ideal],"; _
'        "Min Slope Deviation Vout #2 [% of Ideal],"; _
'        "Peak Hysteresis Vout #2 [%],";
'    'Correlation, Force, & Kickdown
'    Print #lintFileNum, _
'        "Idle Corr Vout #1/Vout #2 [%],"; _
'        "WOT Corr Vout #1/Vout #2 [%],"; _
'        "Max Fwd Output Corr Vout #1/Vout #2 [%],"; _
'        "Min Fwd Output Corr Vout #1/Vout #2 [%],"; _
'        "Max Rev Output Corr Vout #1/Vout #2 [%],"; _
'        "Min Rev Output Corr Vout #1/Vout #2 [%],"; _
'        "Pedal at Rest Location [°],"; _
'        "Average Force Point 1 [N],"; _
'        "Average Force Point 2 [N],"; _
'        "Peak Force [N],"; _
'        "Max Mechanical Hysteresis [%],"; _
'        "Min Mechanical Hysteresis [%],"; _
'        "Kickdown Start Location [°],"; _
'        "Kickdown Peak Location [°],"; _
'        "Kickdown Peak Force [N],"; _
'        "Kickdown Force Span [N],"; _
'        "Kickdown On Location [°],"; _
'        "Kickdown On Span [°],";
'    'Part Status, Comment, Operator Initials, Temperature, and Series
'    Print #lintFileNum, _
'        "Status,"; _
'        "Comment,"; _
'        "Operator,"; _
'        "Temperature,"; _
'        "Series,"
'Else
'    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
'End If
''Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
'Print #lintFileNum, _
'    gstrSerialNumber; ","; _
'    gstrSampleNum; ","; _
'    frmMain.ctrSetupInfo1.TLNum; ","; _
'    gstrDateCode; ","; _
'    DateTime.Now; ","; _
'    App.Major & "." & App.Minor & "." & App.Revision; ","; _
'    gudtMachine.parameterName; ","; _
'    gintPalletNumber; ","; _
'    lstrExp; ","; _
'    lstrExp2; ",";
''Rising Point, Output #1, & Output #2
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).risingPoint.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN0).Index(2).Value, 3), "##0.000", 2); ","; _
'    Format(Round(gudtReading(CHAN0).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).absoluteLin.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).absoluteLin.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).hysteresis.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(1).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).Index(2).Value, 3), "##0.000"); ","; _
'    Format(Round(gudtReading(CHAN1).maxOutput.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).absoluteLin.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).absoluteLin.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).slope.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN1).hysteresis.Value, 2), "##0.00"); ",";
''Correlation, Force, & Kickdown
'Print #lintFileNum, _
'    Format(Round(gudtReading(CHAN0).indexCor(1), 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).indexCor(2), 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).fwdOutputCor.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).revOutputCor.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).revOutputCor.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).pedalAtRestLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).aveForcePt(1).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).aveForcePt(2).Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).peakForce, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).mechHysteresis.high.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtExtreme(CHAN0).mechHysteresis.low.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownStart.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.location, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownPeak.Value, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownForceSpan, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnLoc, 2), "##0.00"); ","; _
'    Format(Round(gudtReading(CHAN0).kickdownOnSpan, 2), "##0.00"); ",";
''Part Status, Comment, Operator Initials, Temperature, and Series
'If gblnScanFailure Then
'    Print #lintFileNum, "REJECT,";
'Else
'    Print #lintFileNum, "PASS,";
'End If
'Print #lintFileNum, _
'    frmMain.ctrSetupInfo1.Comment; ","; _
'    frmMain.ctrSetupInfo1.Operator; ","; _
'    frmMain.ctrSetupInfo1.Temperature; ","; _
'    frmMain.ctrSetupInfo1.Series
''Close the file
'Close #lintFileNum

End Sub

Public Sub Stats703TLLoad()
'
'   PURPOSE:   To input production statistics into the program from
'              a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsLoad_Err
'
''Clear statistics before starting a new lot or resuming an old lot
'Call StatsClear
'
'frmMain.MousePointer = vbHourglass
'
''Get a file number
'lintFileNum = FreeFile
'
''Check to see if file exists if not exit sub
'If gfsoFileSystemObject.FileExists(STATFILEPATH & gstrLotName & STATEXT) Then
'    Open STATFILEPATH & gstrLotName & STATEXT For Input As #lintFileNum
'Else
'    frmMain.MousePointer = vbNormal
'    Exit Sub
'End If
'
''** General Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
''Display to the form
'frmMain.ctrSetupInfo1.Operator = lstrOperator
'frmMain.ctrSetupInfo1.Temperature = lstrTemperature
'frmMain.ctrSetupInfo1.Comment = lstrComment
'frmMain.ctrSetupInfo1.Series = lstrSeries
'frmMain.ctrSetupInfo1.TLNum = lstrTLNum
'frmMain.ctrSetupInfo1.Sample = lstrSample
'
''*** Scan Information ***
''Rising Point
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).risingPoint.failCount.high, gudtScanStats(CHAN0).risingPoint.failCount.low, gudtScanStats(CHAN0).risingPoint.max, gudtScanStats(CHAN0).risingPoint.min, gudtScanStats(CHAN0).risingPoint.sigma, gudtScanStats(CHAN0).risingPoint.sigma2, gudtScanStats(CHAN0).risingPoint.n
''Loop through all channels
'For lintChanNum = 0 To MAXCHANNUM
'    'Index 1 (Idle)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
'    'Index 2 (WOT)
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
'    'Maximum Output
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
'    'Absolute Linearity
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
'    'Slope Max
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
'    'Slope Min
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n
'Next lintChanNum
''Index #1 Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).indexCor(1).failCount.high, gudtScanStats(CHAN0).indexCor(1).failCount.low, gudtScanStats(CHAN0).indexCor(1).max, gudtScanStats(CHAN0).indexCor(1).min, gudtScanStats(CHAN0).indexCor(1).sigma, gudtScanStats(CHAN0).indexCor(1).sigma2, gudtScanStats(CHAN0).indexCor(1).n
''Index #2 Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).indexCor(2).failCount.high, gudtScanStats(CHAN0).indexCor(2).failCount.low, gudtScanStats(CHAN0).indexCor(2).max, gudtScanStats(CHAN0).indexCor(2).min, gudtScanStats(CHAN0).indexCor(2).sigma, gudtScanStats(CHAN0).indexCor(2).sigma2, gudtScanStats(CHAN0).indexCor(2).n
''Forward Output Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
''Reverse Output Correlation
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
''Pedal Zero Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
''Average Force Pt1
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).aveForcePt(1).failCount.high, gudtScanStats(CHAN0).aveForcePt(1).failCount.low, gudtScanStats(CHAN0).aveForcePt(1).max, gudtScanStats(CHAN0).aveForcePt(1).min, gudtScanStats(CHAN0).aveForcePt(1).sigma, gudtScanStats(CHAN0).aveForcePt(1).sigma2, gudtScanStats(CHAN0).aveForcePt(1).n
''Average Force Pt2
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).aveForcePt(2).failCount.high, gudtScanStats(CHAN0).aveForcePt(2).failCount.low, gudtScanStats(CHAN0).aveForcePt(2).max, gudtScanStats(CHAN0).aveForcePt(2).min, gudtScanStats(CHAN0).aveForcePt(2).sigma, gudtScanStats(CHAN0).aveForcePt(2).sigma2, gudtScanStats(CHAN0).aveForcePt(2).n
''Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
''Mechanical Hysteresis
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mechHysteresis(1).failCount.high, gudtScanStats(CHAN0).mechHysteresis(1).failCount.low, gudtScanStats(CHAN0).mechHysteresis(1).max, gudtScanStats(CHAN0).mechHysteresis(1).min, gudtScanStats(CHAN0).mechHysteresis(1).sigma, gudtScanStats(CHAN0).mechHysteresis(1).sigma2, gudtScanStats(CHAN0).mechHysteresis(1).n, gudtScanStats(CHAN0).mechHysteresis(2).failCount.high, gudtScanStats(CHAN0).mechHysteresis(2).failCount.low, gudtScanStats(CHAN0).mechHysteresis(2).max, gudtScanStats(CHAN0).mechHysteresis(2).min, gudtScanStats(CHAN0).mechHysteresis(2).sigma, gudtScanStats(CHAN0).mechHysteresis(2).sigma2, gudtScanStats(CHAN0).mechHysteresis(2).n
''Kickdown Start Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownStartLoc.failCount.high, gudtScanStats(CHAN0).kickdownStartLoc.failCount.low, gudtScanStats(CHAN0).kickdownStartLoc.max, gudtScanStats(CHAN0).kickdownStartLoc.min, gudtScanStats(CHAN0).kickdownStartLoc.sigma, gudtScanStats(CHAN0).kickdownStartLoc.sigma2, gudtScanStats(CHAN0).kickdownStartLoc.n
''Kickdown Peak Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.high, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.low, gudtScanStats(CHAN0).kickdownPeakLoc.max, gudtScanStats(CHAN0).kickdownPeakLoc.min, gudtScanStats(CHAN0).kickdownPeakLoc.sigma, gudtScanStats(CHAN0).kickdownPeakLoc.sigma2, gudtScanStats(CHAN0).kickdownPeakLoc.n
''Kickdown On Location
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownOnLoc.failCount.high, gudtScanStats(CHAN0).kickdownOnLoc.failCount.low, gudtScanStats(CHAN0).kickdownOnLoc.max, gudtScanStats(CHAN0).kickdownOnLoc.min, gudtScanStats(CHAN0).kickdownOnLoc.sigma, gudtScanStats(CHAN0).kickdownOnLoc.sigma2, gudtScanStats(CHAN0).kickdownOnLoc.n
''Kickdown Peak Force
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownPeakForce.failCount.high, gudtScanStats(CHAN0).kickdownPeakForce.failCount.low, gudtScanStats(CHAN0).kickdownPeakForce.max, gudtScanStats(CHAN0).kickdownPeakForce.min, gudtScanStats(CHAN0).kickdownPeakForce.sigma, gudtScanStats(CHAN0).kickdownPeakForce.sigma2, gudtScanStats(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownForceSpan.failCount.high, gudtScanStats(CHAN0).kickdownForceSpan.failCount.low, gudtScanStats(CHAN0).kickdownForceSpan.max, gudtScanStats(CHAN0).kickdownForceSpan.min, gudtScanStats(CHAN0).kickdownForceSpan.sigma, gudtScanStats(CHAN0).kickdownForceSpan.sigma2, gudtScanStats(CHAN0).kickdownForceSpan.n
''Kickdown On Span
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).kickdownOnSpan.failCount.high, gudtScanStats(CHAN0).kickdownOnSpan.failCount.low, gudtScanStats(CHAN0).kickdownOnSpan.max, gudtScanStats(CHAN0).kickdownOnSpan.min, gudtScanStats(CHAN0).kickdownOnSpan.sigma, gudtScanStats(CHAN0).kickdownOnSpan.sigma2, gudtScanStats(CHAN0).kickdownOnSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Idle) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Idle) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (WOT) Values
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (WOT) Locations
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Exposure data
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the file
'Close #lintFileNum
'frmMain.MousePointer = vbNormal
'
'Exit Sub
'StatsLoad_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Lot File!"

End Sub

Public Sub Stats703TLSave()
'
'   PURPOSE:   To write production statistics to a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

'Dim lintFileNum As Integer
'Dim lintChanNum As Integer
'Dim lintProgrammerNum As Integer
'Dim lstrOperator As String
'Dim lstrTemperature As String
'Dim lstrComment As String
'Dim lstrSeries As String
'Dim lstrTLNum As String
'Dim lstrSample As String
'
'On Error GoTo StatsSave_Err
'
''Get a file number
'lintFileNum = FreeFile
''Open the stats file
'Open STATFILEPATH & gstrLotName & STATEXT For Output As #lintFileNum
'
''Take data from the form
'lstrOperator = frmMain.ctrSetupInfo1.Operator
'lstrTemperature = frmMain.ctrSetupInfo1.Temperature
'lstrComment = frmMain.ctrSetupInfo1.Comment
'lstrSeries = frmMain.ctrSetupInfo1.Series
'lstrTLNum = frmMain.ctrSetupInfo1.TLNum
'lstrSample = frmMain.ctrSetupInfo1.Sample
'
''*** General Information ***
'Write #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
'
''Take data from the form
'lstrOperator = frmMain.ctrSetupInfo1.Operator
'lstrTemperature = frmMain.ctrSetupInfo1.Temperature
'lstrComment = frmMain.ctrSetupInfo1.Comment
'
''*** Scan Information ***
''Rising Point
'Write #lintFileNum, gudtScanStats(CHAN0).risingPoint.failCount.high, gudtScanStats(CHAN0).risingPoint.failCount.low, gudtScanStats(CHAN0).risingPoint.max, gudtScanStats(CHAN0).risingPoint.min, gudtScanStats(CHAN0).risingPoint.sigma, gudtScanStats(CHAN0).risingPoint.sigma2, gudtScanStats(CHAN0).risingPoint.n
''Loop through all channels
'For lintChanNum = 0 To MAXCHANNUM
'    'Index 1 (Idle)
'    Write #lintFileNum, gudtScanStats(lintChanNum).Index(1).failCount.high, gudtScanStats(lintChanNum).Index(1).failCount.low, gudtScanStats(lintChanNum).Index(1).max, gudtScanStats(lintChanNum).Index(1).min, gudtScanStats(lintChanNum).Index(1).sigma, gudtScanStats(lintChanNum).Index(1).sigma2, gudtScanStats(lintChanNum).Index(1).n
'    'Index 2 (WOT)
'    Write #lintFileNum, gudtScanStats(lintChanNum).Index(2).failCount.high, gudtScanStats(lintChanNum).Index(2).failCount.low, gudtScanStats(lintChanNum).Index(2).max, gudtScanStats(lintChanNum).Index(2).min, gudtScanStats(lintChanNum).Index(2).sigma, gudtScanStats(lintChanNum).Index(2).sigma2, gudtScanStats(lintChanNum).Index(2).n
'    'Maximum Output
'    Write #lintFileNum, gudtScanStats(lintChanNum).maxOutput.failCount.high, gudtScanStats(lintChanNum).maxOutput.failCount.low, gudtScanStats(lintChanNum).maxOutput.max, gudtScanStats(lintChanNum).maxOutput.min, gudtScanStats(lintChanNum).maxOutput.sigma, gudtScanStats(lintChanNum).maxOutput.sigma2, gudtScanStats(lintChanNum).maxOutput.n
'    'Absolute Linearity
'    Write #lintFileNum, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.high, gudtScanStats(lintChanNum).linDevPerTol(1).failCount.low, gudtScanStats(lintChanNum).linDevPerTol(1).max, gudtScanStats(lintChanNum).linDevPerTol(1).min, gudtScanStats(lintChanNum).linDevPerTol(1).sigma, gudtScanStats(lintChanNum).linDevPerTol(1).sigma2, gudtScanStats(lintChanNum).linDevPerTol(1).n
'    'Slope Max
'    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMax.failCount.high, gudtScanStats(lintChanNum).slopeMax.failCount.low, gudtScanStats(lintChanNum).slopeMax.max, gudtScanStats(lintChanNum).slopeMax.min, gudtScanStats(lintChanNum).slopeMax.sigma, gudtScanStats(lintChanNum).slopeMax.sigma2, gudtScanStats(lintChanNum).slopeMax.n
'    'Slope Min
'    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n
'Next lintChanNum
''Index #1 Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).indexCor(1).failCount.high, gudtScanStats(CHAN0).indexCor(1).failCount.low, gudtScanStats(CHAN0).indexCor(1).max, gudtScanStats(CHAN0).indexCor(1).min, gudtScanStats(CHAN0).indexCor(1).sigma, gudtScanStats(CHAN0).indexCor(1).sigma2, gudtScanStats(CHAN0).indexCor(1).n
''Index #2 Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).indexCor(2).failCount.high, gudtScanStats(CHAN0).indexCor(2).failCount.low, gudtScanStats(CHAN0).indexCor(2).max, gudtScanStats(CHAN0).indexCor(2).min, gudtScanStats(CHAN0).indexCor(2).sigma, gudtScanStats(CHAN0).indexCor(2).sigma2, gudtScanStats(CHAN0).indexCor(2).n
''Forward Output Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(1).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(1).max, gudtScanStats(CHAN0).outputCorPerTol(1).min, gudtScanStats(CHAN0).outputCorPerTol(1).sigma, gudtScanStats(CHAN0).outputCorPerTol(1).sigma2, gudtScanStats(CHAN0).outputCorPerTol(1).n
''Reverse Output Correlation
'Write #lintFileNum, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.high, gudtScanStats(CHAN0).outputCorPerTol(2).failCount.low, gudtScanStats(CHAN0).outputCorPerTol(2).max, gudtScanStats(CHAN0).outputCorPerTol(2).min, gudtScanStats(CHAN0).outputCorPerTol(2).sigma, gudtScanStats(CHAN0).outputCorPerTol(2).sigma2, gudtScanStats(CHAN0).outputCorPerTol(2).n
''Pedal Zero Location
'Write #lintFileNum, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.high, gudtScanStats(CHAN0).pedalAtRestLoc.failCount.low, gudtScanStats(CHAN0).pedalAtRestLoc.max, gudtScanStats(CHAN0).pedalAtRestLoc.min, gudtScanStats(CHAN0).pedalAtRestLoc.sigma, gudtScanStats(CHAN0).pedalAtRestLoc.sigma2, gudtScanStats(CHAN0).pedalAtRestLoc.n
''Average Force Pt1
'Write #lintFileNum, gudtScanStats(CHAN0).aveForcePt(1).failCount.high, gudtScanStats(CHAN0).aveForcePt(1).failCount.low, gudtScanStats(CHAN0).aveForcePt(1).max, gudtScanStats(CHAN0).aveForcePt(1).min, gudtScanStats(CHAN0).aveForcePt(1).sigma, gudtScanStats(CHAN0).aveForcePt(1).sigma2, gudtScanStats(CHAN0).aveForcePt(1).n
''Average Force Pt2
'Write #lintFileNum, gudtScanStats(CHAN0).aveForcePt(2).failCount.high, gudtScanStats(CHAN0).aveForcePt(2).failCount.low, gudtScanStats(CHAN0).aveForcePt(2).max, gudtScanStats(CHAN0).aveForcePt(2).min, gudtScanStats(CHAN0).aveForcePt(2).sigma, gudtScanStats(CHAN0).aveForcePt(2).sigma2, gudtScanStats(CHAN0).aveForcePt(2).n
''Peak Force
'Write #lintFileNum, gudtScanStats(CHAN0).peakForce.failCount.high, gudtScanStats(CHAN0).peakForce.failCount.low, gudtScanStats(CHAN0).peakForce.max, gudtScanStats(CHAN0).peakForce.min, gudtScanStats(CHAN0).peakForce.sigma, gudtScanStats(CHAN0).peakForce.sigma2, gudtScanStats(CHAN0).peakForce.n
''Mechanical Hysteresis
'Write #lintFileNum, gudtScanStats(CHAN0).mechHysteresis(1).failCount.high, gudtScanStats(CHAN0).mechHysteresis(1).failCount.low, gudtScanStats(CHAN0).mechHysteresis(1).max, gudtScanStats(CHAN0).mechHysteresis(1).min, gudtScanStats(CHAN0).mechHysteresis(1).sigma, gudtScanStats(CHAN0).mechHysteresis(1).sigma2, gudtScanStats(CHAN0).mechHysteresis(1).n, gudtScanStats(CHAN0).mechHysteresis(2).failCount.high, gudtScanStats(CHAN0).mechHysteresis(2).failCount.low, gudtScanStats(CHAN0).mechHysteresis(2).max, gudtScanStats(CHAN0).mechHysteresis(2).min, gudtScanStats(CHAN0).mechHysteresis(2).sigma, gudtScanStats(CHAN0).mechHysteresis(2).sigma2, gudtScanStats(CHAN0).mechHysteresis(2).n
''Kickdown Start Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownStartLoc.failCount.high, gudtScanStats(CHAN0).kickdownStartLoc.failCount.low, gudtScanStats(CHAN0).kickdownStartLoc.max, gudtScanStats(CHAN0).kickdownStartLoc.min, gudtScanStats(CHAN0).kickdownStartLoc.sigma, gudtScanStats(CHAN0).kickdownStartLoc.sigma2, gudtScanStats(CHAN0).kickdownStartLoc.n
''Kickdown Peak Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.high, gudtScanStats(CHAN0).kickdownPeakLoc.failCount.low, gudtScanStats(CHAN0).kickdownPeakLoc.max, gudtScanStats(CHAN0).kickdownPeakLoc.min, gudtScanStats(CHAN0).kickdownPeakLoc.sigma, gudtScanStats(CHAN0).kickdownPeakLoc.sigma2, gudtScanStats(CHAN0).kickdownPeakLoc.n
''Kickdown On Location
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownOnLoc.failCount.high, gudtScanStats(CHAN0).kickdownOnLoc.failCount.low, gudtScanStats(CHAN0).kickdownOnLoc.max, gudtScanStats(CHAN0).kickdownOnLoc.min, gudtScanStats(CHAN0).kickdownOnLoc.sigma, gudtScanStats(CHAN0).kickdownOnLoc.sigma2, gudtScanStats(CHAN0).kickdownOnLoc.n
''Kickdown Peak Force
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownPeakForce.failCount.high, gudtScanStats(CHAN0).kickdownPeakForce.failCount.low, gudtScanStats(CHAN0).kickdownPeakForce.max, gudtScanStats(CHAN0).kickdownPeakForce.min, gudtScanStats(CHAN0).kickdownPeakForce.sigma, gudtScanStats(CHAN0).kickdownPeakForce.sigma2, gudtScanStats(CHAN0).kickdownPeakForce.n
''Kickdown Force Span
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownForceSpan.failCount.high, gudtScanStats(CHAN0).kickdownForceSpan.failCount.low, gudtScanStats(CHAN0).kickdownForceSpan.max, gudtScanStats(CHAN0).kickdownForceSpan.min, gudtScanStats(CHAN0).kickdownForceSpan.sigma, gudtScanStats(CHAN0).kickdownForceSpan.sigma2, gudtScanStats(CHAN0).kickdownForceSpan.n
''Kickdown On Span
'Write #lintFileNum, gudtScanStats(CHAN0).kickdownOnSpan.failCount.high, gudtScanStats(CHAN0).kickdownOnSpan.failCount.low, gudtScanStats(CHAN0).kickdownOnSpan.max, gudtScanStats(CHAN0).kickdownOnSpan.min, gudtScanStats(CHAN0).kickdownOnSpan.sigma, gudtScanStats(CHAN0).kickdownOnSpan.sigma2, gudtScanStats(CHAN0).kickdownOnSpan.n
'
''*** Programming Information ***
''Loop through both programmers
'For lintProgrammerNum = 1 To 2
'    'Index #1 (Idle) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(1).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(1).max, gudtProgStats(lintProgrammerNum).indexVal(1).min, gudtProgStats(lintProgrammerNum).indexVal(1).sigma, gudtProgStats(lintProgrammerNum).indexVal(1).sigma2, gudtProgStats(lintProgrammerNum).indexVal(1).n
'    'Index #1 (Idle) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(1).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(1).max, gudtProgStats(lintProgrammerNum).indexLoc(1).min, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma, gudtProgStats(lintProgrammerNum).indexLoc(1).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(1).n
'    'Index #2 (WOT) Values
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.high, gudtProgStats(lintProgrammerNum).indexVal(2).failCount.low, gudtProgStats(lintProgrammerNum).indexVal(2).max, gudtProgStats(lintProgrammerNum).indexVal(2).min, gudtProgStats(lintProgrammerNum).indexVal(2).sigma, gudtProgStats(lintProgrammerNum).indexVal(2).sigma2, gudtProgStats(lintProgrammerNum).indexVal(2).n
'    'Index #2 (WOT) Locations
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.high, gudtProgStats(lintProgrammerNum).indexLoc(2).failCount.low, gudtProgStats(lintProgrammerNum).indexLoc(2).max, gudtProgStats(lintProgrammerNum).indexLoc(2).min, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma, gudtProgStats(lintProgrammerNum).indexLoc(2).sigma2, gudtProgStats(lintProgrammerNum).indexLoc(2).n
'    'Clamp Low
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLow.failCount.high, gudtProgStats(lintProgrammerNum).clampLow.failCount.low, gudtProgStats(lintProgrammerNum).clampLow.max, gudtProgStats(lintProgrammerNum).clampLow.min, gudtProgStats(lintProgrammerNum).clampLow.sigma, gudtProgStats(lintProgrammerNum).clampLow.sigma2, gudtProgStats(lintProgrammerNum).clampLow.n
'    'Clamp High
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHigh.failCount.high, gudtProgStats(lintProgrammerNum).clampHigh.failCount.low, gudtProgStats(lintProgrammerNum).clampHigh.max, gudtProgStats(lintProgrammerNum).clampHigh.min, gudtProgStats(lintProgrammerNum).clampHigh.sigma, gudtProgStats(lintProgrammerNum).clampHigh.sigma2, gudtProgStats(lintProgrammerNum).clampHigh.n
'    'Offset Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).offsetCode.max, gudtProgStats(lintProgrammerNum).offsetCode.min, gudtProgStats(lintProgrammerNum).offsetCode.sigma, gudtProgStats(lintProgrammerNum).offsetCode.sigma2, gudtProgStats(lintProgrammerNum).offsetCode.n
'    'Rough Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).roughGainCode.max, gudtProgStats(lintProgrammerNum).roughGainCode.min, gudtProgStats(lintProgrammerNum).roughGainCode.sigma, gudtProgStats(lintProgrammerNum).roughGainCode.sigma2, gudtProgStats(lintProgrammerNum).roughGainCode.n
'    'Fine Gain Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).fineGainCode.max, gudtProgStats(lintProgrammerNum).fineGainCode.min, gudtProgStats(lintProgrammerNum).fineGainCode.sigma, gudtProgStats(lintProgrammerNum).fineGainCode.sigma2, gudtProgStats(lintProgrammerNum).fineGainCode.n
'    'Clamp Low Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampLowCode.max, gudtProgStats(lintProgrammerNum).clampLowCode.min, gudtProgStats(lintProgrammerNum).clampLowCode.sigma, gudtProgStats(lintProgrammerNum).clampLowCode.sigma2, gudtProgStats(lintProgrammerNum).clampLowCode.n
'    'Clamp High Code
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).clampHighCode.max, gudtProgStats(lintProgrammerNum).clampHighCode.min, gudtProgStats(lintProgrammerNum).clampHighCode.sigma, gudtProgStats(lintProgrammerNum).clampHighCode.sigma2, gudtProgStats(lintProgrammerNum).clampHighCode.n
'    'Offset seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
'    'Rough Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
'    'Fine Gain seedcode
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
'    'MLX Code Failure Counts
'    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
'Next lintProgrammerNum
'
''*** Programming Summary Information ***
'Write #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal
'
''*** Scanning Summary Information ***
'Write #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal
'
''Exposure data
'Write #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
'Write #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
'Write #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
'Write #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
'Write #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
'Write #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
'Write #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
'Write #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
'Write #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
'Write #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
'Write #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance
'
''Close the stats file
'Close #lintFileNum
'Call frmMain.RefreshLotFileList         'Add new files to lot file list
'
'Exit Sub
'StatsSave_Err:
'
'    MsgBox Err.Description, vbOKOnly, "Error Saving Data to Lot File!"

End Sub

Public Sub Save705TLScanResultsToFile()
'
'   PURPOSE: To save the scan results data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
'1.6ANM added MLX Idd
'1.7ANM added Abs Lin
'1.9ANM added MLX WOT

Dim lintFileNum As Integer
Dim lstrFileName As String
Dim lstrExp As String
Dim lstrExp2 As String

lstrExp = ""
lstrExp2 = ""
If gblnDustExp Then
    lstrExp = "Dust "
    lstrExp2 = lstrExp2 & "Dust(" & gudtExposure.Dust.TypeofDust & " " & gudtExposure.Dust.AmountofDust & " " & gudtExposure.Dust.StirTime & " " & gudtExposure.Dust.SettleTime & " " & gudtExposure.Dust.Frequency & " " & gudtExposure.Dust.NumberofCycles & " " & gudtExposure.Dust.Condition & ") "
End If
If gblnVibrationExp Then
    lstrExp = lstrExp & "Vibration "
    lstrExp2 = lstrExp2 & "Vibration(" & gudtExposure.Vibration.Profile & " " & gudtExposure.Vibration.Temperature & " " & gudtExposure.Vibration.Duration & " " & gudtExposure.Vibration.Planes & " " & gudtExposure.Vibration.NumberofCycles & " " & gudtExposure.Vibration.Frequency & ") "
End If
If gblnDitherExp Then lstrExp = lstrExp & "Dither "
If gblnThermalShockExp Then
    lstrExp = lstrExp & "Thermal Shock "
    lstrExp2 = lstrExp2 & "Thermal Shock(" & CStr(gudtExposure.ThermalShock.LowTemp) & " " & gudtExposure.ThermalShock.LowTempTime & " " & CStr(gudtExposure.ThermalShock.HighTemp) & " " & gudtExposure.ThermalShock.HighTempTime & " " & gudtExposure.ThermalShock.NumberofCycles & " " & gudtExposure.ThermalShock.Condition & ") "
End If
If gblnSaltSprayExp Then
    lstrExp = lstrExp & "Salt Spray "
    lstrExp2 = lstrExp2 & "Salt Spray(" & CStr(gudtExposure.SaltSpray.Duration) & " " & gudtExposure.SaltSpray.Condition & ") "
End If
If gblnInitialExp Then lstrExp = lstrExp & "Initial "
If gblnExposure Then
    lstrExp = lstrExp & "Exposure "
    lstrExp2 = lstrExp2 & "Exposure (" & gudtExposure.Exposure.Condition & ") "
End If
If gblnOperStrnExp Then lstrExp = lstrExp & "Operational Strength "
If gblnLateralStrnExp Then lstrExp = lstrExp & "Lateral Strength "
If gblnOpStrnStopExp Then lstrExp = lstrExp & "Operational Strength with Stopper "
If gblnImpactStrnExp Then lstrExp = lstrExp & "Impact Strength "
If gblnOperEndurExp Then
    lstrExp = lstrExp & "Operational Endurance "
    lstrExp2 = lstrExp2 & "Operational Endurance(" & gudtExposure.OperationalEndurance.Temperature & " " & gudtExposure.OperationalEndurance.NewNumberofCycles & " " & gudtExposure.OperationalEndurance.TotalNumberofCycles & " " & gudtExposure.OperationalEndurance.Condition & ") "
End If
If gblnSnapbackExp Then
    lstrExp = lstrExp & "Snapback "
    lstrExp2 = lstrExp2 & "Snapback(" & gudtExposure.SnapBack.Temperature & " " & gudtExposure.SnapBack.NumberofCycles & " " & gudtExposure.SnapBack.Condition & ") "
End If
If gblnHighTempExp Then
    lstrExp = lstrExp & "High Temp Soak "
    lstrExp2 = lstrExp2 & "High Temp Soak(" & CStr(gudtExposure.HighTempSoak.Temperature) & " " & gudtExposure.HighTempSoak.Duration & " " & gudtExposure.HighTempSoak.Condition & ") "
End If
If gblnHighTempHighHumidExp Then
    lstrExp = lstrExp & "High Temp - High Humidity Soak "
    lstrExp2 = lstrExp2 & "High Temp - High Humidity Soak(" & CStr(gudtExposure.HTempHHumiditySoak.Temperature) & " " & CStr(gudtExposure.HTempHHumiditySoak.RelativeHumidity) & " " & gudtExposure.HTempHHumiditySoak.Duration & " " & gudtExposure.HTempHHumiditySoak.Condition & ") "
End If
If gblnLowTempExp Then
    lstrExp = lstrExp & "Low Temp Soak "
    lstrExp2 = lstrExp2 & "Low Temp Soak(" & CStr(gudtExposure.LowTempSoak.Temperature) & " " & gudtExposure.LowTempSoak.Duration & " " & gudtExposure.LowTempSoak.Condition & ") "
End If
If gblnWaterSprayExp Then
    lstrExp = lstrExp & "Water Spray "
    lstrExp2 = lstrExp2 & "Water Spray(" & gudtExposure.WaterSpray.Duration & " " & gudtExposure.WaterSpray.Condition & ") "
End If
If gblnChemResExp Then
    lstrExp = lstrExp & "Chemical Resistance "
    lstrExp2 = lstrExp2 & "Chemical Resistance(" & CStr(gudtExposure.ChemResistance.Temperature) & " " & gudtExposure.ChemResistance.Duration & " " & gudtExposure.ChemResistance.Substance & ") "
End If
If gblnCondenExp Then lstrExp = lstrExp & "Condensation "
If gblnESDElecExp Then lstrExp = lstrExp & "ElectroStatic Discharge "
If gblnEMWaveResElecExp Then lstrExp = lstrExp & "Electromagnetic Wave Resistance "
If gblnBilkCInjElecExp Then lstrExp = lstrExp & "Bilk Current Injection "
If gblnIgnitionNoiseElecExp Then lstrExp = lstrExp & "Ignition Noise "
If gblnNarRadEMEElecExp Then lstrExp = lstrExp & "Narrowband Radiated Electromagnetic Energy "

'Make the results file name
lstrFileName = gstrLotName + " Scan Results" & DATAEXT
'Get a file
lintFileNum = FreeFile

'If file does not exist then add a header
If Not gfsoFileSystemObject.FileExists(PARTSCANDATAPATH + lstrFileName) Then
    Open PARTSCANDATAPATH + lstrFileName For Append As #lintFileNum
    'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
    Print #lintFileNum, _
        "Part Number,"; _
        "Sample Number,"; _
        "TestLog #,"; _
        "Date Code,"; _
        "Date/Time,"; _
        "S/W Revision,"; _
        "Parameter File Name,"; _
        "Pallet Number,"; _
        "Exposures,"; _
        "Exposure Data,";
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
        "Supply Current AP1 [mA],"; _
        "Supply Current AP2 [mA],"; _
        "Supply WOT Current AP1 [mA],"; _
        "Supply WOT Current AP2 [mA],";
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

If frmMain.ctrSetupInfo1.Sample <> "" Then
    gstrSampleNum = frmMain.ctrSetupInfo1.Sample
End If

'Part S/N, Sample, Date Code, Date/Time, Software Revision, Parameter File Name, Pallet Number, Exposures, and Exposure Data
Print #lintFileNum, _
    gstrSerialNumber; ","; _
    gstrSampleNum; ","; _
    frmMain.ctrSetupInfo1.TLNum; ","; _
    gstrDateCode; ","; _
    DateTime.Now; ","; _
    App.Major & "." & App.Minor & "." & App.Revision; ","; _
    gudtMachine.parameterName; ","; _
    gintPalletNumber; ","; _
    lstrExp; ","; _
    lstrExp2; ",";
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
    Format(Round(gudtReading(CHAN1).mlxWCurrent, 2), "##0.00"); ",";
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

Public Sub Stats705TLLoad()
'
'   PURPOSE:   To input production statistics into the program from
'              a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintFileNum As Integer
Dim lintChanNum As Integer
Dim lintProgrammerNum As Integer
Dim lstrOperator As String
Dim lstrTemperature As String
Dim lstrComment As String
Dim lstrSeries As String
Dim lstrTLNum As String
Dim lstrSample As String

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
If Not EOF(lintFileNum) Then Input #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample
'Display to the form
frmMain.ctrSetupInfo1.Operator = lstrOperator
frmMain.ctrSetupInfo1.Temperature = lstrTemperature
frmMain.ctrSetupInfo1.Comment = lstrComment
frmMain.ctrSetupInfo1.Series = lstrSeries
frmMain.ctrSetupInfo1.TLNum = lstrTLNum
frmMain.ctrSetupInfo1.Sample = lstrSample

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
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.max, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n '1.6ANM
    'Full-Close Hysteresis
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(lintChanNum).FullCloseHys.failCount.high, gudtScanStats(lintChanNum).FullCloseHys.failCount.low, gudtScanStats(lintChanNum).FullCloseHys.max, gudtScanStats(lintChanNum).FullCloseHys.min, gudtScanStats(lintChanNum).FullCloseHys.sigma, gudtScanStats(lintChanNum).FullCloseHys.sigma2, gudtScanStats(lintChanNum).FullCloseHys.n '1.6ANM
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
    'Offset seedcode
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
    'Rough Gain seedcode
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
    'Fine Gain seedcode
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
    'MLX Code Failure Counts
    If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
Next lintProgrammerNum

'*** Programming Summary Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal

'*** Scanning Summary Information ***
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal

'Exposure data
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance

'MLX Current '1.6ANM
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mlxCurrent.failCount.high, gudtScanStats(CHAN0).mlxCurrent.failCount.low, gudtScanStats(CHAN0).mlxCurrent.max, gudtScanStats(CHAN0).mlxCurrent.min, gudtScanStats(CHAN0).mlxCurrent.sigma, gudtScanStats(CHAN0).mlxCurrent.sigma2, gudtScanStats(CHAN0).mlxCurrent.n
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN1).mlxCurrent.failCount.high, gudtScanStats(CHAN1).mlxCurrent.failCount.low, gudtScanStats(CHAN1).mlxCurrent.max, gudtScanStats(CHAN1).mlxCurrent.min, gudtScanStats(CHAN1).mlxCurrent.sigma, gudtScanStats(CHAN1).mlxCurrent.sigma2, gudtScanStats(CHAN1).mlxCurrent.n

'MLX WOT Current '1.9ANM
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN0).mlxWCurrent.failCount.high, gudtScanStats(CHAN0).mlxWCurrent.failCount.low, gudtScanStats(CHAN0).mlxWCurrent.max, gudtScanStats(CHAN0).mlxWCurrent.min, gudtScanStats(CHAN0).mlxWCurrent.sigma, gudtScanStats(CHAN0).mlxWCurrent.sigma2, gudtScanStats(CHAN0).mlxWCurrent.n
If Not EOF(lintFileNum) Then Input #lintFileNum, gudtScanStats(CHAN1).mlxWCurrent.failCount.high, gudtScanStats(CHAN1).mlxWCurrent.failCount.low, gudtScanStats(CHAN1).mlxWCurrent.max, gudtScanStats(CHAN1).mlxWCurrent.min, gudtScanStats(CHAN1).mlxWCurrent.sigma, gudtScanStats(CHAN1).mlxWCurrent.sigma2, gudtScanStats(CHAN1).mlxWCurrent.n

'Close the file
Close #lintFileNum
frmMain.MousePointer = vbNormal

Exit Sub
StatsLoad_Err:

    MsgBox Err.Description, vbOKOnly, "Error Retrieving Data from Lot File!"

End Sub

Public Sub Stats705TLSave()
'
'   PURPOSE:   To write production statistics to a disk file.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintFileNum As Integer
Dim lintChanNum As Integer
Dim lintProgrammerNum As Integer
Dim lstrOperator As String
Dim lstrTemperature As String
Dim lstrComment As String
Dim lstrSeries As String
Dim lstrTLNum As String
Dim lstrSample As String

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
lstrTLNum = frmMain.ctrSetupInfo1.TLNum
lstrSample = frmMain.ctrSetupInfo1.Sample

'*** General Information ***
Write #lintFileNum, gstrLotName, lstrOperator, lstrTemperature, lstrComment, lstrSeries, lstrTLNum, lstrSample

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
    Write #lintFileNum, gudtScanStats(lintChanNum).slopeMin.failCount.high, gudtScanStats(lintChanNum).slopeMin.failCount.low, gudtScanStats(lintChanNum).slopeMin.max, gudtScanStats(lintChanNum).slopeMin.min, gudtScanStats(lintChanNum).slopeMin.sigma, gudtScanStats(lintChanNum).slopeMin.sigma2, gudtScanStats(lintChanNum).slopeMin.n '1.6ANM
    'Full-Close Hysteresis
    Write #lintFileNum, gudtScanStats(lintChanNum).FullCloseHys.failCount.high, gudtScanStats(lintChanNum).FullCloseHys.failCount.low, gudtScanStats(lintChanNum).FullCloseHys.max, gudtScanStats(lintChanNum).FullCloseHys.min, gudtScanStats(lintChanNum).FullCloseHys.sigma, gudtScanStats(lintChanNum).FullCloseHys.sigma2, gudtScanStats(lintChanNum).FullCloseHys.n '1.6ANM
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
    'Offset seedcode
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetSeedCode.max, gudtProgStats(lintProgrammerNum).OffsetSeedCode.min, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma, gudtProgStats(lintProgrammerNum).OffsetSeedCode.sigma2, gudtProgStats(lintProgrammerNum).OffsetSeedCode.n
    'Rough Gain seedcode
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.max, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.min, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).RoughGainSeedCode.n
    'Fine Gain seedcode
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).FineGainSeedCode.max, gudtProgStats(lintProgrammerNum).FineGainSeedCode.min, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma, gudtProgStats(lintProgrammerNum).FineGainSeedCode.sigma2, gudtProgStats(lintProgrammerNum).FineGainSeedCode.n
    'MLX Code Failure Counts
    Write #lintFileNum, gudtProgStats(lintProgrammerNum).OffsetDriftCode.failCount.high, gudtProgStats(lintProgrammerNum).AGNDCode.failCount.high, gudtProgStats(lintProgrammerNum).OscillatorAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).CapFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).DACFreqAdjCode.failCount.high, gudtProgStats(lintProgrammerNum).SlowModeCode.failCount.high
Next lintProgrammerNum

'*** Programming Summary Information ***
Write #lintFileNum, gudtProgSummary.totalUnits, gudtProgSummary.totalGood, gudtProgSummary.totalReject, gudtProgSummary.totalNoTest, gudtProgSummary.totalSevere, gudtProgSummary.currentGood, gudtProgSummary.currentTotal

'*** Scanning Summary Information ***
Write #lintFileNum, gudtScanSummary.totalUnits, gudtScanSummary.totalGood, gudtScanSummary.totalReject, gudtScanSummary.totalNoTest, gudtScanSummary.totalSevere, gudtScanSummary.currentGood, gudtScanSummary.currentTotal

'Exposure data
Write #lintFileNum, gudtExposure.Dust.TypeofDust, gudtExposure.Dust.AmountofDust, gudtExposure.Dust.StirTime, gudtExposure.Dust.SettleTime, gudtExposure.Dust.Duration, gudtExposure.Dust.NumberofCycles, gudtExposure.Dust.Condition
Write #lintFileNum, gudtExposure.Vibration.Profile, gudtExposure.Vibration.Temperature, gudtExposure.Vibration.Duration, gudtExposure.Vibration.Planes, gudtExposure.Vibration.NumberofCycles, gudtExposure.Vibration.Frequency
Write #lintFileNum, gudtExposure.ThermalShock.LowTemp, gudtExposure.ThermalShock.LowTempTime, gudtExposure.ThermalShock.HighTemp, gudtExposure.ThermalShock.HighTempTime, gudtExposure.ThermalShock.NumberofCycles, gudtExposure.ThermalShock.Condition
Write #lintFileNum, gudtExposure.SaltSpray.Duration, gudtExposure.SaltSpray.Condition
Write #lintFileNum, gudtExposure.OperationalEndurance.Temperature, gudtExposure.OperationalEndurance.NewNumberofCycles, gudtExposure.OperationalEndurance.TotalNumberofCycles, gudtExposure.OperationalEndurance.Condition
Write #lintFileNum, gudtExposure.SnapBack.Temperature, gudtExposure.SnapBack.NumberofCycles, gudtExposure.SnapBack.Condition
Write #lintFileNum, gudtExposure.HighTempSoak.Temperature, gudtExposure.HighTempSoak.Duration, gudtExposure.HighTempSoak.Condition
Write #lintFileNum, gudtExposure.HTempHHumiditySoak.Temperature, gudtExposure.HTempHHumiditySoak.RelativeHumidity, gudtExposure.HTempHHumiditySoak.Duration, gudtExposure.HTempHHumiditySoak.Condition
Write #lintFileNum, gudtExposure.LowTempSoak.Temperature, gudtExposure.LowTempSoak.Duration, gudtExposure.LowTempSoak.Condition
Write #lintFileNum, gudtExposure.WaterSpray.Duration, gudtExposure.WaterSpray.Condition
Write #lintFileNum, gudtExposure.ChemResistance.Temperature, gudtExposure.ChemResistance.Duration, gudtExposure.ChemResistance.Substance

'MLX Current '1.6ANM
Write #lintFileNum, gudtScanStats(CHAN0).mlxCurrent.failCount.high, gudtScanStats(CHAN0).mlxCurrent.failCount.low, gudtScanStats(CHAN0).mlxCurrent.max, gudtScanStats(CHAN0).mlxCurrent.min, gudtScanStats(CHAN0).mlxCurrent.sigma, gudtScanStats(CHAN0).mlxCurrent.sigma2, gudtScanStats(CHAN0).mlxCurrent.n
Write #lintFileNum, gudtScanStats(CHAN1).mlxCurrent.failCount.high, gudtScanStats(CHAN1).mlxCurrent.failCount.low, gudtScanStats(CHAN1).mlxCurrent.max, gudtScanStats(CHAN1).mlxCurrent.min, gudtScanStats(CHAN1).mlxCurrent.sigma, gudtScanStats(CHAN1).mlxCurrent.sigma2, gudtScanStats(CHAN1).mlxCurrent.n

'MLX WOT Current '1.9ANM
Write #lintFileNum, gudtScanStats(CHAN0).mlxWCurrent.failCount.high, gudtScanStats(CHAN0).mlxWCurrent.failCount.low, gudtScanStats(CHAN0).mlxWCurrent.max, gudtScanStats(CHAN0).mlxWCurrent.min, gudtScanStats(CHAN0).mlxWCurrent.sigma, gudtScanStats(CHAN0).mlxWCurrent.sigma2, gudtScanStats(CHAN0).mlxWCurrent.n
Write #lintFileNum, gudtScanStats(CHAN1).mlxWCurrent.failCount.high, gudtScanStats(CHAN1).mlxWCurrent.failCount.low, gudtScanStats(CHAN1).mlxWCurrent.max, gudtScanStats(CHAN1).mlxWCurrent.min, gudtScanStats(CHAN1).mlxWCurrent.sigma, gudtScanStats(CHAN1).mlxWCurrent.sigma2, gudtScanStats(CHAN1).mlxWCurrent.n

'Close the stats file
Close #lintFileNum
Call frmMain.RefreshLotFileList         'Add new files to lot file list

Exit Sub
StatsSave_Err:

    MsgBox Err.Description, vbOKOnly, "Error Saving Data to Lot File!"

End Sub

Public Sub InitializeTLSensotec()
'
'   PURPOSE:    To initialize communication with the Sensotec SC2000 and set up the
'               limits, as well as define the coefficent between voltage and force.
'
'  INPUT(S):    None.
' OUTPUT(S):    None.

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
    gudtMachine.maxLBF = 25 '1.9aANM
    Call Sensotec.SetDACFullScale(1, gudtMachine.maxLBF)

    'Set Frequency Response
    Call Sensotec.SetFreqResponse(1, 800)

    'Calculate the coefficient for Force -> Voltage from the force cell
    'Call CalcLBFPerVolt

    'Reset sensotec
    Call Sensotec.Reset
    
    'Read force values '1.9aANM
    Call TestLab.ForceLoad
    
    'Assign the gain
    gsngNewtonsPerVolt = gsngForceGain
    'Assign the offset
    gsngForceAmplifierOffset = gsngForceOffset

End If

End Sub
