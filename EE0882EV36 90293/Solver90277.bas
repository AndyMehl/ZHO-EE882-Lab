Attribute VB_Name = "Solver90277"
Option Explicit

'*** Solver Constants ***

'V1.5.0 Private Const HIGHSATURATION = 97               '>=97% of applied represents a saturated output voltage
'V1.5.0 Private Const LOWSATURATION = 3                 '<=3% of applied represents a saturated output voltage

Private Const MAXFGVALUE = 2.59                 'Numerical gain value of FG = 1023
Private Const MINFGVALUE = 1                    'Numerical gain value of FG = 0

'Constants for Clamp Solver
Public Const LOWCLAMPOFFSET = 1023              'Offset bit value to force output to the low clamp
Public Const HIGHCLAMPOFFSET = 0                'Offset bit value to force output to the high clamp
Public Const RGFORCLAMPS = 0                    'Rough Gain bit value used while solving for clamps
Public Const FGFORCLAMPS = 0                    'Fine Gain bit value used while solving for clamps  'V1.3.0
'V1.6.0 Public Const RGFORCLAMPS1 = 14          'Rough Gain output #1 used while solving for clamps 'V1.5.0
'V1.6.0 Public Const FGFORCLAMPS1 = 512         'Fine Gain output #1 used while solving for clamps  'V1.5.0
'V1.6.0 Public Const RGFORCLAMPS2 = 11          'Rough Gain output #2 used while solving for clamps 'V1.5.0
'V1.6.0 Public Const FGFORCLAMPS2 = 512         'Fine Gain output #2 used while solving for clamps  'V1.5.0

Public Const RGFORCLAMPHIGH = 15                'Rough Gain output high used while solving for clamps 'V1.6.0
Public Const FGFORCLAMPHIGH = 1023              'Fine Gain output high used while solving for clamps  'V1.6.0
Public Const RGFORCLAMPLOW = 0                  'Rough Gain output low used while solving for clamps  'V1.6.0
Public Const FGFORCLAMPLOW = 0                  'Fine Gain output low used while solving for clamps   'V1.6.0

Public Const MAXCLAMPCODE = 1023                'Maximum Clamp Code
Public Const MINCLAMPCODE = 0                   'Minimum Clamp Code

Public Const MAXOFFSETCODE = 1023               'Maximum Offset Code
Public Const MINOFFSETCODE = 0                  'Minimum Offset Code

Public Const MAXFINEGAINCODE = 1023             'Maximum Fine Gain Code
Public Const MINFINEGAINCODE = 0                'Minimum Fine Gain Code 'V1.4.0

Public Const MAXCYCLENUM = 2                    'Maximum number of cycles in the Solver
Public Const MAXCLAMPADJUSTCOUNT = 5            'Maximum number of adjustments to Clamps    'V1.4.0 Changed from 3 to 5 'V1.2.0 Added
Public Const NUMBEROFSTEPS = 2                  'Number of Steps in the Solver
'V1.4.0 Public Const MAXHISTORYNUM = 10                 'Number of history codes to use for running average seed code calculations

Type SolverMeasurement
    offset As Integer                           'Offset Code used for measurement
    roughGain As Integer                        'Rough Gain Code used for measurement
    fineGain As Integer                         'Fine Gain Code used for measurement
    MeasuredOutput(1 To 2) As Single            'Measured output at Idle (1) and WOT (2) in % of applied
    CalculatedSlope As Single                   'Calculated slope values
    CalculatedIntercept As Single               'Calcualted interecept values
    MeasurementsOK(1 To 2) As Boolean           'Whether or codes produced non-saturated results
End Type

Type ClampTestType
    Code As Integer                             'Clamp Test Code
    Output As Single                            'Measured Output
End Type

Type ClampType
    IdealValue As Single                             'Ideal clamp value, in % of applied V
    TargetTolerance As Single                        'Target Tolerance around IdealValue, in % of applied V      'V1.2.0
    PassFailTolerance As Single                      'Pass/Fail Tolerance around IdealValue, in % of applied V   'V1.2.0
    InitialCode As Integer                           'Initial code to use for first try
    Test(1 To MAXCLAMPADJUSTCOUNT) As ClampTestType  'Clamp test codes and measured values 'V1.4.0
End Type

Type IndexType
    IdealLocation As Single                     'Ideal Location of each position, in ° from Pedal-At-Rest Location
    IdealValue As Single                        'Ideal Output at each Position, in % of applied V
    TargetTolerance As Single                   'Target Tolerance around IdealValue, in % of applied V      'V1.2.0
    PassFailTolerance As Single                 'Pass/Fail Tolerance around IdealValue, in % of applied V   'V1.2.0
    high As Single                              'High Tolerance, in % of applied V   'V1.9.0
    low As Single                               'Low Tolerance, in % of applied V    'V1.9.0
End Type

Type StepType
    Test(1 To 4) As SolverMeasurement           'Test number for each step of the Solver, 1 to 4    'V1.4.0  Changed from "1 to 3" to "1 to 4"
    MeasuredLocation(1 To 2) As Single          'Measured location for Idle (1) and WOT (2) in °
    NumGoodMeasurements As Integer              'Number of pairs of measurements (Tests) that resulted in non-saturated measurements
End Type

Type CycleType
    Step(1 To NUMBEROFSTEPS) As StepType        'Steps in the solver process, 1 to 4
    OffsetAdjustedOutput(1 To 2) As Single      'Measured Output after an offset adjustment
End Type

Type SolverAttributes
    OSYS1(1 To 2) As Integer
    OSYS2(1 To 2) As Integer
    OSYS3(1 To 2) As Integer
    OSYS5(1 To 2) As Integer
    OSYS6(1 To 2) As Integer
    OSYS7(1 To 2) As Integer
    
    SSYS1(1 To 2) As Single
    SSYS2(1 To 2) As Single
    SSYS3(1 To 2) As Single
    SSYS5(1 To 2) As Single
    SSYS6(1 To 2) As Single
    SSYS7(1 To 2) As Single
    
    Index(1 To 2) As IndexType                      'Two indexes, Idle = 1, WOT = 2
    Cycle(1 To MAXCYCLENUM) As CycleType            'Each Cycle represents one full iteration through the Solver
    Clamp(1 To 2) As ClampType                      'Clamps, 1 = Low Clamp, 2 = High Clamp
    'Pass/Fail Variables
    OffsetNGainGood As Boolean                      'Pass/Fail variable for offset/gain
    ClampsGood As Boolean                           'Pass/Fail variable for clamps
    'Bit Values
    OffsetStep As Single                            'Offset Step per count change of Offset
    ClampStep As Single                             'Clamp Step per count change of Clamp
    'MLX Settings
    Filter As Integer                               'Filter Setting, 0-15
    InvertSlope As Boolean                          'Invert Slope Setting, True/False
    Mode As Integer                                 'Mode Setting, 0-3
    FaultLevel As Boolean                           'FaultLevel Setting, True/False
    MaxOffsetDrift As Integer                       'Maximum Offset Drift Setting, 0-15
    MaxAGND As Integer                              'High AGND Setting, 0-1023
    MinAGND As Integer                              'Low AGND Setting, 0-1023
    FCKADJ As Integer                               'Correct Oscillator Adjust Setting, 0-15
    CKANACH As Integer                              'Correct Capacitor Frequency Adjust Setting, 0-3
    CKDACCH As Integer                              'Correct DAC Frequency Adjust Setting, 0-3
    SlowMode As Boolean                             'Correct Slow Setting, True/False
    'Solver-Gain variables
    CodeRatio(1 To 2, 1 To 3) As Single             '2 sets of 3 ratios of calculated gain codes: test gain codes
    MinRG As Integer                                'Minimum Rough Gain for Solver
    MaxRG As Integer                                'Maximum Rough Gain for Solver
    'Initial gain & offset codes
    HighRGHighFG As Integer                         'Higher Rough Gain Higher Fine Gain setting for Solver  'V1.4.0
    LowRGHighFG As Integer                          'Lower Rough Gain Higher Fine Gain setting for Solver   'V1.4.0
    HighRGLowFG As Integer                          'Higher Rough Gain Lower Fine Gain setting for Solver   'V1.4.0
    LowRGLowFG As Integer                           'Lower Rough Gain Lower Fine Gain setting for Solver    'V1.4.0
    InitialOffset As Integer                        'Initial Offset setting for Solver
'V1.4.0     'Recent gain & offset codes
'V1.4.0     OffsetHistory(1 To MAXHISTORYNUM) As Integer    'Last X Offset codes of successfully programmed parts
'V1.4.0     RGHistory(1 To MAXHISTORYNUM) As Integer        'Last X Rough Gain codes of successfully programmed parts
'V1.4.0     FGHistory(1 To MAXHISTORYNUM) As Integer        'Last X Fine Gain codes of successfully programmed parts
'V1.4.0     NumHistoryCodes As Integer                      'Number of codes in the history arrays
'V1.4.0     NextHistoryCode As Integer                      'Next array location to write to in the history arrays
'V1.4.0     OffsetSeedCode As Integer                       'Seed Offset Code to start Solving with
'V1.4.0     RoughGainSeedCode As Integer                    'Seed Rough Gain Code to start Solving with
'V1.4.0     FineGainSeedCode As Integer                     'Seed Fine Gain Code to start Solving with
    'Final Chosen Codes
    FinalOffsetCode As Integer                      'Final Calculated Offset code to use
    FinalRGCode As Integer                          'Final Calculated Rough Gain code to use
    FinalFGCode As Integer                          'Final Calculated Fine Gain code to use
    FinalClampHighCode As Integer                   'Final Calculated Clamp High code to use
    FinalClampLowCode As Integer                    'Final Calculated Clamp Low code to use
    'Final Measured Values
    FinalIndexVal(1 To 2) As Single                 'Final Measured Index Outputs
    FinalIndexLoc(1 To 2) As Single                 'Final Measured Index Locations
    FinalClampHighVal As Single                     'Final Measured Clamp High Value
    FinalClampLowVal As Single                      'Final Measured Clamp Low Value
    HighSaturation As Single                        'High saturation level 'V1.5.0
    LowSaturation As Single                         'Low saturation level  'V1.5.0
    ZeroGXPos(1 To 2, 1 To 2) As Single             'Zero Gauss X-Pos      'V1.9.4
End Type

Public gblnProgrammingSuccessful As Boolean         'Pass/Fail variable for entire programming process (Offset, Gain, and Clamps, both outputs)
Public gudtSolver(1 To 2) As SolverAttributes       'Instantiate the Solver variables

Public Function AdjustOffset(CycleNum As Integer) As Boolean
'
'   PURPOSE: First, calculate adjusted offset codes.  Next, make
'            measurements with these codes and the best gain codes
'            available.  Determine if the new offset output is acceptable.
'            Note that this routine assumes that it takes place after
'            Step 2 of the Solver.
'
'  INPUT(S): CycleNum    = The current cycle number of the Solver
'
' OUTPUT(S): Function returns true if the best slope was good enough to use
'            as a basis for an offset adjustment (i.e. within spec), returns
'            false otherwise.

Dim lintProgrammerNum As Integer
Dim lintBestTest(1 To 2) As Integer
Dim lsngDeltaAtIndex1 As Single
Dim lsngIndexError As Single
Dim lintIndexNum As Integer
Dim lblnGoodOutput(1 To 2, 1 To 2) As Boolean
Dim lsngIdealSlope As Single

'Initialize the return value of the function
AdjustOffset = False

For lintProgrammerNum = 1 To 2
    'Atempt to calculate adjusted offset codes
    If CalcAdjustedOffsetCodes(lintProgrammerNum, CycleNum, gudtMLX90277(lintProgrammerNum).Write.offset, gudtMLX90277(lintProgrammerNum).Write.RGain, gudtMLX90277(lintProgrammerNum).Write.FGain, lintBestTest(lintProgrammerNum)) Then
        'Encode the new values into the boolean EEPROM array
        Call EncodeEEpromWrite(lintProgrammerNum)
    Else
        'If we were unable to calculate proper codes, exit the sub
        Exit Function
    End If
Next lintProgrammerNum

'Write the new values to Temp RAM (Both IC's)
Call MLX90277.WriteTempRAM

'Delay 50 msec
Call frmSolver90277.KillTime(50)

'Take readings with the new codes applied (Index 1)
Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).Cycle(CycleNum).OffsetAdjustedOutput(1), gudtSolver(2).Cycle(CycleNum).OffsetAdjustedOutput(1))
'Determine index 2 (WOT) output and check both indexes against limits
For lintProgrammerNum = 1 To 2
    'Calculate the change in index 1 output created by the new offset code
    lsngDeltaAtIndex1 = gudtSolver(lintProgrammerNum).Cycle(CycleNum).OffsetAdjustedOutput(1) - gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintBestTest(lintProgrammerNum)).MeasuredOutput(1)
    'Calculate the theoretical index 2 output based on the shift at index 1
    gudtSolver(lintProgrammerNum).Cycle(CycleNum).OffsetAdjustedOutput(2) = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintBestTest(lintProgrammerNum)).MeasuredOutput(2) + lsngDeltaAtIndex1
    'Calculate the ideal slope
    lsngIdealSlope = (gudtSolver(lintProgrammerNum).Index(2).IdealValue - gudtSolver(lintProgrammerNum).Index(1).IdealValue) / (gudtSolver(lintProgrammerNum).Index(2).IdealLocation - gudtSolver(lintProgrammerNum).Index(1).IdealLocation)
    'Check the outputs against the specification
    For lintIndexNum = 1 To 2
        'Calculate the Error at each index: abs(measured - ideal)
        lsngIndexError = Abs(gudtSolver(lintProgrammerNum).Cycle(CycleNum).OffsetAdjustedOutput(lintIndexNum) - (gudtSolver(lintProgrammerNum).Index(lintIndexNum).IdealValue + (gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).MeasuredLocation(lintIndexNum) - gudtSolver(lintProgrammerNum).Index(lintIndexNum).IdealLocation) * lsngIdealSlope))
        'If this is the last cycle, check against the Pass/Fail Tolerances
        If CycleNum = MAXCYCLENUM Then      'V1.2.0  Added If statement to use two different sets of tolerances
            'If the error is less than the Pass/Fail Tolerance, the output is Good
            If (gudtSolver(lintProgrammerNum).Index(lintIndexNum).PassFailTolerance > lsngIndexError) Then  'V1.3.0 Reversed Comparison Statement
                lblnGoodOutput(lintProgrammerNum, lintIndexNum) = True
            Else
                lblnGoodOutput(lintProgrammerNum, lintIndexNum) = False
            End If
        Else
            'If the error is less than the Target Tolerance, the output is Good
            If (gudtSolver(lintProgrammerNum).Index(lintIndexNum).TargetTolerance > lsngIndexError) Then    'V1.3.0 Reversed Comparison Statement
                lblnGoodOutput(lintProgrammerNum, lintIndexNum) = True
            Else
                lblnGoodOutput(lintProgrammerNum, lintIndexNum) = False
            End If
        End If
    Next lintIndexNum
    '*** Define whether or not the outputs were within tolerance ***
    gudtSolver(lintProgrammerNum).OffsetNGainGood = lblnGoodOutput(lintProgrammerNum, 1) And lblnGoodOutput(lintProgrammerNum, 2)
Next lintProgrammerNum

'Save the final codes
For lintProgrammerNum = 1 To 2
'    'If a good solution was not reached...                                   'V1.3.0
'    If Not gudtSolver(lintProgrammerNum).OffsetNGainGood Then                'V1.3.0
'        'Set RG, FG, & Offset to force output into low diagnostic region     'V1.3.0
'        gudtMLX90277(lintProgrammerNum).Write.offset = LOWCLAMPOFFSET        'V1.3.0
'
'        'Set RG & FG to ensure that the offset is able to force to the clamp 'V1.5.0
'        If lintProgrammerNum = 1 Then
'            gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPS1
'            gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPS1
'        Else
'            gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPS2
'            gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPS2
'        End If
'    End If                                                                   'V1.3.0
    'Save Final Codes
    gudtSolver(lintProgrammerNum).FinalOffsetCode = gudtMLX90277(lintProgrammerNum).Write.offset
    gudtSolver(lintProgrammerNum).FinalRGCode = gudtMLX90277(lintProgrammerNum).Write.RGain
    gudtSolver(lintProgrammerNum).FinalFGCode = gudtMLX90277(lintProgrammerNum).Write.FGain
    For lintIndexNum = 1 To 2
        'Measured (Index 1) & Theoretical (Index 2) Output
        gudtSolver(lintProgrammerNum).FinalIndexVal(lintIndexNum) = gudtSolver(lintProgrammerNum).Cycle(CycleNum).OffsetAdjustedOutput(lintIndexNum)
        'Measured Location (AdjustOffset always occurs after Solver Step #2
        gudtSolver(lintProgrammerNum).FinalIndexLoc(lintIndexNum) = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).MeasuredLocation(lintIndexNum)
    Next lintIndexNum
Next lintProgrammerNum

'Return the result of the tests
AdjustOffset = gudtSolver(1).OffsetNGainGood And gudtSolver(2).OffsetNGainGood

End Function

Public Function CalcAdjustedOffsetCodes(ProgrammerNum As Integer, CycleNum As Integer, offset As Integer, RG As Integer, FG As Integer, BestTestNum As Integer) As Boolean
'
'   PURPOSE: First, determine if one of the slopes was within spec.  If not,
'            report this.  If so, calculate an adjustment factor for the
'            offset.  Note that this routine assumes that it takes place
'            after Step 2 of the Solver.
'
'  INPUT(S): ProgrammerNum = the current programmer (Output #)
'            CycleNum      = The current cycle number of the Solver
'
' OUTPUT(S): Offset      = The new calculated offset code
'            RG          = The Rough Gain from the test # with the best slope
'            FG          = The Fine Gain from the test # with the best slope
'            BestTestNum = The test # with the slope closest to ideal slope
'            Function returns true if the best slope was good enough to use
'            as a basis for an offset adjustment (i.e. within spec), returns
'            false otherwise.

Dim lintTestNum As Integer

Dim lsngIdealSlope As Single
Dim lsngHighSlope As Single
Dim lsngLowSlope As Single

Dim lsngMinDiff As Single
Dim lsngDiff As Single

Dim lsngDeltaOffset As Single
Dim lintDeltaOffset As Integer

On Error GoTo AdjustOffsetError

'Initialize the routine to return false
CalcAdjustedOffsetCodes = False

'Initialize the minimum difference to a big number
lsngMinDiff = 10000

'Define the ideal slope
lsngIdealSlope = ((gudtSolver(ProgrammerNum).Index(2).IdealValue) - (gudtSolver(ProgrammerNum).Index(1).IdealValue)) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)

If CycleNum = MAXCYCLENUM Then      'V1.2.0  Added If statement to use two different sets of tolerances
    'Define the highest acceptable slope, using the lowest index 1 & highest index 2 allowable (Pass/Fail Limits):
    lsngHighSlope = ((gudtSolver(ProgrammerNum).Index(2).IdealValue + gudtSolver(ProgrammerNum).Index(2).PassFailTolerance) - (gudtSolver(ProgrammerNum).Index(1).IdealValue - gudtSolver(ProgrammerNum).Index(1).PassFailTolerance)) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)
    'Define the lowest acceptable slope, using the highest index 1 & lowest index 2 allowable (Pass/Fail Limits):
    lsngLowSlope = ((gudtSolver(ProgrammerNum).Index(2).IdealValue - gudtSolver(ProgrammerNum).Index(2).PassFailTolerance) - (gudtSolver(ProgrammerNum).Index(1).IdealValue + gudtSolver(ProgrammerNum).Index(1).PassFailTolerance)) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)
Else
    'Define the highest acceptable slope, using the lowest index 1 & highest index 2 allowable (Target Limits):
    lsngHighSlope = ((gudtSolver(ProgrammerNum).Index(2).IdealValue + gudtSolver(ProgrammerNum).Index(2).TargetTolerance) - (gudtSolver(ProgrammerNum).Index(1).IdealValue - gudtSolver(ProgrammerNum).Index(1).TargetTolerance)) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)
    'Define the lowest acceptable slope, using the highest index 1 & lowest index 2 allowable (Target Limits):
    lsngLowSlope = ((gudtSolver(ProgrammerNum).Index(2).IdealValue - gudtSolver(ProgrammerNum).Index(2).TargetTolerance) - (gudtSolver(ProgrammerNum).Index(1).IdealValue + gudtSolver(ProgrammerNum).Index(1).TargetTolerance)) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)
End If

For lintTestNum = 1 To 3
    'Only check if the step 2 codes provided non-saturated measurements
    If gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).MeasurementsOK(1) And gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).MeasurementsOK(2) Then
        'Sort for the the slope closest to the ideal:
        lsngDiff = Abs(gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).CalculatedSlope - lsngIdealSlope)
        If (lsngDiff < lsngMinDiff) Then
            lsngMinDiff = lsngDiff
            BestTestNum = lintTestNum
        End If
    End If
Next lintTestNum

'None of the slopes were even close, so we need another iteration
If BestTestNum = 0 Then Exit Function

'Test to determine if the best slope is within tolerance
If (gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).CalculatedSlope >= lsngHighSlope) Or (gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).CalculatedSlope <= lsngLowSlope) Then
    'If it's not, the solver will need another iteration to complete; exit the function
    Exit Function
End If

'Calculate the amount to shift the offset by (in % of applied):
'This is the difference between the ideal value and the measured value at index 1
'for the test measurements with the best slope
lsngDeltaOffset = gudtSolver(ProgrammerNum).Index(1).IdealValue - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).MeasuredOutput(1)

'Calculate the amount to shift the offset by (in counts):
lintDeltaOffset = lsngDeltaOffset / gudtSolver(ProgrammerNum).OffsetStep

'Calculate the new offset code
'(old offset + DeltaOffset)
offset = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).offset + lintDeltaOffset

'Insure that the offset code does not go beyond its limits:
If offset > MAXOFFSETCODE Then
    offset = MAXOFFSETCODE
ElseIf offset < MINOFFSETCODE Then
    offset = MINOFFSETCODE
End If

'Assign the RG and FG
RG = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).roughGain
FG = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(2).Test(BestTestNum).fineGain

'The function has completed successfully
CalcAdjustedOffsetCodes = True

Exit Function

AdjustOffsetError:
    CalcAdjustedOffsetCodes = False

End Function

Public Function CalcFineGain(fineGain As Integer) As Single
'
'   PURPOSE:   Returns the value of the coefficient for Fine Gain based on the
'              integer value passed in.
'
'  INPUT(S):   FineGain = integer representation of FG parameter in MLX IC
'
' OUTPUT(S):   The function returns the Fine Gain coefficent

'The FG integer creates a gain between 1.0 and 2.59 as given by the formula:
'                      1
'   Gain = -------------------------
'           1 - (0.614 * (FG/1023))

CalcFineGain = 1 / (1 - (0.614 * (fineGain / MAXFINEGAINCODE)))

End Function

Public Function CalcFineGainCount(gain As Single) As Single
'
'   PURPOSE:   Returns the count value of FineGain used to create the desired gain.
'
'  INPUT(S):   Gain = desired gain
'
' OUTPUT(S):   The function returns the FG count value

'By rearranging the formula for Gain from FG, we get:
'               (Gain - 1)
'   FG = -------------------------
'          (0.614 / 1023) * Gain

'Avoid a divide-by-zero
If gain <> 0 Then
    CalcFineGainCount = (gain - 1) / ((0.614 / MAXFINEGAINCODE) * gain)
End If

End Function

Public Function CalculateGainCodes(ProgrammerNum As Integer, ByVal RG0 As Integer, ByVal FG0 As Integer, Ratio As Single, RG1 As Integer, FG1 As Integer) As Boolean
'
'   PURPOSE: To calculate Gain Codes based on a given Rough Gain, a given Fine
'            Gain, and a ratio of desired gain : given gain.
'
'
'  INPUT(S): RG0    = Current Rough Gain code
'            FG0    = Current Fine Gain Code
'            ratio  = ratio of desired gain to actual gain
'
' OUTPUT(S): RG1    = Calculated Rough Gain
'            FG1    = Calculated Fine Gain

Dim lsngCurrentGain As Single
Dim lsngDesiredGain As Single
Dim lsngFineGainNeeded As Single
Dim lintRG As Integer
Dim lsngrgo
'Initialize to return the current gains in case no new gain is calculated
RG1 = RG0
FG1 = FG0
CalculateGainCodes = False

'Calculate the numerical value of the current gain codes
lsngCurrentGain = CalcRoughGain(RG0) * CalcFineGain(FG0)

'Calculate the numerical value of the desired gain
lsngDesiredGain = lsngCurrentGain * Ratio

'V1.4.0 'Loop through in reverse order to choose the highest RG available
'V1.4.0 For lintRG = gudtSolver(ProgrammerNum).MaxRG To gudtSolver(ProgrammerNum).MinRG Step -1

'Instead of looping...                                                  'V1.4.0
'Always use the Rough Gain Code that was passed in to this routine      'V1.4.0
lintRG = RG0                                                            'V1.4.0

'Calculate what FG would be needed to give the desired gain
lsngFineGainNeeded = lsngDesiredGain / CalcRoughGain(lintRG)

If (lsngFineGainNeeded < MINFGVALUE) Then
    'If the FG necessary is less than the possible FG, set FG = MINFGCODE
    RG1 = lintRG
    FG1 = MINFINEGAINCODE
ElseIf (lsngFineGainNeeded > MAXFGVALUE) Then
    'If the FG necessary is more than the possible FG, set FG = MAXFGCODE
    RG1 = lintRG
    FG1 = MAXFINEGAINCODE
Else
    'If the FG necessary is within the window of possible FG, calculate it
    RG1 = lintRG
    FG1 = CalcFineGainCount(lsngFineGainNeeded)
    'V1.4.0 Exit For    'Done
End If
'SRC lintRG = MsgBox("cg " & CStr(Format(lsngCurrentGain, "###.##")) & " dg " & CStr(Format(lsngDesiredGain, "###.##")) & " rg " & CStr(Format(lintRG, "###.##")), vbOKOnly)
'SRC lintRG = MsgBox("need " & CStr(Format(lsngFineGainNeeded, "###.##")) & " of " & CStr(Format(MAXFGVALUE, "###.##")) & " and " & CStr(Format(MINFGVALUE, "###.##")), vbOKOnly)
'V1.4.0 Next lintRG

CalculateGainCodes = True

End Function

Public Function CalcRoughGain(roughGain As Integer) As Single
'
'   PURPOSE:   Returns the value of the coefficient for Rough Gain based on the
'              integer value passed in.
'
'  INPUT(S):   RoughGain = integer representation of RG parameter in MLX IC
'
' OUTPUT(S):   The function returns the Rough Gain coefficent

Dim lsngGainDIDO As Single
Dim lsngGainDTS As Single

'The RG integer creates a gain between 16 and 820;
'The LSB of the integer represent the "DTS" amplifier,
'and the MSB of the integer represent the "DIDO" amplifier.
'The product of the DIDO amplifier and the DTS amplifier give
'the Rough Gain.

'The two LSB determine the gain of the DTS:
Select Case (roughGain And (BIT1 Or BIT0))
    Case 0               'Values from MLX Documentation
        lsngGainDTS = 1
    Case 1
        lsngGainDTS = 1.5
    Case 2
        lsngGainDTS = 7 / 3
    Case 3
        lsngGainDTS = 4
End Select

'The two MSB determine the gain of the DIDO:
Select Case (roughGain And (BIT3 Or BIT2))
    Case 0               'Values from MLX Documentation
        lsngGainDIDO = 16
    Case 4
        lsngGainDIDO = 39
    Case 8
        lsngGainDIDO = 82
    Case 12
        lsngGainDIDO = 205
End Select

'Multiply the two stages together to get the Rough Gain
CalcRoughGain = lsngGainDIDO * lsngGainDTS

End Function

'V1.4.0 Public Sub CalcSeedCodes(ProgrammerNum)
'V1.4.0 '
'V1.4.0 '   PURPOSE:   Calculate the seed codes for Offset, Rough Gain, and Fine
'V1.4.0 '              Gain, based on the running average of the last X number of
'V1.4.0 '              parts that were successfully programmed.
'V1.4.0 '
'V1.4.0 '  INPUT(S):   None
'V1.4.0 '
'V1.4.0 ' OUTPUT(S):   None
'V1.4.0
'V1.4.0 Dim lintCodeNum As Integer
'V1.4.0 Dim lsngRGSum As Single             'V1.2.0 was long, used to track codes before.
'V1.4.0 Dim lsngFGSum As Single             'V1.2.0 was long, used to track codes before.
'V1.4.0 Dim lsngAverageRG As Single         'V1.2.0
'V1.4.0 Dim lsngAverageFG As Single         'V1.2.0
'V1.4.0 Dim lsngAverageTotalGain As Single  'V1.2.0
'V1.4.0 Dim lintRG As Integer               'V1.2.0
'V1.4.0 Dim lsngFineGainNeeded As Single    'V1.2.0
'V1.4.0 Dim llngOffsetSum As Long
'V1.4.0
'V1.4.0 'Check to see if any codes have been saved yet
'V1.4.0 If gudtSolver(ProgrammerNum).NumHistoryCodes > 0 Then
'V1.4.0     'Sum the codes we've used so far to produce good parts
'V1.4.0     For lintCodeNum = 1 To gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0         llngOffsetSum = llngOffsetSum + gudtSolver(ProgrammerNum).OffsetHistory(lintCodeNum)
'V1.4.0         lsngRGSum = lsngRGSum + CalcRoughGain(gudtSolver(ProgrammerNum).RGHistory(lintCodeNum)) 'V1.2.0 Use calculated gain instead of gain code
'V1.4.0         lsngFGSum = lsngFGSum + CalcFineGain(gudtSolver(ProgrammerNum).FGHistory(lintCodeNum))  'V1.2.0 Use calculated gain instead of gain code
'V1.4.0     Next lintCodeNum
'V1.4.0     'Average the Offset History Codes
'V1.4.0     gudtSolver(ProgrammerNum).OffsetSeedCode = llngOffsetSum / gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0
'V1.4.0     'V1.2.0\/\/\/
'V1.4.0
'V1.4.0     'Calculate the average Rough Gain
'V1.4.0     lsngAverageRG = lsngRGSum / gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0     'Calculate the average Fine Gain
'V1.4.0     lsngAverageFG = lsngFGSum / gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0     'Calculate the average Total Gain
'V1.4.0     lsngAverageTotalGain = lsngAverageRG * lsngAverageFG
'V1.4.0
'V1.4.0     'Calculate the Gain Codes needed to supply the average Total Gain,
'V1.4.0     'looping through in reverse order to choose the highest RG available
'V1.4.0     For lintRG = gudtSolver(ProgrammerNum).MaxRG To gudtSolver(ProgrammerNum).MinRG Step -1
'V1.4.0         'Calculate what FG would be needed to give the average Total Gain
'V1.4.0         lsngFineGainNeeded = lsngAverageTotalGain / CalcRoughGain(lintRG)
'V1.4.0         'If the FG necessary is within the window of possible FG, calculate it
'V1.4.0         If (lsngFineGainNeeded > MINFGVALUE) And (lsngFineGainNeeded <= MAXFGVALUE) Then
'V1.4.0             gudtSolver(ProgrammerNum).RoughGainSeedCode = lintRG
'V1.4.0             gudtSolver(ProgrammerNum).FineGainSeedCode = CalcFineGainCount(lsngFineGainNeeded)
'V1.4.0             Exit For    'Done
'V1.4.0         End If
'V1.4.0     Next lintRG
'V1.4.0
'V1.4.0     'V1.2.0/\/\/\
'V1.4.0
'V1.4.0     'V1.2.0 gudtSolver(ProgrammerNum).RoughGainSeedCode = lsngRGSum / gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0     'V1.2.0 gudtSolver(ProgrammerNum).FineGainSeedCode = lsngFGSum / gudtSolver(ProgrammerNum).NumHistoryCodes
'V1.4.0 Else
'V1.4.0     'If there are no history codes, the seed codes = the codes from the parameter file
'V1.4.0     gudtSolver(ProgrammerNum).OffsetSeedCode = gudtSolver(ProgrammerNum).InitialOffset
'V1.4.0     gudtSolver(ProgrammerNum).RoughGainSeedCode = gudtSolver(ProgrammerNum).InitialRG
'V1.4.0     gudtSolver(ProgrammerNum).FineGainSeedCode = gudtSolver(ProgrammerNum).InitialFG
'V1.4.0     'Initialize the array location of the next history code
'V1.4.0     gudtSolver(ProgrammerNum).NextHistoryCode = 1
'V1.4.0 End If
'V1.4.0
'V1.4.0 End Sub

Public Sub CalculateSlopesAndIntercepts(CycleNum As Integer, StepNum As Integer)
'
'   PURPOSE: To calculate the slopes and the interecepts for the measured test
'            values on the current step of the solver.
'
'  INPUT(S): CycleNum = the current cycle of the solver
'            StepNum  = the current step of the solver
'
' OUTPUT(S): Calculates the Slope and Intercept for the tests on the current step of
'            the solver.
'

Dim lintProgrammerNum As Integer
Dim lintTestNum As Integer
Dim lintNumOfTests As Integer   'V1.4.0

'V1.4.0\/\/\/
If CycleNum = 1 And StepNum = 1 Then
    lintNumOfTests = 4
Else
    lintNumOfTests = 3
End If
'V1.4.0/\/\/\

For lintProgrammerNum = 1 To 2
    For lintTestNum = 1 To lintNumOfTests       'V1.4.0 changed to loop through lintNumOfTests
        'Calculate the Slope    (m)
        gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope = (gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(2) - gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(1)) / (gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(2) - gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(1))
        'Calculate the Intercept    (b)
        gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedIntercept = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(1) - gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope * gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(1)
    Next lintTestNum
Next lintProgrammerNum

End Sub

Public Function FindSaturationLevels(lintClampNum As Integer) As Boolean
'
'   PURPOSE: To find the high or low clamp saturation levels
'
'  INPUT(S): None
'
' OUTPUT(S): Returns whether or not the clamp saturation levels were successfully found
'V1.5.0 new sub

Dim lintProgrammerNum As Integer

FindSaturationLevels = False

For lintProgrammerNum = 1 To 2
    gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).InitialCode
    'Initialize the codes for clamp solving
    Select Case lintClampNum
        Case 1  'Low Clamp
            'Set Offset to force the output to the low clamp
            gudtMLX90277(lintProgrammerNum).Write.offset = LOWCLAMPOFFSET
            
            'Set RG & FG to ensure that the offset is able to force to the clamp 'V1.6.0
            gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPLOW
            gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPLOW
        
            'Set the test clamp codes 'V1.8.0
            gudtMLX90277(lintProgrammerNum).Write.clampLow = MINCLAMPCODE
            gudtMLX90277(lintProgrammerNum).Write.clampHigh = MINCLAMPCODE
        Case 2  'High Clamp
            'Set Offset to force the output to the high clamp
            gudtMLX90277(lintProgrammerNum).Write.offset = HIGHCLAMPOFFSET
                
            'Set RG & FG to ensure that the offset is able to force to the clamp 'V1.6.0
            gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPHIGH
            gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPHIGH
            
            'Set the test clamp codes 'V1.8.0
            gudtMLX90277(lintProgrammerNum).Write.clampLow = MAXCLAMPCODE
            gudtMLX90277(lintProgrammerNum).Write.clampHigh = MAXCLAMPCODE
    End Select

    'Encode the new values into the boolean EEPROM array
    Call EncodeEEpromWrite(lintProgrammerNum)
Next lintProgrammerNum

'Write the new values to Temp RAM (Both IC's)
Call WriteTempRAM

'Delay 50 msec
Call frmSolver90277.KillTime(50)

'Measure the output caused by the test codes
Select Case lintClampNum
    Case 1  'Low Clamp
        Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).LowSaturation, gudtSolver(2).LowSaturation)
    Case 2  'High Clamp
        Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).HighSaturation, gudtSolver(2).HighSaturation)
End Select

'Check saturation levels to spec
Select Case lintClampNum
    Case 1  'Low Clamp
        If (gudtSolver(1).LowSaturation > gudtSolver(1).Clamp(lintClampNum).IdealValue) Or (gudtSolver(2).LowSaturation > gudtSolver(2).Clamp(lintClampNum).IdealValue) Then    '1.9.0A Use just ideal
            gintAnomaly = 173
            'Log the error to the error log and display the error message
            If gblnReClamp Or Not gblnReClampEnable Then 'V1.9.2
                Call ErrorLogFile("Low clamp saturation level reached.", True, True)
            Else
                Call ErrorLogFile("Low clamp saturation level reached.", False, False)
            End If
        End If
    Case 2  'High Clamp
        If (gudtSolver(1).HighSaturation < gudtSolver(1).Clamp(lintClampNum).IdealValue) Or (gudtSolver(2).HighSaturation < gudtSolver(2).Clamp(lintClampNum).IdealValue) Then  '1.9.0A Use just ideal
            gintAnomaly = 174
            'Log the error to the error log and display the error message
            If gblnReClamp Or Not gblnReClampEnable Then 'V1.9.2
                Call ErrorLogFile("High clamp saturation level reached.", True, True)
            Else
                Call ErrorLogFile("High clamp saturation level reached.", False, False)
            End If
        End If
End Select

'Return true if saturation levels were ok
If gintAnomaly = 0 Then FindSaturationLevels = True

End Function

Public Function ClampSolver(lintClampNum As Integer) As Boolean
'
'   PURPOSE: To set the high or low clamp codes according to the requested clamp
'            values.
'
'  INPUT(S): None
'
' OUTPUT(S): Returns whether or not the clamp was successfully programmed
'V1.5.0 new sub

Dim lintProgrammerNum As Integer
Dim lintClampAdjustCount As Integer

Dim lsngDelta(1 To 2) As Single

Dim lblnClampInSpec(1 To 2) As Boolean  'Boolean representing each output (one clamp)
Dim lblnClampGood As Boolean            'Boolean representing each clamp (both outputs)

'NOTE: lblnClampInSpec(1) represents the clamp being adjusted, low or high,
'      on programmer #1.  lblnClampInSpec(2) represents the same for
'      programmer #2.  lblnClampsInSpec(1) represents whether or not the low clamps
'      on programmer #1 and programmer #2 were in spec, while lblnClampsInSpec(2)
'      represents the same for the high clamps.

'Initialize the value of the routine to false
ClampSolver = False

'Initialize to NOT good
lblnClampGood = False

'Make three attempts at adjusting the clamp value
For lintClampAdjustCount = 1 To MAXCLAMPADJUSTCOUNT
    For lintProgrammerNum = 1 To 2

        'If this is the first try, initialize the codes
        If lintClampAdjustCount = 1 Then
            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).InitialCode
            'Initialize the codes for clamp solving
            Select Case lintClampNum
                Case 1  'Low Clamp
                    'Set Offset to force the output to the low clamp
                    gudtMLX90277(lintProgrammerNum).Write.offset = LOWCLAMPOFFSET
                    
                    'Set RG & FG to ensure that the offset is able to force to the clamp 'V1.6.0
                    gudtMLX90277(lintProgrammerNum).Write.RGain = 0 'V1.9.3 gudtSolver(lintProgrammerNum).MinRG       'V1.9.1
                    gudtMLX90277(lintProgrammerNum).Write.FGain = 0 'V1.9.3 gudtSolver(lintProgrammerNum).LowRGHighFG 'V1.9.1
                Case 2  'High Clamp
                    'Set Offset to force the output to the high clamp
                    gudtMLX90277(lintProgrammerNum).Write.offset = gudtSolver(lintProgrammerNum).InitialOffset 'V1.9.1
                    
                    'Set RG & FG to ensure that the offset is able to force to the clamp 'V1.6.0
                    gudtMLX90277(lintProgrammerNum).Write.RGain = gudtSolver(lintProgrammerNum).MinRG       'V1.9.1
                    gudtMLX90277(lintProgrammerNum).Write.FGain = gudtSolver(lintProgrammerNum).LowRGHighFG 'V1.9.1
            End Select
        End If

        'Set the test clamp codes
        Select Case lintClampNum
            Case 1  'Low Clamp
                'Set the test clamp low code
                gudtMLX90277(lintProgrammerNum).Write.clampLow = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code
                gudtMLX90277(lintProgrammerNum).Write.clampHigh = 1023 'V1.9.1
            Case 2  'High Clamp
                'Set the test clamp high code
                gudtMLX90277(lintProgrammerNum).Write.clampLow = 1023  'V1.9.1
                gudtMLX90277(lintProgrammerNum).Write.clampHigh = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code
        End Select

        'Encode the new values into the boolean EEPROM array
        Call EncodeEEpromWrite(lintProgrammerNum)

    Next lintProgrammerNum

    'Write the new values to Temp RAM (Both IC's)
    Call WriteTempRAM

    'Delay 50 msec
    Call frmSolver90277.KillTime(50)

    'Measure the output caused by the test codes
    Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).Clamp(lintClampNum).Test(lintClampAdjustCount).Output, gudtSolver(2).Clamp(lintClampNum).Test(lintClampAdjustCount).Output)
        
    'Determine if the clamps are in spec
    For lintProgrammerNum = 1 To 2
        lsngDelta(lintProgrammerNum) = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).IdealValue - gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Output
        
        If lintClampAdjustCount = MAXCLAMPADJUSTCOUNT Then
            'If the absolute value of the difference is less than the Pass/Fail Tolerance, the clamp is in spec
            lblnClampInSpec(lintProgrammerNum) = (Abs(lsngDelta(lintProgrammerNum)) < gudtSolver(lintProgrammerNum).Clamp(lintClampNum).PassFailTolerance)
        Else
            'If the absolute value of the difference is less than the Target Tolerance, the clamp is in spec
            lblnClampInSpec(lintProgrammerNum) = (Abs(lsngDelta(lintProgrammerNum)) < gudtSolver(lintProgrammerNum).Clamp(lintClampNum).TargetTolerance)
        End If

        'Save the Final clamp codes and outputs if it is in spec
        If lblnClampInSpec(lintProgrammerNum) Then
            If lintClampNum = 1 Then
                gudtSolver(lintProgrammerNum).FinalClampLowCode = gudtSolver(lintProgrammerNum).Clamp(1).Test(lintClampAdjustCount).Code
                gudtSolver(lintProgrammerNum).FinalClampLowVal = gudtSolver(lintProgrammerNum).Clamp(1).Test(lintClampAdjustCount).Output
            Else
                gudtSolver(lintProgrammerNum).FinalClampHighCode = gudtSolver(lintProgrammerNum).Clamp(2).Test(lintClampAdjustCount).Code
                gudtSolver(lintProgrammerNum).FinalClampHighVal = gudtSolver(lintProgrammerNum).Clamp(2).Test(lintClampAdjustCount).Output
            End If
        End If
    Next lintProgrammerNum
    'If the clamp value is in spec for both programmers, we can exit this loop
    'and move on to the next clamp, or to the end of the routine
    If lblnClampInSpec(1) And lblnClampInSpec(2) Then
        lblnClampGood = True
        Exit For
    Else
        'If we've tried adjusting less than MAXCLAMPADJUSTCOUNT times...
        If lintClampAdjustCount < MAXCLAMPADJUSTCOUNT Then
            'Check saturation levels
            Select Case lintClampNum
                Case 1  'Low Clamp
                    If (gudtSolver(1).Clamp(lintClampNum).Test(lintClampAdjustCount).Output < gudtSolver(1).LowSaturation) Or (gudtSolver(2).Clamp(lintClampNum).Test(lintClampAdjustCount).Output < gudtSolver(2).LowSaturation) Then
                        'Calculate next clamp codes to try by adding half the
                        '(difference / clampstep) to the last clamp code
                        For lintProgrammerNum = 1 To 2
                            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code + ((lsngDelta(lintProgrammerNum) / gudtSolver(lintProgrammerNum).ClampStep) / 2)
                            'Make sure we didn't go beyond 0 or 1023:
                            If gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code < MINCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MINCLAMPCODE
                            ElseIf gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code > MAXCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MAXCLAMPCODE
                            End If
                        Next lintProgrammerNum
                    Else
                        'Calculate next clamp codes to try by adding
                        '(difference / clampstep) to the last clamp code
                        For lintProgrammerNum = 1 To 2
                            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code + (lsngDelta(lintProgrammerNum) / gudtSolver(lintProgrammerNum).ClampStep)
                            'Make sure we didn't go beyond 0 or 1023:
                            If gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code < MINCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MINCLAMPCODE
                            ElseIf gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code > MAXCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MAXCLAMPCODE
                            End If
                        Next lintProgrammerNum
                    End If
                Case 2  'High Clamp
                    If (gudtSolver(1).Clamp(lintClampNum).Test(lintClampAdjustCount).Output > gudtSolver(1).HighSaturation) Or (gudtSolver(2).Clamp(lintClampNum).Test(lintClampAdjustCount).Output > gudtSolver(2).HighSaturation) Then
                        'Calculate next clamp codes to try by adding half the
                        '(difference / clampstep) to the last clamp code
                        For lintProgrammerNum = 1 To 2
                            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code + ((lsngDelta(lintProgrammerNum) / gudtSolver(lintProgrammerNum).ClampStep) / 2)
                            'Make sure we didn't go beyond 0 or 1023:
                            If gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code < MINCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MINCLAMPCODE
                            ElseIf gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code > MAXCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MAXCLAMPCODE
                            End If
                        Next lintProgrammerNum
                    Else
                        'Calculate next clamp codes to try by adding
                        '(difference / clampstep) to the last clamp code
                        For lintProgrammerNum = 1 To 2
                            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code + (lsngDelta(lintProgrammerNum) / gudtSolver(lintProgrammerNum).ClampStep)
                            'Make sure we didn't go beyond 0 or 1023:
                            If gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code < MINCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MINCLAMPCODE
                            ElseIf gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code > MAXCLAMPCODE Then
                                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MAXCLAMPCODE
                            End If
                        Next lintProgrammerNum
                    End If
            End Select
        Else
            For lintProgrammerNum = 1 To 2
                'Save the final clamp codes
                If lintClampNum = 1 Then
                    gudtSolver(lintProgrammerNum).FinalClampLowCode = gudtSolver(lintProgrammerNum).Clamp(1).Test(3).Code
                    gudtSolver(lintProgrammerNum).FinalClampLowVal = gudtSolver(lintProgrammerNum).Clamp(1).Test(3).Output
                Else
                    gudtSolver(lintProgrammerNum).FinalClampHighCode = gudtSolver(lintProgrammerNum).Clamp(2).Test(3).Code
                    gudtSolver(lintProgrammerNum).FinalClampHighVal = gudtSolver(lintProgrammerNum).Clamp(2).Test(3).Output
                End If
            Next lintProgrammerNum
        End If
    End If

Next lintClampAdjustCount

'Reset the test clamp codes for Solving
For lintProgrammerNum = 1 To 2
'V1.9.1 Select Case lintClampNum
'V1.9.1      Case 1  'Low Clamp
         'Reset the test clamp low code
         gudtMLX90277(lintProgrammerNum).Write.clampLow = MINCLAMPCODE
'V1.9.1 'SRC     Case 2  'High Clamp
         'Reset the test clamp high code
         gudtMLX90277(lintProgrammerNum).Write.clampHigh = MAXCLAMPCODE
'V1.9.1 End Select
Next lintProgrammerNum

ClampSolver = lblnClampGood

End Function

'Public Function ClampSolver() As Boolean
''
''   PURPOSE: To set the high and low clamp codes according to the requested clamp
''            values.
''
''  INPUT(S): None
''
'' OUTPUT(S): Returns whether or not the clamps were successfully programmed
'
'Dim lintProgrammerNum As Integer
'Dim lintClampNum As Integer
'Dim lintClampAdjustCount As Integer
'
'Dim lsngDelta(1 To 2) As Single
'
'Dim lblnClampInSpec(1 To 2) As Boolean  'Boolean representing each output (one clamp)
'Dim lblnClampsGood(1 To 2) As Boolean   'Boolean representing each clamp (both outputs)
'
''NOTE: lblnClampInSpec(1) represents the clamp being adjusted, low or high,
''      on programmer #1.  lblnClampInSpec(2) represents the same for
''      programmer #2.  lblnClampsInSpec(1) represents whether or not the low clamps
''      on programmer #1 and programmer #2 were in spec, while lblnClampsInSpec(2)
''      represents the same for the high clamps.
'
''Initialize the value of the routine to false
'ClampSolver = False
'
''Loop through both clamps (1 = clamp low, 2 = clamp high)
'For lintClampNum = 1 To 2
'
'    'Initialize to NOT good
'    lblnClampsGood(lintClampNum) = False
'
'    'Make three attempts at adjusting the clamp value
'    For lintClampAdjustCount = 1 To MAXCLAMPADJUSTCOUNT     'V1.2.0 Use new constant
'        For lintProgrammerNum = 1 To 2
'
'            'If this is the first try, initialize the codes
'            If lintClampAdjustCount = 1 Then
'                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).InitialCode
'                'Initialize the codes for clamp solving
'                Select Case lintClampNum
'                    Case 1  'Low Clamp
'                        'Set Offset to force the output to the low clamp
'                        gudtMLX90277(lintProgrammerNum).Write.offset = LOWCLAMPOFFSET
'                        gudtMLX90277(lintProgrammerNum).Write.InvertSlope = False   'V1.4.0
'                    Case 2  'High Clamp
'                        'Set Offset to force the output to the high clamp
'                        gudtMLX90277(lintProgrammerNum).Write.offset = HIGHCLAMPOFFSET
'                        gudtMLX90277(lintProgrammerNum).Write.InvertSlope = True    'V1.4.0
'                End Select
'                'Set RG & FG to ensure that the offset is able to force to the clamp
'                gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPS
'                gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPS   'V1.3.0
'            End If
'
'            'If this is the first try, initialize the codes
'            If lintClampAdjustCount = 1 Then
'                gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).InitialCode
'                'Initialize the codes for clamp solving
'                Select Case lintClampNum
'                    Case 1  'Low Clamp
'                        'Set Offset to force the output to the low clamp
'                        gudtMLX90277(lintProgrammerNum).Write.offset = LOWCLAMPOFFSET
'                    Case 2  'High Clamp
'                        'Set Offset to force the output to the high clamp
'                        gudtMLX90277(lintProgrammerNum).Write.offset = HIGHCLAMPOFFSET
'                End Select
'                'Set RG & FG to ensure that the offset is able to force to the clamp
'                gudtMLX90277(lintProgrammerNum).Write.RGain = RGFORCLAMPS
'                gudtMLX90277(lintProgrammerNum).Write.FGain = FGFORCLAMPS   'V1.3.0
'            End If
'
'            'Set the test clamp codes
'            Select Case lintClampNum
'                Case 1  'Low Clamp
'                    'Set the test clamp low code
'                    gudtMLX90277(lintProgrammerNum).Write.clampLow = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code
'                Case 2  'High Clamp
'                    'Set the test clamp high code
'                    gudtMLX90277(lintProgrammerNum).Write.clampHigh = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code
'            End Select
'
'            'Encode the new values into the boolean EEPROM array
'            Call EncodeEEpromWrite(lintProgrammerNum)
'
'        Next lintProgrammerNum
'
'        'Write the new values to Temp RAM (Both IC's)
'        Call WriteTempRAM
'
'        'Delay 50 msec
'        Call frmSolver90277.KillTime(50)
'
'        'Measure the output caused by the test codes
'        Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).Clamp(lintClampNum).Test(lintClampAdjustCount).Output, gudtSolver(2).Clamp(lintClampNum).Test(lintClampAdjustCount).Output)
'
'        'Determine if the clamps are in spec
'        For lintProgrammerNum = 1 To 2
'            lsngDelta(lintProgrammerNum) = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).IdealValue - gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Output
'
'            If lintClampAdjustCount = MAXCLAMPADJUSTCOUNT Then  'V1.2.0  Added If statement to use two different sets of tolerances
'                'If the absolute value of the difference is less than the Pass/Fail Tolerance, the clamp is in spec
'                lblnClampInSpec(lintProgrammerNum) = (Abs(lsngDelta(lintProgrammerNum)) < gudtSolver(lintProgrammerNum).Clamp(lintClampNum).PassFailTolerance)
'            Else
'                'If the absolute value of the difference is less than the Target Tolerance, the clamp is in spec
'                lblnClampInSpec(lintProgrammerNum) = (Abs(lsngDelta(lintProgrammerNum)) < gudtSolver(lintProgrammerNum).Clamp(lintClampNum).TargetTolerance)
'            End If
'
'            'Save the Final clamp codes and outputs if it is in spec
'            If lblnClampInSpec(lintProgrammerNum) Then
'                If lintClampNum = 1 Then
'                    gudtSolver(lintProgrammerNum).FinalClampLowCode = gudtSolver(lintProgrammerNum).Clamp(1).Test(lintClampAdjustCount).Code
'                    gudtSolver(lintProgrammerNum).FinalClampLowVal = gudtSolver(lintProgrammerNum).Clamp(1).Test(lintClampAdjustCount).Output
'                Else
'                    gudtSolver(lintProgrammerNum).FinalClampHighCode = gudtSolver(lintProgrammerNum).Clamp(2).Test(lintClampAdjustCount).Code
'                    gudtSolver(lintProgrammerNum).FinalClampHighVal = gudtSolver(lintProgrammerNum).Clamp(2).Test(lintClampAdjustCount).Output
'                End If
'            End If
'        Next lintProgrammerNum
'        'If the clamp value is in spec for both programmers, we can exit this loop
'        'and move on to the next clamp, or to the end of the routine
'        If lblnClampInSpec(1) And lblnClampInSpec(2) Then
'            lblnClampsGood(lintClampNum) = True
'            Exit For
'        Else
'            'If we've tried adjusting less than MAXCLAMPADJUSTCOUNT times...    'V1.4.0
'            If lintClampAdjustCount < MAXCLAMPADJUSTCOUNT Then                  'V1.4.0 Corrected to use constant
'                'Calculate next clamp codes to try by adding
'                '(difference / clampstep) to the last clamp code
'                For lintProgrammerNum = 1 To 2
'                    gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount).Code + (lsngDelta(lintProgrammerNum) / gudtSolver(lintProgrammerNum).ClampStep)
'                    'Make sure we didn't go beyond 0 or 1023:
'                    If gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code < MINCLAMPCODE Then
'                        gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MINCLAMPCODE
'                    ElseIf gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code > MAXCLAMPCODE Then
'                        gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintClampAdjustCount + 1).Code = MAXCLAMPCODE
'                    End If
'                Next lintProgrammerNum
'            Else
'                For lintProgrammerNum = 1 To 2
'                    'Save the final clamp codes
'                    If lintClampNum = 1 Then
'                        gudtSolver(lintProgrammerNum).FinalClampLowCode = gudtSolver(lintProgrammerNum).Clamp(1).Test(3).Code
'                        gudtSolver(lintProgrammerNum).FinalClampLowVal = gudtSolver(lintProgrammerNum).Clamp(1).Test(3).Output
'                    Else
'                        gudtSolver(lintProgrammerNum).FinalClampHighCode = gudtSolver(lintProgrammerNum).Clamp(2).Test(3).Code
'                        gudtSolver(lintProgrammerNum).FinalClampHighVal = gudtSolver(lintProgrammerNum).Clamp(2).Test(3).Output
'                    End If
'                Next lintProgrammerNum
'            End If
'        End If
'
'    Next lintClampAdjustCount
'
'Next lintClampNum
'
'ClampSolver = lblnClampsGood(1) And lblnClampsGood(2)
'
''Reset the Offset and Clamp parameters
'For lintProgrammerNum = 1 To 2
'    gudtMLX90277(lintProgrammerNum).Write.offset = gudtMLX90277(lintProgrammerNum).Read.offset
'    gudtMLX90277(lintProgrammerNum).Write.clampLow = gudtMLX90277(lintProgrammerNum).Read.clampLow
'    gudtMLX90277(lintProgrammerNum).Write.clampHigh = gudtMLX90277(lintProgrammerNum).Read.clampHigh
'    gudtMLX90277(lintProgrammerNum).Write.InvertSlope = gudtSolver(lintProgrammerNum).InvertSlope   'V1.4.0
'Next lintProgrammerNum
'
'End Function

Private Sub ClearSolverVariables()
'
'   PURPOSE: To clear (0) the solver process variables before each run of
'            the solver (each new part).
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer
Dim lintCycleNum As Integer
Dim lintStepNum As Integer
Dim lintTestNum As Integer
Dim lintIndexNum As Integer
Dim lintClampNum As Integer

For lintProgrammerNum = 1 To 2
    'Clear the Variables for each Cycle
    For lintCycleNum = 1 To 2
        'Clear the Variables for each Step
        For lintStepNum = 1 To 2
            gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).NumGoodMeasurements = 0
            For lintIndexNum = 1 To 2
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).MeasuredLocation(lintIndexNum) = 0
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).OffsetAdjustedOutput(lintIndexNum) = 0
            Next lintIndexNum
            For lintTestNum = 1 To 4    ''V1.4.0 loop to 4 instead of 3
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).CalculatedIntercept = 0
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).CalculatedSlope = 0
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).offset = 0
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).roughGain = 0
                gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).fineGain = 0
                For lintIndexNum = 1 To 2
                    gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).MeasuredOutput(lintIndexNum) = 0
                    gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).Test(lintTestNum).MeasurementsOK(lintIndexNum) = False
                Next lintIndexNum
            Next lintTestNum
        Next lintStepNum
    Next lintCycleNum
    'Clear the Clamp variables
    gudtSolver(lintProgrammerNum).ClampsGood = False
    For lintClampNum = 1 To 2
        For lintTestNum = 1 To 3
            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintTestNum).Code = 0
            gudtSolver(lintProgrammerNum).Clamp(lintClampNum).Test(lintTestNum).Output = 0
        Next lintTestNum
    Next lintClampNum
    'Clear the Final variables
    gudtSolver(lintProgrammerNum).OffsetNGainGood = False
    gudtSolver(lintProgrammerNum).FinalOffsetCode = 0
    gudtSolver(lintProgrammerNum).FinalRGCode = 0
    gudtSolver(lintProgrammerNum).FinalFGCode = 0
    For lintIndexNum = 1 To 2
        gudtSolver(lintProgrammerNum).FinalIndexVal(lintIndexNum) = 0
        gudtSolver(lintProgrammerNum).FinalIndexLoc(lintIndexNum) = 0
    Next lintIndexNum
    gudtSolver(lintProgrammerNum).FinalClampLowCode = 0
    gudtSolver(lintProgrammerNum).FinalClampHighCode = 0
    gudtSolver(lintProgrammerNum).FinalClampLowVal = 0
    gudtSolver(lintProgrammerNum).FinalClampHighVal = 0
Next lintProgrammerNum

End Sub

Public Sub CountGoodMeasurements(CycleNum As Integer, StepNum As Integer)
'
'   PURPOSE: To count the number of good measurement pairs on the current Step
'
'  INPUT(S): CycleNum = Current Cycle Number
'            StepNum  = Current Step Number
'
' OUTPUT(S): None

Dim lintProgrammerNum As Integer
Dim lintTestNum As Integer
Dim lintNumOfTests As Integer
Dim lintIndex As Integer

If CycleNum = 1 And StepNum = 1 Then
    lintNumOfTests = 4
Else
    lintNumOfTests = 3
End If

'Determine if the measurement was clipped 'V1.5.0 moved from makesolvermeasurements
For lintProgrammerNum = 1 To 2
    For lintTestNum = 1 To lintNumOfTests
        For lintIndex = 1 To 2
            If GoodMeasurements(lintProgrammerNum, CycleNum, StepNum, lintTestNum, lintIndex) Then
                'If the measurements were good, set the boolean that represents that
                gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(lintIndex) = True
            Else
                'If the measurements were not good, set the boolean accordingly
                gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(lintIndex) = False
            End If
        Next lintIndex
    Next lintTestNum
Next lintProgrammerNum
    
For lintProgrammerNum = 1 To 2
    'Make sure the count is reset
    gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).NumGoodMeasurements = 0
    For lintTestNum = 1 To 4    'V1.4.0 Changed from "1 to 3" to "1 to 4"
        If gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(1) And gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(2) Then
            'Increment the counter if the codes were ok at both indexes
            gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).NumGoodMeasurements = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).NumGoodMeasurements + 1
        End If
    Next lintTestNum
Next lintProgrammerNum

End Sub

Public Function EvaluateTests(CycleNum As Integer, StepNum As Integer) As Boolean
'
'   PURPOSE: To evaluate the results of measurements at index 1 & 2:
'            Were there at least two sets of "good" (non-saturated) measurements?
'            If so, calculate new test codes based on the good measurements
'            If not, calculate new test codes based on the intial codes
'
'  INPUT(S): CycleNum = The current cycle of the Solver
'            StepNum  = The current step of the Solver
'
' OUTPUT(S): anomaly  = system status, 0 = ok, <> 0 = system error

Dim lintProgrammerNum As Integer
Dim lintTestNum As Integer

Dim lblnGoodSolverCalculations As Boolean

Dim lintOffset As Integer
Dim lintRG As Integer
Dim lintFG As Integer
Dim lintNumGoodMeasurementsNeeded As Integer    'V1.4.0

On Error GoTo EvaluateTestsError

EvaluateTests = False

For lintProgrammerNum = 1 To 2

    'Cycle 1 Step 1 requires 3 good measurements            'V1.4.0  Added If... Block
    If (CycleNum = 1) And (StepNum = 1) Then
        lintNumGoodMeasurementsNeeded = 3
    Else    'All other steps require 2 good measurements
        lintNumGoodMeasurementsNeeded = 2
    End If

    'Only attempt to make the solver calculations if we have at least lintNumGoodMeasurementsNeeded good sets of measurements   'V1.4.0
    If gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).NumGoodMeasurements >= lintNumGoodMeasurementsNeeded Then    'V1.4.0
        'Loop through all three sets of test codes
        For lintTestNum = 1 To 3
            If (CycleNum = 1) And (StepNum = 1) Then    'V1.4.0  Updated If statement to utilize new Function
                'If we're on Cycle 1, Step 1, then determine which RG to use and set new codes for the step 2 of the current cycle
                If InitSolverCalcs(lintProgrammerNum, CycleNum, StepNum, gudtSolver(lintProgrammerNum).CodeRatio(2, lintTestNum), gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).offset, gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).roughGain, gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).fineGain) Then  'V1.4.0 changed to use init sub
                    lblnGoodSolverCalculations = True
                Else
                    lblnGoodSolverCalculations = False
                    Exit For
                End If
            ElseIf (StepNum = 1) Then   'V1.4.0
                'If we're on Step 1,then set new codes for the step 2 of the current cycle
                If SolverCalculations(lintProgrammerNum, CycleNum, StepNum, gudtSolver(lintProgrammerNum).CodeRatio(2, lintTestNum), gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).offset, gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).roughGain, gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(2).Test(lintTestNum).fineGain) Then
                    lblnGoodSolverCalculations = True
                Else
                    lblnGoodSolverCalculations = False
                    Exit For
                End If
            Else
                'If we're on Step 2,then set new codes for the step 1 of the next cycle
                If SolverCalculations(lintProgrammerNum, CycleNum, StepNum, gudtSolver(lintProgrammerNum).CodeRatio(2, lintTestNum), gudtSolver(lintProgrammerNum).Cycle(CycleNum + 1).Step(1).Test(lintTestNum).offset, gudtSolver(lintProgrammerNum).Cycle(CycleNum + 1).Step(1).Test(lintTestNum).roughGain, gudtSolver(lintProgrammerNum).Cycle(CycleNum + 1).Step(1).Test(lintTestNum).fineGain) Then
                    lblnGoodSolverCalculations = True
                Else
                    lblnGoodSolverCalculations = False
                    Exit For
                End If
            End If
        Next lintTestNum
    Else
        'We don't have good solver calculations if we didn't try
        lblnGoodSolverCalculations = False
    End If

    'V1.4.0 'If we didn't get good solver calculations, we need to reset the Offset & Gain for the next try
    'If we didn't make good calculations, exit (bad part)
    If Not lblnGoodSolverCalculations Then
        EvaluateTests = False
        Exit Function
        'V1.4.0 'Gain was too high; reduce Rough Gain & retry
        'V1.4.0 lintOffset = gudtSolver(lintProgrammerNum).OffsetSeedCode
        'V1.4.0 lintRG = gudtSolver(lintProgrammerNum).RoughGainSeedCode - 1
        'V1.4.0 lintFG = gudtSolver(lintProgrammerNum).FineGainSeedCode
        'V1.4.0 'Set the new codes
        'V1.4.0 If StepNum = 1 Then
        'V1.4.0     'If we're on Step 1, set new codes for the step 2 of the current cycle
        'V1.4.0     Call SetTestCodesByRatio(lintProgrammerNum, CycleNum, 2, lintOffset, lintRG, lintFG, 1)
        'V1.4.0 ElseIf StepNum = 2 Then
        'V1.4.0     'If we're on Step 2, set new codes for the step 1 of the next cycle
        'V1.4.0     Call SetTestCodesByRatio(lintProgrammerNum, CycleNum + 1, 1, lintOffset, lintRG, lintFG, 1)
        'V1.4.0 End If
    End If
Next lintProgrammerNum

'If we made it to the end, there was no software error:
EvaluateTests = True

Exit Function
EvaluateTestsError:
    'Do nothing, exit sub with routine returning False

End Function

Public Function GoodMeasurements(ProgrammerNum As Integer, CycleNum As Integer, StepNum As Integer, TestNum As Integer, IndexNum As Integer) As Boolean
'
'   PURPOSE: To compare the measurement to the contants representing saturated
'            readings and return a boolean representing whether or not the measurement
'            was good, i.e. exhibited no saturation.
'
'  INPUT(S): ProgrammerNum = which programmer
'            CycleNum      = Cycle number, 1-2
'            StepNum       = Step number, 1-2
'            TestNum       = Test number, 1-3
'            IndexNum      = Index number, 1-2
' OUTPUT(S): Function returns a boolean representing whether or not the measurement
'            in question was good.

'Check for saturated output
If (gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(TestNum).MeasuredOutput(IndexNum) < gudtSolver(ProgrammerNum).LowSaturation) Or (gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(TestNum).MeasuredOutput(IndexNum) > gudtSolver(ProgrammerNum).HighSaturation) Then  'V1.5.0 used new measured values instead of constants
    GoodMeasurements = False
Else
    GoodMeasurements = True
End If

End Function

Public Function MakeSolverMeasurements(CycleNum As Integer, StepNum As Integer, IndexNum As Integer) As Boolean
'
'   PURPOSE: To load codes into TempRAM and make measurements accordingly,
'            repeating for three tests
'
'  INPUT(S): CycleNum = Current Cycle number of the Solver process
'            StepNum  = Current Step number of the Solver process
'            IndexNum = Current Index number, i.e. 1 = Idle, 2 = WOT
'
' OUTPUT(S): anomaly  = system status, 0 = ok, <> 0 = system error

Dim lintProgrammerNum As Integer
Dim lintTestNum As Integer
Dim lintNumOfTests As Integer

On Error GoTo SolverMeasurementsError

MakeSolverMeasurements = False

'V1.4.0\/\/\/
If CycleNum = 1 And StepNum = 1 Then
    lintNumOfTests = 4
Else
    lintNumOfTests = 3
End If
'V1.4.0/\/\/\

'Loop and program the test codes, measuring each time   'V1.4.0 was "three test codes"
For lintTestNum = 1 To lintNumOfTests                   'V1.4.0
    For lintProgrammerNum = 1 To 2

        'Set the Offset, Rough Gain, and Fine Gain for the current measurement
        gudtMLX90277(lintProgrammerNum).Write.offset = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).offset
        gudtMLX90277(lintProgrammerNum).Write.RGain = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).roughGain
        gudtMLX90277(lintProgrammerNum).Write.FGain = gudtSolver(lintProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).fineGain

        'Encode the new values into the boolean EEPROM array
        Call MLX90277.EncodeEEpromWrite(lintProgrammerNum)

    Next lintProgrammerNum

    'Write the new values to Temp RAM (Both IC's)
    Call MLX90277.WriteTempRAM

    'Delay 50 msec
    Call frmSolver90277.KillTime(50)

    'Measure the output caused by the test codes
    Call frmSolver90277.ReadSolverVoltages(gudtSolver(1).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(IndexNum), gudtSolver(2).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(IndexNum))
Next lintTestNum

'If we made it to the end, there was no software error:
MakeSolverMeasurements = True

Exit Function
SolverMeasurementsError:
    'Do nothing, exit sub with routine returning False

End Function

Public Sub SaveMLXtoFile()
'
'   PURPOSE: To save the MLX data to a comma delimited file
'
'  INPUT(S): none
' OUTPUT(S): none
' V1.6.0 new sub
' v1.7.0 fixed format

Dim lintFileNum As Integer
Dim lstrFileName As String
Dim lstrTemp As String

'Check SN
gstrSerialNumber = MLX90277.EncodePartID
lstrTemp = MLX90277.EncodePartID2

'Make the results file name
lstrFileName = gstrSerialNumber & " " & Format(Now, "mmddyy (HHMM ampm)") & " MLX Values.csv"
'Get a file
lintFileNum = FreeFile

Open PARTMLXDATAPATH + lstrFileName For Append As #lintFileNum
'Write data to file
Write #lintFileNum, "Parameter", "Output #1", "Output #2"
Write #lintFileNum, "Serial Number", gstrSerialNumber, lstrTemp
Write #lintFileNum, "Clamp Low", gudtMLX90277(1).Write.clampLow, gudtMLX90277(2).Write.clampLow
Write #lintFileNum, "Clamp High", gudtMLX90277(1).Write.clampHigh, gudtMLX90277(2).Write.clampHigh
Write #lintFileNum, "Offset", gudtMLX90277(1).Write.offset, gudtMLX90277(2).Write.offset
Write #lintFileNum, "AGND", gudtMLX90277(1).Write.AGND, gudtMLX90277(2).Write.AGND
Write #lintFileNum, "Rough Gain", gudtMLX90277(1).Write.RGain, gudtMLX90277(2).Write.RGain
Write #lintFileNum, "Fine Gain"; gudtMLX90277(1).Write.FGain, gudtMLX90277(2).Write.FGain
Write #lintFileNum, "Invert Slope", CStr(gudtMLX90277(1).Write.InvertSlope), CStr(gudtMLX90277(2).Write.InvertSlope)
Write #lintFileNum, "MLX Lock", CStr(gudtMLX90277(1).Write.MelexisLock), CStr(gudtMLX90277(2).Write.MelexisLock)
Write #lintFileNum, "Memory Lock", CStr(gudtMLX90277(1).Write.MemoryLock), CStr(gudtMLX90277(2).Write.MemoryLock)
Write #lintFileNum, "TCWin", gudtMLX90277(1).Write.TCWin, gudtMLX90277(2).Write.TCWin
Write #lintFileNum, "TC", gudtMLX90277(1).Write.TC, gudtMLX90277(2).Write.TC
Write #lintFileNum, "TC2nd", gudtMLX90277(1).Write.TC2nd, gudtMLX90277(2).Write.TC2nd
Write #lintFileNum, "Filter", gudtMLX90277(1).Write.Filter, gudtMLX90277(2).Write.Filter
Write #lintFileNum, "Cust. ID", gudtMLX90277(1).Write.CustID, gudtMLX90277(2).Write.CustID
Write #lintFileNum, "Fault Level", CStr(gudtMLX90277(1).Write.FaultLevel), CStr(gudtMLX90277(2).Write.FaultLevel)

'TC Table
If (MLX90277.VerifyMLXCRC(1) And MLX90277.VerifyCustomerCRC(1)) Then
    If (MLX90277.VerifyMLXCRC(2) And MLX90277.VerifyCustomerCRC(2)) Then
        Write #lintFileNum, "TC Table", "1", "1"
    Else
        Write #lintFileNum, "TC Table", "1", "0"
    End If
Else
    If (MLX90277.VerifyMLXCRC(2) And MLX90277.VerifyCustomerCRC(2)) Then
        Write #lintFileNum, "TC Table", "0", "1"
    Else
        Write #lintFileNum, "TC Table", "0", "0"
    End If
End If

Write #lintFileNum, "Mode", gudtMLX90277(1).Write.Mode, gudtMLX90277(2).Write.Mode
Write #lintFileNum, "Oscillator Adj.", gudtMLX90277(1).Read.FCKADJ, gudtMLX90277(2).Read.FCKADJ
Write #lintFileNum, "DAC Freq. Adj.", gudtMLX90277(1).Read.CKDACCH, gudtMLX90277(2).Read.CKDACCH
Write #lintFileNum, "Cap. Freq. Adj.", gudtMLX90277(1).Read.CKANACH, gudtMLX90277(2).Read.CKANACH
Write #lintFileNum, "Slow", CStr(gudtMLX90277(1).Write.SlowMode), CStr(gudtMLX90277(2).Write.SlowMode)
Write #lintFileNum, "SN ID X", gudtMLX90277(1).Write.x, gudtMLX90277(2).Write.x
Write #lintFileNum, "SN ID Y", gudtMLX90277(1).Write.Y, gudtMLX90277(2).Write.Y
Write #lintFileNum, "SN ID Wafer", gudtMLX90277(1).Write.Wafer, gudtMLX90277(2).Write.Wafer
Write #lintFileNum, "SN ID Lot", gudtMLX90277(1).Write.Lot, gudtMLX90277(2).Write.Lot

'Spacer row
Write #lintFileNum, " "

'Close the file
Close #lintFileNum

End Sub

'V1.4.0 Public Sub SetTestCodesByRatio(ProgrammerNum As Integer, CycleNum As Integer, StepNum As Integer, offset As Integer, RG As Integer, FG As Integer, RatioNum As Integer)
'V1.4.0 '
'V1.4.0 '   PURPOSE: To set the three code sets for next program & measure cycle.
'V1.4.0 '
'V1.4.0 '  INPUT(S): ProgrammerNum = Current Programmer number to set codes for
'V1.4.0 '            CycleNum      = Current Cycle number of the Solver process
'V1.4.0 '            StepNum       = Current Step number of the Solver process
'V1.4.0 '            Offset        = Offset to use for all three test codes
'V1.4.0 '            RG            = Rough Gain to use as a basis for ratio calculations
'V1.4.0 '            FG            = Fine Gain to use as a basis for ratio calculations
'V1.4.0 '            RatioNum      = Ratio to use for gain calculations
'V1.4.0 ' OUTPUT(S): None.
'V1.4.0
'V1.4.0 Dim lintTestNum As Integer
'V1.4.0
'V1.4.0 'Define the three sets of codes to use for the next step of the solver
'V1.4.0 For lintTestNum = 1 To 3
'V1.4.0     'Set RG and FG
'V1.4.0     Call CalculateGainCodes(ProgrammerNum, RG, FG, gudtSolver(ProgrammerNum).CodeRatio(RatioNum, lintTestNum), gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).roughGain, gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).fineGain)
'V1.4.0     'Set Offset
'V1.4.0     gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).offset = offset
'V1.4.0 Next lintTestNum
'V1.4.0
'V1.4.0 End Sub

Public Function SolverCalculations(ProgrammerNum As Integer, CycleNum As Integer, StepNum As Integer, SlopeRatio As Single, offset As Integer, RG As Integer, FG As Integer) As Boolean
'
'   PURPOSE: To calculate new Gain and Offset codes based on the three sets
'            of programmed codes and corresponding measurements.
'
'  INPUT(S): ProgrammerNum = Programmer (output) we're working with
'            CycleNum      = Current Cycle Number
'            StepNum       = Current Step Number
'            SlopeRatio    = Ratio of target slope to ideal slope
'
' OUTPUT(S): Offset = Calculated Offset Value
'            RG     = Calculated Rough Gain Value
'            FG     = Calculated Fine Gain Value
'            Function returns True or False, representing whether the routine
'            finished successfully.

Dim lintTestNum As Integer

Dim lsngXa As Single                'Measured position, Xa, for two best lines
Dim lsngXb As Single                'Measured position, Xb, for two best lines
Dim lsngYa(1 To 2) As Single        'Measured Y values at Xa for two best lines
Dim lsngYb(1 To 2) As Single        'Measured Y values at Xb for two best lines
Dim lsngM(1 To 2) As Single         'Slopes of two best lines
Dim lsngB(1 To 2) As Single         'Intercepts of two best lines

Dim lsngMIdeal As Single            'Ideal slope
Dim lsngXZeroGauss As Single        'X position of Zero Gauss point
Dim lsngDeltaYa As Single           'Amount to shift output at Xa from Ya(1)
Dim lsngDeltaOffset As Single       'Amount to shift from Ya(1) to Ideal Index 1

Dim lintSelectedMeasurement(1 To 2) As Integer

On Error GoTo SolverCalculationError

'Initialize variables
SolverCalculations = False  'Initialize to not completed successfully
lsngM(1) = 0
lsngM(2) = 0
lintSelectedMeasurement(1) = 0
lintSelectedMeasurement(2) = 0

'Calculate the the ideal slope, M(ideal):
lsngMIdeal = (gudtSolver(ProgrammerNum).Index(2).IdealValue - gudtSolver(ProgrammerNum).Index(1).IdealValue) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)

'The measurements for the current step were taken were taken at Xa and Xb:
lsngXa = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(1)
lsngXb = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(2)

'Skew that slope by the ratio requested of the routine
'(This creates a target slope that is SlopeRatio * the ideal slope)
lsngMIdeal = lsngMIdeal * SlopeRatio

'Loop through each of the test measurements for the current programmer,
'searching for the two slopes closest to the ideal slope
For lintTestNum = 1 To 3
    'Only attempt to sort if the codes were ok (not saturated)
    If gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(1) And gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasurementsOK(2) Then
        'Sort for the least difference between the ideal slope and the
        'measured slope, then the next least difference (bubble sort)
        'Define M(i), Ya(i), Yb(i), and B(i) for each of the two lines
        If Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope) < Abs(lsngMIdeal - lsngM(1)) Then
            DoEvents 'V1.6.0
            'Shift what were the closest values into the second slot
            lsngM(2) = lsngM(1)             'Define M(2)
            lsngYa(2) = lsngYa(1)           'Define Ya(2)
            lsngYb(2) = lsngYb(1)           'Define Yb(2)
            lsngB(2) = lsngB(1)             'Define B(2)
            'Save which measurement was selected for slot 2
            lintSelectedMeasurement(2) = lintSelectedMeasurement(1)
            'Shift the new closest values into the first slot
            lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope                                'Define M(1)
            lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(1)    'Define Ya(1)
            lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(2)    'Define Yb(1)
            lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedIntercept   'Define B(1)
            'Save which measurement was selected for slot 1
            lintSelectedMeasurement(1) = lintTestNum
        ElseIf Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope) < Abs(lsngMIdeal - lsngM(2)) Then
            DoEvents 'V1.6.0
            'Shift the new second closest values into the second spot
            lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedSlope                                'Define M(2)
            lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(1)    'Define Ya(2)
            lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).MeasuredOutput(2)    'Define Yb(2)
            lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintTestNum).CalculatedIntercept   'Define B(2)
            'Save which measurement was selected for slot 2
            lintSelectedMeasurement(2) = lintTestNum
        End If
    End If
Next lintTestNum

'We've determined that M(1) was the slope of the line that was closest to M(ideal),
'so use the ratio of M(1) to M(ideal) to calculate the best set of gain codes:
If Not CalculateGainCodes(ProgrammerNum, gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).roughGain, gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).fineGain, lsngMIdeal / lsngM(1), RG, FG) Then
    'Unsuccessful Calculation
    SolverCalculations = False
    Exit Function
End If

'   We have now sorted for the sets of measurements representing the two closest
'   lines to the ideal slope:
'
'   Y(1) = M(1)*X + B(1)    and
'   Y(2) = M(2)*X + B(2)
'
'   Y(1), M(1), & B(1) represent the line with the slope closest to ideal, while
'   Y(2), M(2), & B(2) represent the line with the next closest slope to ideal.
'   (Xa(i), Ya(i)) represent the measured X and Y values at position
'   one, while (Xb(i), Yb(i)) represent the measured X and Y values at position
'   two, for each line, (i).

'   Because we know that these lines intersect at the zero Gauss position, X(0G)
'   We can solve the two equations for Y(1) = Y(2):
'
'   M(1)*X(0G) + B(1) = M(2)*X(0G) + B(2)
'
'   Solving for X(0G), we show that:
'
'   X(0G) = (B(2) - B(1)) / (M(1) - M(2))

'Trap a divide-by-zero before attempting to calculate the Zero Gauss X-Position:
If (lsngM(1) - lsngM(2)) = 0 Then
    'Unsuccessful Calculation
    SolverCalculations = False
    Exit Function
End If

'Calculate the Zero Gauss X-Position:
lsngXZeroGauss = (lsngB(2) - lsngB(1)) / (lsngM(1) - lsngM(2))
gudtSolver(ProgrammerNum).ZeroGXPos(CycleNum, StepNum) = lsngXZeroGauss 'V1.9.4

'   The X-position that all lines interesect at allows us to calculate the
'   amount of shift the new gain will cause at position Xa (idle):
'                        /    /
'               M(1)--->/   /
'                      /  /
'                     / /<---M(ideal)
'                    //
'                   /
'                 //|
'               / / |          DeltaYa = (X(0G) - Xa) * (M(ideal) - M(1))
'             /  /  |
'           /   /   |
'         /    /    |
'        |    /     |
'        |   /      |
'DeltaYa{|  /       |
'        | /        |
'        |/         |
'        |----------|
'        Xa        X(0G)

'Calculate the amount of shift the new slope will cause at Xa from Ya(1):
lsngDeltaYa = (lsngXZeroGauss - lsngXa) * (lsngMIdeal - lsngM(1))

'Shifting by this amount, when using the new gain, will effectively shift Ya(newGain) back to
'where it was with line 1 (Ya(1)).

'Now, calculate the amount of shift that is necessary to get from the measured
'Ya(1) to the ideal output value at Xa (ideal index 1 output):
lsngDeltaOffset = gudtSolver(ProgrammerNum).Index(1).IdealValue - lsngYa(1)

'These shifts, together, are the amount that the index point (idle, Xa) will need
'to shift to work with the new gain codes (theoretically).  We calculate the
'new offset code, starting at the offset code from the line with the slope
'closest to ideal (line 1).  Added to this is the amount of shift needed, divided
'by the amount of offset change provided by each step of offset code change:
'New Offset Code = Offset(1) + ((DeltaYa + DeltaOffset) / OffsetStep)
offset = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).offset + ((lsngDeltaYa + lsngDeltaOffset) / gudtSolver(ProgrammerNum).OffsetStep)
If (offset < MINOFFSETCODE) Or (offset > MAXOFFSETCODE) Then
    'Unsuccessful Calculation
    SolverCalculations = False
    Exit Function
End If

'If we made it this far, the routine completed successfully
SolverCalculations = True

Exit Function

SolverCalculationError:

SolverCalculations = False

End Function

Public Function InitSolverCalcs(ProgrammerNum As Integer, CycleNum As Integer, StepNum As Integer, SlopeRatio As Single, offset As Integer, RG As Integer, FG As Integer) As Boolean
'
'   PURPOSE: To calculate new Gain and Offset codes based on the four sets
'            of programmed codes and corresponding measurements.
'
'  INPUT(S): ProgrammerNum = Programmer (output) we're working with
'            CycleNum      = Current Cycle Number
'            StepNum       = Current Step Number
'            SlopeRatio    = Ratio of target slope to ideal slope
'
' OUTPUT(S): Offset = Calculated Offset Value
'            RG     = Calculated Rough Gain Value
'            FG     = Calculated Fine Gain Value
'            Function returns True or False, representing whether the routine
'            finished successfully.
'V1.4.0  New Function

Dim lintTestNum As Integer

Dim lsngXa As Single                'Measured position, Xa, for two best lines
Dim lsngXb As Single                'Measured position, Xb, for two best lines
Dim lsngYa(1 To 2) As Single        'Measured Y values at Xa for two best lines
Dim lsngYb(1 To 2) As Single        'Measured Y values at Xb for two best lines
Dim lsngM(1 To 2) As Single         'Slopes of two best lines
Dim lsngB(1 To 2) As Single         'Intercepts of two best lines

Dim lsngMIdeal As Single            'Ideal slope
Dim lsngXZeroGauss As Single        'X position of Zero Gauss point
Dim lsngDeltaYa As Single           'Amount to shift output at Xa from Ya(1)
Dim lsngDeltaOffset As Single       'Amount to shift from Ya(1) to Ideal Index 1

Dim lintSelectedMeasurement(1 To 2) As Integer

Dim lblnHigherRGNotSaturated As Boolean    'V1.4.0

On Error GoTo InitSolverCalcsError

'Initialize variables
InitSolverCalcs = False  'Initialize to not completed successfully
lsngM(1) = 0
lsngM(2) = 0
lintSelectedMeasurement(1) = 0
lintSelectedMeasurement(2) = 0

'Calculate the the ideal slope, M(ideal):
lsngMIdeal = (gudtSolver(ProgrammerNum).Index(2).IdealValue - gudtSolver(ProgrammerNum).Index(1).IdealValue) / (gudtSolver(ProgrammerNum).Index(2).IdealLocation - gudtSolver(ProgrammerNum).Index(1).IdealLocation)

'The measurements for the current step were taken were taken at Xa and Xb:
lsngXa = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(1)
lsngXb = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).MeasuredLocation(2)

'Check if the highest gain test (test #2) was saturated
lblnHigherRGNotSaturated = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasurementsOK(1) And gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasurementsOK(2)

'Choose between the higher and lower RG if the higher RG was not saturated
If lblnHigherRGNotSaturated Then
    'MsgBox "Calculated Slope " & CStr(Format(gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedSlope, "###.###")) & " <  Ideal Slope " & CStr(Format(lsngMIdeal, "###.###")) & "???"
    'If test 1 is less than the ideal slope, use tests 1 and 2
    If gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedSlope < lsngMIdeal Then
        DoEvents  'Must have this statement for above if to work correctly (WIERD COMPILER ISSUE)
        'Check which point is closer test 1 or test 2
        If Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedSlope) < Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).CalculatedSlope) Then
            lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedSlope       'Define M(1)
            lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).MeasuredOutput(1)    'Define Ya(1)
            lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).MeasuredOutput(2)    'Define Yb(1)
            lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedIntercept   'Define B(1)
            lintSelectedMeasurement(1) = 1                                                                   'Save which measurement was selected for slot 1
            lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).CalculatedSlope       'Define M(2)
            lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasuredOutput(1)    'Define Ya(2)
            lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasuredOutput(2)    'Define Yb(2)
            lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).CalculatedIntercept   'Define B(2)
            lintSelectedMeasurement(2) = 2                                                                   'Save which measurement was selected for slot 2
        Else
            lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).CalculatedSlope       'Define M(1)
            lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasuredOutput(1)    'Define Ya(1)
            lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).MeasuredOutput(2)    'Define Yb(1)
            lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(2).CalculatedIntercept   'Define B(1)
            lintSelectedMeasurement(1) = 2                                                                   'Save which measurement was selected for slot 1
            lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedSlope       'Define M(2)
            lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).MeasuredOutput(1)    'Define Ya(2)
            lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).MeasuredOutput(2)    'Define Yb(2)
            lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(1).CalculatedIntercept   'Define B(2)
            lintSelectedMeasurement(2) = 1                                                                   'Save which measurement was selected for slot 2
        End If
    Else    'If test 1 is greater than the ideal slope, use tests 3 and 4
        DoEvents  'Must have this statement for above if to work correctly (WIERD COMPILER ISSUE)
        'Check which point is closer test 3 or test 4
        If Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope) < Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope) Then
            lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope       'Define M(1)
            lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(1)    'Define Ya(1)
            lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(2)    'Define Yb(1)
            lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedIntercept   'Define B(1)
            lintSelectedMeasurement(1) = 3                                                                   'Save which measurement was selected for slot 1
            lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope       'Define M(2)
            lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(1)    'Define Ya(2)
            lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(2)    'Define Yb(2)
            lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedIntercept   'Define B(2)
            lintSelectedMeasurement(2) = 4                                                                   'Save which measurement was selected for slot 2
        Else
            lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope       'Define M(1)
            lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(1)    'Define Ya(1)
            lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(2)    'Define Yb(1)
            lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedIntercept   'Define B(1)
            lintSelectedMeasurement(1) = 4                                                                   'Save which measurement was selected for slot 1
            lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope       'Define M(2)
            lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(1)    'Define Ya(2)
            lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(2)    'Define Yb(2)
            lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedIntercept   'Define B(2)
            lintSelectedMeasurement(2) = 3                                                                   'Save which measurement was selected for slot 2
        End If
    End If
Else    'Choose the lower RG
    'ANM 'Verify that the lower RG has enough gain (last test is the higher gain of the two lower RG tests)
    'ANM If gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope > lsngMIdeal Then
    'Check which point is closer test 3 or test 4
    If Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope) < Abs(lsngMIdeal - gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope) Then
        lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope       'Define M(1)
        lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(1)    'Define Ya(1)
        lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(2)    'Define Yb(1)
        lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedIntercept   'Define B(1)
        lintSelectedMeasurement(1) = 3                                                                   'Save which measurement was selected for slot 1
        lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope       'Define M(2)
        lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(1)    'Define Ya(2)
        lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(2)    'Define Yb(2)
        lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedIntercept   'Define B(2)
        lintSelectedMeasurement(2) = 4                                                                   'Save which measurement was selected for slot 2
    Else
        lsngM(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedSlope       'Define M(1)
        lsngYa(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(1)    'Define Ya(1)
        lsngYb(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).MeasuredOutput(2)    'Define Yb(1)
        lsngB(1) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(4).CalculatedIntercept   'Define B(1)
        lintSelectedMeasurement(1) = 4                                                                   'Save which measurement was selected for slot 1
        lsngM(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedSlope       'Define M(2)
        lsngYa(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(1)    'Define Ya(2)
        lsngYb(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).MeasuredOutput(2)    'Define Yb(2)
        lsngB(2) = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(3).CalculatedIntercept   'Define B(2)
        lintSelectedMeasurement(2) = 3                                                                   'Save which measurement was selected for slot 2
    End If
    'ANM Else
    'ANM     'Unsuccessful Calculation
    'ANM     InitSolverCalcs = False
    'ANM     Exit Function
    'ANM End If
End If

'Skew that slope by the ratio requested of the routine
'(This creates a target slope that is SlopeRatio * the ideal slope)
'NOTE: Do this AFTER choosing the RG code!!!
lsngMIdeal = lsngMIdeal * SlopeRatio

'We've determined that M(1) was the slope of the line that was closest to M(ideal),
'so use the ratio of M(1) to M(ideal) to calculate the best set of gain codes:
If Not CalculateGainCodes(ProgrammerNum, gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).roughGain, gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).fineGain, lsngMIdeal / lsngM(1), RG, FG) Then
    'Unsuccessful Calculation
    InitSolverCalcs = False
    Exit Function
End If

'   We have now sorted for the sets of measurements representing the two closest
'   lines to the ideal slope:
'
'   Y(1) = M(1)*X + B(1)    and
'   Y(2) = M(2)*X + B(2)
'
'   Y(1), M(1), & B(1) represent the line with the slope closest to ideal, while
'   Y(2), M(2), & B(2) represent the line with the next closest slope to ideal.
'   (Xa(i), Ya(i)) represent the measured X and Y values at position
'   one, while (Xb(i), Yb(i)) represent the measured X and Y values at position
'   two, for each line, (i).

'   Because we know that these lines intersect at the zero Gauss position, X(0G)
'   We can solve the two equations for Y(1) = Y(2):
'
'   M(1)*X(0G) + B(1) = M(2)*X(0G) + B(2)
'
'   Solving for X(0G), we show that:
'
'   X(0G) = (B(2) - B(1)) / (M(1) - M(2))

'Trap a divide-by-zero before attempting to calculate the Zero Gauss X-Position:
If (lsngM(1) - lsngM(2)) = 0 Then
    'Unsuccessful Calculation
    InitSolverCalcs = False
    Exit Function
End If

'Calculate the Zero Gauss X-Position:
lsngXZeroGauss = (lsngB(2) - lsngB(1)) / (lsngM(1) - lsngM(2))
gudtSolver(ProgrammerNum).ZeroGXPos(CycleNum, StepNum) = lsngXZeroGauss 'V1.9.4

'   The X-position that all lines interesect at allows us to calculate the
'   amount of shift the new gain will cause at position Xa (idle):
'                        /    /
'               M(1)--->/   /
'                      /  /
'                     / /<---M(ideal)
'                    //
'                   /
'                 //|
'               / / |          DeltaYa = (X(0G) - Xa) * (M(ideal) - M(1))
'             /  /  |
'           /   /   |
'         /    /    |
'        |    /     |
'        |   /      |
'DeltaYa{|  /       |
'        | /        |
'        |/         |
'        |----------|
'        Xa        X(0G)

'Calculate the amount of shift the new slope will cause at Xa from Ya(1):
lsngDeltaYa = (lsngXZeroGauss - lsngXa) * (lsngMIdeal - lsngM(1))

'Shifting by this amount, when using the new gain, will effectively shift Ya(newGain) back to
'where it was with line 1 (Ya(1)).

'Now, calculate the amount of shift that is necessary to get from the measured
'Ya(1) to the ideal output value at Xa (ideal index 1 output):
lsngDeltaOffset = gudtSolver(ProgrammerNum).Index(1).IdealValue - lsngYa(1)

'These shifts, together, are the amount that the index point (idle, Xa) will need
'to shift to work with the new gain codes (theoretically).  We calculate the
'new offset code, starting at the offset code from the line with the slope
'closest to ideal (line 1).  Added to this is the amount of shift needed, divided
'by the amount of offset change provided by each step of offset code change:
'New Offset Code = Offset(1) + ((DeltaYa + DeltaOffset) / OffsetStep)
offset = gudtSolver(ProgrammerNum).Cycle(CycleNum).Step(StepNum).Test(lintSelectedMeasurement(1)).offset + ((lsngDeltaYa + lsngDeltaOffset) / gudtSolver(ProgrammerNum).OffsetStep)
If (offset < MINOFFSETCODE) Or (offset > MAXOFFSETCODE) Then
    'Unsuccessful Calculation
    InitSolverCalcs = False
    Exit Function
End If

'If we made it this far, the routine completed successfully
InitSolverCalcs = True

Exit Function

InitSolverCalcsError:

InitSolverCalcs = False

End Function

Public Sub SolverInitialization()
'
'   PURPOSE: To determine the initial best-guess offset and gain codes.  Also
'            eliminate clamping.
'
'  INPUT(S): None.
' OUTPUT(S): None.
'
Dim lintProgrammerNum As Integer
Dim lintCycleNum As Integer
Dim lintStepNum As Integer

'Zero all the process variables
Call ClearSolverVariables

'Set up codes for both programmers
For lintProgrammerNum = 1 To 2

    'Eliminate Clamping
    gudtMLX90277(lintProgrammerNum).Write.clampHigh = MAXCLAMPCODE
    gudtMLX90277(lintProgrammerNum).Write.clampLow = MINCLAMPCODE

    'Set other MLX variables
    gudtMLX90277(lintProgrammerNum).Write.Filter = gudtSolver(lintProgrammerNum).Filter
    gudtMLX90277(lintProgrammerNum).Write.InvertSlope = gudtSolver(lintProgrammerNum).InvertSlope
    gudtMLX90277(lintProgrammerNum).Write.Mode = gudtSolver(lintProgrammerNum).Mode
    gudtMLX90277(lintProgrammerNum).Write.FaultLevel = gudtSolver(lintProgrammerNum).FaultLevel

    'Reset CountOK variables
    For lintCycleNum = 1 To 2
        For lintStepNum = 1 To 2
            gudtSolver(lintProgrammerNum).Cycle(lintCycleNum).Step(lintStepNum).NumGoodMeasurements = 0
        Next lintStepNum
    Next lintCycleNum

    'V1.4.0\/\/\/
    'During the first step of the first cycle, we should determine which Rough Gain code
    'is appropriate.  Do this by performing four test measurements with four different
    'gains:

    ' TestNum     RG     FG
    '  (1)      MaxRG   MinFG
    '  (2)      MaxRG   MidFG
    '  (3)      MinRG   MidFG
    '  (4)      MinRG   MaxFG

    'This combination should provide  enough information to properly characterize the two
    'different Rough Gain settings that might be appropriate for the part (MinRG and MaxRG)
    'MaxRG and MinRG are set in the parameter file.  MinFG, MidFG, and MaxFG are constants
    
    'Test 1
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(1).roughGain = gudtSolver(lintProgrammerNum).MaxRG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(1).fineGain = gudtSolver(lintProgrammerNum).HighRGLowFG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(1).offset = gudtSolver(lintProgrammerNum).InitialOffset
    'Test 2
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(2).roughGain = gudtSolver(lintProgrammerNum).MaxRG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(2).fineGain = gudtSolver(lintProgrammerNum).HighRGHighFG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(2).offset = gudtSolver(lintProgrammerNum).InitialOffset
    'Test 3
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(3).roughGain = gudtSolver(lintProgrammerNum).MinRG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(3).fineGain = gudtSolver(lintProgrammerNum).LowRGLowFG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(3).offset = gudtSolver(lintProgrammerNum).InitialOffset
    'Test 4
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(4).roughGain = gudtSolver(lintProgrammerNum).MinRG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(4).fineGain = gudtSolver(lintProgrammerNum).LowRGHighFG
    gudtSolver(lintProgrammerNum).Cycle(1).Step(1).Test(4).offset = gudtSolver(lintProgrammerNum).InitialOffset

    'V1.4.0/\/\/\

    'V1.4.0 Call CalcSeedCodes(lintProgrammerNum)
    'V1.4.0
    'V1.4.0 'Set the test codes for Cycle 1, Step 1 based on the initial best-guess codes
    'V1.4.0 Call SetTestCodesByRatio(lintProgrammerNum, 1, 1, gudtSolver(lintProgrammerNum).OffsetSeedCode, gudtSolver(lintProgrammerNum).RoughGainSeedCode, gudtSolver(lintProgrammerNum).FineGainSeedCode, 1)

Next lintProgrammerNum

End Sub

'V1.4.0 Public Sub UpdateHistoryCodes()
'V1.4.0 '
'V1.4.0 '   PURPOSE: To update the history code arrays for calculation of new seed codes
'V1.4.0 '            based on good programming results.
'V1.4.0 '
'V1.4.0 '  INPUT(S): None.
'V1.4.0 ' OUTPUT(S): None.
'V1.4.0 '
'V1.4.0 Dim lintProgrammerNum As Integer
'V1.4.0
'V1.4.0 'Update the history array that will define Seed Codes
'V1.4.0 For lintProgrammerNum = 1 To 2
'V1.4.0     gudtSolver(lintProgrammerNum).OffsetHistory(gudtSolver(lintProgrammerNum).NextHistoryCode) = gudtSolver(lintProgrammerNum).FinalOffsetCode
'V1.4.0     gudtSolver(lintProgrammerNum).RGHistory(gudtSolver(lintProgrammerNum).NextHistoryCode) = gudtSolver(lintProgrammerNum).FinalRGCode
'V1.4.0     gudtSolver(lintProgrammerNum).FGHistory(gudtSolver(lintProgrammerNum).NextHistoryCode) = gudtSolver(lintProgrammerNum).FinalFGCode
'V1.4.0     'Increment to find the next history code number
'V1.4.0     gudtSolver(lintProgrammerNum).NextHistoryCode = gudtSolver(lintProgrammerNum).NextHistoryCode + 1
'V1.4.0     'If it's beyond the last allowable code number, rollover to 1
'V1.4.0     If gudtSolver(lintProgrammerNum).NextHistoryCode > MAXHISTORYNUM Then
'V1.4.0         gudtSolver(lintProgrammerNum).NextHistoryCode = 1
'V1.4.0     End If
'V1.4.0     'If we haven't already filled the history code array, increment the
'V1.4.0     'variable tracking the number of codes in the array
'V1.4.0     If gudtSolver(lintProgrammerNum).NumHistoryCodes < MAXHISTORYNUM Then
'V1.4.0         gudtSolver(lintProgrammerNum).NumHistoryCodes = gudtSolver(lintProgrammerNum).NumHistoryCodes + 1
'V1.4.0     End If
'V1.4.0 Next lintProgrammerNum
'V1.4.0
'V1.4.0 End Sub

