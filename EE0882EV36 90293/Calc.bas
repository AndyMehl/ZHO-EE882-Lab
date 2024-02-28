Attribute VB_Name = "Calc"
Option Explicit

Public Sub CalcControlLimits(nominal As Single, minLimit As Single, maxLimit As Single, minControlLimit As Single, maxControlLimit As Single)
'
'   PURPOSE:    To calculate high and low control limits based on the
'               specified limits for a parameter.
'
'  INPUT(S):    nominal        : Specified nominal for parameter
'               minLimit       : Specified minimum limit for parameter
'               maxLimit       : Specified maximum limit for parameter
'
' OUTPUT(S):    minControlLimit: Calculated min control limit for parameter
'               maxControlLimit: Calculated max control limit for parameter
'

Dim lsngControlBand As Single

On Error GoTo calc_Err

'The control band is based on the span between the limit value and the nominal
lsngControlBand = (maxLimit - nominal) * 2
'Calculate the high control limit
maxControlLimit = nominal + lsngControlBand
'The control band is based on the span between the limit value and the nominal
lsngControlBand = (nominal - minLimit) * 2
'Calculate the high control limit
minControlLimit = nominal - lsngControlBand

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcControlLimits: " & Err.Description, True, True)
End Sub

Public Sub CalcHysteresis(forwardArray() As Single, reverseArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, calcDataArray() As Single)
'
'   PURPOSE:    To calculate the hysteresis.  The hysteresis is the based on the
'               difference between the reverse data array and the forward
'               data array.
'
'  INPUT(S):    forwardArray    : Scaled forward data array for the current channel
'               reverseArray    : Scaled reverse data array for the current channel
'               evaluateStart   : Start point of data for current channel
'               evaluateStop    : End point of data for current channel
'               resolution      : Scan resolution
'
' OUTPUT(S):    calcDataArray : Output array of hysteresis values.
'

Dim i As Integer

On Error GoTo calc_Err

'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Hysteresis array
For i = evaluateStart To evaluateStop
    calcDataArray(i) = (reverseArray(i) - forwardArray(i))
Next i

Exit Sub
calc_Err:

    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcHysteresis: " & Err.Description, True, True)
End Sub

Public Sub CalcIndexCor(ByVal indexActual1 As Single, ByVal indexActual2 As Single, ByVal corCoefficient As Single, indexCor12 As Single)
'
'   PURPOSE:    To calculate the Index Correlation.
'
'  INPUT(S):    indexActual1    : Actual index value for output 1
'               indexActual2    : Actual index value for output 2
'               corCoefficient  : mathematical multiplication factor for correlation
'
' OUTPUT(S):    indexCor12      : Index Correlation of Output 1 & 2
'

On Error GoTo calc_Err
  
'Calculate the Index Correlation
'   scaled output 1 index = indexActual1,
'   scaled output 2 index = correlationFactor * indexActual2
'   index corellation = scaled output 1 index - scaled output 2 index
indexCor12 = Abs(indexActual1 - (corCoefficient * indexActual2))
    
Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcIndexCor: " & Err.Description, True, True)
End Sub

Public Sub CalcKneeLoc(DataArray() As Single, ByVal TransitionSlope As Single, PositiveTransition As Boolean, ByVal TransitionPercentage As Single, ByVal TransitionWindow As Single, ByVal StartLoc As Single, ByVal StopLoc As Single, ByVal resolution As Single, KneeLoc As Single, KneeVal As Single, KneeFound As Boolean)
'
'   PURPOSE:    To calculate the rising point.
'
'
'  INPUT(S):    DataArray            : Scaled data array, binary counts
'               TransitionSlope      : Slope after which the knee is considered found
'               PositiveTransition   : True if the routine should look for a slope greater than the TransitionSlope; False if the routine should look for a slope less than the TransitionSlope
'               TransitionPercentage : Percentage of point-to-point checks that must have a slope greater than (less than) than the TransitionSlope
'               TransitionWindow     : Length of area to check for TransitionPercentage within
'               StartLoc             : Location to start looking for knee
'               StopLoc              : Location to stop looking for knee
'               Resolution           : Scan resolution
'
' OUTPUT(S):    KneeLoc              : Location of Knee
'               KneeVal              : DataArray value at KneeLoc
'
'      NOTE:    This routine looks for a "knee" or a transition within a set
'               of data by looking for for the first point-to-point slope
'               greater than (less than) TransitionSlope.  Once this is
'               found, the data from there over the TransitionLength is
'               evaluated.  If at least TransitionPercentage of the point-to-
'               point slope checks are greater than (less than) TransitionSlope,
'               then the second point (first point) from that first point-to-point
'               check is considered to be the "knee" location.  If not, the
'               algorithm searches for the next point-to-point slope greater than
'               (less than) TransitionSlope, and again evaluates as before.  This
'               is continued until a knee is found.  Note that "greater than" is
'               used in this description, which assumes the variable
'               PositiveTransition is True.  If it is False, read the
'               description as "less than."
'

On Error GoTo calc_Err
    
Dim i As Integer
Dim j As Integer
Dim lintStart As Integer
Dim lintStop As Integer
Dim lsngPointToPointSlope As Single
Dim lintNumPointsInTransitionLength As Integer
Dim lintNumBadPoints As Integer
Dim lsngPercentage As Single
Dim lsngScaledTransitionSlope As Single

'Initialize the variables to zeroes
KneeLoc = 0
KneeVal = 0

'Initialize the boolean to not found
KneeFound = False

'Calculate the number of points in the Transition length
lintNumPointsInTransitionLength = CInt(TransitionWindow * resolution)

'Define the Start & Stop locations in terms of array locations
lintStart = StartLoc * resolution
lintStop = StopLoc * resolution

'Loop through the array
For i = lintStart To lintStop

    'Calculate the point-to-point slope
    lsngPointToPointSlope = (DataArray(i + 1) - DataArray(i)) * resolution

    'Check for slope greater than if PositiveTransition is true
    If PositiveTransition Then
        'Is the calculated slope greater than the Transition Slope?
        If lsngPointToPointSlope >= TransitionSlope Then
            'Initialize the bad point counter
            lintNumBadPoints = 0
            For j = (i + 1) To (i + lintNumPointsInTransitionLength)
                'Calculate the point-to-point slope
                lsngPointToPointSlope = (DataArray(j + 1) - DataArray(j)) * resolution
                'Count the number of points with slope less than the Transition Slope
                If lsngPointToPointSlope < TransitionSlope Then
                    lintNumBadPoints = lintNumBadPoints + 1
                End If
            Next j
            'Calculate the percentage of good points
            lsngPercentage = ((lintNumPointsInTransitionLength - lintNumBadPoints) / lintNumPointsInTransitionLength) * 100
            'Is this beyond the Transition percentage?
            If lsngPercentage > TransitionPercentage Then
                'The knee is the second point of the first increased point-to-point slope
                KneeLoc = (i + 1) / resolution
                'Return the value at the knee
                KneeVal = DataArray(i + 1)
                'Routine complete
                KneeFound = True
                Exit For
            End If
        End If
    Else
        'Is the calculated slope less than the Transition Slope?
        If lsngPointToPointSlope <= TransitionSlope Then
            'Initialize the bad point counter
            lintNumBadPoints = 0
            For j = (i + 1) To (i + lintNumPointsInTransitionLength)
                'Calculate the point-to-point slope
                lsngPointToPointSlope = (DataArray(j + 1) - DataArray(j)) * resolution
                'Count the number of points with slope greater than the Transition Slope
                If lsngPointToPointSlope > TransitionSlope Then
                    lintNumBadPoints = lintNumBadPoints + 1
                End If
            Next j
            'Calculate the percentage of good points
            lsngPercentage = ((lintNumPointsInTransitionLength - lintNumBadPoints) / lintNumPointsInTransitionLength) * 100
            'Is this beyond the Transition percentage?
            If lsngPercentage > TransitionPercentage Then
                'The knee is the first point of the first decreased point-to-point slope
                KneeLoc = i / resolution
                'Return the value at the knee
                KneeVal = DataArray(i)
                'Routine complete
                KneeFound = True
                Exit For
            End If
        End If
    End If
Next i

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcKneeLoc: " & Err.Description, True, True)
End Sub

Public Sub calcLimitArray(ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal m As Single, ByVal b As Single, ByVal resolution As Single, calcLimitArray() As Single)
'
'   PURPOSE:    To calculate a single limit array based on the equation
'               of a line.
'
'     NOTES:    The limit array is calculated using the equation of a
'               line:  y(x) = (m * x) + B, where:
'
'               y(x) =  limit array value @ point x
'                 m  =  slope of limit line
'                 x  =  location of each data point
'                 B  =  specified limit of parameter
'
'
'  INPUT(S):    evaluateStart : Start point of data for current channel
'               evaluateStop  : End point of data for current channel
'               m             : Slope of the limit line
'               B             : Specified limit of region (Y-intercept)
'               resolution    : Scan resolution
'
' OUTPUT(S):    calcLimitArray: Calculated limit array
'

Dim i As Integer

On Error GoTo calc_Err

'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
m = m / resolution

For i = evaluateStart To evaluateStop
    calcLimitArray(i) = (m * i) + b
Next i

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLimitArray: " & Err.Description, True, True)
End Sub

Public Sub CalcLinAbsolute(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, indexIdeal As Single, ByVal indexLoc As Single, ByVal mIdeal As Single, ByVal resolution As Single, calcDataArray() As Single)
'
'   PURPOSE:    To calculate the absolute linearity deviation.  The absolute
'               linearity is based on the ideal index output and the ideal slope
'               of the part.
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
'  INPUT(S):    scaledDataArray : Scaled data array
'               evaluateStart   : Start point of data
'               evaluateStop    : End point of data
'               indexIdeal      : Ideal index value
'               indexLoc        : Index location
'               mIdeal          : Ideal slope
'               resolution      : Scan resolution
'
' OUTPUT(S):    calcDataArray   : Output array of absolute linearity deviation
'                                 values.
'

Dim lintPointNum As Integer
Dim lsngIdealOutput As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
indexLoc = indexLoc * resolution
mIdeal = mIdeal / resolution

'Calculate the Absolute Linearity Deviation array
For lintPointNum = evaluateStart To evaluateStop
    lsngIdealOutput = indexIdeal + (mIdeal * (lintPointNum - indexLoc))
    calcDataArray(lintPointNum) = (ScaledDataArray(lintPointNum)) - lsngIdealOutput
Next

Exit Sub

calc_Err:
    
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLinAbsolute: " & Err.Description, True, True)
End Sub

Public Sub CalcLinAbsoluteSegment(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, indexIdeal As Single, ByVal indexLoc As Single, ByVal mIdeal As Single, ByVal resolution As Single, calcDataArray() As Single)
'   PURPOSE:    To calculate the absolute linearity deviation.  The absolute
'               linearity is based on the ideal index output and the ideal slope
'               of the part.
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
'  INPUT(S):    scaledDataArray : Scaled data array
'               evaluateStart   : Start point of data
'               evaluateStop    : End point of data
'               indexIdeal      : Ideal index value
'               indexLoc        : Index location
'               mIdeal          : Ideal slope
'               resolution      : Scan resolution
'
' OUTPUT(S):    calcDataArray   : Output array of absolute linearity deviation
'                                 values.
'

Dim i As Integer
Dim lintPointNum As Integer
Dim lsngIdealOutput As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
indexLoc = indexLoc * resolution
mIdeal = mIdeal / resolution

'Calculate the Absolute Linearity Deviation array
For i = evaluateStart To evaluateStop
    lsngIdealOutput = indexIdeal + (mIdeal * (i - indexLoc))
    calcDataArray(i) = (ScaledDataArray(i)) - lsngIdealOutput
Next

Exit Sub

calc_Err:
    
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLinAbsoluteSegment: " & Err.Description, True, True)
End Sub

Public Sub CalcLinSinglePoint(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, indexMeasured As Single, ByVal indexLoc As Single, ByVal mIdeal As Single, ByVal resolution As Single, calcDataArray() As Single)
'
'   PURPOSE:    To calculate the single-point linearity deviation.  The
'               single-point linearity is based on the measured index
'               output and the ideal slope of the part.
'
'     NOTES:    For the linear portion of the curve, the ideal value follows
'               the equation of a line:  y(x) = m * (x - n) + b, where:
'
'               y(x) =  measured value @ point x
'                 m  =  ideal slope
'                 x  =  location of ideal value point
'                 n  =  location of index point
'                 b  =  output at index point
'
'               The linearity checks are typically performed on only the forward data.
'
'  INPUT(S):    scaledDataArray : Scaled data array
'               evaluateStart   : Start point of data
'               evaluateStop    : End point of data
'               indexMeasured   : Measured index value
'               indexLoc        : Index location
'               mIdeal          : Ideal slope
'               resolution      : Scan resolution
'
' OUTPUT(S):    calcDataArray   : Output array of SinglePoint linearity deviation
'                                 values.
'

Dim lintPointNum As Integer
Dim lsngIdealOutput As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
indexLoc = indexLoc * resolution
mIdeal = mIdeal / resolution

'Calculate the SinglePoint Linearity Deviation array
For lintPointNum = evaluateStart To evaluateStop
    lsngIdealOutput = indexMeasured + (mIdeal * (lintPointNum - indexLoc))
    calcDataArray(lintPointNum) = (ScaledDataArray(lintPointNum)) - lsngIdealOutput
Next

Exit Sub

calc_Err:
    
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLinSinglePoint: " & Err.Description, True, True)
End Sub

Public Sub CalcLinSinglePointwBend(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal indexMeasured As Single, ByVal indexLoc As Single, ByVal mIdealA As Single, ByVal mIdealB As Single, ByVal resolution As Single, calcDataArray() As Single)
'
'   PURPOSE:    To calculate the single-point linearity deviation.  The
'               single-point linearity is based on the measured index
'               output and the ideal slope of the part.
'
'     NOTES:    For the linear portion of the curve, the ideal value follows
'               the equation of a line:  y(x) = m * (x - n) + b, where:
'
'               y(x) =  measured value @ point x
'                 m  =  ideal slope
'                 x  =  location of ideal value point
'                 n  =  location of index point
'                 b  =  output at index point
'
'               The linearity checks are typically performed on only the forward data.
'
'  INPUT(S):    scaledDataArray : Scaled data array
'               evaluateStart   : Start point of data
'               evaluateStop    : End point of data
'               indexMeasured   : Measured index value
'               indexLoc        : Index location
'               mIdealA         : First Ideal Slope
'               mIdealB         : Second Ideal Slope
'               resolution      : Scan resolution
'
' OUTPUT(S):    calcDataArray   : Output array of SinglePoint linearity deviation
'                                 values.
'
'1.1ANM new sub

Dim lintPointNum As Integer
Dim lsngIdealOutput As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
indexLoc = indexLoc * resolution
mIdealB = mIdealB / resolution
mIdealA = mIdealA / resolution

'Calculate the SinglePoint Linearity Deviation array
For lintPointNum = evaluateStart To (gudtTest(CHAN0).slope.start * resolution)
    lsngIdealOutput = indexMeasured + (mIdealA * (lintPointNum - indexLoc))
    calcDataArray(lintPointNum) = (ScaledDataArray(lintPointNum)) - lsngIdealOutput
Next

indexMeasured = lsngIdealOutput - (mIdealB * ((lintPointNum - 1) - indexLoc)) '3.6iANM

'Calculate the SinglePoint Linearity Deviation array
For lintPointNum = ((gudtTest(CHAN0).slope.start * resolution) + 1) To evaluateStop
    lsngIdealOutput = indexMeasured + (mIdealB * (lintPointNum - indexLoc))
    calcDataArray(lintPointNum) = (ScaledDataArray(lintPointNum)) - lsngIdealOutput
Next

Exit Sub

calc_Err:
    
    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLinSinglePointwBend: " & Err.Description, True, True)
End Sub

Public Sub CalcLinTwoPoint(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal index1Val As Single, ByVal index1Loc As Single, ByVal index2Val As Single, ByVal index2Loc As Single, ByVal resolution As Single, IdealSlope As Single, calcDataArray() As Single)
'
'   PURPOSE:    To calculate two-point linearity deviation.  The two-point
'               linearity is based on the two points passed in to the routine
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
'               The ideal line is defined by the two reference points passed into the routine.
'               The linearity checks are typically performed on only the forward data.
'
'  INPUT(S):    scaledDataArray : Scaled data array
'               evaluateStart   : Start point of data
'               evaluateStop    : End point of data
'               index1Val       : Index one value
'               index1Loc       : Index one location
'               index2Val       : Index two value
'               index2Loc       : Index two location
'               resolution      : Scan resolution
'
' OUTPUT(S):    IdealSlope      : Calculated ideal slope
'               calcDataArray   : Output array of absolute linearity deviation
'                                 values.
'

Dim lintPointNum As Integer
Dim lsngIdealM As Single
Dim lsngIdealB As Single
Dim lsngIdealOutput As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
index1Loc = index1Loc * resolution
index2Loc = index2Loc * resolution

'Define the ideal slope (in %/data resolution)
lsngIdealM = (index2Val - index1Val) / (index2Loc - index1Loc)
'Define the ideal Y-Intercept
lsngIdealB = index1Val - (index1Loc * lsngIdealM)

'Calculate the Absolute Linearity Deviation Array
For lintPointNum = evaluateStart To evaluateStop
    lsngIdealOutput = (lsngIdealM * lintPointNum) + lsngIdealB
    calcDataArray(lintPointNum) = (ScaledDataArray(lintPointNum)) - lsngIdealOutput
Next

'Return the Calculated Ideal Slope (in %/°)
IdealSlope = lsngIdealM * resolution

Exit Sub

calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcLinTwoPoint: " & Err.Description, True, True)
End Sub

Public Sub CalcMechanicalHysteresis(forwardForce() As Single, reverseForce() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, mechHysteresis() As Single)
'   PURPOSE:    To calculate the Mechanical Hysteresis
'
'      NOTE:    Mechanical Hysteresis is defined as the the difference
'               between the measured force on the forward and reverse scans.
'
'  INPUT(S):    forwardForce  : Forward scan force data
'               reverseForce  : reverse scan force data
'               evaluateStart : location to start evaluating
'               evaluateStop  : location to stop evaluating
'               resolution    : Scan resolution
'
' OUTPUT(S):    mechHysteresis: calculated Mechanical Hysteresis data
'

Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Iterate through the data
For i = evaluateStart To evaluateStop
    'Calculate the difference between forward & reverse
    mechHysteresis(i) = forwardForce(i) - reverseForce(i)
Next i

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcMechanicalHysteresis: " & Err.Description, True, True)
End Sub

Public Sub CalcMechanicalHysteresisPercentage(forwardForce() As Single, reverseForce() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, mechHysteresis() As Single)
'   PURPOSE:    To calculate the Mechanical Hysteresis
'
'      NOTE:    Mechanical Hysteresis is defined as the the difference
'               between the measured force on the forward and reverse scans,
'               expressed as a percentage of the forward scan force.
'
'  INPUT(S):    forwardForce  : Forward scan force data
'               reverseForce  : reverse scan force data
'               evaluateStart : location to start evaluating
'               evaluateStop  : location to stop evaluating
'               resolution    : Scan resolution
'
' OUTPUT(S):    mechHysteresis: calculated Mechanical Hysteresis data
'

Dim i As Integer
Dim lsngDifference As Single

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Iterate through the data
For i = evaluateStart To evaluateStop
    'Calculate the difference between forward & reverse
    lsngDifference = forwardForce(i) - reverseForce(i)
    'Avoid division by zero
    If forwardForce(i) <> 0 Then
        mechHysteresis(i) = (lsngDifference / forwardForce(i)) * HUNDREDPERCENT
    Else
        mechHysteresis(i) = 0
    End If
Next i

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcMechanicalHystersisPercentage: " & Err.Description, True, True)
End Sub

Public Sub CalcMinMax(DataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, minValue As Single, minLocation As Single, maxValue As Single, maxLocation As Single)
'
'   PURPOSE:    To calculate the peak minimum and peak maximum value of the
'               calculated data array.  Also, determine the location of the
'               peak values.
'
'  INPUT(S):    dataArray     : Calculated data array for current channel
'               evaluateStart : Start point of data for current channel
'               evaluateStop  : End point of data for current channel
'               resolution    : Scan resolution
'
' OUTPUT(S):    minValue      : Peak minimum value of calculated data array
'               minLocation   : Peak minimum location of calculated data array
'               maxValue      : Peak maximum value of calculated data array
'               maxLocation   : Peak maximum location of calculated data array
'

Dim i As Integer

On Error GoTo calc_Err
    
'Initialize the minimums and maximums
maxValue = -10000
minValue = 10000
maxLocation = 0
minLocation = 0

'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
    
'Find peak max & min points
For i = evaluateStart To evaluateStop
    If DataArray(i) > maxValue Then       'Look for peak positive point
       maxValue = DataArray(i)
       maxLocation = i
    End If
    If DataArray(i) < minValue Then       'Look for peak negative point
       minValue = DataArray(i)
       minLocation = i
    End If
Next

'Convert locations using resolution
maxLocation = maxLocation / resolution
minLocation = minLocation / resolution

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcMinMax: " & Err.Description, True, True)
End Sub

Public Sub CalcOutputCor(ByVal corCoefficient As Single, scaledDataArray1() As Single, scaledDataArray2() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, outputCor() As Single)
'
'   PURPOSE:    To calculate the Output Correlation.
'
'  INPUT(S):    corCoefficient   : Output Correlation Coefficient
'               scaledDataArray1 : Output 1 data array
'               scaledDataArray2 : Output 2 data array
'               evaluateStart    : Start point of data for current channel
'               evaluateStop     : End point of data for current channel
'               resolution       : Scan resolution
'
' OUTPUT(S):    OutputCor        : Output Correlation array values
'

Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Slope Correlation
For i = evaluateStart To evaluateStop
     outputCor(i) = scaledDataArray1(i) - (corCoefficient * scaledDataArray2(i))
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcOutputCor: " & Err.Description, True, True)
End Sub

Public Sub CalcOutputCor702(scaledDataArray1() As Single, scaledDataArray2() As Single, ByVal Index1Output1 As Single, ByVal Index1Output2, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, outputCor() As Single)
'
'   PURPOSE:    To calculate the 702 Output Correlation.
'
'  INPUT(S):    scaledDataArray1 : Output 1 data array
'               scaledDataArray2 : Output 2 data array
'               Index1Output1    : Index 1 Output #1
'               Index1Output2    : Index 1 Output #2
'               evaluateStart    : Start point of data for current channel
'               evaluateStop     : End point of data for current channel
'               resolution       : Scan resolution
'
' OUTPUT(S):    OutputCor        : Output Correlation array values
'

Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Slope Correlation
For i = evaluateStart To evaluateStop
     outputCor(i) = (scaledDataArray2(i) - Index1Output2) - (scaledDataArray1(i) - Index1Output1)
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcOutputCor702: " & Err.Description, True, True)
End Sub

Public Sub CalcOutputCor705(ByVal corCoefficient As Single, scaledDataArray1() As Single, scaledDataArray2() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, outputCor() As Single)
'
'   PURPOSE:    To calculate the Output Correlation.
'
'  INPUT(S):    corCoefficient   : Output Correlation Coefficient
'               scaledDataArray1 : Output 1 data array
'               scaledDataArray2 : Output 2 data array
'               evaluateStart    : Start point of data for current channel
'               evaluateStop     : End point of data for current channel
'               resolution       : Scan resolution
'
' OUTPUT(S):    OutputCor        : Output Correlation array values
'

Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Slope Correlation
For i = evaluateStart To evaluateStop
     outputCor(i) = (scaledDataArray1(i) / corCoefficient) - scaledDataArray2(i)
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcOutputCor705: " & Err.Description, True, True)
End Sub

Public Sub CalcPercentTol(calcDataArray() As Single, minLimitArray() As Single, maxLimitArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, perTolArray() As Single)
'
'   PURPOSE:    To calculate the percent tolerance of the linearity deviation
'               array.  The percent tolerance shall vary based on the values
'               passed into the routine (i.e. absolute linearity, single-point
'               linearity, etc.).
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
'  INPUT(S):    calcDataArray : Calculated input data array for current channel
'               minLimitArray : Calculated min limit array for current channel
'               maxLimitArray : Calculated max limit array for current channel
'               evaluateStart : Start point of data for current channel
'               evaluateStop  : End point of data for current channel
'               resolution    : Scan resolution
'
' OUTPUT(S):    perTolArray   : Output array of percent tolerance values.
'

Dim i As Integer
  
On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
    
'Calculate the Percent Tolerance array
For i = evaluateStart To evaluateStop
    If calcDataArray(i) > 0 Then
        perTolArray(i) = (calcDataArray(i) / maxLimitArray(i)) * HUNDREDPERCENT
    ElseIf calcDataArray(i) < 0 Then
        perTolArray(i) = (calcDataArray(i) / minLimitArray(i)) * HUNDREDPERCENT
    End If
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcPercentTol: " & Err.Description, True, True)
End Sub

Public Sub CalcRatiometricDataArray(ByVal chanNum As Integer, rawDataArray() As Integer, supplyDataArray() As Integer, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal resolution As Single, ScaledDataArray() As Single)
'
'   PURPOSE:    To calculate the voltage gradient.
'
'  INPUT(S):    chanNum         : Channel number for current output
'               rawDataArray    : Raw data array for the current channel
'               supplyArray     : Supply data array for the current channel
'               evaluateStart   : Start point of data for current channel
'               evaluateStop    : End point of data for current channel
'               resolution      : Scan resolution
'
' OUTPUT(S):    scaledDataArray : Output array of voltage gradient data
'                                 values.
'

Dim i As Integer

On Error GoTo calc_Err

'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Voltage Gradient array
For i = evaluateStart To evaluateStop
    'Return scaledDataArray as a percentage
    ScaledDataArray(i) = (rawDataArray(chanNum, i) / supplyDataArray(i)) * 100
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcRatiometricDataArray: " & Err.Description, True, True)
End Sub

Public Sub CalcRefPointByLoc(ScaledDataArray() As Single, ByVal testLocation As Single, ByVal resolution As Single, refPointValue As Single, refPointLocation As Single)
'
'   PURPOSE:    To calculate the value of a reference point based on its
'               location.
'
'  INPUT(S):    scaledDataArray  : Scaled data array
'               testLocation     : Reference point location
'               resolution       : Scan resolution
'
' OUTPUT(S):    refPointValue    : Reference point value
'               refPointLocation : Reference point location
'
'      NOTE:    The reference point value will need to be passed into both
'               the maxValue and the minValue in the subroutines checkFault
'               and checkSevere.
'

On Error GoTo calc_Err

'Return the location the reference point was measured
refPointLocation = testLocation
'Account for data resolution
testLocation = testLocation * resolution
'Get the value of the reference point
refPointValue = ScaledDataArray(testLocation)

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcRefPointByLoc: " & Err.Description, True, True)
End Sub

Public Sub CalcRefPointByVal(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal refPointIdeal As Single, ByVal mIdeal As Single, ByVal resolution As Single, refPointValue As Single, refPointLoc As Single)
'
'   PURPOSE:    To calculate the value of a reference point based on its
'               ideal value.
'
'  INPUT(S):    scaledDataArray : Scaled data array for current channel
'               evaluateStart   : Start point of data for current channel
'               evaluateStop    : End point of data for current channel
'               refPointIdeal   : Ideal reference point value for the current channel
'               mIdeal          : Ideal slope for current channel
'               resolution      : Scan resolution
'
' OUTPUT(S):    refPointLoc     : Reference point location for current channel
'               refPointValue   : Reference point value for the current channel
'
'      NOTE:    The reference point value will need to be passed into both
'               the maxValue and the minValue in the subroutines checkFault
'               and checkSevere.
'
Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
mIdeal = mIdeal / resolution
    
'Calculate the value and location of the reference point (positive slope)
If (mIdeal > 0) Then
    For i = evaluateStart To evaluateStop
        'Find the first point above the ideal reference
        If ScaledDataArray(i) > refPointIdeal Then
            If i = 0 Then i = 1         'prevents negative array values
            'Check if the previous point is closer to the ideal reference
            If (refPointIdeal - ScaledDataArray(i - 1)) < (ScaledDataArray(i) - refPointIdeal) Then
                refPointLoc = (i - 1)
                refPointValue = ScaledDataArray(i - 1)
            Else
                refPointLoc = i
                refPointValue = ScaledDataArray(i)
            End If
            Exit For
        End If
    Next
'Calculate the value and location of the reference point (negative slope)
Else
    For i = evaluateStart To evaluateStop
        'Find the first point below the ideal reference
        If ScaledDataArray(i) < refPointIdeal Then
            If i = 0 Then i = 1         'Prevents negative array values
            'Check if the previous point is closer to the ideal reference
            If (ScaledDataArray(i - 1) - refPointIdeal) < (refPointIdeal - ScaledDataArray(i)) Then
                refPointLoc = (i - 1)
                refPointValue = ScaledDataArray(i - 1)
            Else
                refPointLoc = i
                refPointValue = ScaledDataArray(i)
            End If
            Exit For
        End If
    Next
End If

'Convert location to degrees
refPointLoc = refPointLoc / resolution

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcRefPointByVal: " & Err.Description, True, True)
End Sub

Public Sub CalcSlopeDev(ScaledDataArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal slopeInterval As Single, ByVal slopeIncrement As Single, ByVal mIdeal As Single, ByVal resolution As Single, RunRatioMethod As Boolean, calcDataArray() As Single)
'
'   PURPOSE:    To calculate the slope deviation.  The slope deviation is specified
'               by Slope Interval and Slope Increment.  The Slope Interval represents
'               the number of data points between the two points used for the slope
'               calculation.  The Slope Increment represents the step size between
'               the slope checks.
'
'     NOTES:    The slope deviation is typically performed on only the forward data.
'
'  INPUT(S):    scaledDataArray : Scaled data array for the current channel
'               evaluateStart   : Start point of data for current channel
'               evaluateStop    : End point of data for current channel
'               slopeInterval   : Number of data points between slope points
'               slopeIncrement  : Step size between slope checks
'               mIdeal          : Ideal slope for current channel
'               resolution      : Scan resolution
'               RunRatioMethod  : TRUE  = Calculate slope deviation via ratio method
'                               : FALSE = Calculate slope deviation via subtraction method
'
' OUTPUT(S):    calcDataArray   : Output array of slope deviation values.
'                                 other = system error
'
Dim lintPointNum As Integer

On Error GoTo calc_Err
    
'Detmerine startpoint, endpoint, & slope based on resolution
evaluateStart = (evaluateStart * resolution)
evaluateStop = (evaluateStop * resolution)
mIdeal = mIdeal / resolution

'Calculate the Slope Deviation array based on method selected
For lintPointNum = evaluateStart To (evaluateStop - slopeInterval) Step slopeIncrement
    If (RunRatioMethod) Then            'Slope dev via ratio method
        calcDataArray(lintPointNum / slopeIncrement) = (ScaledDataArray(lintPointNum + slopeInterval) - ScaledDataArray(lintPointNum)) / (mIdeal * slopeInterval)
    Else                                'Slope dev via subtraction method
        calcDataArray(lintPointNum / slopeIncrement) = (ScaledDataArray(lintPointNum + slopeInterval) - ScaledDataArray(lintPointNum)) - (mIdeal * slopeInterval)
    End If
Next lintPointNum

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcSlopeDev: " & Err.Description, True, True)
End Sub

Public Sub CalcRisingPoint(ScaledDataArray() As Single, ByVal idleOutput As Single, ByVal LineStartVal As Single, ByVal LineStartLoc As Single, ByVal LineStopVal As Single, ByVal LineStopLoc As Single, risingPointVal As Single, risingPointLoc As Single)
'
'   PURPOSE:    To calculate the rising point.
'
'  INPUT(S):    scaledDataArray : Scaled data array, binary counts
'               idleOutput      : Output at Idle, defines first line
'               LineStartVal    : Start Location for Line representing linear pedal output
'               LineStartLoc    : Stop Location for Line representing linear pedal output
'               LineStopVal     : Start Location for Line representing linear pedal output
'               LineStopLoc     : Stop Location for Line representing linear pedal output
'
' OUTPUT(S):    risingPointVal  : Value of Rising Point
'               risingPointLoc  : Location of Rising Point
'               anomaly:        : 0     = subroutine completed successfully
'                                 other = system error
'

Dim i As Integer
Dim lsngM As Single
Dim lsngB As Single

On Error GoTo calc_Err

'Initialize the variables to zeroes
risingPointLoc = 0
risingPointVal = 0

'Calculate the best-fit slope of line 2
lsngM = (LineStopVal - LineStartVal) / (LineStopLoc - LineStartLoc)

'Calculate the intercept of line 2
lsngB = LineStartVal - (LineStartLoc * lsngM)

If lsngM <> 0 Then
    'Calculate the Rising Point using Line 2 and the Idle Output
    risingPointLoc = (idleOutput - lsngB) / lsngM
    'If the Rising Point is less than zero, or the rising point is beyond the
    'start of Line 2, the part's output is BAD
    If risingPointLoc < 0 Or risingPointLoc > LineStartLoc Then
        risingPointLoc = 0
    Else
        'Determine the output at the Rising Point Location
        risingPointVal = ScaledDataArray(Round(risingPointLoc))
    End If
End If

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcRisingPoint: " & Err.Description, True, True)
End Sub

Public Sub CalcScaledDataArray(ByVal chanNum As Integer, rawDataArray() As Integer, ByVal evaluateStart As Single, ByVal evaluateStop As Single, ByVal scaler As Single, ByVal offset As Single, ByVal resolution As Single, ScaledDataArray() As Single)
'
'   PURPOSE:    To calculate the voltage gradient.
'
'  INPUT(S):    chanNum         : Channel number for current output
'               rawDataArray    : Raw data array for the current channel
'               evaluateStart   : Start point of data for current channel
'               evaluateStop    : End point of data for current channel
'               scaler          : Scale Factor
'               offset          : Offset (Applied after scaling)
'               resolution      : Scan resolution
'
' OUTPUT(S):    scaledDataArray : Output array of voltage gradient data
'                                 values.
'

Dim i As Integer

On Error GoTo calc_Err

'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution

'Calculate the Scaled & Offset Data Array
For i = evaluateStart To evaluateStop
    ScaledDataArray(i) = (rawDataArray(chanNum, i) * scaler) + offset
Next

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CalcScaledDataArray: " & Err.Description, True, True)
End Sub

Public Sub CheckFault(ByVal chanNum As Integer, ByVal minValue As Single, ByVal maxValue As Single, ByVal minLimit As Single, ByVal maxLimit As Single, ByVal lowBit As Long, ByVal highBit As Long, Failure() As Integer)
'
'   PURPOSE:    To check the minimum and maximum values to the specified
'               limits.  If either limit is exceeded, a fault will be set
'               for the current parameter.  Otherwise, the fault will be
'               cleared.
'
'  INPUT(S):    chanNum       : Channel number of current failure
'               minValue      : Peak minimum value
'               maxValue      : Peak maximum value
'               minLimit      : Specified minimum limit for parameter
'               maxLimit      : Specified maximum limit for parameter
'               lowBit        : Identifies low  failure bit for parameter
'               highBit       : Identifies high failure bit for parameter
'
' OUTPUT(S):    failure       : Fault status for parameter
'

On Error GoTo calc_Err
    
'Set failure based on high limit value
If maxValue > maxLimit Then
    Failure(chanNum, highBit) = True
Else
    Failure(chanNum, highBit) = False
End If

'Set failure based on low limit value
If minValue < minLimit Then
    Failure(chanNum, lowBit) = True
Else
    Failure(chanNum, lowBit) = False
End If

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CheckFault: " & Err.Description, True, True)
End Sub

Public Sub CheckFaultArray(calcDataArray() As Single, minLimitArray() As Single, maxLimitArray() As Single, ByVal evaluateStart As Single, ByVal evaluateStop As Single, chanNum As Integer, lowBit As Long, highBit As Long, ByVal resolution As Single, Failure() As Integer)
'
'   PURPOSE:    To check an array of data to specified min and max limit
'               arrays.  If either limit is exceeded, a fault will be set
'               for the current parameter.  Otherwise, the fault will be
'               cleared.
'
'  INPUT(S):    calcDataArray : Calculated data array for current channel
'               minLimitArray : Calculated min limit array for current channel
'               maxLimitArray : Calculated max limit array for current channel
'               evaluateStart : Start point of data for current channel
'               evaluateStop  : End point of data for current channel
'               chanNum       : Channel number of current failure
'               lowBit        : Identifies low  failure bit for parameter
'               highBit       : Identifies high failure bit for parameter
'               resolution    : Scan resolution
'
' OUTPUT(S):    failure       : Fault status for parameter
'

Dim lblnFailHigh As Boolean, lblnFailLow As Boolean
Dim i As Integer

On Error GoTo calc_Err
    
'Account for data resolution
evaluateStart = evaluateStart * resolution
evaluateStop = evaluateStop * resolution
    
'Check data array to limit array
For i = evaluateStart To evaluateStop
    If calcDataArray(i) > maxLimitArray(i) Then
        lblnFailHigh = True             'Set TRUE when max limit exceeded
        'Exit loop if both failures set, otherwise continue...
        If (lblnFailHigh) And (lblnFailLow) Then Exit For
    End If
    If calcDataArray(i) < minLimitArray(i) Then
        lblnFailLow = True              'Set TRUE when min limit exceeded
        'Exit loop if both failures set, otherwise continue...
        If (lblnFailHigh) And (lblnFailLow) Then Exit For
    End If
Next i

'Set failure based on logical for high limit
If (lblnFailHigh) Then
    Failure(chanNum, highBit) = True
Else
    Failure(chanNum, highBit) = False
End If

'Set failure based on logical for low limit
If (lblnFailLow) Then
    Failure(chanNum, lowBit) = True
Else
    Failure(chanNum, lowBit) = False
End If

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CheckFaultArray: " & Err.Description, True, True)
End Sub

Public Sub CheckSevere(chanNum As Integer, minValue As Single, maxValue As Single, minControlLimit As Single, maxControlLimit As Single, lowBit As Long, highBit As Long, severe() As Integer)
'
'   PURPOSE:    To check the minimum and maximum values to the specified
'               control limits.  If either control limit is exceeded, a
'               severe fault will be set for the current parameter.
'               Otherwise, the severe fault will be cleared.
'
'  INPUT(S):    chanNum         : Channel number of current failure
'               minValue        : Peak minimum value
'               maxValue        : Peak maximum value
'               minControlLimit : Specified minimum control limit for parameter
'               maxControlLimit : Specified maximum control limit for parameter
'               lowBit          : Identifies low  severe bit for parameter
'               highBit         : Identifies high severe bit for parameter
'
' OUTPUT(S):    severe          : Severe status for current channel
'

On Error GoTo calc_Err
    
'Set severe based on high control limit value
If maxValue > maxControlLimit Then
    severe(chanNum, highBit) = True
Else
    severe(chanNum, highBit) = False
End If
    
'Set severe based on low control limit value
If minValue < minControlLimit Then
    severe(chanNum, lowBit) = True
Else
    severe(chanNum, lowBit) = False
End If

Exit Sub
calc_Err:

    gintAnomaly = 1
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Calc.CheckSevere: " & Err.Description, True, True)
End Sub
