VERSION 5.00
Object = "{8C7A5A52-105F-11CF-9BE5-0020AF6845F6}#1.4#0"; "cwdaq.ocx"
Begin VB.Form frmSolver90277 
   Caption         =   "Solver"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1995
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   1995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   360
   End
   Begin CWDAQControlsLib.CWAIPoint cwaiSolver 
      Left            =   360
      Top             =   360
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AIPoint_0       =   1
      ClassName_1     =   "CCWAIPoint"
      opts_1          =   2
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   0
      ChannelClock_1  =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   2
      Frequency_2     =   100
      Period_2        =   0.01
      InternalClockMode_2=   1
      Buffer_1        =   0
      Channels_1      =   3
      ClassName_3     =   "CCWAIChannelArray"
      Editor_3        =   4
      ClassName_4     =   "CCWAIChannelsArrayEditor"
      Owner_4         =   1
      StartCond_1     =   0
      PauseCond_1     =   0
      StopCond_1      =   0
      HoldoffClock_1  =   0
   End
End
Attribute VB_Name = "frmSolver90277"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************** Melexis 90277 Programming Interface *****************
'
'   Scott R Calkins
'   CTS Corporation Automotive Products
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (574) 295-3575
'
'Ver      Date      By   Purpose of modification
'1.0.0  07/15/2004  SRC  First release per PR12722-A & PR12722-B.
'
'1.0.1  01/31/2005  SRC  Changed comments slightly for release of
'                        EE904B & EE905B.
'
'1.1.0  02/14/2005  SRC  Changed High/Low Saturation level
'                        constants from 96%/4% to 97%/3%.
'                        Change verified by Melexis.  Filter/
'                        Load circuits with pull-down resistors
'                        less than 10 K-ohms may need to be
'                        evaluated for compatibility with these
'                        constants.
'1.2.0  03/16/2005  SRC  Separated Index 1 & 2 and Hi/Lo Clamp tolerances
'                        into pass/fail tolerances and target tolerances.
'                        Improved programming history code tracking.
'                        Added ability to perform second offset adjustment
'                        at the end of each cycle.  Changes per PR13321-B.
'                        Changes tagged as 'V1.2.0.
'1.3.0  05/27/2005  SRC  Modified software to force output of programming
'                        failures into low diagnostic region.  Added constant
'                        FGFORCLAMPS to be consistent with RG settings.
'                        Changes tagged as 'V1.3.0.
'1.4.0  08/22/2006  SRC  Updated ClampSolver to make 5 attempts to reach the
'                        correct clamp, and to utilize the InvertSlope parameter
'                        to help force the output to the high clamp.
'                        Update to use 4 points on initial solver calculations.
'                        Changes tagged as 'V1.4.0.
'1.5.0  11/20/2006  ANM  Updated solver for new clamp routine tagged as 'V1.5.0
'1.6.0  03/08/2007  ANM  Updated solver FG/RG clamps per TR8501-E tagged as 'V1.6.0
'1.7.0  05/01/2007  ANM  Updated solver for MLX data dump per SCN# MISC-102 'V1.7.0
'1.8.0  05/17/2007  ANM  Updated solver for clamps after MLX visit 'V1.8.0
'1.9.0  07/23/2007  ANM  Updated solver for new AMAD 'V1.9.0
'1.9.1  09/24/2007  ANM  Updated solver for prog/scan offset per SCN# 705F-008 (3979). 'V1.9.1
'1.9.2  01/22/2008  ANM  Updated solver for 2nd attempts per SCN#s 4066 & 4067. 'V1.9.2
'1.9.3  01/30/2008  ANM  Updated solver for reverse clamp low.                  'V1.9.3
'1.9.4  02/06/2008  ANM  Updated solver to save ZG value per SCN# 4087.         'V1.9.4
'

Option Explicit

Private mblnKillTimeDone As Boolean

Public Sub KillTime(milliSecDelay As Integer)
'
'   PURPOSE:   Delays a set amount of time(user-specified) using a timer
'              event.  The delay time is in milliseconds.
'
'  INPUT(S):   milliSecDelay : Delay time in milliseconds
' OUTPUT(S):   None
    
mblnKillTimeDone = False
tmrKillTime.Interval = milliSecDelay
tmrKillTime.Enabled = True
Do
    DoEvents
Loop Until mblnKillTimeDone

tmrKillTime.Enabled = False

End Sub

Public Sub ReadSolverVoltages(Output1AsPercentOfApplied As Single, Output2AsPercentOfApplied As Single)
'
'   PURPOSE: Reads Vout & Vdd while the DUT is connected to the programmer.

'            Returns Vout as a percentage of the applied voltage(Vdd).
'
'  INPUT(S): none
' OUTPUT(S): Output1AsPercentOfApplied = Vout1 as a percentage of Vdd1
'            Output2AsPercentOfApplied = Vout2 as a percentage of Vdd2

Dim lintIndex As Integer
Dim llngVout1 As Long
Dim llngVout2 As Long
Dim llngSupply1 As Long
Dim llngSupply2 As Long
Dim lvntData As Variant

'Make the read of Vout1, Vdd1, Vout2, & Vdd2
cwaiSolver.SingleRead lvntData

'Loop through the ten pieces of data summing data from the variant
For lintIndex = 0 To 9
    llngVout1 = llngVout1 + lvntData(lintIndex * 4)             '0,4,8,...36
    llngVout2 = llngVout2 + lvntData((lintIndex * 4) + 1)       '1,5,9,...37
    llngSupply1 = llngSupply1 + lvntData((lintIndex * 4) + 2)   '2,6,10,...38
    llngSupply2 = llngSupply2 + lvntData((lintIndex * 4) + 3)   '3,7,11,...39
Next lintIndex

'Determine the average of the Supply & Vout readings
llngVout1 = llngVout1 / 10
llngVout2 = llngVout2 / 10
llngSupply1 = llngSupply1 / 10
llngSupply2 = llngSupply2 / 10

'Get the ratiometric output
Output1AsPercentOfApplied = (llngVout1 / llngSupply1) * HUNDREDPERCENT
Output2AsPercentOfApplied = (llngVout2 / llngSupply2) * HUNDREDPERCENT

End Sub

Public Sub SolverDAQSetup()
'
'   PURPOSE:    Initializes the data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

'Initialize the parameters for the Solver DAQ Control
cwaiSolver.Device = 1
cwaiSolver.Channels.RemoveAll
cwaiSolver.Channels.Add "0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3"
cwaiSolver.ReturnDataType = cwaiBinaryCodes
cwaiSolver.ChannelClock.InternalClockMode = cwaiPeriod
cwaiSolver.ChannelClock.Period = 0.0005

End Sub

Private Sub Form_Load()
'
'   PURPOSE:    Initializes the form when it is loaded
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

'Setup the Solver Data Acquisition Properties
If InStr(command$, "NOHARDWARE") = 0 Then
    'Initialize the Solver data acquisition properties
    Call SolverDAQSetup
End If

End Sub

Private Sub tmrKillTime_Timer()
'
'   PURPOSE:   Event triggered when when timer, tmrKillTime, is complete.
'
'  INPUT(S):   None
' OUTPUT(S):   None

mblnKillTimeDone = True

End Sub
