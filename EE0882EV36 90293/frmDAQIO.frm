VERSION 5.00
Object = "{8C7A5A52-105F-11CF-9BE5-0020AF6845F6}#1.4#0"; "cwdaq.ocx"
Object = "{E7BC3920-33D4-11D0-8B73-0020AF31CEF9}#1.4#0"; "cwanalysis.ocx"
Begin VB.Form frmDAQIO 
   Caption         =   "Data Acquisition & I/O Controls"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin CWDAQControlsLib.CWAIPoint cwaiMonitorDAQ 
      Left            =   120
      Top             =   2520
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
   Begin CWDAQControlsLib.CWAI cwaiRDAQ 
      Left            =   120
      Top             =   1800
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AITask_0        =   1
      ClassName_1     =   "CCWAITask"
      opts_1          =   2
      ErrorEventMask_1=   14336
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   1
      Frequency_2     =   100
      Period_2        =   0.01
      ClockSourceType_2=   1
      ChannelClock_1  =   3
      ClassName_3     =   "CCWAIClock"
      ClockType_3     =   2
      Frequency_3     =   100
      Period_3        =   0.01
      InternalClockMode_3=   1
      Buffer_1        =   4
      ClassName_4     =   "CCWAIBuffer"
      Channels_1      =   5
      ClassName_5     =   "CCWAIChannelArray"
      Editor_5        =   6
      ClassName_6     =   "CCWAIChannelsArrayEditor"
      Owner_6         =   1
      StartCond_1     =   7
      ClassName_7     =   "CCWAICondition"
      WhichCondition_7=   1
      PreTriggerScans_7=   3
      PauseCond_1     =   8
      ClassName_8     =   "CCWAICondition"
      WhichCondition_8=   2
      TrigPauseMode_8 =   7
      PreTriggerScans_8=   3
      StopCond_1      =   9
      ClassName_9     =   "CCWAICondition"
      WhichCondition_9=   3
      PreTriggerScans_9=   3
      HoldoffClock_1  =   10
      ClassName_10    =   "CCWAIClock"
      ClockType_10    =   3
      Frequency_10    =   100
      Period_10       =   0.01
      InternalClockMode_10=   1
   End
   Begin CWDAQControlsLib.CWDAQTools cwDAQTools1 
      Left            =   3600
      Top             =   3960
      _Version        =   65540
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin CWDAQControlsLib.CWDIO cwDIO2 
      Left            =   3600
      Top             =   1800
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      DIO_0           =   1
      ClassName_1     =   "CDigital"
      opts_1          =   2
      Device_1        =   1
      prts_1          =   2
      ClassName_2     =   "CCWDIOPorts"
      Editor_2        =   3
      ClassName_3     =   "CCWReadOnlyArrayEditor"
      Owner_3         =   1
      chans_1         =   4
      ClassName_4     =   "CCWDIOChannels"
      Editor_4        =   5
      ClassName_5     =   "CCWDIOChannelArrayEditor"
      Owner_5         =   1
   End
   Begin CWDAQControlsLib.CWAIPoint cwaiPeakForce 
      Left            =   120
      Top             =   5400
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
   Begin CWDAQControlsLib.CWAOPoint cwaoVRef 
      Left            =   3600
      Top             =   360
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AOPointTask_0   =   1
      ClassName_1     =   "CCWAOPointTask"
      opts_1          =   2
   End
   Begin CWAnalysisControlsLib.CWStat cwStat1 
      Left            =   3600
      Top             =   2520
      _Version        =   65540
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   3240
   End
   Begin CWDAQControlsLib.CWDIO cwDIO1 
      Left            =   3600
      Top             =   1080
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      DIO_0           =   1
      ClassName_1     =   "CDigital"
      opts_1          =   2
      Device_1        =   1
      prts_1          =   2
      ClassName_2     =   "CCWDIOPorts"
      Editor_2        =   3
      ClassName_3     =   "CCWReadOnlyArrayEditor"
      Owner_3         =   1
      chans_1         =   4
      ClassName_4     =   "CCWDIOChannels"
      Editor_4        =   5
      ClassName_5     =   "CCWDIOChannelArrayEditor"
      Owner_5         =   1
   End
   Begin CWDAQControlsLib.CWAIPoint cwaiVRef 
      Left            =   120
      Top             =   3240
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
   Begin CWDAQControlsLib.CWAI cwaiFDAQ 
      Left            =   120
      Top             =   1080
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AITask_0        =   1
      ClassName_1     =   "CCWAITask"
      opts_1          =   2
      ErrorEventMask_1=   14336
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   1
      Frequency_2     =   100
      Period_2        =   0.01
      ClockSourceType_2=   1
      ChannelClock_1  =   3
      ClassName_3     =   "CCWAIClock"
      ClockType_3     =   2
      Frequency_3     =   100
      Period_3        =   0.01
      InternalClockMode_3=   1
      Buffer_1        =   4
      ClassName_4     =   "CCWAIBuffer"
      Channels_1      =   5
      ClassName_5     =   "CCWAIChannelArray"
      Editor_5        =   6
      ClassName_6     =   "CCWAIChannelsArrayEditor"
      Owner_6         =   1
      StartCond_1     =   7
      ClassName_7     =   "CCWAICondition"
      WhichCondition_7=   1
      PreTriggerScans_7=   3
      PauseCond_1     =   8
      ClassName_8     =   "CCWAICondition"
      WhichCondition_8=   2
      TrigPauseMode_8 =   7
      PreTriggerScans_8=   3
      StopCond_1      =   9
      ClassName_9     =   "CCWAICondition"
      WhichCondition_9=   3
      PreTriggerScans_9=   3
      HoldoffClock_1  =   10
      ClassName_10    =   "CCWAIClock"
      ClockType_10    =   3
      Frequency_10    =   100
      Period_10       =   0.01
      InternalClockMode_10=   1
   End
   Begin CWDAQControlsLib.CWAI cwaiPreScanDAQ 
      Left            =   120
      Top             =   360
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AITask_0        =   1
      ClassName_1     =   "CCWAITask"
      opts_1          =   2
      ErrorEventMask_1=   14336
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   1
      Frequency_2     =   100
      Period_2        =   0.01
      ClockSourceType_2=   1
      ChannelClock_1  =   3
      ClassName_3     =   "CCWAIClock"
      ClockType_3     =   2
      Frequency_3     =   100
      Period_3        =   0.01
      InternalClockMode_3=   1
      Buffer_1        =   4
      ClassName_4     =   "CCWAIBuffer"
      Channels_1      =   5
      ClassName_5     =   "CCWAIChannelArray"
      Editor_5        =   6
      ClassName_6     =   "CCWAIChannelsArrayEditor"
      Owner_6         =   1
      StartCond_1     =   7
      ClassName_7     =   "CCWAICondition"
      WhichCondition_7=   1
      PreTriggerScans_7=   3
      PauseCond_1     =   8
      ClassName_8     =   "CCWAICondition"
      WhichCondition_8=   2
      TrigPauseMode_8 =   7
      PreTriggerScans_8=   3
      StopCond_1      =   9
      ClassName_9     =   "CCWAICondition"
      WhichCondition_9=   3
      PreTriggerScans_9=   3
      HoldoffClock_1  =   10
      ClassName_10    =   "CCWAIClock"
      ClockType_10    =   3
      Frequency_10    =   100
      Period_10       =   0.01
      InternalClockMode_10=   1
   End
   Begin CWDAQControlsLib.CWAIPoint cwaiVOut 
      Left            =   120
      Top             =   3960
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
   Begin CWDAQControlsLib.CWAIPoint cwaiForce 
      Left            =   3600
      Top             =   4680
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
   Begin CWDAQControlsLib.CWAIPoint cwaiFOut 
      Left            =   120
      Top             =   4680
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
   Begin CWDAQControlsLib.CWAI cwaiFTO 
      Left            =   600
      Top             =   1080
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AITask_0        =   1
      ClassName_1     =   "CCWAITask"
      opts_1          =   2
      ErrorEventMask_1=   14336
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   1
      Frequency_2     =   100
      Period_2        =   0.01
      ClockSourceType_2=   1
      ChannelClock_1  =   3
      ClassName_3     =   "CCWAIClock"
      ClockType_3     =   2
      Frequency_3     =   100
      Period_3        =   0.01
      InternalClockMode_3=   1
      Buffer_1        =   4
      ClassName_4     =   "CCWAIBuffer"
      Channels_1      =   5
      ClassName_5     =   "CCWAIChannelArray"
      Editor_5        =   6
      ClassName_6     =   "CCWAIChannelsArrayEditor"
      Owner_6         =   1
      StartCond_1     =   7
      ClassName_7     =   "CCWAICondition"
      WhichCondition_7=   1
      PreTriggerScans_7=   3
      PauseCond_1     =   8
      ClassName_8     =   "CCWAICondition"
      WhichCondition_8=   2
      TrigPauseMode_8 =   7
      PreTriggerScans_8=   3
      StopCond_1      =   9
      ClassName_9     =   "CCWAICondition"
      WhichCondition_9=   3
      PreTriggerScans_9=   3
      HoldoffClock_1  =   10
      ClassName_10    =   "CCWAIClock"
      ClockType_10    =   3
      Frequency_10    =   100
      Period_10       =   0.01
      InternalClockMode_10=   1
   End
   Begin CWDAQControlsLib.CWAI cwaiRFTO 
      Left            =   600
      Top             =   1800
      _Version        =   393219
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Reset_0         =   0   'False
      CompatibleVers_0=   393219
      AITask_0        =   1
      ClassName_1     =   "CCWAITask"
      opts_1          =   2
      ErrorEventMask_1=   14336
      Device_1        =   1
      TotalScansToAcquire_1=   100
      ScanClock_1     =   2
      ClassName_2     =   "CCWAIClock"
      ClockType_2     =   1
      Frequency_2     =   100
      Period_2        =   0.01
      ClockSourceType_2=   1
      ChannelClock_1  =   3
      ClassName_3     =   "CCWAIClock"
      ClockType_3     =   2
      Frequency_3     =   100
      Period_3        =   0.01
      InternalClockMode_3=   1
      Buffer_1        =   4
      ClassName_4     =   "CCWAIBuffer"
      Channels_1      =   5
      ClassName_5     =   "CCWAIChannelArray"
      Editor_5        =   6
      ClassName_6     =   "CCWAIChannelsArrayEditor"
      Owner_6         =   1
      StartCond_1     =   7
      ClassName_7     =   "CCWAICondition"
      WhichCondition_7=   1
      PreTriggerScans_7=   3
      PauseCond_1     =   8
      ClassName_8     =   "CCWAICondition"
      WhichCondition_8=   2
      TrigPauseMode_8 =   7
      PreTriggerScans_8=   3
      StopCond_1      =   9
      ClassName_9     =   "CCWAICondition"
      WhichCondition_9=   3
      PreTriggerScans_9=   3
      HoldoffClock_1  =   10
      ClassName_10    =   "CCWAIClock"
      ClockType_10    =   3
      Frequency_10    =   100
      Period_10       =   0.01
      InternalClockMode_10=   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Input Control for FOut (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   13
      Left            =   960
      TabIndex        =   14
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Input Control for Force"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   25
      Left            =   4440
      TabIndex        =   13
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Input Control for VOut (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   12
      Left            =   960
      TabIndex        =   12
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "PreScan Analog Input  Control for Data Acquisition (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   9
      Left            =   960
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Monitor DAQ Analog  Input Control for Data Acquisition (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   7
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Reverse Analog Input  Control for Data Acquisition (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   11
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "DAQ Tools for A/D Convert Signal"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   10
      Left            =   4440
      TabIndex        =   8
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Forward Analog Input  Control for Data Acquisition (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   8
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Input Control for Peak Force (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   6
      Left            =   960
      TabIndex        =   6
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Output Control for VRef"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Single Point Analog Input Control for VRef (Device #1)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Digital I/O Control for PT Board (Device #2)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Digital I/O Control for Digitizer / PLC (Device #3)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Component Works Statistical Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "KillTime Delay Timer Control (msec)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmDAQIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnKillTimeDone As Boolean

Public Sub FoutDAQSetup()
'
'   PURPOSE:    Initializes the Fout data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None
'1.3ANM new sub

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    cwaiFOut.Device = 1                             'DAQ Board is device #1

    cwaiFOut.Channels.RemoveAll                     'Empty the channel string
    cwaiFOut.Channels.Add ("4,4,4")                 'Set up the channel string

    cwaiFOut.ReturnDataType = cwaiScaledData        'Return data is binary counts

End If
    
Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 120
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in FoutDAQSetup: " & Err.Description, True, True)
    
End Sub

Public Sub ReadFout(FOut As Single)
'
'   PURPOSE: Reads the FOut on Ch #4
'
'  INPUT(S): none
' OUTPUT(S): none
'1.3ANM new sub

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim lvntVoltage As Variant
Dim lsngFout As Single

'Read the Peak Force Channel
cwaiFOut.SingleRead lvntVoltage

lsngFout = (lvntVoltage(0) + lvntVoltage(1) + lvntVoltage(2)) / 3

'Calculate Fout
FOut = (lsngFout * gsngNewtonsPerVolt) + gsngForceAmplifierOffset

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 121
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ReadFout: " & Err.Description, True, True)

End Sub

Private Sub cwaiFDAQ_AcquiredData(Voltages As Variant, BinaryData As Variant)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
'
'   NOTE:  Voltages and BinaryData are two-dimensional arrays. The
'          first dimension in the array represents the channel
'          string.  The second dimension in the array represents the
'          number of scans to acquire for each channel.  Hence, the
'          arrays can be identified as follows:
'

'               Voltages(channel number, scan number)
'                               OR
'               BinaryData(channel number, scan number)
'

Dim i As Integer
Dim llngTriggerCnt As Long

ReDim gintForward(CHAN2, gintMaxData - 1)   'Raw forward data
ReDim gintForSupply(gintMaxData - 1)        'Raw forward data for supply

Call KillTime(100)      'Delay to allow counts to be updated

'Get the number of trigger counts for the forward scan
llngTriggerCnt = TriggerCnt

'Check number of triggers to see if data is valid
'Note:  Need to convert to long because the types are different
If llngTriggerCnt <> CLng(((gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution) + 1) Then
    gintAnomaly = 108
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Incorrect # of Triggers counted during the Forward Scan.", True, True)
End If

If gintAnomaly = 0 Then
    For i = 0 To (gintMaxData - 1)
        'Fill Output #1 Array
        gintForward(CHAN0, i) = (BinaryData(0, i) + BinaryData(1, i) + BinaryData(2, i)) \ 3
        'Fill Output #2 Array
        gintForward(CHAN1, i) = (BinaryData(3, i) + BinaryData(4, i) + BinaryData(5, i)) \ 3
        'Fill Force Array
        gintForward(CHAN2, i) = (BinaryData(6, i) + BinaryData(7, i) + BinaryData(8, i)) \ 3
        'Fill VRef Array
        gintForSupply(i) = (BinaryData(9, i) + BinaryData(10, i) + BinaryData(11, i)) \ 3
    Next i
End If

gblnAnalogDone = True

End Sub

Private Sub cwaiFTO_AcquiredData(ScaledData As Variant, BinaryData As Variant)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
'
'   NOTE:  Voltages and BinaryData are two-dimensional arrays. The
'          first dimension in the array represents the channel
'          string.  The second dimension in the array represents the
'          number of scans to acquire for each channel.  Hence, the
'          arrays can be identified as follows:
'

'               Voltages(channel number, scan number)
'                               OR
'               BinaryData(channel number, scan number)
'
'1.8ANM new sub

Dim i As Integer
Dim llngTriggerCnt As Long

ReDim gintForward(CHAN2, gintMaxData - 1)   'Raw forward data
ReDim gintForSupply(gintMaxData - 1)        'Raw forward data for supply

Call KillTime(100)      'Delay to allow counts to be updated

'Get the number of trigger counts for the forward scan
llngTriggerCnt = TriggerCnt

'Check number of triggers to see if data is valid
'Note:  Need to convert to long because the types are different
If llngTriggerCnt <> CLng(((gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution) + 1) Then
    gintAnomaly = 108
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Incorrect # of Triggers counted during the Force Scan.", True, True)
End If

If gintAnomaly = 0 Then
    For i = 0 To (gintMaxData - 1)
        'Fill Force Array
        gintForward(CHAN2, i) = (BinaryData(0, i) + BinaryData(1, i) + BinaryData(2, i)) \ 3
    Next i
End If

gblnAnalogDone = True

End Sub

Private Sub cwaiPreScanDAQ_AcquiredData(Voltages As Variant, BinaryData As Variant)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
'
'   NOTE:  Voltages and BinaryData are two-dimensional arrays. The
'          first dimension in the array represents the channel
'          string.  The second dimension in the array represents the
'          number of scans to acquire for each channel.  Hence, the
'          arrays can be identified as follows:
'
'               Voltages(channel number, scan number)
'                               OR
'               BinaryData(channel number, scan number)
'

Dim i As Integer
Dim llngTriggerCnt As Long

ReDim gintPreScanForce(gintMaxData - 1)   'Raw PreScan Force Data

Call KillTime(100)      'Delay to allow counts to be updated

'Get the number of trigger counts for the reverse scan
llngTriggerCnt = TriggerCnt

'Check number of triggers to see if the acquisition went as expected
'Note:  Need to convert to long because the types are different
If llngTriggerCnt <> CLng(((gudtMachine.preScanStop - gudtMachine.preScanStart) * gsngResolution) + 1) Then
    gintAnomaly = 107
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Incorrect # of Triggers counted during the Pre-Scan.", True, True)
End If

If gintAnomaly = 0 Then
    'Fill the PreScan Force Array
    For i = 0 To (gintMaxData - 1)
        gintPreScanForce(i) = (BinaryData(0, i) + BinaryData(1, i) + BinaryData(2, i)) \ 3
    Next i
End If

gblnAnalogDone = True

End Sub

Private Sub cwaiRDAQ_AcquiredData(Voltages As Variant, BinaryData As Variant)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
'
'   NOTE:  Voltages and BinaryData are two-dimensional arrays. The
'          first dimension in the array represents the channel
'          string.  The second dimension in the array represents the
'          number of scans to acquire for each channel.  Hence, the
'          arrays can be identified as follows:
'
'               Voltages(channel number, scan number)
'                               OR
'               BinaryData(channel number, scan number)
'

Dim i As Integer
Dim llngTriggerCnt As Long

ReDim gintReverse(CHAN2, gintMaxData - 1)   'Raw reverse data
ReDim gintRevSupply(gintMaxData - 1)        'Raw reverse data for supply

'NOTE:  After the reverse scan, the trigger count value reflects the counts
'       of both the forward and reverse scans (i.e. double the counts expected)

Call KillTime(100)      'Delay to allow counts to be updated

'Get the number of trigger counts for the reverse scan
llngTriggerCnt = TriggerCnt / 2

'*** Check number of triggers to see if data is valid ***
'Note:  Need to convert to long because the types are different
If llngTriggerCnt <> CLng(((gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution) + 1) Then
    gintAnomaly = 109
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Incorrect # of Triggers counted during the Reverse Scan.", True, True)
End If

If gintAnomaly = 0 Then
    'Fill the Reverse arrays in reverse order by using ((gintMaxData - 1) - i) as the index
    For i = 0 To (gintMaxData - 1)
        'Fill Output #1 Array
        gintReverse(CHAN0, (gintMaxData - 1) - i) = (BinaryData(0, i) + BinaryData(1, i) + BinaryData(2, i)) \ 3
        'Fill Output #2 Array
        gintReverse(CHAN1, (gintMaxData - 1) - i) = (BinaryData(3, i) + BinaryData(4, i) + BinaryData(5, i)) \ 3
        'Fill Force Array
        gintReverse(CHAN2, (gintMaxData - 1) - i) = (BinaryData(6, i) + BinaryData(7, i) + BinaryData(8, i)) \ 3
        'Fill VRef Array
        gintRevSupply((gintMaxData - 1) - i) = (BinaryData(9, i) + BinaryData(10, i) + BinaryData(11, i)) \ 3
    Next i
End If

gblnAnalogDone = True

End Sub

Public Sub KillTime(milliSecDelay As Integer)
'
'   PURPOSE:   Delays a set amount of time(user-specified) using a timer
'              event.  The delay time is in milliseconds.
'
'  INPUT(S):   milliSecDelay : Delay time in milliseconds
' OUTPUT(S):   None
    
mblnKillTimeDone = False
tmrKillTime.Enabled = False
tmrKillTime.Interval = milliSecDelay
tmrKillTime.Enabled = True
Do
    DoEvents
Loop Until mblnKillTimeDone
    
End Sub

Private Sub cwaiRFTO_AcquiredData(ScaledData As Variant, BinaryData As Variant)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
'
'   NOTE:  Voltages and BinaryData are two-dimensional arrays. The
'          first dimension in the array represents the channel
'          string.  The second dimension in the array represents the
'          number of scans to acquire for each channel.  Hence, the
'          arrays can be identified as follows:
'
'               Voltages(channel number, scan number)
'                               OR
'               BinaryData(channel number, scan number)
'
'1.8ANM new sub

Dim i As Integer
Dim llngTriggerCnt As Long

ReDim gintReverse(CHAN2, gintMaxData - 1)   'Raw reverse data
ReDim gintRevSupply(gintMaxData - 1)        'Raw reverse data for supply

'NOTE:  After the reverse scan, the trigger count value reflects the counts
'       of both the forward and reverse scans (i.e. double the counts expected)

Call KillTime(100)      'Delay to allow counts to be updated

'Get the number of trigger counts for the reverse scan
llngTriggerCnt = TriggerCnt / 2

'*** Check number of triggers to see if data is valid ***
'Note:  Need to convert to long because the types are different
If llngTriggerCnt <> CLng(((gudtMachine.scanEnd - gudtMachine.scanStart) * gsngResolution) + 1) Then
    gintAnomaly = 109
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Incorrect # of Triggers counted during the Reverse Force Scan.", True, True)
End If

If gintAnomaly = 0 Then
    'Fill the Reverse arrays in reverse order by using ((gintMaxData - 1) - i) as the index
    For i = 0 To (gintMaxData - 1)
        'Fill Force Array
        gintReverse(CHAN2, (gintMaxData - 1) - i) = (BinaryData(0, i) + BinaryData(1, i) + BinaryData(2, i)) \ 3
    Next i
End If

gblnAnalogDone = True

End Sub

Private Sub tmrKillTime_Timer()
'
'   PURPOSE:   Event triggered when when timer, tmrKillTime, is complete.
'
'  INPUT(S):   None
' OUTPUT(S):   None
    
mblnKillTimeDone = True

End Sub

Public Sub OffPort1(Port As Integer, Bits As Integer)
'
'   PURPOSE:   Turns "OFF" bit locations selected (SET) & leaves other bits
'              unchanged on DIO board #1.
'
'  INPUT(S):   Port : Digital I/O Port
'              Bits : Bit locations selected for digital I/O Port
' OUTPUT(S):   None

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim Data As Variant
Dim Value As Integer

If InStr(command$, "NOHARDWARE") = 0 Then
    cwDIO1.Ports.Item(Port).SingleRead Data
    Value = Data And (Not Bits)
    cwDIO1.Ports.Item(Port).SingleWrite Value
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 122
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in OffPort1: " & Err.Description, True, True)

End Sub

Public Sub OffPort2(Port As Integer, Bits As Integer)
'
'   PURPOSE:   Turns "OFF" bit locations selected (SET) & leaves other bits
'              unchanged on DIO Board #2.
'
'  INPUT(S):   Port : Digital I/O Port
'              Bits : Bit locations selected for digital I/O Port
' OUTPUT(S):   None

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim Data As Variant
Dim Value As Integer

If InStr(command$, "NOHARDWARE") = 0 Then
    cwDIO2.Ports.Item(Port).SingleRead Data
    Value = Data And (Not Bits)
    cwDIO2.Ports.Item(Port).SingleWrite Value
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 123
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in OffPort2: " & Err.Description, True, True)

End Sub

Public Sub OnPort1(Port As Integer, Bits As Integer)
'
'   PURPOSE:   Turns "ON" bit locations selected (SET) & leaves other bits
'              unchanged on DIO Board #1.
'
'  INPUT(S):   Port : Digital I/O Port
'              Bits : Bit locations selected for digital I/O Port
' OUTPUT(S):   None

Dim Data As Variant
Dim Value As Integer

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    cwDIO1.Ports.Item(Port).SingleRead Data
    Value = Data Or Bits
    cwDIO1.Ports.Item(Port).SingleWrite Value
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 124
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in OnPort1: " & Err.Description, True, True)

End Sub

Public Sub OnPort2(Port As Integer, Bits As Integer)
'
'   PURPOSE:   Turns "ON" bit locations selected (SET) & leaves other bits
'              unchanged on DIO board #2.
'
'  INPUT(S):   Port : Digital I/O Port
'              Bits : Bit locations selected for digital I/O Port
' OUTPUT(S):   None

Dim Data As Variant
Dim Value As Integer

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    cwDIO2.Ports.Item(Port).SingleRead Data
    Value = Data Or Bits
    cwDIO2.Ports.Item(Port).SingleWrite Value
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 125
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in OnPort2: " & Err.Description, True, True)

End Sub

Public Sub ScanDAQSetup()
'
'   PURPOSE:    Initializes the data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    'Setup the device number for the DAQ board
    cwaiFDAQ.Device = 1                                             'Device number of DAQ card
    cwaiFTO.Device = 1                                              'Device number of DAQ card  '1.8ANM
    cwaiRDAQ.Device = 1                                             'Device number of DAQ card
    cwaiRFTO.Device = 1                                             'Device number of DAQ card  '1.8ANM
    cwaiPreScanDAQ.Device = 1                                       'Device number of DAQ card

    'Setup the buffer information for the A/D (Forward DAQ)
    cwaiFDAQ.AllocationMode = cwaiHostMemory                        'Specifies internal buffer
    cwaiFDAQ.AcquiredDataEnabled = True                             'Enables AcquiredData event
    cwaiFDAQ.ReturnDataType = cwaiScaledDataAndBinaryCodes          'Specifies data as unscaled

    'Setup the buffer information for the A/D (Forward DAQ)                                     '1.8ANM \/\/
    cwaiFTO.AllocationMode = cwaiHostMemory                         'Specifies internal buffer
    cwaiFTO.AcquiredDataEnabled = True                              'Enables AcquiredData event
    cwaiFTO.ReturnDataType = cwaiScaledDataAndBinaryCodes           'Specifies data as unscaled '1.8ANM /\/\

    'Setup the buffer information for the A/D (Reverse DAQ)
    cwaiRDAQ.AllocationMode = cwaiHostMemory                        'Specifies internal buffer
    cwaiRDAQ.AcquiredDataEnabled = True                             'Enables AcquiredData event
    cwaiRDAQ.ReturnDataType = cwaiScaledDataAndBinaryCodes          'Specifies data as unscaled

    'Setup the buffer information for the A/D (Reverse DAQ)                                     '1.8ANM \/\/
    cwaiRFTO.AllocationMode = cwaiHostMemory                        'Specifies internal buffer
    cwaiRFTO.AcquiredDataEnabled = True                             'Enables AcquiredData event
    cwaiRFTO.ReturnDataType = cwaiScaledDataAndBinaryCodes          'Specifies data as unscaled '1.8ANM /\/\

    'Setup the buffer information for the A/D (PreScan DAQ)
    cwaiPreScanDAQ.AllocationMode = cwaiHostMemory                  'Specifies internal buffer
    cwaiPreScanDAQ.AcquiredDataEnabled = True                       'Enables AcquiredData event
    cwaiPreScanDAQ.ReturnDataType = cwaiScaledDataAndBinaryCodes    'Specifies data as unscaled

    'Setup the clock information for the A/D  (Forward DAQ)
    cwaiFDAQ.ScanClock.ClockSourceType = cwaiPFILoToHiCS            'Use low to high PFI transition for clock
    cwaiFDAQ.ScanClock.ClockSourceSignal = "PFI0"                   'Specifies source of clock
    cwaiFDAQ.ChannelClock.ClockSourceType = cwaiInternalCS          'Use internal clock
    cwaiFDAQ.ChannelClock.Period = 0.00001                          'Specifies interchannel delay (sec)
    cwaiFDAQ.HoldoffClock.ClockSourceType = cwaiNIDAQChoosesCS      'NI-DAQ selects clock source

    'Setup the clock information for the A/D  (Force Test Only)                                                 '1.8ANM \/\/
    cwaiFTO.ScanClock.ClockSourceType = cwaiPFILoToHiCS             'Use low to high PFI transition for clock
    cwaiFTO.ScanClock.ClockSourceSignal = "PFI0"                    'Specifies source of clock
    cwaiFTO.ChannelClock.ClockSourceType = cwaiInternalCS           'Use internal clock
    cwaiFTO.ChannelClock.Period = 0.00001                           'Specifies interchannel delay (sec)
    cwaiFTO.HoldoffClock.ClockSourceType = cwaiNIDAQChoosesCS       'NI-DAQ selects clock source                '1.8ANM /\/\

    'Setup the clock information for the A/D  (Reverse DAQ)
    cwaiRDAQ.ScanClock.ClockSourceType = cwaiPFILoToHiCS            'Use low to high PFI transition for clock
    cwaiRDAQ.ScanClock.ClockSourceSignal = "PFI0"                   'Specifies source of clock
    cwaiRDAQ.ChannelClock.ClockSourceType = cwaiInternalCS          'Use internal clock
    cwaiRDAQ.ChannelClock.Period = 0.00001                          'Specifies interchannel delay (sec)
    cwaiRDAQ.HoldoffClock.ClockSourceType = cwaiNIDAQChoosesCS      'NI-DAQ selects clock source

    'Setup the clock information for the A/D  (Reverse DAQ)                                                     '1.8ANM \/\/
    cwaiRFTO.ScanClock.ClockSourceType = cwaiPFILoToHiCS            'Use low to high PFI transition for clock
    cwaiRFTO.ScanClock.ClockSourceSignal = "PFI0"                   'Specifies source of clock
    cwaiRFTO.ChannelClock.ClockSourceType = cwaiInternalCS          'Use internal clock
    cwaiRFTO.ChannelClock.Period = 0.00001                          'Specifies interchannel delay (sec)
    cwaiRFTO.HoldoffClock.ClockSourceType = cwaiNIDAQChoosesCS      'NI-DAQ selects clock source                '1.8ANM /\/\

    'Setup the clock information for the A/D  (PreScan DAQ)
    cwaiPreScanDAQ.ScanClock.ClockSourceType = cwaiPFILoToHiCS       'Use low to high PFI transition for clock
    cwaiPreScanDAQ.ScanClock.ClockSourceSignal = "PFI0"              'Specifies source of clock
    cwaiPreScanDAQ.ChannelClock.ClockSourceType = cwaiInternalCS     'Use internal clock
    cwaiPreScanDAQ.ChannelClock.Period = 0.00001                     'Specifies interchannel delay (sec)
    cwaiPreScanDAQ.HoldoffClock.ClockSourceType = cwaiNIDAQChoosesCS 'NI-DAQ selects clock source

    'Setup the condition information for the A/D (Forward DAQ)
    cwaiFDAQ.StartCondition.Type = cwaiHWDigital                    'Start immediately
    cwaiFDAQ.StartCondition.Mode = cwaiRising                       'Specifies trigger on rising edge
    cwaiFDAQ.StartCondition.Source = "PFI0"                         'Specifies source of trigger
    cwaiFDAQ.StopCondition.Type = cwaiNoActiveCondition             'Stop when all scans acquired
    cwaiFDAQ.Channels.RemoveAll                                     'Remove all channels
    cwaiFDAQ.Channels.Add ("0,0,0,1,1,1,4,4,4,7,7,7")               'Setup the channel string

    'Setup the condition information for the A/D (Forward DAQ)                                         '1.8ANM \/\/
    cwaiFTO.StartCondition.Type = cwaiHWDigital                     'Start immediately
    cwaiFTO.StartCondition.Mode = cwaiRising                        'Specifies trigger on rising edge
    cwaiFTO.StartCondition.Source = "PFI0"                          'Specifies source of trigger
    cwaiFTO.StopCondition.Type = cwaiNoActiveCondition              'Stop when all scans acquired
    cwaiFTO.Channels.RemoveAll                                      'Remove all channels
    cwaiFTO.Channels.Add ("4,4,4")                                  'Setup the channel string          '1.8ANM /\/\

    'Setup the condition information for the A/D (Reverse DAQ)
    cwaiRDAQ.StartCondition.Type = cwaiHWDigital                    'Start immediately
    cwaiRDAQ.StartCondition.Mode = cwaiRising                       'Specifies trigger on rising edge
    cwaiRDAQ.StartCondition.Source = "PFI0"                         'Specifies source of trigger
    cwaiRDAQ.StopCondition.Type = cwaiNoActiveCondition             'Stop when all scans acquired
    cwaiRDAQ.Channels.RemoveAll                                     'Remove all channels
    cwaiRDAQ.Channels.Add ("0,0,0,1,1,1,4,4,4,7,7,7")               'Setup the channel string

    'Setup the condition information for the A/D (Reverse DAQ)                                           '1.8ANM \/\/
    cwaiRFTO.StartCondition.Type = cwaiHWDigital                      'Start immediately
    cwaiRFTO.StartCondition.Mode = cwaiRising                         'Specifies trigger on rising edge
    cwaiRFTO.StartCondition.Source = "PFI0"                           'Specifies source of trigger
    cwaiRFTO.StopCondition.Type = cwaiNoActiveCondition               'Stop when all scans acquired
    cwaiRFTO.Channels.RemoveAll                                       'Remove all channels
    cwaiRFTO.Channels.Add ("4,4,4")                                   'Setup the channel string          '1.8ANM /\/\
    
    'Setup the condition information for the A/D (Prescan DAQ)
    cwaiPreScanDAQ.StartCondition.Type = cwaiHWDigital              'Start immediately
    cwaiPreScanDAQ.StartCondition.Mode = cwaiRising                 'Specifies trigger on rising edge
    cwaiPreScanDAQ.StartCondition.Source = "PFI0"                   'Specifies source of trigger
    cwaiPreScanDAQ.StopCondition.Type = cwaiNoActiveCondition       'Stop when all scans acquired
    cwaiPreScanDAQ.Channels.RemoveAll                               'Remove all channels
    cwaiPreScanDAQ.Channels.Add ("4,4,4")                           'Setup the channel string

End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 126
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ScanDAQSetup: " & Err.Description, True, True)

End Sub

Public Sub DIO1_Setup()
'
'   PURPOSE:    Initializes the digital I/O properties for the Position
'               Trigger Board (DIO Board #1).
'
'  INPUT(S):    None
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    'Setup the device number for DIO board #1
    cwDIO1.Device = 2
    'Configure Port 0 for OUTPUT
    cwDIO1.Ports.Item(0).Assignment = cwdioOutput
    'Configure Port 1 for OUTPUT
    cwDIO1.Ports.Item(1).Assignment = cwdioOutput
    'Configure Port 2 for OUTPUT
    cwDIO1.Ports.Item(2).Assignment = cwdioOutput
    'Configure Port 3 for OUTPUT
    cwDIO1.Ports.Item(3).Assignment = cwdioOutput
    'Configure Port 4 for OUTPUT
    cwDIO1.Ports.Item(4).Assignment = cwdioOutput
    'Configure Port 5 for INPUT
    cwDIO1.Ports.Item(5).Assignment = cwdioInput
    'Configure Port 6 for OUTPUT
    cwDIO1.Ports.Item(6).Assignment = cwdioOutput
    'Configure Port 7 for OUTPUT
    cwDIO1.Ports.Item(7).Assignment = cwdioOutput
    'Configure Port 8 for OUTPUT
    cwDIO1.Ports.Item(8).Assignment = cwdioOutput
    'Configure Port 9 for INPUT
    cwDIO1.Ports.Item(9).Assignment = cwdioInput
    'Configure Port 10 for INPUT
    cwDIO1.Ports.Item(10).Assignment = cwdioInput
    'Configure Port 11 for INPUT
    cwDIO1.Ports.Item(11).Assignment = cwdioInput
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 127
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in DIO1_Setup: " & Err.Description, True, True)

End Sub

Public Sub DIO2_Setup()
'
'   PURPOSE:    Initializes the digital I/O properties for the Digitizer and
'               PLC (DIO Board #2).
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    'Set CWDIO2 as device #3
    frmDAQIO.cwDIO2.Device = 3
    'Assign Port 0 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(0).Assignment = cwdioOutput
    'Assign Port 1 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(1).Assignment = cwdioOutput
    'Assign Port 2 as INPUT
    frmDAQIO.cwDIO2.Ports.Item(2).Assignment = cwdioOutput
    'Assign Port 3 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(3).Assignment = cwdioOutput
    'Assign Port 4 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(4).Assignment = cwdioOutput
    'Assign Port 5 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(5).Assignment = cwdioOutput
    'Assign Port 6 as INPUT
    frmDAQIO.cwDIO2.Ports.Item(6).Assignment = cwdioInput
    'Assign Port 7 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(7).Assignment = cwdioOutput
    'Assign Port 8 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(8).Assignment = cwdioOutput
    'Assign Port 9 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(9).Assignment = cwdioOutput
    'Assign Port 10 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(10).Assignment = cwdioOutput
    'Assign Port 11 as OUTPUT
    frmDAQIO.cwDIO2.Ports.Item(11).Assignment = cwdioOutput
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 128
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in DIO2_Setup: " & Err.Description, True, True)

End Sub

Public Sub Force_Setup()
'
'   PURPOSE:    Initializes the force data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    'Setup the device number and channel string
    cwaiForce.Device = 1                                   'Device number of DAQ card
    cwaiForce.Channels.RemoveAll                           'Clear all channels
    cwaiForce.Channels.Add ("4,4,4,4,4,4,4,4,4,4")         'Add the desired channels
    cwaiForce.ReturnDataType = cwaiScaledData              'Set the controle to return ScaledData
    'Reset the data acquisition control to free resources (memory)
    cwaiForce.Reset

End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 129
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in Force_Setup: " & Err.Description, True, True)

End Sub

Public Sub MonitorDAQRead(Voltages() As Single)
'
'   PURPOSE: Reads channels 0-7
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim lvntVoltage As Variant
Dim lintIndex As Integer

'Reset the Monitor DAQ Control
cwaiMonitorDAQ.Reset
'Read Channels 0-7
Call cwaiMonitorDAQ.SingleRead(lvntVoltage)

'Average the three readings from each channel
For lintIndex = CHAN0 To CHAN7
    Voltages(lintIndex) = (lvntVoltage(3 * lintIndex) + lvntVoltage((3 * lintIndex) + 1) + lvntVoltage((3 * lintIndex) + 2)) / 3
Next lintIndex

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 130
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in MonitorDAQRead: " & Err.Description, True, True)

End Sub

Public Sub MonitorDAQSetup()
'
'   PURPOSE:    Initializes the monitor DAQ data acquisition control properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    cwaiMonitorDAQ.Device = 1                           'DAQ Board is device #1

    cwaiMonitorDAQ.Channels.RemoveAll                    'Empty the channel string
    cwaiMonitorDAQ.Channels.Add ("0,0,0,1,1,1,2,2,2,3,3,3,4,4,4,5,5,5,6,6,6,7,7,7")  'Set up the channel string

    cwaiMonitorDAQ.ReturnDataType = cwaiScaledData       'Return data is Scaled Data

End If
    
Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 131
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in MonitorDAQSetup: " & Err.Description, True, True)
    
End Sub

Public Sub PeakForceDAQSetup()
'
'   PURPOSE:    Initializes the peak force data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    cwaiPeakForce.Device = 1                            'DAQ Board is device #1

    cwaiPeakForce.Channels.RemoveAll                    'Empty the channel string
    cwaiPeakForce.Channels.Add ("5,5,5,5,5,5,5,5,5,5")  'Set up the channel string

    cwaiPeakForce.ReturnDataType = cwaiBinaryCodes      'Return data is binary counts

End If
    
Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 132
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in PeakForceDAQSetup: " & Err.Description, True, True)
    
End Sub

Public Sub VRefDAQSetup()
'
'   PURPOSE:    Initializes the VRef data acquisition control properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    cwaiVRef.Device = 1                             'DAQ Board is device #1

    cwaiVRef.Channels.RemoveAll                     'Empty the channel string
    cwaiVRef.Channels.Add ("7,7,7,7,7,7,7,7,7,7")   'Set up the channel string

    cwaiVRef.ReturnDataType = cwaiScaledData        'Return data is Scaled Data

End If
    
Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 133
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in VRefDAQSetup: " & Err.Description, True, True)
    
End Sub

Public Sub VoutDAQSetup()
'
'   PURPOSE:    Initializes the Vout data acquisition properties.
'
'  INPUT(S):    None
'
' OUTPUT(S):    None

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then

    cwaiVOut.Device = 1                             'DAQ Board is device #1

    cwaiVOut.Channels.RemoveAll                     'Empty the channel string
    cwaiVOut.Channels.Add ("0,0,0,1,1,1,7,7,7")     'Set up the channel string

    cwaiVOut.ReturnDataType = cwaiBinaryCodes       'Return data is binary counts

End If
    
Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 134
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in VoutDAQSetup: " & Err.Description, True, True)
    
End Sub

Public Function ReadDIOLine1(PortNum As Integer, LineNum As Integer) As Boolean
'
'   PURPOSE:   Returns the value of the DIO board #1 port & line indicated.
'
'  INPUT(S):   PortNum : The port number on the DIO board to read
'              LineNum : The line number on the DIO board DIO port to read
'
' OUTPUT(S):   None

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim lvntbitState As Variant

cwDIO1.Ports.Item(PortNum).Lines.Item(LineNum).SingleRead lvntbitState

ReadDIOLine1 = CBool(lvntbitState)

Exit Function '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 135
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ReadDIOLine1: " & Err.Description, True, True)

End Function

Public Function ReadDIOLine2(PortNum As Integer, LineNum As Integer) As Boolean
'
'   PURPOSE:   Returns the value of the DIO board #2 port & line indicated.
'
'  INPUT(S):   PortNum : The port number on the DIO board to read
'              LineNum : The line number on the DIO board DIO port to read
'
' OUTPUT(S):   None

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim lvntbitState As Variant

cwDIO2.Ports.Item(PortNum).Lines.Item(LineNum).SingleRead lvntbitState

ReadDIOLine2 = CBool(lvntbitState)

Exit Function '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 136
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ReadDIOLine2: " & Err.Description, True, True)

End Function

Public Function ReadVRef() As Single
'
'   PURPOSE: Reads VRef (A/D Channel #7)
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lvntVoltage As Variant

On Error GoTo VRefReadError

'Clear data on the VRef channel
cwaiVRef.Reset

'Read the VRef Channel
cwaiVRef.SingleRead lvntVoltage

'Calculate and return the VRef
ReadVRef = cwStat1.Mean(lvntVoltage)

VRefReadError:
    
End Function

Public Sub ReadPTBoardData(ByVal LSBAddr As Integer, ByVal MSBAddr As Integer, LSBData As Variant, MIDData As Variant, MSBData As Variant)
'
'     PURPOSE:  To read data from the PT Board.
'
'    INPUT(S):  LSBAddr => LSB Address sent to the Address Bus
'               MSBAddr => MSB Address sent to the Address Bus
'   OUTPUT(S):  LSBdata => LSB Data from the Read Data bus
'               MIDdata => MID Data from the Read Data bus
'               MSBdata => MSB Data from the Read Data bus

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    Call frmDAQIO.OffPort1(PORT3, BIT0)                 'Set ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT1 + BIT2)           'Set WRbar & RDbar high
    cwDIO1.Ports.Item(0).SingleWrite LSBAddr            'Send LSB Address to Address bus
    cwDIO1.Ports.Item(1).SingleWrite MSBAddr            'Send MSB Address to Address bus
    Call frmDAQIO.OnPort1(PORT3, BIT0 + BIT1 + BIT2)    'Set WRbar, RDbar & ALE high
    Call frmDAQIO.OffPort1(PORT3, BIT0 + BIT1)          'Set RDbar & ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT2)                  'Set WRbar high
    cwDIO1.Ports.Item(9).SingleRead LSBData             'Read LSB Data to Data bus
    cwDIO1.Ports.Item(10).SingleRead MIDData            'Read MID Data to Data bus
    cwDIO1.Ports.Item(11).SingleRead MSBData            'Read MSB Data to Data bus
    Call frmDAQIO.OffPort1(PORT3, BIT0)                 'Set ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT1 + BIT2)           'Set WRbar & RDbar high
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 137
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ReadPTBoardData: " & Err.Description, True, True)

End Sub

Public Sub ReadVout(Vout1 As Single, Vout2 As Single)
'
'   PURPOSE: Reads the VOut Ch #1 & Vout Ch #2
'
'  INPUT(S): none
' OUTPUT(S): none

On Error GoTo DAQ_Err      '2.5ANM added error trap

Dim lvntVoltage As Variant
Dim llngVout1 As Long
Dim llngVout2 As Long
Dim llngVRef As Long

'Read the Peak Force Channel
cwaiVOut.SingleRead lvntVoltage

llngVout1 = (lvntVoltage(0) + lvntVoltage(1) + lvntVoltage(2)) / 3
llngVout2 = (lvntVoltage(3) + lvntVoltage(4) + lvntVoltage(5)) / 3
llngVRef = (lvntVoltage(6) + lvntVoltage(7) + lvntVoltage(8)) / 3

'Calculate Vout1
Vout1 = (llngVout1 / llngVRef) * HUNDREDPERCENT
'Calculate Vout2
Vout2 = (llngVout2 / llngVRef) * HUNDREDPERCENT

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 138
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in ReadVout: " & Err.Description, True, True)

End Sub

Public Sub WritePTBoardData(ByVal LSBAddr As Integer, ByVal MSBAddr As Integer, LSBData As Integer, MIDData As Integer, MSBData As Integer)
'
'     PURPOSE:  To write data to the PT Board (DIO Card #1).
'
'    INPUT(S):  LSBAddr => LSB Address sent to the Address Bus
'               MSBAddr => MSB Address sent to the Address Bus
'   OUTPUT(S):  LSBdata => LSB Data sent to the Write Data bus
'               MIDdata => MID Data sent to the Write Data bus
'               MSBdata => MSB Data sent to the Write Data bus

On Error GoTo DAQ_Err      '2.5ANM added error trap

If InStr(command$, "NOHARDWARE") = 0 Then
    Call frmDAQIO.OffPort1(PORT3, BIT0)                 'Set ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT1 + BIT2)           'Set WRbar & RDbar high
    cwDIO1.Ports.Item(0).SingleWrite LSBAddr            'Send LSB Address to Address bus
    cwDIO1.Ports.Item(1).SingleWrite MSBAddr            'Send MSB Address to Address bus
    cwDIO1.Ports.Item(6).SingleWrite LSBData            'Write LSB Data to Data bus
    cwDIO1.Ports.Item(7).SingleWrite MIDData            'Write MID Data to Data bus
    cwDIO1.Ports.Item(8).SingleWrite MSBData            'Write MSB Data to Data bus
    Call frmDAQIO.OnPort1(PORT3, BIT0 + BIT1 + BIT2)    'Set WRbar, RDbar & ALE high
    Call frmDAQIO.OffPort1(PORT3, BIT0 + BIT2)          'Set WRbar & ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT1)                  'Set RDbar high
    Call frmDAQIO.OffPort1(PORT3, BIT0)                 'Set ALE low
    Call frmDAQIO.OnPort1(PORT3, BIT1 + BIT2)           'Set WRbar & RDbar high
End If

Exit Sub '2.5ANM \/\/
DAQ_Err:

    gintAnomaly = 139
    'Log the error to the error log and display the error message
    Call ErrorLogFile("Software error in WritePTBoardData: " & Err.Description, True, True)

End Sub
