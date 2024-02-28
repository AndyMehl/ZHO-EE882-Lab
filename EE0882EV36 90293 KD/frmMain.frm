VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "CWUI.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "n/a"
   ClientHeight    =   10455
   ClientLeft      =   1965
   ClientTop       =   1560
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin EE0882EV36.ctrResultsTabs ctrResultsTabs1 
      Height          =   8415
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   14843
   End
   Begin MSComctlLib.StatusBar staMessage 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   9960
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16854
            Text            =   "System Message: OK!"
            TextSave        =   "System Message: OK!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Cycle Time"
            TextSave        =   "Cycle Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2/25/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "1:37 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EE0882EV36.ctrStatus ctrStatus1 
      Height          =   4815
      Left            =   13995
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   8493
   End
   Begin EE0882EV36.ctrSetupInfo ctrSetupInfo1 
      Height          =   855
      Left            =   4560
      TabIndex        =   14
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1508
   End
   Begin EE0882EV36.ctrLotSummary ctrScanSummary 
      Height          =   975
      Left            =   7680
      TabIndex        =   13
      Top             =   9330
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1720
   End
   Begin EE0882EV36.ctrLotSummary ctrProgSummary 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   9330
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1720
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2235
      Left            =   240
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   3135
      TabIndex        =   11
      Top             =   6720
      Width           =   3195
   End
   Begin VB.TextBox txtPosition 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   9000
      Width           =   1215
   End
   Begin CWUIControlsLib.CWKnob CWPosition 
      Height          =   1575
      Left            =   14040
      TabIndex        =   9
      Top             =   7320
      Width           =   1215
      _Version        =   393218
      _ExtentX        =   2143
      _ExtentY        =   2778
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   393218
      Radial_0        =   1
      ClassName_1     =   "CCWKnob"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Position"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   52
      SclRef_10.l     =   12
      SclRef_10.t     =   26
      SclRef_10.r     =   73
      SclRef_10.b     =   97
      Scl_10.l        =   41
      Scl_10.t        =   60
      Scl_10.r        =   43
      Scl_10.b        =   62
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   91136
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   122880
      rMax_14         =   12288
      dMax_14         =   360
      discInterval_14 =   1
      Radial_12       =   15
      ClassName_15    =   "CCWRadial"
      Center_15.cx    =   40
      Center_15.cy    =   59
      start_15        =   1.5707963267949
      length_15       =   6.26573201465964
      radius_15       =   31
      Enum_12         =   16
      ClassName_16    =   "CCWEnum"
      Editor_16       =   17
      ClassName_17    =   "CCWEnumArrayEditor"
      Owner_17        =   12
      Font_12         =   0
      tickopts_12     =   2702
      major_12        =   90
      minor_12        =   45
      Caption_12      =   18
      ClassName_18    =   "CCWDrawObj"
      opts_18         =   62
      C[0]_18         =   -2147483640
      Image_18        =   19
      ClassName_19    =   "CCWTextImage"
      style_19        =   6
      font_19         =   0
      Animator_18     =   0
      Blinker_18      =   0
      DrawLst_1       =   20
      ClassName_20    =   "CDrawList"
      count_20        =   7
      list[7]_20      =   8
      list[6]_20      =   21
      ClassName_21    =   "CCWNeedle"
      opts_21         =   63
      Name_21         =   "Pointer-1"
      C[0]_21         =   -2147483640
      C[2]_21         =   -2147483635
      Image_21        =   22
      ClassName_22    =   "CCWPictImage"
      opts_22         =   1280
      Rows_22         =   1
      Cols_22         =   1
      Pict_22         =   286
      B_22            =   -2147483633
      ColorReplaceWith_22=   8421504
      ColorReplace_22 =   8421504
      Tolerance_22    =   2
      Animator_21     =   0
      Blinker_21      =   0
      style_21        =   5
      mode_21         =   1
      Len_21          =   31
      Radl_21         =   15
      InsideFillBnd_21=   23
      ClassName_23    =   "CCWGuiObject"
      opts_23         =   52
      SclRef_23.l     =   12
      SclRef_23.t     =   26
      SclRef_23.r     =   73
      SclRef_23.b     =   97
      Scl_23.l        =   11
      Scl_23.t        =   30
      Scl_23.r        =   73
      Scl_23.b        =   92
      list[5]_20      =   12
      list[4]_20      =   6
      list[3]_20      =   11
      list[2]_20      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   62
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   19
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[1]_20      =   2
      Ptrs_1          =   26
      ClassName_26    =   "CCWPointerArray"
      Array_26        =   1
      Editor_26       =   27
      ClassName_27    =   "CCWPointerArrayEditor"
      Owner_27        =   1
      Array[0]_26     =   21
      Bindings_1      =   28
      ClassName_28    =   "CCWBindingHolderArray"
      Editor_28       =   29
      ClassName_29    =   "CCWBindingHolderArrayEditor"
      Owner_29        =   1
      Stats_1         =   30
      ClassName_30    =   "CCWStats"
      Knob_1          =   24
      InsideFill_1    =   23
      radial_1        =   15
      labOffset_1     =   0
   End
   Begin VB.TextBox txtTrigCount 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Timer tmrPollPLC_IO 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   14880
      Top             =   10200
   End
   Begin VB.FileListBox filLotFile 
      Height          =   300
      Left            =   2400
      Pattern         =   "*.lot"
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame fraLotFile 
      Caption         =   "Lot File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboLotFile 
         Height          =   330
         ItemData        =   "frmMain.frx":163F6
         Left            =   60
         List            =   "frmMain.frx":163F8
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.FileListBox filParameterFiles 
      Height          =   300
      Left            =   240
      Pattern         =   "*.csv"
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fraParameterFile 
      Caption         =   "Parameter File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboParameterFileName 
         Height          =   330
         ItemData        =   "frmMain.frx":163FA
         Left            =   60
         List            =   "frmMain.frx":163FC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Parameter File"
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   14760
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line line 
      Index           =   0
      X1              =   7680
      X2              =   7695
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Label lblTriggerCounts 
      Alignment       =   2  'Center
      Caption         =   "Trigger Counts"
      Height          =   495
      Index           =   1
      Left            =   14040
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "&Function"
      Begin VB.Menu mnuFunctionTest 
         Caption         =   "&Test"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFunctionProgram 
         Caption         =   "&Program"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFunctionProgramAndTest 
         Caption         =   "Program &And Test"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFunctionSampleNum 
         Caption         =   "&Sample Number"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFunctionCustomerName 
         Caption         =   "Customer Name"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFunctionCustomerPartNumber 
         Caption         =   "Customer Part Number"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuFunctionCTSPartNumber 
         Caption         =   "CTS Part Number"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFunctionPartName 
         Caption         =   "Part Name"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFunctionSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionResetPartCount 
         Caption         =   "&Reset Part Count"
      End
      Begin VB.Menu mnuFunctionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFunctionPrintOpt 
            Caption         =   "Scan Results"
            Index           =   1
         End
         Begin VB.Menu mnuFunctionPrintOpt 
            Caption         =   "Programming Results"
            Index           =   2
         End
         Begin VB.Menu mnuFunctionPrintOpt 
            Caption         =   "Scan Graphs"
            Index           =   3
         End
         Begin VB.Menu mnuFunctionPrintOpt 
            Caption         =   "Scan Statistics"
            Index           =   4
         End
         Begin VB.Menu mnuFunctionPrintOpt 
            Caption         =   "Programming Statistics"
            Index           =   5
         End
      End
      Begin VB.Menu mnuFunctionAutoPrint 
         Caption         =   "&Auto Print"
         Begin VB.Menu mnuFunctionAutoPrintProgResults 
            Caption         =   "&Programming Results"
         End
         Begin VB.Menu mnuFunctionAutoPrintScanResults 
            Caption         =   "&Scan Results"
         End
         Begin VB.Menu mnuFunctionAutoPrintGraphs 
            Caption         =   "&ScanGraphs"
         End
         Begin VB.Menu mnuFunctionAutoSavePDFs 
            Caption         =   "&Auto Save PDFs"
         End
      End
      Begin VB.Menu mnuFunctionPrintPreview 
         Caption         =   "Prin&t Preview"
         Begin VB.Menu mnuFunctionPreviewOpt 
            Caption         =   "Scan Results"
            Index           =   0
         End
         Begin VB.Menu mnuFunctionPreviewOpt 
            Caption         =   "Programming Results"
            Index           =   1
         End
         Begin VB.Menu mnuFunctionPreviewOpt 
            Caption         =   "Scan Statistics"
            Index           =   2
         End
         Begin VB.Menu mnuFunctionPreviewOpt 
            Caption         =   "Programming Stats"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFunctionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionHomeMotor 
         Caption         =   "Home Motor"
      End
      Begin VB.Menu mnuFunctionInitializeSensotec 
         Caption         =   "Initialize Force Amplifier Communication"
      End
      Begin VB.Menu mnuFunctionInitializeProgrammers 
         Caption         =   "Initialize Programmer Communications"
      End
      Begin VB.Menu mnuFunctionResetVRef 
         Caption         =   "Reset Voltage Reference"
      End
      Begin VB.Menu mnuFunctionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTab 
         Caption         =   "Scan Results"
         Index           =   0
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "Programming Results"
         Index           =   1
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "Scan Graphs"
         Index           =   2
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "Scan Statistics"
         Index           =   3
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "Programming Statistics"
         Index           =   4
      End
      Begin VB.Menu mnuViewSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewParameterFile 
         Caption         =   "&Parameter File"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsMonitorDAQ 
         Caption         =   "Monitor &Data Acquisition"
      End
      Begin VB.Menu mnuToolsProgrammerUtility 
         Caption         =   "&Programmer Utility"
      End
      Begin VB.Menu mnuMLXVI 
         Caption         =   "Melexis &VI "
      End
      Begin VB.Menu mnuToolsVIX500IEUtility 
         Caption         =   "&Motor Test Utility"
      End
      Begin VB.Menu mnuToolsSensotec 
         Caption         =   "Sensotec Force Amplifier Test Utility"
      End
      Begin VB.Menu mnuToolsDDE 
         Caption         =   "PLC DDE Communication Utility"
      End
      Begin VB.Menu mnuToolsCycle 
         Caption         =   "Cycle Part"
      End
      Begin VB.Menu mnuToolsMotor 
         Caption         =   "Move to Location"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAutoPrintFailureResults 
         Caption         =   "Auto-Print Failure Results"
      End
      Begin VB.Menu mnuOptionsPrintType 
         Caption         =   "Print Type"
         Begin VB.Menu mnuOptionsPrintTypeTestlab 
            Caption         =   "Test Lab"
         End
         Begin VB.Menu mnuOptionsPrintTypeCustomer 
            Caption         =   "Customer"
         End
      End
      Begin VB.Menu mnuOptionsSaveRawData 
         Caption         =   "Save Raw Data"
      End
      Begin VB.Menu mnuOptionsScanResultsData 
         Caption         =   "Save Scan Results Data"
      End
      Begin VB.Menu mnuOptionsProgResultsData 
         Caption         =   "Save Programming Results Data"
      End
      Begin VB.Menu mnuOptionsGraphEnable 
         Caption         =   "Enable Graphs"
      End
      Begin VB.Menu mnuOptionsLock 
         Caption         =   "Lock Chip after Programming"
      End
      Begin VB.Menu mnuOptionsLockRejects 
         Caption         =   "Lock Reject Parts"
      End
      Begin VB.Menu mnuOptionFTO 
         Caption         =   "Force Test Only"
      End
      Begin VB.Menu mnuOptionBMT 
         Caption         =   "Benchmark Test"
      End
      Begin VB.Menu mnuOptionMSE 
         Caption         =   "Multi-Scan Enable"
      End
      Begin VB.Menu mnuOptionsPLCStart 
         Caption         =   "On PLC Start"
         Begin VB.Menu mnuOptionsPLCStartProgram 
            Caption         =   "Program"
         End
         Begin VB.Menu mnuOptionsPLCStartScan 
            Caption         =   "Scan"
         End
      End
      Begin VB.Menu mnuOptionsSlow 
         Caption         =   "Slow Scan"
         Begin VB.Menu mnuOptionsSSN 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuOptionsSSD 
            Caption         =   "Dual"
         End
      End
   End
   Begin VB.Menu mnuExp 
      Caption         =   "&Exposure"
      Begin VB.Menu mnuExpDust 
         Caption         =   "&Dust"
      End
      Begin VB.Menu mnuExpVibration 
         Caption         =   "&Vibration"
      End
      Begin VB.Menu mnuExpDither 
         Caption         =   "D&ither"
      End
      Begin VB.Menu mnuExpThermalShock 
         Caption         =   "&Thermal Shock"
      End
      Begin VB.Menu mnuExpSaltSpray 
         Caption         =   "Salt S&pray"
      End
      Begin VB.Menu mnuExpInitial 
         Caption         =   "Initial"
      End
      Begin VB.Menu mnuExpExposure 
         Caption         =   "Exposure"
      End
      Begin VB.Menu mnuExpStrength 
         Caption         =   "Strength"
         Begin VB.Menu mnuExpStrengthOperational 
            Caption         =   "Operational"
         End
         Begin VB.Menu mnuExpStrengthLateral 
            Caption         =   "Lateral"
         End
         Begin VB.Menu mnuExpStrengthOpwithStop 
            Caption         =   "Operational w/Stopper"
         End
         Begin VB.Menu mnuExpStrengthImpact 
            Caption         =   "Impact"
         End
      End
      Begin VB.Menu mnuExpOperationalEndurance 
         Caption         =   "Operational Endurance"
      End
      Begin VB.Menu mnuExpSnapback 
         Caption         =   "Snapback"
      End
      Begin VB.Menu mnuExpHighTemp 
         Caption         =   "High Temp Soak"
      End
      Begin VB.Menu mnuExpHighTempHighHumidity 
         Caption         =   "High Temp - High Humidity Soak"
      End
      Begin VB.Menu mnuExpLowTemp 
         Caption         =   "Low Temp Soak"
      End
      Begin VB.Menu mnuExpWaterSpray 
         Caption         =   "Water Spray"
      End
      Begin VB.Menu mnuExpChemicalResistance 
         Caption         =   "Chemical Resistance"
      End
      Begin VB.Menu mnuExpCondensation 
         Caption         =   "Condensation"
      End
      Begin VB.Menu mnuExpElectrical 
         Caption         =   "Electrical"
         Begin VB.Menu mnuExpElecElectroStaticDischarge 
            Caption         =   "ElectroStatic Discharge"
         End
         Begin VB.Menu mnuExpElecEMWaveResistance 
            Caption         =   "EM Wave Resistance"
         End
         Begin VB.Menu mnuExpElecBilkCurrentInjection 
            Caption         =   "Bilk Current Injection"
         End
         Begin VB.Menu mnuExpElecIgnitionNoise 
            Caption         =   "Ignition Noise"
         End
         Begin VB.Menu mnuExpElecNarrowbandRadiatedEMEnergy 
            Caption         =   "Narrowband Radiated EME"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About!"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********  705 Series Test Lab Programming & Scanning System *********
'
'   Andrew N. Mehl
'   CTS Automotive
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (574) 295-3575
'
'   Originally written by Andrew N Mehl during October of
' 2005, based on EE0882BV10, the 702 Test Lab Prog & Scan System.
' This system uses the Position Trigger Board (CC256) developed by CTS
' Elkhart Electronics Engineering Group for the 5th Generation system.
'
' Performs programming of 705 Series parts using MLX 90277 ICs:
'   Offset                          Calibrate/Program/Verify/Record
'   Rough Gain                      Calibrate/Program/Verify/Record
'   Fine Gain                       Calibrate/Program/Verify/Record
'   Filter                                   /Program/Verify/
'   Invert                                   /Program/Verify/
'   Mode                                     /Program/Verify/
'   FaultLevel                               /Program/Verify/
'   ClampHigh                       Calibrate/Program/Verify/Record
'   ClampLow                        Calibrate/Program/Verify/Record
'   CustomerID                      Calibrate/Program/Verify/Record
'   Offset Drift                             /       /Verify/
'   AGND                                     /       /Verify/
'   Oscillator Adjust                        /       /Verify/
'   Capacitor Frequency Adjust               /       /Verify/
'   DAC Frequency Adjust                     /       /Verify/
'   Slow                                     /       /Verify/
'
' Performs testing of 705 Series parts:
'   Voltage Gradient        Outputs 1 & 2       /       /Graph/Record
'   Index 1 (Full-Close)    Outputs 1 & 2   Test/Display/     /Record
'   Index 2 (Mid-Point)     Outputs 1 & 2   Test/Display/     /Record
'   Index 3 (Full-Open)     Outputs 1 & 2   Test/Display/     /Record
'   Maximum Output          Outputs 1 & 2   Test/Display/     /Record
'   Linearity Deviation     Outputs 1 & 2   Test/Display/Graph/Record
'   Slope Deviation         Outputs 1 & 2   Test/Display/Graph/Record
'   Hysteresis              Outputs 1 & 2       /       /Graph/Record
'   Forward Output Correlation              Test/Display/Graph/Record
'   Reverse Output Correlation              Test/Display/Graph/Record
'   Pedal-At-Rest Location                      /Display/     /Record
'   Pedal Force                             Test/Display/Graph/Record
'   Mechanical Hysteresis                   Test/Display/Graph/Record
'
'   This program consists of the following components:  form module(s),
' standard module(s) and user control(s).  A form module contains
' event procedures (i.e. sections of code that will execute in response to
' specific events).  Forms can contain any number of controls or graphical
' data for display.  All the forms for this program are contained in the
' Forms folder.  A standard module contains only code.  Standard modules
' were previously called code modules.  All the standard modules for this
' program are contained in the Modules folder.  The user controls are
' components built for use on the GUI.
'
'   The subroutines found in this program are designed to be modular and
' reuseable code segments.  The intent here is to reduce design time for
' future programs.  Specifically, the Calc.bas module, the MLX90277.bas
' module, the VIX500IE.bas module, and the Sensotec.bas module, as well as
' the forms associated with these modules, are intended to be useable in any
' project.  They are version controlled as such.  Also, the Pedal.bas module,
' the Series705.bas module, and the Solver90277 module are intended to be
' useable within the appropriate pedal projects.  All 705 series programming
' & scanning systems should use the same modules, with minor changes to
' frmMain.frm to distingish between systems.
'
'   If you need to make a change to one of the subroutines or code segments,
' simply make the change to the source code and make the executable.
' To make the executable, click File on the menu bar and then click Make
' EE#####V***.exe under the File pulldown menu. Next, enter the correct
' version for the executable and click OK. This will compile the code and
' create the executable.
'
'   To make the code more readable, please follow the convention of
' making variable names all lower case (except for multiple word names
' like datumZero) and constant names (anything assigned with the Const
' keyword) all upper case.  It's a big help to be able to tell the
' difference between constants and variables and this is a widely used
' convention.
'
'   In addition to making variable names lower case, please follow the
' naming convention for scope and type identification.  This convention
' is taken from Microsoft's recommended standard for naming variables
' in VB.  The first letter should identify the scope of the variable.
' Use "g" for global, "m" for module level or "l" for local. The next
' three letters identify the type of the variable. Examples of this are
' "bln" for boolean, "int" for integer, and "sng" for single. Follow this
' with the name of the variable.  For instance, a global integer used to
' detect failures could be named: gintAnomaly.
'
'   Also, please help us track this program through its useful life by
' changing the version number of the program and placing a brief description
' of your modification here.  The version number of the program should
' have a digit and a decimal value (X.X).  The "ID" is a stamp that you
' should keep close to if not on same lines of any and all modifications
' you make. It should appear exactly as typed in the table so others can
' find all your modifications by clicking Edit on the menu bar and using
' the Find function.  The format that is being used for the ID is the version
' number and initials.  A hypothetical example of an ID is "1.1SRC".  Within
' the version controlled modules, follow the version control numbering system
' used there.
'
'Ver    Date     By  Purpose of modification                                  ID
'1.0   12/16/2005 ANM First release per SCN# 704T-001 (3102)                   1.0ANM
'1.1   01/02/2006 ANM Updates per SRC.                                         1.1ANM
'1.2   01/13/2006 ANM Updated pedal.bas for 704 items, added testlab features, 1.2ANM
'                     change 704 to 705, change slope to have two slopes, and
'                     updated single pt. lin per SCN# 704T-003 (3258).
'1.3   01/20/2006 ANM Update to check force calib. per SCN# 705T-001 (3296).   1.3ANM
'1.4   01/31/2006 ANM Update items in excel file headers for customer.         1.4ANM
'1.5   02/22/2006 ANM Added AMAD705 to program, and updated for testlab        1.5ANM
'                     features per PR 11801-K.
'1.6   02/28/2006 ANM Update for separate print results for customer/TL        1.6ANM
'                     per PR 11801-K.
'1.7   05/03/2006 ANM Updates per SCN# MISC-092 (3365) for updates to Canada   1.7ANM
'                     and making one pedal.bas again. Also per SCN# 705T-004
'                     (3439) for scan only items.
'1.8   05/04/2006 ANM Updates for force only, rawdata, customer #, smart KD,   1.8ANM
'                     and delays per SCN# MISC-094 (3423).
'1.9   06/01/2006 ANM Updates for programming per SCN# 705T-005 (3481).        1.9ANM
'2.0   08/30/2006 ANM Updates per SCN# MISC-100 (3521).                        2.0ANM
'2.1   12/05/2006 ANM Updates per SCN# MISC-101 (3636).                        2.1ANM
'2.2   01/18/2006 ANM Updates per SCN# MISC-102 (3702).                        2.2ANM
'2.3   02/09/2007 ANM Updates per SCN# 705F-001 (3741).                        2.3ANM
'2.4   03/15/2007 ANM Updates per SCN# MISC-104 (3789).                        2.4ANM
'2.5   05/01/2007 ANM Updates per SCN# MISC-102 (3702) (more changes).         2.5ANM
'2.6   05/17/2007 ANM Updates per SCN# 705T-010 (3876).                        2.6ANM
'2.7   08/15/2007 ANM Updates per SCN# BNCH-023 (3938).                        2.7ANM
'                     New modules, benchmarking, and auto-run
'2.8   08/30/2007 ANM Updates per SCN# 705F-007 (3973).                        2.8ANM
'                     Removed force knee items.
'2.9   09/25/2007 ANM Updates to solver per SCN# 705F-008 (3979).              2.9ANM
'3.0   11/02/2007 ANM Updates per SCN# 705F-011 (4018).                        3.0ANM
'3.1   01/31/2008 ANM Updated to new modules per SCN# 4066 & 4067.             3.1ANM
'3.2   02/26/2008 ANM Update to save ZG point per SCN# 4087.                   3.2ANM
'3.3   04/29/2008 ANM Update for fixes and MLX current check per SCN# 4124.    3.3ANM
'3.3a  05/02/2008 ANM New modules for SCN# 4139.                               3.3aANM
'3.4   06/05/2008 ANM Add ABS Lin per SCN# 4167.                               3.4ANM
'3.4a  07/29/2008 ANM Update to restict commas per SCN# 4186.                  3.4aANM
'3.5   10/31/2008 ANM Update to add P@R per SCN# 4236.                         3.5ANM
'3.6   01/19/2009 ANM Update for PDF prints per SCN# 4258.                     3.6ANM
'3.6   02/09/2011 ANM Update for PDF prints D4.                                3.6ANM
'3.6a  04/16/2009 ANM Update for ENR fixes per SCN# 4317.                      3.6aANM
'3.6b  08/14/2009 ANM Update for FPFL message per SCN# 4372.                   3.6bANM
'3.6c  10/02/2009 ANM MLX VI/ABSLin M/MLX I/Checks per 4391/4392/4401/4403.    3.6cANM
'3.6d  10/13/2009 ANM Update MLX Checks per 4422.                              3.6dANM
'3.6e  11/18/2009 ANM Update slow speed & Bnmk MLX Idd per 4420/4428.          3.6eANM
'3.6f  06/25/2010 ANM Update for 90277 SSPSS per SCN# 4585.                    3.6fANM
'3.6g  01/07/2011 ANM Update for force offset per SCN# 4698.                   3.6gANM
'3.6h  01/30/2012 ANM Update for filter 0 per SCN# 4933.                       3.6hANM
'3.6i  02/24/2012 ANM Update for lin per SCN# 4955.                            3.6iANM
'3.6i  12/10/2015 ANM Update for motor control per SCN# 5275.                  3.6iANM
'3.6j  10/30/2012 ANM Update for TO per SCN# 5106.                             3.6jANM
'3.6*  03/27/2019 ANM Update for 90293 MLX.                                    3.6*ANM
'3.6** 06/19/2019 ANM Update for MLX errors.                                   3.6**ANM
'3.6l  01/07/2020 ANM Update for KD.                                           3.6lANM
'3.6k  02/25/2020 ANM Update MLX90293 forms per SCN# 6185.                     3.6kANM
'

Option Explicit

Private mintLastParameterComboBoxIndex As Integer
Private mintLastLotComboBoxIndex As Integer

Private Sub RefreshParameterFileList()
'
'   PURPOSE: To populate the file list box with parameter files which are
'            located in the Parameter subdirectory.
'
'  INPUT(S): none
' OUTPUT(S): none

'Parameter file list
filParameterFiles.Path = App.Path + PARPATH
filParameterFiles.Refresh                   'Refresh parameter file list
    
Dim i As Integer

For i = 0 To filParameterFiles.ListCount - 1
    cboParameterFileName.AddItem (filParameterFiles.List(i))
Next i

End Sub

Public Sub RampToVref()
'
'   PURPOSE: To ramp the voltage from 0V to Vref
'
'  INPUT(S): none
' OUTPUT(S): none
'2.2ANM new sub '3.6ANM fix

Dim X As Single
Dim Y As Long

'Ramp based on speed of system (delay set by gdblDelay)
For X = 0.01 To (gsngVRefSetPoint / 2) Step 0.01
    Call frmDAQIO.cwaoVRef.SingleWrite(X)
    For Y = 1 To gdblDelay
    Next Y
Next X

'Set to Vref after ramp
Call frmDAQIO.cwaoVRef.SingleWrite(gsngVRefSetPoint / 2)

End Sub

Public Sub RefreshLotFileList()
'
'   PURPOSE: To populate file list box with lot files which are located
'            in the lot file subdirectory
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lstrCurrentLotFile As String
Dim i, j As Integer
Dim ldteDateTime() As Date
Dim ldteTempDateTime As Date
Dim lintOrderNum() As Integer
Dim lintTempOrderNum As Integer
Dim lfleFile As File

filLotFile.Path = STATFILEPATH
filLotFile.Refresh                          'Refresh lot file list

lstrCurrentLotFile = cboLotFile.Text        'Save current lot file name
cboLotFile.Clear                            'Clear lot file list

'If there are any lot files, sort them by date modified
If filLotFile.ListCount <> 0 Then
    ReDim lintOrderNum(0 To filLotFile.ListCount - 1)
    ReDim ldteDateTime(0 To filLotFile.ListCount - 1)

    'Loop through obtaining the lot file date/time stamps
    For i = 0 To filLotFile.ListCount - 1
        Set lfleFile = gfsoFileSystemObject.GetFile(STATFILEPATH & filLotFile.List(i))
        ldteDateTime(i) = lfleFile.DateLastModified
        lintOrderNum(i) = i
    Next i

    'Loop through performing a bubble sort
    For i = 0 To filLotFile.ListCount - 1
        For j = 0 To filLotFile.ListCount - 2 - i
            'Order according to the latest date/time stamp
            If ldteDateTime(j + 1) > ldteDateTime(j) Then
                'Save the first item
                ldteTempDateTime = ldteDateTime(j)
                lintTempOrderNum = lintOrderNum(j)
                'Re-order the first element
                ldteDateTime(j) = ldteDateTime(j + 1)
                lintOrderNum(j) = lintOrderNum(j + 1)
                'Re-order the second element
                ldteDateTime(j + 1) = ldteTempDateTime
                lintOrderNum(j + 1) = lintTempOrderNum
            End If
        Next j
    Next i

    For i = 0 To filLotFile.ListCount - 1
        'Add item to lot file combo box according to the sorted order
        cboLotFile.AddItem (filLotFile.List(lintOrderNum(i)))
    Next i
End If
cboLotFile.Text = lstrCurrentLotFile        'Place current lot file name in the combo box

End Sub

Private Sub cboLotFile_GotFocus()
'
'   PURPOSE: To allow selection of lot file.  This subroutine will be initiated by
'            a got focus event of the combo box
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintResponse As Integer

If ((Not gblnStartUpDone) Or gblnLotFileSelected) Then
    staMessage.Panels(1).Text = "Operator Input:  Select or Enter a New Lot File & Press <ENTER> "
Else
    staMessage.Panels(1).Text = "Operator Input:  You MUST Press <ENTER> after Selecting a Lot File..."
End If

tmrPollPLC_IO.Enabled = False                       'Make sure the Timer to look for startscan is disabled
mintLastLotComboBoxIndex = cboLotFile.ListIndex     'Save the Index referring to the currently selected Lot File
gblnLotFileSelected = False                         'Indicate that no Lot File has been selected
cboParameterFileName.Enabled = False                'Disable the Parameter File box while the Lot File box has focus

End Sub

Private Sub cboLotFile_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To allow selection of lot file.  This subroutine will be initiated by
'            a key pressed event
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none

'When both parameter file and lot file selected, display the Lot file box and the
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
    'Exit if nothing in combo box
    If cboLotFile.Text = "" Then Exit Sub

    'Add the Stat File Extension if it isn't already there
    If UCase$(right(cboLotFile.Text, 4)) <> UCase(STATEXT) Then
        cboLotFile.Text = cboLotFile.Text + STATEXT
    End If

    gstrLotName = left(cboLotFile.Text, Len(cboLotFile.Text) - 4)   'Save the Lot Name

    ctrResultsTabs1.Visible = False         'Make the Results Tabs Invisible
    cboLotFile.Enabled = False              'Disable the combo box
    gblnLotFileSelected = True              'Indicate the the Lot File is loaded

    Call AMAD705.SetLot                                     'Set the database link to the Lot Name
    Call TestLab.Stats705TLLoad                             '1.7ANM  'Load the Lot file
    Call Pedal.DisplayInitialization                        'Initialize Control Grids
    Call Pedal.SummaryInitialization                        'Initialize the Summary controls
    Call Series705.InitializeAndMaskScanFailures            'Initialize and mask failures before displaying
    Call Pedal.DisplayProgStatisticsData                    'Display any Prog stats data
    Call Pedal.DisplayProgStatisticsCountsPrioritized       'Display any Prog stats data
    Call Pedal.DisplayProgResultsCountsPrioritized          'Display any Prog stats data
    Call Series705.DisplayScanStatisticsData                'Display any Scan stats data
    Call Series705.DisplayScanStatisticsCountsPrioritized   'Display any Scan stats data
    Call Series705.DisplayScanResultsCountsPrioritized      'Display any Scan stats data
    Call Pedal.DisplayProgSummary                           'Display any Prog Summary Data
    Call Series705.DisplayScanSummary                       'Display any Scan Summary Data
    Call ctrResultsTabs1.ClearData(SCANRESULTSGRID, 1, 2)   'Clear the Scan Results Control
    Call ctrResultsTabs1.ClearData(PROGRESULTSGRID, 1, 2)   'Clear the Programming Results Control
    frmDAQIO.KillTime (500)                     'Delay to keep screen from jumping
    ctrResultsTabs1.Visible = True              'Make the results control visible
    ctrStatus1.Visible = True                   'Make the status control visible
    ctrSetupInfo1.SetFocus                      'Set the setup control to have focus and prompt for user input
    cboParameterFileName.Enabled = True         'Enable the Parameter File box
    cboLotFile.Enabled = True                   'Enable the Lot File box

    staMessage.Panels(1).Text = "System Initialized.  Enter Setup Information If Desired..."

    'StartUp is done once the Parameters and Lot File have been loaded
    If Not gblnStartUpDone Then
        gblnStartUpDone = True
        'Enable the menus
        mnuFunctionTest.Enabled = True
        mnuFunctionProgram.Enabled = True
        mnuFunctionProgramAndTest.Enabled = True
        mnuFunctionHomeMotor.Enabled = True
        mnuFunctionInitializeProgrammers.Enabled = True
        mnuFunctionResetVRef.Enabled = True
        mnuToolsMonitorDAQ.Enabled = True
        mnuToolsProgrammerUtility.Enabled = True
        mnuToolsVIX500IEUtility.Enabled = True
        mnuToolsSensotec.Enabled = True
    End If

    tmrPollPLC_IO.Enabled = True        'Enable timer to poll for a start scan

Else
    'Accept only letters, numbers, & appropriate characters
    Select Case KeyAscii
        Case 8              'backspace
            'Accept the character
        Case 32 To 33       'space and !
            'Accept the character
        Case 35 To 41       '# $ % & ` ( )
            'Accept the character
        Case 43 To 45       '+ , -
            'Accept the character
        Case 48 To 57       '0-9
            'Accept the character
        Case 64 To 90       '@ and A-Z (upper case)
            'Accept the character
        Case 94 To 95       '^ and underscore
            'Accept the character
        Case 97 To 122      'a-z (lower case)
            'Accept the character
        Case Else
            KeyAscii = 0    ' Cancel the character.
            Beep            ' Sound error signal.
    End Select
End If

End Sub

Private Sub cboLotFile_LostFocus()
'
'   PURPOSE: To allow selection of lot file.  This subroutine will be initiated by
'            the lost focus event.
'
'  INPUT(S): none
' OUTPUT(S): none

If gblnLotFileSelected = False Then
    'Put the previously selected Parameter File back in the box
    cboLotFile.ListIndex = mintLastLotComboBoxIndex
    'Force the user to enter a lot file
    cboLotFile.SetFocus
End If

End Sub

Private Sub cboParameterFileName_GotFocus()
'
'   PURPOSE: To ensure that the current lot file name is displayed, and to
'            force the user to select a lot file
'
'  INPUT(S): none
' OUTPUT(S): none


If ((Not gblnStartUpDone) Or gblnParFileSelected) Then
    staMessage.Panels(1).Text = "Operator Input:  Select Parameter File & Press <ENTER>..."
Else
    staMessage.Panels(1).Text = "Operator Input:  You MUST Press <ENTER> after Selecting a Parameter File..."
End If

tmrPollPLC_IO.Enabled = False                           'Make sure the Timer to look for startscan is disabled
mintLastParameterComboBoxIndex = cboLotFile.ListIndex   'Save the Index referring to the currently selected Parameter File
gblnParFileSelected = False                             'Indicate that no Parameter File has been selected
cboLotFile.Enabled = False                              'Disable the Lot File box while the Lot File box has focus

End Sub

Private Sub cboParameterFileName_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To allow selection of parameter file.  This event is triggered by
'            the key press event.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintResponse As Integer
Dim PSFMan As PSF090293AAMLXManager
Dim DevicesCol As ObjectCollection
Dim PSFMan2 As PSF090293AAMLXManager
Dim DevicesCol2 As ObjectCollection
Dim i As Long

If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then

    If cboParameterFileName.Text = "" Then Exit Sub 'Exit if there is nothing in the combo box

    ctrResultsTabs1.Visible = False         'Make the Results Tabs Invisible
    cboParameterFileName.Enabled = False    'Disable the combo box
    gblnParFileSelected = True              'Indicate the the Parameter File is loaded

    Call Series705.LoadParameters            'Load the parameter file
    
    '3.6bANM \/\/
    If gudtMachine.filterLoc(CHAN0) = 0 Then
        MsgBox "External Card Required! Make sure front panel filter load card is installed.", vbOKOnly, "External Filter Needed!"
    Else
        MsgBox "Internal filters are used. Please remove external card if it is installed.", vbOKOnly, "NO EXTERNAL CARDS"
    End If
    '3.6bANM /\/\
    
    Call AMAD705.SetMachineParameters        'Update the database link to that parameter file
    Call AMAD705.SetProgrammingParameters    'Update the database link to that parameter file
    Call AMAD705.SetScanParameters           'Update the database link to that parameter file
    Call Series705.LoadControlLimits         'Load the control limits

    'Scan Only
    frmMain.mnuOptionsPLCStartProgram.Checked = False
    frmMain.mnuOptionsPLCStartScan.Checked = False
    frmMain.mnuOptionsPLCStart.Enabled = False
    frmMain.mnuToolsDDE.Enabled = False
    frmMain.ctrResultsTabs1.ActiveTab = SCANRESULTSGRID
    
    If InStr(command$, "NOHARDWARE") = 0 Then

        frmMain.staMessage.Panels(1).Text = "System Message:  Initializing Force Amplifier... Please Wait         "
        
        Call TestLab.InitializeTLSensotec                    '1.7ANM 'Re-Establish Communication
        
        'Update the database link to the Force Cal Information
        Call AMAD705.SetForceCal
        
        frmMain.staMessage.Panels(1).Text = "System Message:  Force Amplifier Initialization Complete.             "
        
        'Establish Communication with MLX Programmers
        gudtPTC04(1).CommPortNum = PTC04PORT1
        gudtPTC04(2).CommPortNum = PTC04PORT2
        staMessage.Panels(1).Text = "Establishing Communication with MLX Programmers"
        gblnGoodPTC04Link = True
        'Call MLX90293.EstablishCommunication
        
'        Set PSFMan = CreateObject("MPT.PSF090293AAMLXManager")
'        Set DevicesCol = PSFMan.ScanStandalone(dtAll)
'        If DevicesCol.Count <= 0 Then
'            MsgBox ("No 90293 were found!")
'            gblnGoodPTC04Link = False
'        End If
'
'        ReDim MyDev(0 To DevicesCol.Count - 1)
'        For i = 0 To DevicesCol.Count - 1
'            Set MyDev(i) = DevicesCol(i)
'        Next i
'
'        If DevicesCol.Count > 1 Then
'            For i = 0 To DevicesCol.Count - 1
'                'Set MyDev3 = DevicesCol(i)
'                If MyDev(i).channel.Name = "COM1" Then
'                    lintDev1 = i
'                    Call MyDev(i).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
'                    MyDev(i).Advanced.ChipVersion = 2
'                ElseIf MyDev(i).channel.Name = "COM7" Then
'                    lintDev2 = i
'                    Call MyDev(i).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
'                    MyDev(i).Advanced.ChipVersion = 2
'                Else
'                    Call DevicesCol(i).Destroy(True) 'We are responsible to call Destroy(True) on device objects we do not need
'                End If
'            Next i
'        End If
        
        Set MyDev(0) = CreateObject("MPT.PSF090293AAMLXDevice")
        Call MyDev(0).ConnectChannel(CVar(CLng(1)), dtSerial)
        Call MyDev(0).CheckSetup(False)
        Call MyDev(0).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
        MyDev(0).Advanced.ChipVersion = 2
        lintDev1 = 0
        
        Set MyDev(1) = CreateObject("MPT.PSF090293AAMLXDevice")
        Call MyDev(1).ConnectChannel(CVar(CLng(7)), dtSerial)
        Call MyDev(1).CheckSetup(False)
        Call MyDev(1).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
        MyDev(1).Advanced.ChipVersion = 2
        lintDev2 = 1
        
        If gblnGoodPTC04Link Then
            staMessage.Panels(1).Text = "Communication Established with MLX Programmers."
        Else
            staMessage.Panels(1).Text = "Error Establishing Communication with Programmers!"
        End If
        'Call the routine which will Home the Motor
        Call Pedal.HomeMotor

    End If

    staMessage.Panels(1).Text = "Operator Input:  Select or Enter a New Lot File & Press <ENTER>"

    cboLotFile.Enabled = True       'Enable the Lot File box
    cboLotFile.SetFocus             'Set Focus on the Lot File box

End If

'Enable the combo box
cboParameterFileName.Enabled = True

End Sub

Private Sub cboParameterFileName_LostFocus()
'
'   PURPOSE: To ensure that the current parameter file name is displayed, and to
'            force the user to select a parameter file
'
'  INPUT(S): none
' OUTPUT(S): none

If gblnParFileSelected = False Then
    'Put the previously selected Parameter File back in the box
    cboParameterFileName.ListIndex = mintLastParameterComboBoxIndex
    'Force the user to enter a parameter file
    cboParameterFileName.SetFocus
End If

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To initialize form
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintTitleLen As Integer

'Do not run if another instance of program is currently running
If App.PrevInstance Then
    MsgBox "There is a previous instance of this program already running." & vbCrLf & vbCrLf & _
           "       You cannot run another instance of this program !", vbOKOnly, _
           "Cannot Run Multiple Instances of Program"
    End                     'End Program
End If

'Show the Startup Form
frmStartUp.Show vbModal

'Define the System Name
'NOTE: This is done here in an attempt to keep SUBS identical for EE882A-D
'      and EE882E.  Since Public Constants are not allowed in object modules,
'      we use a global string variable here instead.
gstrSystemName = "EE882E 705 Test Lab System #0"
lintTitleLen = Len(gstrSystemName)

Caption = left$(gstrSystemName, lintTitleLen - 3) & " Version " & App.Major & "." & App.Minor & "." & App.Revision
cboLotFile.Enabled = False  'Disable combo lot file box until parameter file and system homed
CWPosition.Pointers(1).Visible = False  'Position is not yet meaningful

'Initialize menus to not enabled
mnuFunctionTest.Enabled = False
mnuFunctionProgram.Enabled = False
mnuFunctionProgramAndTest.Enabled = False
mnuFunctionHomeMotor.Enabled = False
mnuFunctionInitializeProgrammers.Enabled = False
mnuFunctionResetVRef.Enabled = False
mnuToolsMonitorDAQ.Enabled = False
mnuToolsProgrammerUtility.Enabled = False
mnuToolsVIX500IEUtility.Enabled = False
mnuToolsSensotec.Enabled = False

'Initialize variables for start-up
gblnStartUpDone = False
gblnParFileSelected = False
gblnLotFileSelected = False
gblnGraphEnable = True                      'Default to graphs enabled
mnuOptionsGraphEnable.Checked = True
gblnSaveRawData = True                      'Default to save Scan Raw Data '2.1ANM
mnuOptionsSaveRawData.Checked = True
gblnSaveScanResultsToFile = True            'Default to save Scan Results
mnuOptionsScanResultsData.Checked = True
gblnSaveProgResultsToFile = True            'Default to save Programming Results
mnuOptionsProgResultsData.Checked = True
gblnLockICs = False                         'Default to lock the MLX chips
mnuOptionsLock.Checked = False
gblnLockRejects = False                     'Default to lock rejects '2.1ANM
mnuOptionsLockRejects.Checked = False
gsngVRefSetPoint = SUPPLYIDEAL              'Default VRef set point
gstrSampleNum = ""                          '1.2ANM
gblnGraphsLoaded = False                    '1.5ANM
gblnTLPrintType = True                      '1.6ANM
mnuOptionsPrintTypeTestlab.Checked = True   '1.6ANM
gblnForceOnly = False                       '1.8ANM
frmMain.mnuOptionFTO.Checked = False        '1.8ANM
gsngForceOffset = -0.5                      '2.1ANM '3.6gANM
gsngForceGain = -11.086                     '2.1ANM '3.6gANM
gdblDelay = 400000                          '2.3ANM
gblnTLScanner = True                        '2.4ANM
gblnBnmkTest = False                        '2.7ANM
frmMain.mnuOptionBMT.Checked = False        '2.7ANM
frmMain.mnuOptionMSE.Checked = False        '2.7ANM
gblnUseNewAmad = False                      '2.7ANM
gblnLockSkip = False                        '3.1ANM \/\/
gblnReClamp = False
gblnReScanRun = False
gblnReScanEnable = False
gblnReClampEnable = False                   '3.1ANM /\/\
mnuFunctionAutoSavePDFs.Checked = True      '3.6ANM
gblnMLXVI = False                           '3.6cANM
mnuOptionsSSN.Checked = False               '3.6eANM
mnuOptionsSSD.Checked = False               '3.6eANM

If InStr(command$, "NOHARDWARE") = 0 Then

    '*** DAQ & DIO Board Properties ***

    'Initialize DAQ properties
    Call frmDAQIO.ScanDAQSetup
    Call frmDAQIO.Force_Setup
    Call frmDAQIO.PeakForceDAQSetup
    Call frmDAQIO.VoutDAQSetup
    Call frmDAQIO.FoutDAQSetup                              '1.3ANM
    Call frmDAQIO.MonitorDAQSetup
    Call frmDAQIO.VRefDAQSetup
    'Call frmSolver90277.SolverDAQSetup

    'Initalize DIO #1 & DIO #2 properties
    Call frmDAQIO.DIO1_Setup
    Call frmDAQIO.DIO2_Setup

    'Send Reset to PT Board microprocessor
    Call frmDAQIO.OffPort1(PORT2, BIT2)
    frmDAQIO.KillTime (100)                                 'Delay 100 msec
    Call frmDAQIO.OnPort1(PORT2, BIT2)

    'Initialize the Digitizer I/O board
    frmDAQIO.cwDIO2.Ports.Item(PORT0).SingleWrite &H0       'Disable OA, Series, Swap
    frmDAQIO.cwDIO2.Ports.Item(PORT1).SingleWrite &H0       'Disable CC222 #4
    frmDAQIO.cwDIO2.Ports.Item(PORT2).SingleWrite &H80      '3.6*ANM Set to 90293
    frmDAQIO.cwDIO2.Ports.Item(PORT3).SingleWrite &H0       'Disable CC222 #1
    frmDAQIO.cwDIO2.Ports.Item(PORT4).SingleWrite &H0       'Disable CC222 #2
    frmDAQIO.cwDIO2.Ports.Item(PORT5).SingleWrite &H0       'Disable CC222 #3
    'PORT6 is not used
    'PORT7 is not used
    'PORT8 is not used
    'PORT9 is not used
    'PORT10 is not used
    'PORT11 is not used

    'Enable the VRef
    Call frmDAQIO.OnPort1(PORT4, BIT1)
    'Enable Programming paths
    Call frmDAQIO.OnPort1(PORT4, BIT2)  'Output #1
    Call frmDAQIO.OnPort1(PORT4, BIT3)  'Output #2

End If

'*** Check that Directories Exist ***
'    Note: Exit program if directories don't exist...
If Not (DirectoryExists(DATAPATH)) Then End             'This is the root directory and must be checked first
If Not (DirectoryExists(DATA705PATH)) Then End          'This is the next directory and must be checked second
If Not (DirectoryExists(RAWDATAPATH)) Then End
If Not (DirectoryExists(ERRORPATH)) Then End
If Not (DirectoryExists(STATFILEPATH)) Then End
If Not (DirectoryExists(PARTDATAPATH)) Then End
If Not (DirectoryExists(PARTPROGDATAPATH)) Then End
If Not (DirectoryExists(PARTSCANDATAPATH)) Then End
If Not (DirectoryExists(PARTRAWDATAPATH)) Then End
If Not (DirectoryExists(PARTMLXDATAPATH)) Then End      '2.2ANM

Call RefreshParameterFileList
Call RefreshLotFileList

'Make user controls not visible until parameter and lot file are selected.
ctrResultsTabs1.Visible = False
ctrStatus1.Visible = False

'Set database start and end numbers
gintDatabaseStartNum = 1
gintDatabaseStopNum = 2
Call AMAD705.InitializeDatabaseConnection      'Open a connection to the database

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: To unload form
'
'  INPUT(S): Cancel = if cancel true the exit
' OUTPUT(S): none

If InStr(command$, "NOHARDWARE") = 0 Then                   'If hardware is not present bypass logic
    '*** Reset any TTL outputs that are still active ***
    frmDAQIO.cwDIO2.Ports.Item(PORT0).SingleWrite &H0       'Disable OA, Series, Swap
    frmDAQIO.cwDIO2.Ports.Item(PORT1).SingleWrite &H0       'Disable CC222 #4
    frmDAQIO.cwDIO2.Ports.Item(PORT2).SingleWrite &H0       'Disable 90293
    frmDAQIO.cwDIO2.Ports.Item(PORT3).SingleWrite &H0       'Disable CC222 #1
    frmDAQIO.cwDIO2.Ports.Item(PORT4).SingleWrite &H0       'Disable CC222 #2
    frmDAQIO.cwDIO2.Ports.Item(PORT5).SingleWrite &H0       'Disable CC222 #3
    'PORT6 is not used
    'PORT7 is not used
    'PORT8 is not used
    'PORT9 is not used
    'PORT10 is not used
    'PORT11 is not used

    '*** Reset PLC DDE ***
    If gudtMachine.PLCCommType Then
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
    End If
End If

'Destory 90293 devices
MyDev(0).Destroy (True)
MyDev(1).Destroy (True)

'Close the connection to the database
Call AMAD705.CloseDatabaseConnection

'Unload all forms before exiting the program
Unload frmAbout
Unload frmDAQIO
Unload frmDDE
Unload frmMLX90293
Unload frmMonitorDAQ
Unload frmParamViewer
Unload frmPassword
Unload frmPrintPreview
Unload frmSensotec
'Unload frmSolver90277
Unload frmStartUp
Unload frmTimeSettings
Unload frmVIX500IE

tmrPollPLC_IO.Enabled = False

'Terminate the program
End

End Sub

Private Sub mnuExpChemicalResistance_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpChemicalResistance.Checked = True Then
    mnuExpChemicalResistance.Checked = False
    gblnChemResExp = False
Else
    mnuExpChemicalResistance.Checked = True
    gblnChemResExp = True
    
    'Show the Chemical Resistance Data Form
    frmCRData.Show vbModal
End If

End Sub

Private Sub mnuExpCondensation_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpCondensation.Checked = True Then
    mnuExpCondensation.Checked = False
    gblnCondenExp = False
Else
    mnuExpCondensation.Checked = True
    gblnCondenExp = True
End If

End Sub

Private Sub mnuExpDither_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpDither.Checked = True Then
    mnuExpDither.Checked = False
    gblnDitherExp = False
Else
    mnuExpDither.Checked = True
    gblnDitherExp = True
End If

End Sub

Private Sub mnuExpDust_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpDust.Checked = True Then
    mnuExpDust.Checked = False
    gblnDustExp = False
Else
    mnuExpDust.Checked = True
    gblnDustExp = True
    
    'Show the Dust Data Form
    frmDustData.Show vbModal
End If

End Sub

Private Sub mnuExpElecBilkCurrentInjection_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpElecBilkCurrentInjection.Checked = True Then
    mnuExpElecBilkCurrentInjection.Checked = False
    gblnBilkCInjElecExp = False
Else
    mnuExpElecBilkCurrentInjection.Checked = True
    gblnBilkCInjElecExp = True
End If

End Sub

Private Sub mnuExpElecElectroStaticDischarge_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpElecElectroStaticDischarge.Checked = True Then
    mnuExpElecElectroStaticDischarge.Checked = False
    gblnESDElecExp = False
Else
    mnuExpElecElectroStaticDischarge.Checked = True
    gblnESDElecExp = True
End If

End Sub

Private Sub mnuExpElecEMWaveResistance_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpElecEMWaveResistance.Checked = True Then
    mnuExpElecEMWaveResistance.Checked = False
    gblnEMWaveResElecExp = False
Else
    mnuExpElecEMWaveResistance.Checked = True
    gblnEMWaveResElecExp = True
End If

End Sub

Private Sub mnuExpElecIgnitionNoise_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpElecIgnitionNoise.Checked = True Then
    mnuExpElecIgnitionNoise.Checked = False
    gblnIgnitionNoiseElecExp = False
Else
    mnuExpElecIgnitionNoise.Checked = True
    gblnIgnitionNoiseElecExp = True
End If

End Sub

Private Sub mnuExpElecNarrowbandRadiatedEMEnergy_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpElecNarrowbandRadiatedEMEnergy.Checked = True Then
    mnuExpElecNarrowbandRadiatedEMEnergy.Checked = False
    gblnNarRadEMEElecExp = False
Else
    mnuExpElecNarrowbandRadiatedEMEnergy.Checked = True
    gblnNarRadEMEElecExp = True
End If

End Sub

Private Sub mnuExpExposure_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.5ANM new menu

If mnuExpExposure.Checked = True Then
    mnuExpExposure.Checked = False
    gblnExposure = False
Else
    mnuExpExposure.Checked = True
    gblnExposure = True
    
    gudtExposure.Exposure.Condition = InputBox("Enter Comment", "Exposure")
End If

End Sub

Private Sub mnuExpHighTemp_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpHighTemp.Checked = True Then
    mnuExpHighTemp.Checked = False
    gblnHighTempExp = False
Else
    mnuExpHighTemp.Checked = True
    gblnHighTempExp = True
    
    'Show the High Temp Soak Data Form
    frmHTSData.Show vbModal
End If

End Sub

Private Sub mnuExpHighTempHighHumidity_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpHighTempHighHumidity.Checked = True Then
    mnuExpHighTempHighHumidity.Checked = False
    gblnHighTempHighHumidExp = False
Else
    mnuExpHighTempHighHumidity.Checked = True
    gblnHighTempHighHumidExp = True
    
    'Show the High Temp High Humidity Soak Data Form
    frmHTHHSData.Show vbModal
End If

End Sub

Private Sub mnuExpInitial_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpInitial.Checked = True Then
    mnuExpInitial.Checked = False
    gblnInitialExp = False
Else
    mnuExpInitial.Checked = True
    gblnInitialExp = True
End If

End Sub

Private Sub mnuExpLowTemp_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpLowTemp.Checked = True Then
    mnuExpLowTemp.Checked = False
    gblnLowTempExp = False
Else
    mnuExpLowTemp.Checked = True
    gblnLowTempExp = True
    
    'Show the Low Temp Soak Data Form
    frmLTSData.Show vbModal
End If

End Sub

Private Sub mnuExpOperationalEndurance_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpOperationalEndurance.Checked = True Then
    mnuExpOperationalEndurance.Checked = False
    gblnOperEndurExp = False
Else
    mnuExpOperationalEndurance.Checked = True
    gblnOperEndurExp = True
    
    'Show the Operational Endurance Data Form
    frmOEData.Show vbModal
End If

End Sub

Private Sub mnuExpSaltSpray_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpSaltSpray.Checked = True Then
    mnuExpSaltSpray.Checked = False
    gblnSaltSprayExp = False
Else
    mnuExpSaltSpray.Checked = True
    gblnSaltSprayExp = True
    
    'Show the Salt Spray Data Form
    frmSSData.Show vbModal
End If

End Sub

Private Sub mnuExpSnapback_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpSnapback.Checked = True Then
    mnuExpSnapback.Checked = False
    gblnSnapbackExp = False
Else
    mnuExpSnapback.Checked = True
    gblnSnapbackExp = True
    
    'Show the Snapback Data Form
    frmSBData.Show vbModal
End If

End Sub

Private Sub mnuExpStrengthImpact_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpStrengthImpact.Checked = True Then
    mnuExpStrengthImpact.Checked = False
    gblnImpactStrnExp = False
Else
    mnuExpStrengthImpact.Checked = True
    gblnImpactStrnExp = True
End If

End Sub

Private Sub mnuExpStrengthLateral_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpStrengthLateral.Checked = True Then
    mnuExpStrengthLateral.Checked = False
    gblnLateralStrnExp = False
Else
    mnuExpStrengthLateral.Checked = True
    gblnLateralStrnExp = True
End If

End Sub

Private Sub mnuExpStrengthOperational_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpStrengthOperational.Checked = True Then
    mnuExpStrengthOperational.Checked = False
    gblnOperStrnExp = False
Else
    mnuExpStrengthOperational.Checked = True
    gblnOperStrnExp = True
End If

End Sub

Private Sub mnuExpStrengthOpwithStop_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpStrengthOpwithStop.Checked = True Then
    mnuExpStrengthOpwithStop.Checked = False
    gblnOpStrnStopExp = False
Else
    mnuExpStrengthOpwithStop.Checked = True
    gblnOpStrnStopExp = True
End If

End Sub

Private Sub mnuExpThermalShock_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpThermalShock.Checked = True Then
    mnuExpThermalShock.Checked = False
    gblnThermalShockExp = False
Else
    mnuExpThermalShock.Checked = True
    gblnThermalShockExp = True
    
    'Show the Thermal Shock Data Form
    frmTSData.Show vbModal
End If

End Sub

Private Sub mnuExpVibration_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpVibration.Checked = True Then
    mnuExpVibration.Checked = False
    gblnVibrationExp = False
Else
    mnuExpVibration.Checked = True
    gblnVibrationExp = True
    
    'Show the Vibration Data Form
    frmVibData.Show vbModal
End If

End Sub

Private Sub mnuExpWaterSpray_Click()
'
'   PURPOSE: To allow exposure to be selected
'
'  INPUT(S): none
' OUTPUT(S): none
'1.1ANM new menu

If mnuExpWaterSpray.Checked = True Then
    mnuExpWaterSpray.Checked = False
    gblnWaterSprayExp = False
Else
    mnuExpWaterSpray.Checked = True
    gblnWaterSprayExp = True
    
    'Show the Water Spray Data Form
    frmWSData.Show vbModal
End If

End Sub

Private Sub mnuFunctionAutoPrintProgResults_Click()
'
'   PURPOSE: To allow results to be printed after every scan
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuFunctionAutoPrintProgResults.Checked Then
    mnuFunctionAutoPrintProgResults.Checked = False      'Toggle Auto Print Results
Else
    mnuFunctionAutoPrintProgResults.Checked = True       'Toggle Auto Print Results
End If

End Sub

Private Sub mnuFunctionAutoPrintScanResults_Click()
'
'   PURPOSE: To allow results to be printed after every scan
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuFunctionAutoPrintScanResults.Checked Then
    mnuFunctionAutoPrintScanResults.Checked = False      'Toggle Auto Print Results
Else
    mnuFunctionAutoPrintScanResults.Checked = True       'Toggle Auto Print Results
End If

End Sub

Private Sub mnuAbout_Click()
'
'   PURPOSE: To display the information regrading this program to the user
'
'  INPUT(S): none
' OUTPUT(S): none

frmAbout.Show vbModal
    
End Sub

Private Sub mnuFunctionAutoPrintGraphs_Click()
'
'   PURPOSE: To allow graphs to be printed after every scan
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuFunctionAutoPrintGraphs.Checked Then
    mnuFunctionAutoPrintGraphs.Checked = False      'Toggle Auto Print Graphs
Else
    mnuFunctionAutoPrintGraphs.Checked = True       'Toggle Auto Print Graphs
End If

End Sub

Private Sub mnuFunctionAutoSavePDFs_Click()
'
'   PURPOSE: To allow the user to auto-save pdfs
'
'  INPUT(S): none
' OUTPUT(S): none
'3.6ANM new menu

If mnuFunctionAutoSavePDFs.Checked Then
    mnuFunctionAutoSavePDFs.Checked = False      'Toggle
Else
    mnuFunctionAutoSavePDFs.Checked = True       'Toggle
End If

End Sub

Private Sub mnuFunctionCTSPartNumber_Click()
'
'   PURPOSE: To allow the user to enter CTS part number
'
'  INPUT(S): none
' OUTPUT(S): none
'1.6ANM new menu

Dim lstrCTSPartNum As String            'CTS Part Number

lstrCTSPartNum = InputBox("Enter the CTS Part Number:", "CTS Part Number", gstrCTSPartNum)
gstrCTSPartNum = lstrCTSPartNum

End Sub

Private Sub mnuFunctionCustomerName_Click()
'
'   PURPOSE: To allow the user to enter customer name
'
'  INPUT(S): none
' OUTPUT(S): none
'1.6ANM new menu

Dim lstrCustomerName As String          'Customer Name

lstrCustomerName = InputBox("Enter the Customer Name:", "Customer Name", gstrCustomerName)
gstrCustomerName = lstrCustomerName

End Sub

Private Sub mnuFunctionCustomerPartNumber_Click()
'
'   PURPOSE: To allow the user to enter customer part number
'
'  INPUT(S): none
' OUTPUT(S): none
'1.6ANM new menu

Dim lstrCustomerPartNum As String       'Customer Part Number

lstrCustomerPartNum = InputBox("Enter the Customer Part Number:", "Customer Part Number", gudtMachine.CustomerPartNum)
gudtMachine.CustomerPartNum = lstrCustomerPartNum

End Sub

Private Sub mnuFunctionExit_Click()
'
'   PURPOSE: To allow the user to exit the program
'
'  INPUT(S): none
' OUTPUT(S): none

Unload Me    'Unload Main form

End Sub

Private Sub mnuFunctionHomeMotor_Click()
'
'   PURPOSE: To initiate a home motor command
'
'  INPUT(S): none
' OUTPUT(S): none

'Change mouse pointer to hourglass while homing motor
MousePointer = vbHourglass

'Don't execute if in NOHARDWARE mode
If InStr(command$, "NOHARDWARE") = 0 Then

    'Disable the timer to poll for PLC IO
    tmrPollPLC_IO.Enabled = False

    'Tell the PLC that the Scanner is not Initialized
    If gudtMachine.PLCCommType Then Call frmDDE.WriteDDEOutput(ScannerInit, 0)

    'Call the routine which will Home the Motor
    Call Pedal.HomeMotor

    'Enable the timer to poll for PLC IO
    tmrPollPLC_IO.Enabled = True

End If
'Change mouse pointer back to default
MousePointer = vbNormal

End Sub

Private Sub mnuFunctionInitializeProgrammers_Click()
'
'   PURPOSE: To allow the user to initialize the programmers
'
'  INPUT(S): None
' OUTPUT(S): None

Dim PSFMan As PSF090293AAMLXManager
Dim DevicesCol As ObjectCollection
Dim i As Long

'Disable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = False

'Establish Communication with MLX Programmers
gudtPTC04(1).CommPortNum = PTC04PORT1
gudtPTC04(2).CommPortNum = PTC04PORT2
staMessage.Panels(1).Text = "Establishing Communication with MLX Programmers"

'Set PSFMan = CreateObject("MPT.PSF090293AAMLXManager")
'Set DevicesCol = PSFMan.ScanStandalone(dtAll)
'If DevicesCol.Count <= 0 Then
'    MsgBox ("No 90293 were found!")
'    gblnGoodPTC04Link = False
'Else
    gblnGoodPTC04Link = True
'End If
'
'ReDim MyDev(0 To DevicesCol.Count - 1)
'For i = 0 To DevicesCol.Count - 1
'    Set MyDev(i) = DevicesCol(i)
'Next i
'
'If DevicesCol.Count > 1 Then
'    For i = 0 To DevicesCol.Count - 1
'        Set MyDev3 = DevicesCol(i)
'        If MyDev3.channel.Name = "COM1" Then
'            lintDev1 = i
'            Call MyDev(i).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
'            MyDev(i).Advanced.ChipVersion = 2
'        ElseIf MyDev3.channel.Name = "COM7" Then
'            lintDev2 = i
'            Call MyDev(i).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
'            MyDev(i).Advanced.ChipVersion = 2
'        Else
'            MyDev(i) = Nothing
'            Call DevicesCol(i).Destroy(True) 'We are responsible to call Destroy(True) on device objects we do not need
'        End If
'    Next i
'End If
 
Set MyDev(0) = CreateObject("MPT.PSF090293AAMLXDevice")
Call MyDev(0).ConnectChannel(CVar(CLng(1)), dtSerial)
Call MyDev(0).CheckSetup(False)
Call MyDev(0).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
MyDev(0).Advanced.ChipVersion = 2
lintDev1 = 0

Set MyDev(1) = CreateObject("MPT.PSF090293AAMLXDevice")
Call MyDev(1).ConnectChannel(CVar(CLng(7)), dtSerial)
Call MyDev(1).CheckSetup(False)
Call MyDev(1).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
MyDev(1).Advanced.ChipVersion = 2
lintDev2 = 1
        
If gblnGoodPTC04Link Then
    staMessage.Panels(1).Text = "Communication Established with MLX Programmers."
Else
    staMessage.Panels(1).Text = "Error Establishing Communication with Programmers!"
End If

'Enable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = True

End Sub

Private Sub mnuFunctionInitializeSensotec_Click()
'
'   PURPOSE: To allow the user to initialize the programmers
'
'  INPUT(S): None
' OUTPUT(S): None

'Disable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = False

frmMain.staMessage.Panels(1).Text = "System Message:  Initializing Force Amplifier... Please Wait         "
Call TestLab.InitializeTLSensotec                    '1.7ANM 'Re-Establish Communication
'Update the database link to the Force Cal Information
Call AMAD705.SetForceCal
frmMain.staMessage.Panels(1).Text = "System Message:  Force Amplifier Initialization Complete.             "

'Enable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = True

End Sub

Private Sub mnuFunctionPartName_Click()
'
'   PURPOSE: To allow the user to enter part name
'
'  INPUT(S): none
' OUTPUT(S): none
'1.6ANM new menu

Dim lstrPartName As String              'Part Name

lstrPartName = InputBox("Enter the Part Name:", "Part Name", gstrPartName)
gstrPartName = lstrPartName

End Sub

Private Sub mnuFunctionPreviewOpt_Click(Index As Integer)
'
'   PURPOSE: To allow the user to preview the scanning/programming results/stats
'
'  INPUT(S): Index = Which menu option was selected
' OUTPUT(S): none
'1.6ANM \/\/

If gblnTLPrintType Then
    Select Case Index
        Case 0  'Scan Results
            Call frmPrintPreview2.DisplayData(SCANRESULTSGRID)
            gstrType = " Scan " '3.6ANM
            frmPrintPreview2.Visible = True
        Case 1  'Programming Results
            Call frmPrintPreview2.DisplayData(PROGRESULTSGRID)
            gstrType = " Prog " '3.6ANM
            frmPrintPreview2.Visible = True
        Case 2  'Scan Statistics
            Call frmPrintPreview.DisplayData(SCANSTATSGRID)
            gstrType = " Scan " '3.6ANM
            frmPrintPreview.Visible = True
        Case 3  'Programming Statistics
            Call frmPrintPreview.DisplayData(PROGSTATSGRID)
            gstrType = " Prog " '3.6ANM
            frmPrintPreview.Visible = True
    End Select
Else
    Select Case Index
        Case 0  'Scan Results
            Call frmPrintPreview3.DisplayData(SCANRESULTSGRID)
            gstrType = " Scan " '3.6ANM
            frmPrintPreview3.Visible = True
        Case 1  'Programming Results
            Call frmPrintPreview3.DisplayData(PROGRESULTSGRID)
            gstrType = " Prog " '3.6ANM
            frmPrintPreview3.Visible = True
        Case 2  'Scan Statistics
            Call frmPrintPreview4.DisplayData(SCANSTATSGRID)
            gstrType = " Scan " '3.6ANM
            frmPrintPreview4.Visible = True
        Case 3  'Programming Statistics
            Call frmPrintPreview4.DisplayData(PROGSTATSGRID)
            gstrType = " Prog " '3.6ANM
            frmPrintPreview4.Visible = True
    End Select
End If

End Sub

Private Sub mnuFunctionPrintOpt_Click(Index As Integer)
'
'   PURPOSE: To allow the user to print results/stats/graphs
'
'  INPUT(S): Index = Which menu option was selected
' OUTPUT(S): None
'1.6ANM \/\/

If gblnTLPrintType Then
    Select Case Index
        Case 1    'Scan Results
            Call frmPrintPreview2.DisplayData(SCANRESULTSGRID)
            gstrType = " Scan " '3.6ANM
            Call frmPrintPreview2.PrintDisplay
        Case 2    'Programming Results
            Call frmPrintPreview2.DisplayData(PROGRESULTSGRID)
            gstrType = " Prog " '3.6ANM
            Call frmPrintPreview2.PrintDisplay
        Case 3  'Graphs
            gstrType = " Graph " '3.6ANM
            Call ctrResultsTabs1.PrintAllGraphsInWindow
        Case 4   'Scan Stats
            Call frmPrintPreview.DisplayData(SCANSTATSGRID)
            gstrType = " Scan " '3.6ANM
            Call frmPrintPreview.PrintDisplay
        Case 5   'Programming Stats
            Call frmPrintPreview.DisplayData(PROGSTATSGRID)
            gstrType = " Prog " '3.6ANM
            Call frmPrintPreview.PrintDisplay
    End Select
Else
    Select Case Index
        Case 1    'Scan Results
            Call frmPrintPreview3.DisplayData(SCANRESULTSGRID)
            gstrType = " Scan " '3.6ANM
            Call frmPrintPreview3.PrintDisplay
        Case 2    'Programming Results
            Call frmPrintPreview3.DisplayData(PROGRESULTSGRID)
            gstrType = " Prog " '3.6ANM
            Call frmPrintPreview3.PrintDisplay
        Case 3  'Graphs
            gstrType = " Graph " '3.6ANM
            Call ctrResultsTabs1.PrintAllGraphsInWindow
        Case 4   'Scan Stats
            Call frmPrintPreview4.DisplayData(SCANSTATSGRID)
            gstrType = " Scan " '3.6ANM
            Call frmPrintPreview4.PrintDisplay
        Case 5   'Programming Stats
            Call frmPrintPreview4.DisplayData(PROGSTATSGRID)
            gstrType = " Prog " '3.6ANM
            Call frmPrintPreview4.PrintDisplay
    End Select
End If

End Sub

Private Sub mnuFunctionProgram_Click()
'
'   PURPOSE: To allow user to initiate programming
'
'  INPUT(S): none
' OUTPUT(S): none
    
'Only enable if the timer to poll PLC IO is enabled
If tmrPollPLC_IO.Enabled = True Then
    gblnProgramStart = True                 'Enable manual program
End If

End Sub

Private Sub mnuFunctionProgramAndTest_Click()
'
'   PURPOSE: To allow user to initiate programming and scanning
'
'  INPUT(S): none
' OUTPUT(S): none

'Only enable if the timer to poll PLC IO is enabled
If tmrPollPLC_IO.Enabled = True Then
    gblnProgramStart = True                 'Enable manual program
    gblnScanStart = True                    'Enable manual run test
End If

End Sub

Private Sub mnuFunctionResetPartCount_Click()
'
'   PURPOSE: To allow the reset of the current yield
'
'  INPUT(S): none
' OUTPUT(S): none

gudtProgSummary.currentGood = 0                'Last xxx parts good
gudtProgSummary.currentTotal = 0               'Last xxx parts
gudtScanSummary.currentGood = 0                'Last xxx parts good
gudtScanSummary.currentTotal = 0               'Last xxx parts
MsgBox "Current part Counter has been Reset"

End Sub

Private Sub mnuFunctionResetVRef_Click()
'
'   PURPOSE: To allow the user to reset the voltage reference
'
'  INPUT(S): none
' OUTPUT(S): none

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic
    Call Pedal.CalculateVRefParameters(1)
End If

End Sub

Private Sub mnuFunctionSampleNum_Click()
'
'   PURPOSE: To allow user to enter sample number
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lstrSampleNum As String     'Sample number

lstrSampleNum = InputBox("Enter the next Sample Number:", "Sample Number", gstrSampleNum)
gstrSampleNum = lstrSampleNum
frmMain.ctrSetupInfo1.Sample = gstrSampleNum

End Sub

Private Sub mnuFunctionTest_Click()
'
'   PURPOSE: To allow user to initiate scanning
'
'  INPUT(S): none
' OUTPUT(S): none

'Only enable if the timer to poll PLC IO is enabled
If tmrPollPLC_IO.Enabled = True Then
    gblnScanStart = True                 'Enable manual run test
End If

End Sub

Private Sub mnuMLXVI_Click()
'
'   PURPOSE: To allow user to adjust MLX VI
'
'  INPUT(S): none
' OUTPUT(S): none
'3.6cANM

Dim lintChanNum As Integer

If InStr(command$, "NOHARDWARE") = 0 Then

    'Make sure the Timer to look for startscan is disabled
    tmrPollPLC_IO.Enabled = False
    gblnMLXVI = True
    
    'Enable the proper filter-loads & paths
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
    Next lintChanNum

    Call frmDAQIO.OnPort1(PORT4, BIT2)  'Enable the programming path for Checkhead #1
    Call frmDAQIO.OnPort1(PORT4, BIT3)  'Enable the programming path for Checkhead #2

    'Delay for the relays to debounce
    Call frmDAQIO.KillTime(50)

    'Show the Programmer Utility Form
    frmMLXVI.Show vbModal

    'Disable the paths
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
    Next lintChanNum

    Call frmDAQIO.OffPort1(PORT4, BIT2)  'Enable the scanning path for Checkhead #1
    Call frmDAQIO.OffPort1(PORT4, BIT3)  'Enable the scanning path for Checkhead #2

    'Enable the timer that polls for PLC IO
    tmrPollPLC_IO.Enabled = True
    gblnMLXVI = False
End If

End Sub

Private Sub mnuOptionBMT_Click()
'
'   PURPOSE: To allow user to enable/disable force only scan
'
'  INPUT(S): none
' OUTPUT(S): none
'2.7ANM new sub

If gblnBnmkTest = True Then
    gblnBnmkTest = False
    frmMain.mnuOptionBMT.Checked = False
Else
    gblnBnmkTest = True
    frmMain.mnuOptionBMT.Checked = True
End If

End Sub

Private Sub mnuOptionFTO_Click()
'
'   PURPOSE: To allow user to enable/disable force only scan
'
'  INPUT(S): none
' OUTPUT(S): none
'1.8ANM new sub

If gblnForceOnly = True Then
    gblnForceOnly = False
    frmMain.mnuOptionFTO.Checked = False
Else
    gblnForceOnly = True
    frmMain.mnuOptionFTO.Checked = True
End If

End Sub

Private Sub mnuOptionMSE_Click()
'
'   PURPOSE: To allow user to enable/disable multiple scans
'
'  INPUT(S): none
' OUTPUT(S): none
'2.7ANM new sub

If frmMain.mnuOptionMSE.Checked = True Then
    frmMain.mnuOptionMSE.Checked = False
Else
    frmMain.mnuOptionMSE.Checked = True
End If

End Sub

Private Sub mnuOptionsAutoPrintFailureResults_Click()
'
'   PURPOSE: To allow user to enable/disable auto-printing of failure results
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsAutoPrintFailureResults.Checked = True Then
    'Disable Auto Printing of Failure Results
    mnuOptionsAutoPrintFailureResults.Checked = False
Else
    'Enable Auto Printing of Failure Results
    mnuOptionsAutoPrintFailureResults.Checked = True
End If

End Sub

Private Sub mnuOptionsGraphEnable_Click()
'
'   PURPOSE: To allow user to enable/disable graphs
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsGraphEnable.Checked = True Then
    'Disable Graphing
    mnuOptionsGraphEnable.Checked = False
    gblnGraphEnable = False
Else
    'Enable Graphing
    mnuOptionsGraphEnable.Checked = True
    gblnGraphEnable = True
End If

End Sub

Private Sub mnuOptionsLock_Click()
'
'   PURPOSE: To allow the user to enable/disable locking of parts
'
'  INPUT(S): none
' OUTPUT(S): none
'2.1ANM \/\/ updated for lock rejects option

'Show the password form
Beep
If gblnAdministrator = False Then frmPassword.Show vbModal
'Exit if the password was not correct
If Not gblnAdministrator Then Exit Sub

If mnuOptionsLock.Checked = True Then
    gblnLockICs = False
    mnuOptionsLock.Checked = False
    mnuOptionsLockRejects.Checked = False
    mnuOptionsLockRejects.Enabled = False
Else
    gblnLockICs = True
    mnuOptionsLock.Checked = True
    mnuOptionsLockRejects.Enabled = True
    If gblnLockRejects Then mnuOptionsLockRejects.Checked = True
End If

'Reset the permissions
gblnAdministrator = False

End Sub

Private Sub mnuOptionsLockRejects_Click()
'
'   PURPOSE: To allow the user to enable/disable locking of reject parts
'
'  INPUT(S): none
' OUTPUT(S): none

'Show the password form
Beep
If gblnAdministrator = False Then frmPassword.Show vbModal
'Exit if the password was not correct
If Not gblnAdministrator Then Exit Sub

If mnuOptionsLockRejects.Checked = True Then
    gblnLockRejects = False
    mnuOptionsLockRejects.Checked = False
Else
    gblnLockRejects = True
    mnuOptionsLockRejects.Checked = True
End If

'Reset the permissions
gblnAdministrator = False

End Sub

Private Sub mnuOptionsPLCStartProgram_Click()
'
'   PURPOSE: To allow the user to enable/disable Programming of parts after
'            a PLC StartScan signal.
'
'  INPUT(S): none
' OUTPUT(S): none

'Show the password form
Beep
If gblnAdministrator = False Then frmPassword.Show vbModal
'Exit if the password was not correct
If Not gblnAdministrator Then Exit Sub

If mnuOptionsPLCStartProgram.Checked = True Then
    mnuOptionsPLCStartProgram.Checked = False
Else
    mnuOptionsPLCStartProgram.Checked = True
End If

'Reset the permissions
gblnAdministrator = False

End Sub

Private Sub mnuOptionsPLCStartScan_Click()
'
'   PURPOSE: To allow the user to enable/disable Scanning of parts after
'            a PLC StartScan signal.
'
'  INPUT(S): none
' OUTPUT(S): none

'Show the password form
Beep
If gblnAdministrator = False Then frmPassword.Show vbModal
'Exit if the password was not correct
If Not gblnAdministrator Then Exit Sub

If mnuOptionsPLCStartScan.Checked = True Then
    mnuOptionsPLCStartScan.Checked = False
Else
    mnuOptionsPLCStartScan.Checked = True
End If

'Reset the permissions
gblnAdministrator = False

End Sub

Private Sub mnuOptionsPrintTypeCustomer_Click()
'
'   PURPOSE: To allow the user to select customer type result printouts
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsPrintTypeCustomer.Checked = True Then
    gblnCustPrintType = False
    gblnTLPrintType = True
    mnuOptionsPrintTypeCustomer.Checked = False
    mnuOptionsPrintTypeTestlab.Checked = True
Else
    gblnCustPrintType = True
    gblnTLPrintType = False
    mnuOptionsPrintTypeCustomer.Checked = True
    mnuOptionsPrintTypeTestlab.Checked = False
End If

End Sub

Private Sub mnuOptionsPrintTypeTestlab_Click()
'
'   PURPOSE: To allow the user to select testlab type result printouts
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsPrintTypeTestlab.Checked = False Then
    gblnCustPrintType = False
    gblnTLPrintType = True
    mnuOptionsPrintTypeCustomer.Checked = False
    mnuOptionsPrintTypeTestlab.Checked = True
Else
    gblnCustPrintType = True
    gblnTLPrintType = False
    mnuOptionsPrintTypeCustomer.Checked = True
    mnuOptionsPrintTypeTestlab.Checked = False
End If

End Sub

Private Sub mnuOptionsProgResultsData_Click()
'
'   PURPOSE: To allow the user to enable/disable programming results data
'            being saved to a file
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsProgResultsData.Checked = True Then
    gblnSaveProgResultsToFile = False
    mnuOptionsProgResultsData.Checked = False
Else
    gblnSaveProgResultsToFile = True
    mnuOptionsProgResultsData.Checked = True
End If

End Sub

Private Sub mnuOptionsSaveRawData_Click()
'
'   PURPOSE: To allow the user to enable/disable raw data being saved
'            to a file
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsSaveRawData.Checked = True Then
    gblnSaveRawData = False
    mnuOptionsSaveRawData.Checked = False
Else
    gblnSaveRawData = True
    mnuOptionsSaveRawData.Checked = True
End If

End Sub

Private Sub mnuOptionsScanResultsData_Click()
'
'   PURPOSE: To allow the user to enable/disable scan results data
'            being saved to a file
'
'  INPUT(S): none
' OUTPUT(S): none

If mnuOptionsScanResultsData.Checked = True Then
    gblnSaveScanResultsToFile = False
    mnuOptionsScanResultsData.Checked = False
Else
    gblnSaveScanResultsToFile = True
    mnuOptionsScanResultsData.Checked = True
End If

End Sub

Private Sub mnuOptionsSSD_Click()
'
'   PURPOSE: To allow the user to enable/disable slow scan.
'
'  INPUT(S): none
' OUTPUT(S): none
'3.6eANM new sub

If mnuOptionsSSD.Checked = True Then
    mnuOptionsSSD.Checked = False
    mnuOptionsSSN.Checked = False
    If Len(frmMain.ctrSetupInfo1.Comment) > 22 Then frmMain.ctrSetupInfo1.Comment = left(frmMain.ctrSetupInfo1.Comment, Len(frmMain.ctrSetupInfo1.Comment) - 23)
Else
    mnuOptionsSSD.Checked = True
    If mnuOptionsSSN.Checked = False Then frmMain.ctrSetupInfo1.Comment = frmMain.ctrSetupInfo1.Comment + "   ***Slow Scan Mode***"
    mnuOptionsSSN.Checked = False
End If

End Sub

Private Sub mnuOptionsSSN_Click()
'
'   PURPOSE: To allow the user to enable/disable slow scan.
'
'  INPUT(S): none
' OUTPUT(S): none
'3.6eANM new sub

If mnuOptionsSSN.Checked = True Then
    mnuOptionsSSN.Checked = False
    mnuOptionsSSD.Checked = False
    If Len(frmMain.ctrSetupInfo1.Comment) > 22 Then frmMain.ctrSetupInfo1.Comment = left(frmMain.ctrSetupInfo1.Comment, Len(frmMain.ctrSetupInfo1.Comment) - 23)
Else
    mnuOptionsSSN.Checked = True
    If mnuOptionsSSD.Checked = False Then frmMain.ctrSetupInfo1.Comment = frmMain.ctrSetupInfo1.Comment + "   ***Slow Scan Mode***"
    mnuOptionsSSD.Checked = False
End If

End Sub

Private Sub mnuToolsCycle_Click()
'
'   PURPOSE: To allow the user to cycle parts
'
'  INPUT(S): none
' OUTPUT(S): none
'2.5ANM new sub

'Show the password form
Beep
If gblnAdministrator = False Then frmPassword.Show vbModal
'Exit if the password was not correct
If Not gblnAdministrator Then Exit Sub

frmCycle.Show vbModal

gblnAdministrator = False

End Sub

Private Sub mnuToolsDDE_Click()
'
'   PURPOSE: To allow the user to use the PLC DDE Test Panel
'
'  INPUT(S): None
' OUTPUT(S): None

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    'Disable the timer that polls for PLC I/O
    tmrPollPLC_IO.Enabled = False

    If gudtMachine.PLCCommType Then

        'Show the password form if the user is not already administrator
        If gblnAdministrator = False Then
            frmPassword.Show vbModal
        End If
    
        'If the password was incorrect, do NOT allow use of the test utility
        If Not gblnAdministrator Then
            MsgBox "INCORRECT PASSWORD."
            Exit Sub
        End If
    
        'Show the DDE Form
        frmDDE.Show
    
        gblnAdministrator = False
    
        'Re-Setup PLC DDE Topics and Items
        Call frmDDE.PLCDDESetup
    Else
        MsgBox "No PLC on this system!  Option Not Available!"
    End If

    'Enable the timer that polls for PLC I/O
    tmrPollPLC_IO.Enabled = True
    
End If

End Sub

Private Sub mnuToolsMonitorDAQ_Click()
'
'   PURPOSE: To allow user to view the voltages for each A/D channel
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintChanNum As Integer

'Disable the timer that polls for PLC I/O
tmrPollPLC_IO.Enabled = False

'Enable VRef
Call frmDAQIO.OnPort1(PORT4, BIT1)
'Disable Programming paths
Call frmDAQIO.OffPort1(PORT4, BIT2)  'Output #1
Call frmDAQIO.OffPort1(PORT4, BIT3)  'Output #2

'Enable the proper filter-loads & paths
For lintChanNum = CHAN0 To CHAN3
    Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
Next lintChanNum

'Set Force Amplifier to operate mode
'3.6aANM Call frmDAQIO.OffPort2(PORT0, BIT7)

'Display Monitor DAQ form
frmMonitorDAQ.Show vbModal

'Set Force Amplifier to reset mode
'3.6aANM Call frmDAQIO.OnPort2(PORT0, BIT7)

'Enable the timer that polls for PLC I/O
tmrPollPLC_IO.Enabled = True

End Sub

Private Sub mnuToolsMotor_Click()
'
'   PURPOSE:  To show the Motor Control Form
'
'  INPUT(S): none
' OUTPUT(S): none
'3.6iANM new

If InStr(command$, "NOHARDWARE") = 0 Then
    'Make sure the Timer to look for startscan is disabled
    tmrPollPLC_IO.Enabled = False
    'Display Motor Utility form
    frmMotorControl.Show vbModal
    'Make sure the Timer to look for startscan is enabled
    tmrPollPLC_IO.Enabled = True
End If

End Sub

Private Sub mnuToolsProgrammerUtility_Click()
'
'   PURPOSE:  To show the Programmer Utility Form
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lintProgrammerNum As Integer
Dim lintPositionNum As Integer
Dim lintChanNum As Integer
'Dim lstrMLX90277Revision As String

If InStr(command$, "NOHARDWARE") = 0 Then

    'Make sure the Timer to look for startscan is disabled
    tmrPollPLC_IO.Enabled = False

    'Enable the proper filter-loads & paths
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
    Next lintChanNum

    Call frmDAQIO.OnPort1(PORT4, BIT2)  'Enable the programming path for Checkhead #1
    Call frmDAQIO.OnPort1(PORT4, BIT3)  'Enable the programming path for Checkhead #2

    'Delay for the relays to debounce
    Call frmDAQIO.KillTime(50)

    'Enable the timer to update the Melexis Form display
    frmMLX90293.tmrMLX.Enabled = True

    'Add the proper Com Port Numbers to the form
    frmMLX90293.cboComPortNum(1).Text = CStr(PTC04PORT1)
    frmMLX90293.cboComPortNum(2).Text = CStr(PTC04PORT2)

'    'Save the Revision level in case it is changed on the form
'    lstrMLX90277Revision = gstrMLX90277Revision
'    'Set up the form to work with the selected Revision Level
'    If gstrMLX90277Revision = "Cx" Then
'        frmMLX90277.mnuMLX90277RevisionCx.Checked = True
'        frmMLX90277.mnuMLX90277RevisionFA.Checked = False
'    ElseIf gstrMLX90277Revision = "FA" Then
'        frmMLX90277.mnuMLX90277RevisionCx.Checked = False
'        frmMLX90277.mnuMLX90277RevisionFA.Checked = True
'    End If

    'Show the Programmer Utility Form
    frmMLX90293.Show vbModal

    'Disable the paths
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), False)
    Next lintChanNum

    Call frmDAQIO.OffPort1(PORT4, BIT2)  'Enable the scanning path for Checkhead #1
    Call frmDAQIO.OffPort1(PORT4, BIT3)  'Enable the scanning path for Checkhead #2

    'Disable the timer to update the Melexis Form display
    frmMLX90293.tmrMLX.Enabled = False

    'Reset the Revision level in case it was changed on the form
    'gstrMLX90277Revision = lstrMLX90277Revision
    
    'Establish Communication with MLX Programmers
'    gudtPTC04(1).CommPortNum = PTC04PORT1
'    gudtPTC04(2).CommPortNum = PTC04PORT2
'    staMessage.Panels(1).Text = "Establishing Communication with MLX Programmers"
'    Call MLX90293.EstablishCommunication
'    If gblnGoodPTC04Link Then
'        staMessage.Panels(1).Text = "Communication Established with MLX Programmers."
'    Else
'        staMessage.Panels(1).Text = "Error Establishing Communication with Programmers!"
'    End If

    'Enable the timer that polls for PLC IO
    tmrPollPLC_IO.Enabled = True

End If

End Sub

Private Sub mnuToolsSensotec_Click()
'
'   PURPOSE:  To show the Sensotec Utility Form
'
'  INPUT(S): none
' OUTPUT(S): none

If InStr(command$, "NOHARDWARE") = 0 Then
    'Make sure the Timer to look for startscan is disabled
    tmrPollPLC_IO.Enabled = False
    'Put the appropriate Com Port Number in the combo box on the utility form
    frmSensotec.cboComPortNum.Text = CStr(SENSOTECPORT)
    'Display Sensotec form
    frmSensotec.Show vbModal
End If

End Sub

Private Sub mnuToolsVIX500IEUtility_Click()
'
'   PURPOSE:  To show the Motor Utility Form
'
'  INPUT(S): none
' OUTPUT(S): none

If InStr(command$, "NOHARDWARE") = 0 Then
    'Make sure the Timer to look for startscan is disabled
    tmrPollPLC_IO.Enabled = False
    'Put the appropriate Com Port Number in the combo box on the utility form
    frmVIX500IE.cboComPortNum.Text = CStr(VIX500IEPORT)
    'Set the Gear Ratio to the System Gear Ratio
    frmVIX500IE.txtGearRatio = CStr(gudtMachine.gearRatio)
    'Set the Comm Active Button on the Motor Utility Form
    frmVIX500IE.cwbtnCommunicationActive = VIX500IE.GetLinkStatus
    'Display Motor Utility form
    frmVIX500IE.Show vbModal
    'Make sure the Timer to look for startscan is enabled
    tmrPollPLC_IO.Enabled = True
End If

End Sub

Private Sub mnuViewParameterFile_Click()
'
'   PURPOSE: To allow user to view parameter file
'
'  INPUT(S): none
' OUTPUT(S): none

Dim fileName As String
Dim lstrRc(999, 6) As String    'Parameter file converted to table format (MAX ROWS,MAX COLUMNS)

'Disable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = False

'Get parameter file name which was selected
fileName = App.Path + PARPATH + cboParameterFileName.Text

'convert parameter file into a table
Call TabulateFile(fileName, lstrRc())
frmParamViewer.Visible = True

'Enable the timer to poll for PLC IO
tmrPollPLC_IO.Enabled = True

End Sub

Private Sub mnuViewTab_Click(Index As Integer)
'
'   PURPOSE: To change the selected tab
'
'  INPUT(S): Index
' OUTPUT(S):

'Set active tab
ctrResultsTabs1.ActiveTab = Index

End Sub

Public Function GetShiftLetter() As String
'
'   PURPOSE: To build the shift letter based on the time of day.
'
'  INPUT(S): None
' OUTPUT(S): returns the Shift Letter (String)
'3.1ANM moved here from pedal.bas

Dim lintHours As Integer
Dim lintMinutes As Integer
Dim lintMinuteOfDay As Integer
Dim lstrShift As String

'Determine the minute of the day (x/1439)
lintHours = DateTime.Hour(DateTime.Now)         'Get the hour
lintMinutes = DateTime.Minute(DateTime.Now)     'Get the minutes
lintMinuteOfDay = lintHours * 60 + lintMinutes  'Calculate the current minute of the day

'Select the shift based on what minute of the day it is
Select Case lintMinuteOfDay
    '12:00AM to 6:59AM, 11:00PM to 11:59PM
    Case 0 To 419, 1380 To 1439
        lstrShift = "C"
    '7:00AM to 2:59PM
    Case 420 To 899
        lstrShift = "A"
    '3:00PM to 10:59PM
    Case 900 To 1379
        lstrShift = "B"
End Select

'Return the selected shift letter
GetShiftLetter = lstrShift

End Function

Private Sub tmrPollPLC_IO_Timer()
'
'   PURPOSE: Polls for PLC Start Signal and acts as the executive
'            for both Programming and Scanning
'
'  INPUT(S): none
' OUTPUT(S): none
'

Dim lsngCurrentTime As Single
Dim lblnScanPart As Boolean
Dim X As Integer
Dim lintS As Integer
Dim lblnOff As Boolean
Dim PSFMan As PSF090293AAMLXManager
Dim DevicesCol As ObjectCollection
Dim i As Long

On Error GoTo TMR_ERR

tmrPollPLC_IO.Enabled = False               'Disable timer
lblnOff = False

If InStr(command$, "NOHARDWARE") = 0 Then   'If hardware is not present bypass logic

    If gudtMachine.PLCCommType Then
        'Check for PLC Start Scan
        gblnPLCStart = frmDDE.ReadDDEInput(StartScan)
    End If

    Call Pedal.Position

    If gblnPLCStart Then
        'Use pull-down options to determine whether to program, scan, or program & scan
        gblnScanStart = mnuOptionsPLCStartScan.Checked
        gblnProgramStart = mnuOptionsPLCStartProgram.Checked
        'Acknowledge Start Scan: Scan Complete High
        Call frmDDE.WriteDDEOutput(StartScanAck, 1)
        'Read the Pallet Number from the PLC
        gintPalletNumber = frmDDE.ReadDDEInput(PalletNum)
    Else
        'Manual Test; Assume that the PLC does not know what pallet is in place
        gintPalletNumber = 0
    End If

    If gblnScanStart Or gblnProgramStart Then
    
        If mnuOptionsSSD.Checked = True Then '3.6eANM
            lintS = 2
        Else
            lintS = 1
        End If
            
        For X = 1 To lintS
            'Turn Off Slow for Second Scan
            If X = 2 Then
                mnuOptionsSSD.Checked = False
                If Len(frmMain.ctrSetupInfo1.Comment) > 22 Then frmMain.ctrSetupInfo1.Comment = left(frmMain.ctrSetupInfo1.Comment, Len(frmMain.ctrSetupInfo1.Comment) - 23)
                lblnOff = True
            End If
            
            If gblnScanStart Or gblnProgramStart Then
        
                'Turn Off the Status Buttons
                ctrStatus1.StatusValue(1) = False
                ctrStatus1.StatusValue(2) = False
        
                'Initialize the system anomaly to zero
                gintAnomaly = 0
                'Initialize the flag representing whether we had a good S/N to False
                gblnGoodSerialNumber = False
                gblnSkipPDF = False      '3.6ANM
                
                'Set items to be sent to results file
                If frmMain.ctrSetupInfo1.Sample <> "" Then
                    gstrSampleNum = frmMain.ctrSetupInfo1.Sample
                Else
                    If frmMain.ctrSetupInfo1.Sample <> gstrSampleNum Then
                        gstrSampleNum = frmMain.ctrSetupInfo1.Sample
                    End If
                End If
                
                'Clear the Scan Results Control
                Call ctrResultsTabs1.ClearData(SCANRESULTSGRID, 1, 2)
                'Clear the Programming Results Control
                Call ctrResultsTabs1.ClearData(PROGRESULTSGRID, 1, 2)
        
                'Display the Cycle Time
                lsngCurrentTime = Timer
                If (lsngCurrentTime - gsngCycleTimerStart) < 999 Then
                    staMessage.Panels(2).Text = "Cycle Time = " & Format(lsngCurrentTime - gsngCycleTimerStart, "#0.#0") & " S"
                Else
                    staMessage.Panels(2).Text = "Cycle Time"
                End If
                'Restart cycle timer
                gsngCycleTimerStart = Timer
        
                'Check to be sure that Programmers are Initialized:
                If Not gblnGoodPTC04Link Then
                    If gudtMachine.PLCCommType Then Call frmDDE.WriteDDEOutput(WatchdogDisable, 1)  'Disable the PLC's watchdog timer
                    'Call MLX90277.EstablishCommunication            'Re-Establish Communication
        
                    gblnGoodPTC04Link = True
                    
                    Set MyDev(0) = CreateObject("MPT.PSF090293AAMLXDevice")
                    Call MyDev(0).ConnectChannel(CVar(CLng(1)), dtSerial)
                    Call MyDev(0).CheckSetup(False)
                    Call MyDev(0).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP1.ini")
                    MyDev(0).Advanced.ChipVersion = 2
                    lintDev1 = 0
                    
                    Set MyDev(1) = CreateObject("MPT.PSF090293AAMLXDevice")
                    Call MyDev(1).ConnectChannel(CVar(CLng(7)), dtSerial)
                    Call MyDev(1).CheckSetup(False)
                    Call MyDev(1).Advanced.OpenProfile("C:\EE's\EE882\705\EE0882EV36 90293\705 TYF AP2.ini")
                    MyDev(1).Advanced.ChipVersion = 2
                    lintDev2 = 1
        
                    If gblnGoodPTC04Link Then
                        'Re-Enable the PLC Watchdog Timer
                        If gudtMachine.PLCCommType Then Call frmDDE.WriteDDEOutput(WatchdogDisable, 0)
                    Else
                        'Error Re-Initializing Programmers
                        gintAnomaly = 168
                        'Log the error to the error log and display the error message
                        Call Pedal.ErrorLogFile("Programmer Communication Error: Error during Initialization." & vbCrLf & _
                                               "Verify Connections to Programmer.", True, True)
                    End If
                End If
        
                'Check to be sure that the Sensotec SC2000 is Initialized:
                If Not Sensotec.GetLinkStatus Then
                    If gudtMachine.PLCCommType Then Call frmDDE.WriteDDEOutput(WatchdogDisable, 1)  'Disable the PLC's watchdog timer
                    
                    frmMain.staMessage.Panels(1).Text = "System Message:  Initializing Force Amplifier... Please Wait         "
        
                    Call TestLab.InitializeTLSensotec                    '1.7ANM 'Re-Establish Communication
                    'Update the database link to the Force Cal Information
                    Call AMAD705.SetForceCal
                    frmMain.staMessage.Panels(1).Text = "System Message:  Force Amplifier Initialization Complete.             "
                    
                    If Sensotec.GetLinkStatus Then
                        'Re-Enable the PLC Watchdog Timer
                        If gudtMachine.PLCCommType Then Call frmDDE.WriteDDEOutput(WatchdogDisable, 0)
                    Else
                        'Error Re-Initializing the Sensotec SC2000
                        gintAnomaly = 50
                        'Log the error to the error log and display the error message
                        Call ErrorLogFile("Serial Communication Error:  Sensotec SC2000 Not Initialized!", True, True)
                    End If
        
                End If
        
            End If
        
            'If Programming Requested
            If gblnProgramStart And Not gblnBnmkTest Then '3.6eANM
        
                'Call the executive that handles Programming
                If gintAnomaly = 0 Then Call Pedal.TunePedal90293
                
                'Move to the Load Location
                Call MoveToLoadLocation
        
                'If not a lock skip part
                If Not gblnLockSkip Then '3.2ANM
                    frmMain.ctrSetupInfo1.DateCode = gstrDateCode '1.7ANM
                    
                    'If (ONLY) Programming was requested by the PLC...
                    If gblnPLCStart And (Not gblnScanStart) Then
                        'Send Scan Complete (Start Scan Ack Low)
                        Call frmDDE.WriteDDEOutput(StartScanAck, 0)
                        'Send Programming Part Results to the PLC
                        Call Pedal.SendProgrammingResultsToPLC
                    End If
            
                    'Update the Display
                    Call Pedal.StatsUpdateProgCounts                     'Update Prog Summary Counts
                    Call Pedal.DisplayProgSummary                        'Display Prog Summary Display
                    Call Pedal.DisplayProgResultsCountsPrioritized       'Display Prog Counts
                    Call Pedal.DisplayProgStatisticsCountsPrioritized    'Display Prog Counts
                    'Don't update statistical sums on system faults or programming failures
                    If (gintAnomaly = 0) And Not gblnProgFailure Then
                        Call Pedal.StatsUpdateProgSums                   'Update Prog Stats
                        Call Pedal.DisplayProgStatisticsData             'Display Prog Stats
                    End If
            
                    'Display results when no system fault
                    If (gintAnomaly = 0) Then
                        Call Pedal.DisplayProgResultsData                'Display Prog Results
                        'Save the programming results if called for
                        If gblnSaveProgResultsToFile Then Call TestLab.SaveTLProgResultsToFile '1.7ANM
                    End If
                    'If we read a S/N for the current device, save the results
                    If gblnGoodSerialNumber Then
                        'Set the Serial Number ID
                        Call AMAD705.SetSerialNumber
                        'Add a record to the database for the current test
                        Call AMAD705.AddProgrammingResultsRecord
                    Else
                        'We didn't read a serial number, save the info
                        Call AMAD705.AddUnserializedProgRecord
                    End If
                End If
            End If
        
            'Determine whether or not to Scan the Part
            lblnScanPart = gblnScanStart And (gintAnomaly = 0) And Not (gblnProgramStart And gblnProgFailure)
        
            If lblnScanPart Then
        
                'If the part was programmed prior to test, then the serial
                'number has already been read.  If not, it needs to be read now.
                If Not gblnForceOnly And Not gblnBnmkTest Then    '2.7ANM '1.8ANM added if block
                    If (Not gblnProgramStart Or gblnLockSkip) And (gintAnomaly = 0) Then '3.2ANM
                        'Read the Serial Number and Date Code
                        Call Pedal.ReadSerialNumberAndDateCode90293
                        If gblnGoodDateCode = True Then  '1.7ANM \/\/
                            frmMain.ctrSetupInfo1.DateCode = gstrDateCode
                        End If                           '1.7ANM /\/\
                    End If
                Else
                    gstrSerialNumber = ""
                    gstrDateCode = ""
                    frmMain.ctrSetupInfo1.PartNum = ""
                    frmMain.ctrSetupInfo1.DateCode = ""
                    gudtReading(0).mlxCurrent = 0 '3.6eANM
                    gudtReading(1).mlxCurrent = 0 '3.6eANM
                End If
                
                If frmMain.mnuOptionMSE.Checked = True Then
                    frmMultiScan.Show vbModal
                Else
                    'Call the executive that handles Scanning
                    If (gintAnomaly = 0) Then Call Pedal.RunTest
            
                    'Set graph variable
                    If gblnGraphEnable Then gblnGraphsLoaded = True                    '1.7ANM
            
                    'Save the results data if called for '1.7ANM moved from pedal.bas
                    If ((gintAnomaly = 0) And gblnSaveScanResultsToFile) Then Call TestLab.Save705TLScanResultsToFile
            
                    'Move to the Load Location
                    Call MoveToLoadLocation
                End If
            End If
        
            'If Scanning was Requested by the PLC, Send Results to the PLC
            If (gblnScanStart And gblnPLCStart) Then
                'Send Scan Complete (Start Scan Ack Low)
                Call frmDDE.WriteDDEOutput(StartScanAck, 0)
                'Send the Serial Number & Date Code Info to the PLC
                If ((Not gblnScanFailure) And (Not gblnProgFailure) And (Not gblnSevere)) And (gintAnomaly = 0) Then
                    Call Pedal.SendSerialNumberToPLC
                End If
                'Send Programming Part Results to the PLC
                Call Pedal.SendScanResultsToPLC
            End If
        
            'If the part was scanned...
            If lblnScanPart Then
        
                'Update Stats & Stats Display
                Call Series705.StatsUpdateScanCounts                     'Update Scan Summary Counts
                Call Series705.DisplayScanSummary                        'Display Scan Summary Display
                Call Series705.DisplayScanResultsCountsPrioritized       'Display Scan Counts
                Call Series705.DisplayScanStatisticsCountsPrioritized    'Display Scan Counts
                'Don't update statistical sums on system faults or severe failures
                If (gintAnomaly = 0) And Not gblnSevere Then
                    Call Series705.StatsUpdateScanSums                   'Update Scan Stats
                    Call Series705.DisplayScanStatisticsData             'Display Scan Stats
                End If
        
                'Display results when no system fault
                If (gintAnomaly = 0) Then
                    Call Series705.DisplayScanResultsData                'Display Scan Data
                End If
                'If we read a S/N for the current device, save the results
                If (Not gblnForceOnly) And (Not gblnBnmkTest) Then       '2.7ANM '1.8ANM added if block
                    If gblnGoodSerialNumber Then
                        'Set the Serial Number ID
                        Call AMAD705.SetSerialNumber
                        'Add a record to the database for the current test
                        Call AMAD705.AddScanResultsRecord
                    Else
                        'We didn't read a serial number, save the info
                        Call AMAD705.AddUnserializedScanRecord
                    End If
                End If
            End If
        
            'Print Programming Results data when Auto Print flag enabled
            If gblnProgramStart Then
                If (gintAnomaly = 0 And (mnuFunctionAutoPrintProgResults.Checked Or (gblnProgFailure And mnuOptionsAutoPrintFailureResults.Checked))) Then
                    gblnSkipPDF = False      '3.6ANM \/\/
                    If mnuFunctionAutoPrintScanResults.Checked Or mnuFunctionAutoPrintGraphs.Checked Then
                        gblnSkipPDF = True
                    Else
                        gblnSkipPDF = False
                    End If
                    gstrType = " Prog "      '3.6ANM /\/\
                    
                    If gblnTLPrintType Then  '1.6ANM \/\/
                        Call frmPrintPreview2.DisplayData(PROGRESULTSGRID)
                        Call frmPrintPreview2.PrintDisplay
                    Else
                        Call frmPrintPreview3.DisplayData(PROGRESULTSGRID)
                        Call frmPrintPreview3.PrintDisplay
                    End If                   '1.6ANM /\/\
                End If
            End If
        
            'Print Scan Results data when Auto Print flag enabled
            If gblnScanStart Then
                If (gintAnomaly = 0 And (mnuFunctionAutoPrintScanResults.Checked Or (gblnScanFailure And mnuOptionsAutoPrintFailureResults.Checked))) Then
                    If gblnSkipPDF Then      '3.6ANM \/\/
                        gstrType = gstrType & " & Scan "
                    Else
                        gstrType = " Scan "
                    End If
                    
                    If mnuFunctionAutoPrintGraphs.Checked Then
                        gblnSkipPDF = True
                    Else
                        gblnSkipPDF = False
                    End If                   '3.6ANM /\/\
                    
                    If gblnTLPrintType Then  '1.6ANM \/\/
                        Call frmPrintPreview2.DisplayData(SCANRESULTSGRID)
                        Call frmPrintPreview2.PrintDisplay
                    Else
                        Call frmPrintPreview3.DisplayData(SCANRESULTSGRID)
                        Call frmPrintPreview3.PrintDisplay
                    End If                   '1.6ANM /\/\
                End If
                'Print Graph data when Auto Print flag enabled
                If (gintAnomaly = 0 And (mnuFunctionAutoPrintGraphs.Checked Or (gblnScanFailure And mnuOptionsAutoPrintFailureResults.Checked))) Then
                    If gblnSkipPDF Then      '3.6ANM \/\/
                        gstrType = gstrType & " & Graph "
                    Else
                        gstrType = " Graph "
                    End If
                    gblnSkipPDF = False      '3.6ANM /\/\
                    
                    Call ctrResultsTabs1.PrintAllGraphsInWindow
                End If
            End If
        
            If (Not gblnForceOnly) And (Not gblnBnmkTest) Then      '2.7ANM '1.8ANM added if block
                If gblnProgramStart Or gblnScanStart Then
            
                    'Verify that the stat file name is good
                    If gstrLotName <> "" Then
                        MousePointer = vbHourglass
                        staMessage.Panels(1).Text = "System Message:  Saving Lot Data..."
                        Call TestLab.Stats705TLSave     '1.7ANM
                        staMessage.Panels(1).Text = "System Message:  Lot Data Saved."
                        MousePointer = vbNormal
                    Else
                        MsgBox "No Lot File Selected.  The Current Data Did Not Save!", vbCritical, "Lot File"
                    End If
                    
                    'Check to see if the database needs to be switched
                    If AMAD705.IsDatabaseTooBig Then
                        'Loop until there is a connection to the database
                        Do
                            Call AMAD705.SwitchDatabase
                        Loop While Not AMAD705.GetConnectionStatus
                       'Now set all the database IDs
                        Call AMAD705.SetLot
                        Call AMAD705.SetMachineParameters
                        Call AMAD705.SetProgrammingParameters
                        Call AMAD705.SetScanParameters
                    End If
                End If
            End If
            
            'Always re-enable the PLC Watchdog Timer
            If gblnPLCStart Then
                'Enable the PLC Watchdog Timer
                Call frmDDE.WriteDDEOutput(WatchdogDisable, 0)
            End If
        
            'Update the System Status Box
            If gblnProgramStart And gblnScanStart Then
                staMessage.Panels(1).Text = "System Message:  Programming & Scanning Complete ... Ready for Next Part"
            ElseIf gblnProgramStart Then
                staMessage.Panels(1).Text = "System Message:  Programming Complete ... Ready for Next Part"
            ElseIf gblnScanStart Then
                staMessage.Panels(1).Text = "System Message:  Scanning Complete ... Ready for Next Part"
            End If
            
            'Exit on error
            If gintAnomaly <> 0 Then Exit For
        Next X
        
        'Turn On Slow for Dual Mode
        If lblnOff Then
            mnuOptionsSSD.Checked = True
            frmMain.ctrSetupInfo1.Comment = frmMain.ctrSetupInfo1.Comment + "   ***Slow Scan Mode***"
        End If
    End If
End If

gblnScanStart = False                           'Disable manual run test
gblnProgramStart = False                        'Disable manual program
tmrPollPLC_IO.Enabled = True                    'Re-enable timer

Exit Sub

TMR_ERR:
    Call frmDAQIO.KillTime(50)
    MsgBox Err.Description, vbOKOnly, "Error in Main Loop!"
    Err.Clear
    'Destory 90293 devices
    MyDev(0).Destroy (True)
    MyDev(1).Destroy (True)
    gblnGoodPTC04Link = False
    gblnScanStart = False                           'Disable manual run test
    gblnProgramStart = False                        'Disable manual program
    tmrPollPLC_IO.Enabled = True                    'Re-enable timer

End Sub

Private Sub txtPosition_Change()
'
'   PURPOSE: To monitor and display the change event of the position from the encoder
'
'  INPUT(S): none
' OUTPUT(S): none

CWPosition.Pointers(1).Value = CSng(left(txtPosition.Text, Len(txtPosition.Text) - 1)) Mod DEGPERREV     'Current position

End Sub

