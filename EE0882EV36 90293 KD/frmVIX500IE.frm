VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "CWUI.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmVIX500IE 
   BackColor       =   &H8000000A&
   Caption         =   "VIX500IE Motor Controller Test Panel"
   ClientHeight    =   9915
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12825
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
   ScaleHeight     =   9915
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLimits 
      Caption         =   "CW/CCW Limits"
      Height          =   855
      Left            =   240
      TabIndex        =   65
      Top             =   8760
      Width           =   3735
      Begin CWUIControlsLib.CWButton cwbtnLimitEnable 
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   3255
         _Version        =   393218
         _ExtentX        =   5741
         _ExtentY        =   450
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
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483643
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "CW/CCW Limits Enabled"
         style_8         =   63986560
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "CW/CCW Limits Disabled"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   -2147483633
         B_12            =   -2147483633
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   -2147483633
         B_19            =   -2147483633
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   1
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
   End
   Begin VB.Frame fraMotorOnOff 
      Caption         =   "Motor On/Off"
      Height          =   735
      Left            =   240
      TabIndex        =   60
      Top             =   2520
      Width           =   3735
      Begin VB.OptionButton optMotorOnOff 
         Caption         =   "Off"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   62
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optMotorOnOff 
         Caption         =   "On"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   61
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraTerminal 
      Caption         =   "Terminal"
      Height          =   3975
      Left            =   4200
      TabIndex        =   51
      Top             =   5640
      Width           =   8295
      Begin VB.TextBox txtTerminal 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   3720
         Width           =   7815
      End
      Begin VB.TextBox txtTerminalHistory 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame fraContinuousMode 
      Caption         =   "Continuous Mode Control"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   8160
      TabIndex        =   48
      Top             =   4200
      Width           =   4335
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   2400
         TabIndex        =   67
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   255
         Left            =   2400
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Counterclockwise"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   50
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Clockwise"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   49
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraPositionMode 
      Caption         =   "Position Mode Control"
      Height          =   2175
      Left            =   4200
      TabIndex        =   41
      Top             =   3240
      Width           =   3735
      Begin VB.CommandButton cmdRelativeMove 
         Caption         =   "Go"
         Height          =   255
         Left            =   2160
         TabIndex        =   47
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtRelativeMove 
         Height          =   315
         Left            =   240
         TabIndex        =   45
         Text            =   "0"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdAbsoluteMove 
         Caption         =   "Go"
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtAbsoluteMove 
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblRelativePosition 
         Caption         =   "Relative Move (Degrees)"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblAbsolutePosition 
         Caption         =   "Absolute Move (Degrees)"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraServoMode 
      Caption         =   "Servo Mode"
      Height          =   1455
      Left            =   240
      TabIndex        =   34
      Top             =   7080
      Width           =   3735
      Begin VB.OptionButton optMode 
         Caption         =   "Incremental Mode"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   37
         Top             =   960
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Absolute Mode"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Continuous Mode"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraPIDParameters 
      Caption         =   "PID Parameters"
      Height          =   3975
      Left            =   8160
      TabIndex        =   16
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtFeedForwardGain 
         Height          =   315
         Left            =   120
         TabIndex        =   74
         Text            =   "5"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdSetPIDParameters 
         Caption         =   "Set PID Parameters"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox txtFilterTime 
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         Text            =   "0"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtInPositionTime 
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Text            =   "10"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtPositionErrorWindow 
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Text            =   "50"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtPositionError 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtIntegralWindow 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Text            =   "25"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox txtVelocityGain 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Text            =   "5"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtIntegralGain 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Text            =   "0"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtProportionalGain 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "10"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblFeedFwdGain 
         Caption         =   "Feed Forward Gain"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblFilterTime 
         Caption         =   "Filter Time"
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblInPositionTime 
         Caption         =   "In Position Time"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPositionErrorWindow 
         Caption         =   "Position Error Window"
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblCurrentGain 
         Caption         =   "Position Error"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblIntegralWindow 
         Caption         =   "Integral Window"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblVelGain 
         Caption         =   "Velocity Gain"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblIntegralGain 
         Caption         =   "Integral Gain"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblProportionalGain 
         Caption         =   "Proportional Gain"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "Current Position"
      Height          =   2895
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtTestPanelUpdateRate 
         Height          =   315
         Left            =   240
         TabIndex        =   53
         Text            =   "1000"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdUpdateOnce 
         Caption         =   "Update Once"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtCurrentPosition 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin CWUIControlsLib.CWButton cwbtnContinuousRead 
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   2280
         Width           =   3255
         _Version        =   393218
         _ExtentX        =   5741
         _ExtentY        =   450
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
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483643
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         style_6         =   63971720
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Continuous Update On"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "Continuous Update Off"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   12632256
         C[1]_11         =   12632256
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   12632256
         B_12            =   12632256
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   -2147483633
         B_19            =   -2147483633
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   1
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin VB.Label lblTestPanelUpdateRate 
         Caption         =   "Update Rate"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUpdateRateUnits 
         Caption         =   "milliseconds ( >1000 )"
         Height          =   495
         Left            =   2040
         TabIndex        =   54
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPosCounts 
         Caption         =   "Degrees"
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCurrentPosition 
         Caption         =   "Current Position (Read)"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraCommunicationStatus 
      Caption         =   "Communication Status"
      Height          =   2175
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox cboControllerNumber 
         Height          =   330
         ItemData        =   "frmVIX500IE.frx":0000
         Left            =   2040
         List            =   "frmVIX500IE.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cboComPortNum 
         Height          =   330
         ItemData        =   "frmVIX500IE.frx":0034
         Left            =   240
         List            =   "frmVIX500IE.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdResetCommunication 
         Caption         =   "Reset Communication"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   3255
      End
      Begin CWUIControlsLib.CWButton cwbtnCommunicationActive 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   375
         _Version        =   393218
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483643
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         style_8         =   1816403032
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   32768
         C[1]_11         =   32768
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   32768
         B_12            =   32768
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         UsePiccList_12  =   -1  'True
         PiccList_12     =   13
         ClassName_13    =   "CPiccListRoundLED"
         count_13        =   2
         list[2]_13      =   14
         ClassName_14    =   "CCWPicc"
         opts_14         =   62
         Image_14        =   0
         Animator_14     =   0
         Blinker_14      =   0
         Size_14.cx      =   21
         Size_14.cy      =   21
         Model_14.r      =   21
         Model_14.b      =   21
         Actual_14.r     =   25
         Actual_14.b     =   25
         Picc_14         =   411
         Color_14        =   32768
         Name_14         =   "Divot"
         list[1]_13      =   15
         ClassName_15    =   "CCWPicc"
         opts_15         =   62
         Image_15        =   0
         Animator_15     =   0
         Blinker_15      =   0
         Size_15.cx      =   21
         Size_15.cy      =   21
         Model_15.l      =   2
         Model_15.t      =   2
         Model_15.r      =   19
         Model_15.b      =   19
         Actual_15.l     =   2
         Actual_15.t     =   2
         Actual_15.r     =   22
         Actual_15.b     =   22
         Picc_15         =   404
         Color_15        =   32768
         Name_15         =   "Light"
         AllowSetColor_15=   -1  'True
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   16
         ClassName_16    =   "CCWEnumElt"
         opts_16         =   1
         Name_16         =   "On"
         frame_16        =   286
         DrawList_16     =   17
         ClassName_17    =   "CDrawList"
         count_17        =   4
         list[4]_17      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         C[1]_18         =   -2147483640
         Image_18        =   6
         Animator_18     =   0
         Blinker_18      =   0
         list[3]_17      =   19
         ClassName_19    =   "CCWDrawObj"
         opts_19         =   62
         C[0]_19         =   -2147483640
         C[1]_19         =   -2147483640
         Image_19        =   8
         Animator_19     =   0
         Blinker_19      =   0
         list[2]_17      =   20
         ClassName_20    =   "CCWDrawObj"
         opts_20         =   60
         C[0]_20         =   -2147483640
         C[1]_20         =   -2147483640
         Image_20        =   10
         Animator_20     =   0
         Blinker_20      =   0
         list[1]_17      =   21
         ClassName_21    =   "CCWDrawObj"
         opts_21         =   62
         C[0]_21         =   65380
         C[1]_21         =   65380
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   65380
         B_22            =   65380
         ColorReplaceWith_22=   8421504
         ColorReplace_22 =   8421504
         Tolerance_22    =   2
         UsePiccList_22  =   -1  'True
         PiccList_22     =   23
         ClassName_23    =   "CPiccListRoundLED"
         count_23        =   2
         list[2]_23      =   24
         ClassName_24    =   "CCWPicc"
         opts_24         =   62
         Image_24        =   0
         Animator_24     =   0
         Blinker_24      =   0
         Size_24.cx      =   21
         Size_24.cy      =   21
         Model_24.r      =   21
         Model_24.b      =   21
         Actual_24.r     =   25
         Actual_24.b     =   25
         Picc_24         =   411
         Color_24        =   65380
         Name_24         =   "Divot"
         list[1]_23      =   25
         ClassName_25    =   "CCWPicc"
         opts_25         =   62
         Image_25        =   0
         Animator_25     =   0
         Blinker_25      =   0
         Size_25.cx      =   21
         Size_25.cy      =   21
         Model_25.l      =   2
         Model_25.t      =   2
         Model_25.r      =   19
         Model_25.b      =   19
         Actual_25.l     =   2
         Actual_25.t     =   2
         Actual_25.r     =   22
         Actual_25.b     =   22
         Picc_25         =   404
         Color_25        =   65380
         Name_25         =   "Light"
         AllowSetColor_25=   -1  'True
         Animator_21     =   0
         Blinker_21      =   0
         varVarType_16   =   5
         Bindings_1      =   26
         ClassName_26    =   "CCWBindingHolderArray"
         Editor_26       =   27
         ClassName_27    =   "CCWBindingHolderArrayEditor"
         Owner_27        =   1
         Style_1         =   18
         mechAction_1    =   1
         BGImg_1         =   28
         ClassName_28    =   "CCWDrawObj"
         opts_28         =   62
         Image_28        =   29
         ClassName_29    =   "CCWPictImage"
         opts_29         =   1280
         Rows_29         =   1
         Cols_29         =   1
         Pict_29         =   286
         F_29            =   -2147483633
         B_29            =   -2147483633
         ColorReplaceWith_29=   8421504
         ColorReplace_29 =   8421504
         Tolerance_29    =   2
         Animator_28     =   0
         Blinker_28      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   21
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   30
         ClassName_30    =   "CCWDrawObj"
         opts_30         =   62
         Image_30        =   8
         Animator_30     =   0
         Blinker_30      =   0
         Array[5]_1      =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   62
         Image_31        =   10
         Animator_31     =   0
         Blinker_31      =   0
         Label_1         =   32
         ClassName_32    =   "CCWDrawObj"
         opts_32         =   62
         C[0]_32         =   -2147483640
         Image_32        =   6
         Animator_32     =   0
         Blinker_32      =   0
      End
      Begin VB.Label lblConrtollerNumber 
         Caption         =   "Controller #"
         Height          =   255
         Left            =   2040
         TabIndex        =   63
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblComm 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblComPortNumber 
         Caption         =   "Comm Port #"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Timer tmrVIX500IE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   -240
   End
   Begin VB.Frame fraMovementParameters 
      Caption         =   "Motion Parameters"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   3735
      Begin VB.TextBox txtGearRatio 
         Height          =   315
         Left            =   240
         TabIndex        =   69
         Text            =   "1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtStepsPerRev 
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Text            =   "4000"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtAcceleration 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Text            =   "0.25"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdSetMotionParameters 
         Caption         =   "Set  Motion Parameters"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txtVelocity 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "0.01"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblRevperRev 
         Caption         =   "In Rev/Out Rev"
         Height          =   255
         Left            =   2040
         TabIndex        =   71
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblGearRatio 
         Caption         =   "Gear Ratio"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblResolution 
         Caption         =   "Motor Encoder Resolution"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblStepsPerRev 
         Caption         =   "Steps/Rev"
         Height          =   255
         Left            =   2040
         TabIndex        =   58
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbThouPerSecondSquared 
         Caption         =   "Rev/Sec/Sec"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblThouPerSecond 
         Caption         =   "Rev/Sec"
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblFrequencyResponse 
         Caption         =   "Motor Acceleration"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblMotorVelocity 
         Caption         =   "Motor Velocity"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSCommLib.MSComm SerialPort 
      Left            =   5280
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4200
      Top             =   -240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmVIX500IE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************  VIX500IE Motor Controller Serial Communication Module  ************
'
'   Andrew N Mehl
'   CTS Automotive
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (219) 295-3575
'
'   The following program was written as a test panel / plug-in module
'   to allow serial communication with the Parker Compumotor VIX500IE Motor
'   Controller.  The form & module can be added to a project and used to
'   communicate with a VIX500IE Motor Controller.
'
'Ver    Date       By   Purpose of modification
'1.0.0 08/12/2004  ANM  First release of VIX500IE software module.
'
'1.1.0 09/28/2004  SRC  Made default Communication Setting 9600 Baud instead of
'                       19200 to address communication errors.  Added delays in
'                       InitializeCommunication and SetPIDParameters.
'

Option Explicit

Private mblnKillTimeDone As Boolean

Private Sub cmdAbsoluteMove_Click()

On Error GoTo BadAbsolutePositionBox

If VIX500IE.GetLinkStatus Then
    'Send the absolute move
    Call VIX500IE.AbsoluteMoveTo(CSng(txtAbsoluteMove.Text))
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadAbsolutePositionBox:

    MsgBox "Invalid Absolute Move. " _
           & vbCrLf & "Try again.", vbOKOnly + vbCritical, "Error"
End Sub

Private Sub cmdRelativeMove_Click()

On Error GoTo BadRelativePositionBox

If VIX500IE.GetLinkStatus Then
    'Send the Relative Move
    Call VIX500IE.RelativeMove(CSng(txtRelativeMove.Text))
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadRelativePositionBox:

    MsgBox "Invalid Relative Move. " _
           & vbCrLf & "Try again.", vbOKOnly + vbCritical, "Error"

End Sub

Private Sub cmdResetCommunication_Click()

On Error GoTo BadComboBox

tmrVIX500IE.Enabled = True

Call VIX500IE.SetControllerNumber(cboControllerNumber.ListIndex)

'The motor is now Off
optMotorOnOff(0).Value = False
optMotorOnOff(1).Value = True

'Set the GearRatio
Call VIX500IE.SetGearRatio(CLng(txtGearRatio.Text))

'Establish Communication with the VIX500IE
Call VIX500IE.InitializeCommunication(CInt(cboComPortNum.Text))

Exit Sub
'Error Trap
BadComboBox:

    MsgBox "Invalid Com Port Number. " _
           & vbCrLf & "Try again.", vbOKOnly + vbCritical, "Error"

End Sub

Private Sub cmdSetMotionParameters_Click()

On Error GoTo BadMotionTextBox

If VIX500IE.GetLinkStatus Then
    'Set the accel and velocity
    Call VIX500IE.SetVelocity(CSng(txtVelocity.Text))
    Call VIX500IE.SetAcceleration(CSng(txtAcceleration.Text))
    Call VIX500IE.SetDeceleration(CSng(txtAcceleration.Text))
    'Set the encoder resolution
    Call VIX500IE.SetStepsPerRev(CInt(txtStepsPerRev.Text))
    'Set the GearRatio
    Call VIX500IE.SetGearRatio(CInt(txtGearRatio.Text))
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadMotionTextBox:

    MsgBox "At least one of the text boxes contains an invalid" _
       & vbCrLf & "value.  Please correct the error and try again." _
       , vbOKOnly + vbCritical, "Velocity/Acceleration/Force Error"

End Sub

Private Sub cmdSetPIDParameters_Click()

On Error GoTo BadPIDTextBox

If VIX500IE.GetLinkStatus Then
     Call VIX500IE.SetPIDParameters(CInt(txtFeedForwardGain.Text), CInt(txtProportionalGain.Text), CInt(txtIntegralGain.Text), CInt(txtVelocityGain.Text), CInt(txtIntegralWindow.Text), CInt(txtPositionError.Text), CInt(txtPositionErrorWindow.Text), CInt(txtInPositionTime.Text), CInt(txtFilterTime.Text))
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadPIDTextBox:

    MsgBox "At least one of the text boxes contains an invalid" _
       & vbCrLf & "value.  Please correct the error and try again." _
       , vbOKOnly + vbCritical, "PID Error"

End Sub

Private Sub cmdStart_Click()

If VIX500IE.GetLinkStatus Then
    'Start the motor
    Call VIX500IE.StartMotor
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdStop_Click()

If VIX500IE.GetLinkStatus Then
    'Start the motor
    Call VIX500IE.StopMotor
Else
    'Re-Initialize communication to send a Reset command and clear a runaway situation
    Call VIX500IE.InitializeCommunication(CInt(frmVIX500IE.cboComPortNum.Text))
End If

End Sub

Private Sub cmdUpdateOnce_Click()

Dim i As Integer
Dim lsngPosition As Single

On Error GoTo BadUpdateBox

'Make sure the update period is at least 1000 milliseconds
If CInt(txtTestPanelUpdateRate) < 1000 Then
    txtTestPanelUpdateRate.Text = "1000"
End If
'Set the update period
tmrVIX500IE.Interval = CInt(txtTestPanelUpdateRate.Text)

If VIX500IE.GetLinkStatus Then
    If VIX500IE.ReadPosition(lsngPosition) Then
        'Update the read position value
        txtCurrentPosition.Text = CStr(lsngPosition)
        Call Pedal.Position
    End If
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadUpdateBox:
    MsgBox "The Update Rate text box contains an invalid" _
       & vbCrLf & "value.  Please correct the error and try again." _
       , vbOKOnly + vbCritical, "Update Rate Error"

End Sub

Private Sub cwbtnLimitEnable_Click()

If VIX500IE.GetLinkStatus Then
    If cwbtnLimitEnable.Value = True Then
        'Enable the limits
        Call VIX500IE.EnableLimits
        cwbtnLimitEnable.Value = False
        cwbtnLimitEnable.Refresh
    ElseIf cwbtnLimitEnable.Value = False Then
        'Disable the limits
        Call VIX500IE.DisableLimits
        cwbtnLimitEnable.Value = True
        cwbtnLimitEnable.Refresh
    End If
End If

End Sub

Private Sub cwbtnContinuousRead_ValueChanged(ByVal Value As Boolean)

'While continuously updating, don't allow other button presses
    
    If Value Then
        cmdUpdateOnce.Enabled = False
        fraCommunicationStatus.Enabled = False
        fraServoMode.Enabled = False
        fraPositionMode.Enabled = False
        frmVIX500IE.fraContinuousMode.Enabled = False
        fraMovementParameters.Enabled = False
        fraPIDParameters.Enabled = False
        fraTerminal.Enabled = False
        
    Else
        cmdUpdateOnce.Enabled = True
        fraCommunicationStatus.Enabled = True
        fraServoMode.Enabled = True
        If optMode(0).Enabled Then fraPositionMode.Enabled = True
        If Not optMode(0).Enabled Then fraContinuousMode.Enabled = True
        fraMovementParameters.Enabled = True
        fraPIDParameters.Enabled = True
        fraTerminal.Enabled = True

    End If

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To initialize form
'
'  INPUT(S): none
' OUTPUT(S): none

'Initialize the combo boxes
cboControllerNumber.ListIndex = 1       'Controller # = 1
cboComPortNum.ListIndex = 2             'Comm Port # = 3

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: To unload form
'
'  INPUT(S): none
' OUTPUT(S): none

'Disable timers and close the serial port
tmrVIX500IE.Enabled = False
tmrKillTime.Enabled = False
If SerialPort.PortOpen = True Then SerialPort.PortOpen = False

'Set the link status to false
VIX500IE.SetLinkStatus (False)

Unload Me

End Sub

Public Sub KillTime(milliSecDelay As Integer)
'
'   PURPOSE:   Delays using a timer event.
'
'  INPUT(S):   milliSecDelay : Delay time in milliseconds
'
' OUTPUT(S):   None

mblnKillTimeDone = False
tmrKillTime.Interval = milliSecDelay
tmrKillTime.Enabled = True

Do
    DoEvents
Loop Until mblnKillTimeDone
   
tmrKillTime.Enabled = False

End Sub

Private Sub mnuFileExit_Click()
'
'   PURPOSE:  Makes the form invisible
'
'  INPUT(S):   None
' OUTPUT(S):   None

Visible = False

End Sub

Private Sub optDirection_Click(Index As Integer)
Dim lmtTemp As ModeType

lmtTemp = Index

'Set the direction
Call VIX500IE.SetDirection(lmtTemp)

End Sub

Private Sub optMode_Click(Index As Integer)

If VIX500IE.GetLinkStatus Then
    
    'Set the servo mode
    Call VIX500IE.SetServoMode(Index)
    
    'Enable the proper buttons based on the current servo mode
    Select Case Index
        Case 0
            fraPositionMode.Enabled = False
            fraContinuousMode.Enabled = True
        Case 1, 2
            fraPositionMode.Enabled = True
            fraContinuousMode.Enabled = False
    End Select
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub optMotorOnOff_Click(Index As Integer)

If VIX500IE.GetLinkStatus Then
    Select Case Index
        Case 0
            'Turn the motor on
            Call VIX500IE.EnergizeMotor
        Case 1
            'Turn the motor off
            Call VIX500IE.DeEnergizeMotor
    End Select
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication to the Motor Controller." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub tmrKillTime_Timer()

mblnKillTimeDone = True

End Sub

Private Sub tmrVIX500IE_Timer()

If VIX500IE.GetLinkStatus Then
    'Update the Communication Active LED
    cwbtnCommunicationActive.Value = True

    If cwbtnContinuousRead.Value Then Call cmdUpdateOnce_Click

Else
    cwbtnCommunicationActive.Value = False
End If

End Sub

Private Sub txtTerminal_KeyPress(KeyAscii As Integer)

Dim lintStringLength As Integer
Dim lstrCommand As String
Dim lstrResponse As String

If KeyAscii = 13 Then   '13 = Enter Key

    lintStringLength = Len(txtTerminal.Text)
    lstrCommand = UCase(txtTerminal.Text)

    'Clear the terminal box
    txtTerminal.Text = ""

    Call VIX500IE.SendDataGetResponse(lstrCommand, lstrResponse)

    'Make sure that if the mode or motor on/off state changes
    'the form changes to reperesent the change
    Select Case lstrCommand
        Case "ON"
            optMotorOnOff(0).Value = True
            optMotorOnOff(1).Value = False
        Case "OFF"
            optMotorOnOff(0).Value = False
            optMotorOnOff(1).Value = True
    End Select

    txtTerminalHistory.Text = txtTerminalHistory.Text & vbCr & vbLf & lstrResponse
End If

End Sub

