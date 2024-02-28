VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSensotec 
   BackColor       =   &H8000000A&
   Caption         =   "Sensotec Instrumentation Test Panel"
   ClientHeight    =   8055
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   12465
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
   ScaleHeight     =   8055
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraForceDACOutput 
      Caption         =   "Forced DAC Output"
      Height          =   2055
      Left            =   240
      TabIndex        =   68
      Top             =   5640
      Width           =   3735
      Begin VB.CommandButton cmdNormalDAC 
         Caption         =   "Auto DAC"
         Height          =   495
         Left            =   1800
         TabIndex        =   83
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdForceDACOutput 
         Caption         =   "Force DAC"
         Height          =   495
         Left            =   240
         TabIndex        =   82
         Top             =   1320
         Width           =   1575
      End
      Begin MSComctlLib.Slider sldPercentage 
         Height          =   630
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1111
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Min             =   -100
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   10
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   67
      Top             =   240
      Width           =   3735
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "Values And Limit Satus (Channel 1)"
      Height          =   5175
      Left            =   4320
      TabIndex        =   59
      Top             =   2520
      Width           =   3735
      Begin VB.CommandButton cmdUpdateOnce 
         Caption         =   "Update Once"
         Height          =   495
         Left            =   240
         TabIndex        =   69
         Top             =   3240
         Width           =   3255
      End
      Begin VB.CommandButton cmdClearPeakValley 
         Caption         =   "Clear Peak and Valley"
         Height          =   495
         Left            =   240
         TabIndex        =   66
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox txtPeak 
         Height          =   315
         Left            =   240
         TabIndex        =   62
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtValley 
         Height          =   315
         Left            =   240
         TabIndex        =   61
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtTracking 
         Height          =   315
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   1695
      End
      Begin CWUIControlsLib.CWButton cwbtnContinuousRead 
         Height          =   495
         Left            =   240
         TabIndex        =   70
         Top             =   3840
         Width           =   3255
         _Version        =   393218
         _ExtentX        =   5741
         _ExtentY        =   873
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
         C[0]_7          =   0
         C[1]_7          =   0
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Continuous Update On"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   0
         C[1]_9          =   0
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
         C[0]_16         =   0
         C[1]_16         =   0
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   0
         C[1]_17         =   0
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483644
         C[1]_18         =   -2147483644
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   -2147483644
         B_19            =   -2147483644
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
      Begin CWUIControlsLib.CWButton cwbtnLimit 
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   73
         Top             =   2280
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
         C[0]_11         =   128
         C[1]_11         =   128
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   128
         B_12            =   128
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
         Color_14        =   128
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
         Color_15        =   128
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
         C[0]_21         =   255
         C[1]_21         =   255
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   255
         B_22            =   255
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
         Color_24        =   255
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
         Color_25        =   255
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
         mechAction_1    =   3
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
      Begin CWUIControlsLib.CWButton cwbtnLimit 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   74
         Top             =   2280
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
         C[0]_11         =   128
         C[1]_11         =   128
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   128
         B_12            =   128
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
         Color_14        =   128
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
         Color_15        =   128
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
         C[0]_21         =   255
         C[1]_21         =   255
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   255
         B_22            =   255
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
         Color_24        =   255
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
         Color_25        =   255
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
         mechAction_1    =   3
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
      Begin CWUIControlsLib.CWButton cwbtnLimit 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   75
         Top             =   2760
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
         C[0]_11         =   128
         C[1]_11         =   128
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   128
         B_12            =   128
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
         Color_14        =   128
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
         Color_15        =   128
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
         C[0]_21         =   255
         C[1]_21         =   255
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   255
         B_22            =   255
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
         Color_24        =   255
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
         Color_25        =   255
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
         mechAction_1    =   3
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
      Begin CWUIControlsLib.CWButton cwbtnLimit 
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   76
         Top             =   2760
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
         C[0]_11         =   128
         C[1]_11         =   128
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   128
         B_12            =   128
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
         Color_14        =   128
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
         Color_15        =   128
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
         C[0]_21         =   255
         C[1]_21         =   255
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   255
         B_22            =   255
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
         Color_24        =   255
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
         Color_25        =   255
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
         mechAction_1    =   3
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
      Begin VB.Label lblLimit 
         Caption         =   "Limit #1"
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
         Index           =   0
         Left            =   840
         TabIndex        =   80
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblLimit 
         Caption         =   "Limit #3"
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
         Index           =   1
         Left            =   840
         TabIndex        =   79
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblLimit 
         Caption         =   "Limit #2"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   78
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblLimit 
         Caption         =   "Limit #4"
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
         Index           =   3
         Left            =   2400
         TabIndex        =   77
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblPeak 
         Caption         =   "Peak Value"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblValley 
         Caption         =   "Valley Value"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblTracking 
         Caption         =   "Tracking Value"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraCommunicationStatus 
      Caption         =   "Communication Status"
      Height          =   1815
      Left            =   240
      TabIndex        =   52
      Top             =   960
      Width           =   3735
      Begin VB.ComboBox cboComPortNum 
         Height          =   330
         Left            =   1920
         TabIndex        =   57
         Text            =   "5"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdResetSensotecCommunication 
         Caption         =   "Reset Sensotec Communication"
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   1080
         Width           =   3255
      End
      Begin CWUIControlsLib.CWButton cwbtnCommunicationActive 
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   480
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
         style_6         =   47095284
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
         style_8         =   6
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
         style_10        =   6
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
         C[0]_21         =   65280
         C[1]_21         =   65280
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   65280
         B_22            =   65280
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
         Color_24        =   65280
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
         Color_25        =   65280
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
         mechAction_1    =   3
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
         TabIndex        =   56
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblComPortNumber 
         Caption         =   "Comm Port #"
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer tmrSensotec 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   7920
   End
   Begin VB.Frame fraParameters 
      Caption         =   "Parameters"
      Height          =   2415
      Left            =   240
      TabIndex        =   47
      Top             =   3000
      Width           =   3735
      Begin VB.TextBox txtTestPanelUpdateRate 
         Height          =   315
         Left            =   2160
         TabIndex        =   71
         Text            =   "1000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cboFrequencyResponse 
         Height          =   330
         Left            =   240
         TabIndex        =   58
         Text            =   "100"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdSetParameters 
         Caption         =   "Set Parameters"
         Height          =   495
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtDACFullScale 
         Height          =   315
         Left            =   240
         TabIndex        =   48
         Text            =   "50"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblTestPanelUpdateRate 
         Caption         =   "Test Panel Update Rate (milliseconds) minimum 1000"
         Height          =   855
         Left            =   2160
         TabIndex        =   72
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblFrequencyResponse 
         Caption         =   "Frequency Reponse"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblDACFullScale 
         Caption         =   "DAC Full Scale Output"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraTare 
      Caption         =   "Tare (Channel 1)"
      Height          =   2055
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   3735
      Begin CWUIControlsLib.CWButton cwbtnTare 
         Height          =   375
         Left            =   240
         TabIndex        =   5
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
         style_6         =   47095284
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
         style_8         =   6
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
         style_10        =   6
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   32896
         C[1]_11         =   32896
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   32896
         B_12            =   32896
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
         Color_14        =   32896
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
         Color_15        =   32896
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
         C[0]_21         =   65535
         C[1]_21         =   65535
         Image_21        =   22
         ClassName_22    =   "CCWPiccListImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         F_22            =   65535
         B_22            =   65535
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
         Color_24        =   65535
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
         Color_25        =   65535
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
         mechAction_1    =   3
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
      Begin VB.CommandButton cmdDeActivateTare 
         Caption         =   "De-Activate Tare"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton cmdActivateTare 
         Caption         =   "Activate Tare"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblTare 
         Caption         =   "Tare Activated"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frLimitOperation 
      Caption         =   "Limit Operation (Channel 1)"
      Height          =   7455
      Left            =   8400
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.TextBox txtLimitReturnPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtLimitSetPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Frame fraLimitMode 
         Caption         =   "Limit Mode"
         Enabled         =   0   'False
         Height          =   1455
         Index           =   4
         Left            =   1560
         TabIndex        =   32
         Top             =   5040
         Width           =   1935
         Begin VB.OptionButton optOutsideWindow 
            Caption         =   "Outside Window"
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optInsideWindow 
            Caption         =   "Inside Window"
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optLess 
            Caption         =   "Less Than"
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optGreater 
            Caption         =   "Greater Than"
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame fraLimitMode 
         Caption         =   "Limit Mode"
         Enabled         =   0   'False
         Height          =   1455
         Index           =   3
         Left            =   1560
         TabIndex        =   29
         Top             =   3480
         Width           =   1935
         Begin VB.OptionButton optOutsideWindow 
            Caption         =   "Outside Window"
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optInsideWindow 
            Caption         =   "Inside Window"
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optLess 
            Caption         =   "Less Than"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optGreater 
            Caption         =   "Greater Than"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox txtLimitReturnPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtLimitSetPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Frame fraLimitMode 
         Caption         =   "Limit Mode"
         Enabled         =   0   'False
         Height          =   1455
         Index           =   2
         Left            =   1560
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
         Begin VB.OptionButton optOutsideWindow 
            Caption         =   "Outside Window"
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optInsideWindow 
            Caption         =   "Inside Window"
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optGreater 
            Caption         =   "Greater Than"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optLess 
            Caption         =   "Less Than"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.TextBox txtLimitReturnPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtLimitSetPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame fraLimitMode 
         Caption         =   "Limit Mode"
         Enabled         =   0   'False
         Height          =   1455
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton optOutsideWindow 
            Caption         =   "Outside Window"
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optInsideWindow 
            Caption         =   "Inside Window"
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optLess 
            Caption         =   "Less Than"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optGreater 
            Caption         =   "Greater Than"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox txtLimitReturnPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtLimitSetPoint 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit #4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit #3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSetLimitOperation 
         Caption         =   "Set Limit Operation"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   6720
         Width           =   3255
      End
      Begin VB.Label lblLimit4ReturnPoint 
         Caption         =   "Return Point"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label lblLimit4SetPoint 
         Caption         =   "Set Point"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblLimit3ReturnPoint 
         Caption         =   "Return Point"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label lblLimit3SetPoint 
         Caption         =   "Set Point"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblLimit2ReturnPoint 
         Caption         =   "Return Point"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblLimit2SetPoint 
         Caption         =   "Set Point"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLimit1ReturnPoint 
         Caption         =   "Return Point"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLimit1SetPoint 
         Caption         =   "Set Point"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSCommLib.MSComm SerialPort 
      Left            =   7440
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6960
      Tag             =   "done"
      Top             =   7920
   End
End
Attribute VB_Name = "frmSensotec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnKillTimeDone As Boolean

Private Sub chkLimit_Click(Index As Integer)

    If chkLimit(Index).Value = Checked Then
        txtLimitSetPoint(Index).Enabled = True
        txtLimitReturnPoint(Index).Enabled = True
        fraLimitMode(Index).Enabled = True
    Else
        txtLimitSetPoint(Index).Enabled = False
        txtLimitReturnPoint(Index).Enabled = False
        fraLimitMode(Index).Enabled = False
    End If

End Sub

Private Sub cmdActivateTare_Click()

If Sensotec.GetLinkStatus Then
    'Turn the tare on
    Call Sensotec.ActivateTare(1)
    cwbtnTare.Value = True
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdClearPeakValley_Click()

If Sensotec.GetLinkStatus Then
    Call Sensotec.ClearPeakAndValley(1)
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdDeActivateTare_Click()

If Sensotec.GetLinkStatus Then
    'Turn the tare off
    Call Sensotec.DeActivateTare(1)
    cwbtnTare.Value = False
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdExit_Click()

tmrSensotec.Enabled = False

'Close the window
Unload Me

End Sub

Private Sub cmdForceDACOutput_Click()

If Sensotec.GetLinkStatus Then
    Call Sensotec.ForceDACOutput(1, sldPercentage.Value, True)
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdNormalDAC_Click()

If Sensotec.GetLinkStatus Then
    Call Sensotec.ForceDACOutput(1, 0, False)
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

End Sub

Private Sub cmdResetSensotecCommunication_Click()

On Error GoTo BadComboBox

'Establish Communication with the SC Instrument
Call Sensotec.InitializeCommunication(CInt(cboComPortNum.Text))

cwbtnLimit(1).Value = False
cwbtnLimit(2).Value = False
cwbtnLimit(3).Value = False
cwbtnLimit(4).Value = False

tmrSensotec.Enabled = True

Exit Sub
'Error Trap
BadComboBox:

    MsgBox "Invalid Com Port Number. " _
           & vbCrLf & "Try again.", vbOKOnly + vbCritical, "Error"

End Sub

Private Sub cmdSetLimitOperation_Click()

Dim i As Integer

On Error GoTo BadTextBox1

If Sensotec.GetLinkStatus Then
    'Loop through all the limits
    For i = 1 To 4
        If chkLimit(i).Value = Checked Then
            'Set the limit operation according to the radio buttons
            If optGreater(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, True, leSignalGreaterThanSetPoint)
            ElseIf optLess(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, True, leSignalLessThanSetPoint)
            ElseIf optInsideWindow(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, True, leSignalInside)
            ElseIf optOutsideWindow(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, True, leSignalOutside)
            End If
            
            'Set the limits
            Call SetLimitPoints(i, CSng(txtLimitSetPoint(i)), CSng(txtLimitReturnPoint(i)))
    
        Else
            'Set the limit operation according to the radio buttons
            If optGreater(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, False, leSignalGreaterThanSetPoint)
            ElseIf optLess(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, False, leSignalLessThanSetPoint)
            ElseIf optInsideWindow(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, False, leSignalInside)
            ElseIf optOutsideWindow(i).Value Then
                Call Sensotec.SetLimitOperation(i, 1, False, leSignalOutside)
            End If
        
        End If
    
    Next i
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub
'Error Trap
BadTextBox1:
    MsgBox "At least one of the text boxes contains an invalid" _
       & vbCrLf & "value.  Please correct the error and try again." _
       , vbOKOnly + vbCritical, "Set Point/Return Point Error"

End Sub

Private Sub cmdSetParameters_Click()

On Error GoTo BadTextBox2

'Make sure the update period is at least 1000 milliseconds
If CInt(txtTestPanelUpdateRate) < 1000 Then
    tmrSensotec.Interval = 1000
    txtTestPanelUpdateRate.Text = "1000"
End If
'Set the update period
tmrSensotec.Interval = CInt(txtTestPanelUpdateRate.Text)

If Sensotec.GetLinkStatus Then
    'Set the DAC full scale
    Call Sensotec.SetDACFullScale(1, CInt(txtDACFullScale.Text))
    
    'Set the frequency response
    Call Sensotec.SetFreqResponse(1, CInt(cboFrequencyResponse.Text))
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

Exit Sub

BadTextBox2:

    MsgBox "At least one of the text boxes contains an invalid" _
       & vbCrLf & "value.  Please correct the error and try again." _
       , vbOKOnly + vbCritical, "DAC Full Scale / Frequency Response Error"

End Sub

Private Sub cmdUpdateOnce_Click()

Dim i As Integer
Dim lsngTracking As Single
Dim lsngPeak As Single
Dim lsngValley As Single
Dim lblnLimit(1 To 4) As Boolean

If Sensotec.GetLinkStatus Then
    'Update the tracking value
    Call Sensotec.ReadTracking(1, lsngTracking)
    txtTracking.Text = CStr(lsngTracking)
    
    'Update the peak value
    Call Sensotec.ReadPeak(1, lsngPeak)
    txtPeak.Text = CStr(lsngPeak)
    
    'Update the valley value
    Call Sensotec.ReadValley(1, lsngValley)
    txtValley.Text = CStr(lsngValley)
    
    'Read the status of the limits
    Call Sensotec.ReadLimitStatus(lblnLimit())
Else
    MsgBox "Communication is not established." _
       & vbCrLf & "Please reset communication with the Sensotec Force Amplifier." _
       , vbOKOnly + vbCritical, "Bad Comm Link"
End If

If Sensotec.GetLinkStatus Then
    
    'Turn on the proper indicators
    For i = 1 To 4
            cwbtnLimit(i).Value = lblnLimit(i)
    Next i

End If

End Sub

Private Sub cwbtnContinuousRead_ValueChanged(ByVal Value As Boolean)

'While continuously updating, don't allow other button presses
    
    If Value Then
        cmdUpdateOnce.Enabled = False
        cmdActivateTare.Enabled = False
        cmdDeActivateTare.Enabled = False
        cmdResetSensotecCommunication.Enabled = False
        cmdSetLimitOperation.Enabled = False
        cmdSetParameters.Enabled = False
        cmdForceDACOutput.Enabled = False
        cmdNormalDAC.Enabled = False
        
        
    Else
        cmdUpdateOnce.Enabled = True
        cmdActivateTare.Enabled = True
        cmdDeActivateTare.Enabled = True
        cmdResetSensotecCommunication.Enabled = True
        cmdSetLimitOperation.Enabled = True
        cmdSetParameters.Enabled = True
        cmdForceDACOutput.Enabled = True
        cmdNormalDAC.Enabled = True
    
    End If

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To initialize form
'
'  INPUT(S): none
' OUTPUT(S): none

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: To unload form
'
'  INPUT(S): none
' OUTPUT(S): none

tmrSensotec.Enabled = False
tmrKillTime.Enabled = False
SerialPort.PortOpen = False
'Set the link status to false (the port is closed when the form is unloaded)
Sensotec.SetLinkStatus (False)

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

Private Sub tmrSensotec_Timer()

Dim i As Integer

If Sensotec.GetLinkStatus Then
    'Update the Communication Active LED
    cwbtnCommunicationActive.Value = True
    
    If cwbtnContinuousRead.Value Then Call cmdUpdateOnce_Click
        
Else
    cwbtnCommunicationActive.Value = False
    
    'Turn limit indicators off
    For i = 1 To 4
        cwbtnLimit(i).Value = False
    Next i
    'Turn tare indicator off
    cwbtnTare.Value = False

End If

End Sub

Private Sub tmrKillTime_Timer()

mblnKillTimeDone = True

End Sub
