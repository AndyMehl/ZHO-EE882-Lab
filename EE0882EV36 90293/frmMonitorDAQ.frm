VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form frmMonitorDAQ 
   Caption         =   "Monitor Data Aqcuisition"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrADReadings 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   20
      Top             =   3840
      Width           =   3615
   End
   Begin CWUIControlsLib.CWNumEdit CWNumEditDA0 
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   480
      Width           =   1695
      _Version        =   393218
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   393218
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   393278
      BorderStyle_1   =   1
      ButtonPosition_1=   1
      TextAlignment_1 =   2
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".##0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      IncValueVarType_1=   5
      IncValue_Val_1  =   0.1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   100
      ButtonStyle_1   =   0
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWNumEdit CWNumEditDA1 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   1200
      Width           =   1695
      _Version        =   393218
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   393218
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   393278
      BorderStyle_1   =   1
      ButtonPosition_1=   1
      TextAlignment_1 =   2
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".##0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      IncValueVarType_1=   5
      IncValue_Val_1  =   0.1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   100
      ButtonStyle_1   =   0
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWButton EnablePath 
      Height          =   615
      Left            =   4560
      TabIndex        =   40
      Top             =   3000
      Width           =   3615
      _Version        =   393218
      _ExtentX        =   6376
      _ExtentY        =   1085
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      C[0]_1          =   12632256
      Enum_1          =   2
      ClassName_2     =   "CCWEnum"
      Array_2         =   2
      Editor_2        =   0
      Array[0]_2      =   3
      ClassName_3     =   "CCWEnumElt"
      opts_3          =   1
      Name_3          =   "Off"
      frame_3         =   535
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
      szText_8        =   "Thru Path Enabled"
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
      szText_10       =   "Parameter File Filter/Load"
      font_10         =   0
      Animator_9      =   0
      Blinker_9       =   0
      list[1]_4       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   255
      C[1]_11         =   255
      Image_11        =   12
      ClassName_12    =   "CCWPictImage"
      opts_12         =   1280
      Rows_12         =   1
      Cols_12         =   1
      Pict_12         =   286
      F_12            =   255
      B_12            =   255
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
      frame_13        =   536
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
      C[0]_18         =   65280
      C[1]_18         =   65280
      Image_18        =   19
      ClassName_19    =   "CCWPictImage"
      opts_19         =   1280
      Rows_19         =   1
      Cols_19         =   1
      Pict_19         =   286
      F_19            =   65280
      B_19            =   65280
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
      Style_1         =   24
      frameStyle_1    =   2
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "FORCE (CH4) ="
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
      Index           =   27
      Left            =   4440
      TabIndex        =   39
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   7920
      TabIndex        =   38
      Top             =   2430
      Width           =   255
   End
   Begin VB.Label lblForce 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   7920
      TabIndex        =   36
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   3720
      TabIndex        =   35
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   3720
      TabIndex        =   34
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   3720
      TabIndex        =   33
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   3720
      TabIndex        =   32
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   3720
      TabIndex        =   31
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   3720
      TabIndex        =   30
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   8160
      TabIndex        =   29
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   8160
      TabIndex        =   28
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   3720
      TabIndex        =   27
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   26
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Supply)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "D / A Chan 1  ="
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
      Index           =   15
      Left            =   4560
      TabIndex        =   24
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Not Used)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4560
      TabIndex        =   23
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Supply Control)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   22
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "D / A Chan 0  ="
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
      Index           =   13
      Left            =   4560
      TabIndex        =   21
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "POSITION  ="
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
      Index           =   11
      Left            =   4560
      TabIndex        =   16
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2160
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2160
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblADChanNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 7  ="
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
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 2  ="
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
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 3  ="
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
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 4  ="
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
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 5  ="
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
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 6  ="
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
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 1  ="
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
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A / D Chan 0  ="
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmMonitorDAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

'Unload Monitor DAQ form
Unload Me

End Sub

Private Sub CWNumEditDA0_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

'Update the D/A output
Call frmDAQIO.cwaoVRef.SingleWrite(Value)

End Sub

Private Sub EnablePath_ValueChanged(ByVal Value As Boolean)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

Dim lintChanNum As Integer

If EnablePath.Value = True Then
    'Enable the straight-thru path
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, 1, True)
    Next lintChanNum
Else
    'Enable the proper filter
    For lintChanNum = CHAN0 To CHAN3
        Call SelectFilter(lintChanNum, gudtMachine.filterLoc(lintChanNum), True)
    Next lintChanNum
End If

End Sub

Private Sub Form_Load()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

Me.top = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2

tmrADReadings.Enabled = True
'Set the VRef Point
frmMonitorDAQ.CWNumEditDA0.Value = gsngVRefSetPoint

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

tmrADReadings.Enabled = False

End Sub

Private Sub tmrADReadings_Timer()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

Dim lintChanNum As Integer

tmrADReadings.Enabled = False

Call frmDAQIO.MonitorDAQRead(gsngMonitorData())

'Display A/D channels (0-7)
For lintChanNum = CHAN0 To CHAN7
    lblADChanNum(lintChanNum).Caption = Format(gsngMonitorData(lintChanNum), "#0.000")
Next lintChanNum

'Display current position
lblPosition.Caption = Format(Pedal.Position, "#0.00")

'Display Force
lblForce.Caption = Format((gsngMonitorData(CHAN4) * gsngNewtonsPerVolt) + gsngForceAmplifierOffset, "##0.##0")

tmrADReadings.Enabled = True

End Sub
