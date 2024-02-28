VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.UserControl ctrStatus 
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   LockControls    =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   1350
   Begin CWUIControlsLib.CWButton cwbtnStatus 
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _Version        =   393218
      _ExtentX        =   2355
      _ExtentY        =   2143
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
      szText_6        =   "Status 1"
      style_6         =   1076101120
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
      font_10         =   0
      Animator_9      =   0
      Blinker_9       =   0
      list[1]_4       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   12632256
      C[1]_11         =   12632256
      Image_11        =   12
      ClassName_12    =   "CCWPiccListImage"
      opts_12         =   1280
      Rows_12         =   1
      Cols_12         =   1
      F_12            =   12632256
      B_12            =   12632256
      ColorReplaceWith_12=   8421504
      ColorReplace_12 =   8421504
      Tolerance_12    =   2
      UsePiccList_12  =   -1  'True
      PiccList_12     =   13
      ClassName_13    =   "CPiccListSquareLED"
      count_13        =   2
      list[2]_13      =   14
      ClassName_14    =   "CCWPicc"
      opts_14         =   62
      Image_14        =   0
      Animator_14     =   0
      Blinker_14      =   0
      Size_14.cx      =   30
      Size_14.cy      =   14
      Model_14.r      =   30
      Model_14.b      =   14
      Actual_14.t     =   20
      Actual_14.r     =   89
      Actual_14.b     =   81
      Picc_14         =   412
      Color_14        =   12632256
      Name_14         =   "Divot"
      list[1]_13      =   15
      ClassName_15    =   "CCWPicc"
      opts_15         =   62
      Image_15        =   0
      Animator_15     =   0
      Blinker_15      =   0
      Size_15.cx      =   30
      Size_15.cy      =   14
      Model_15.l      =   3
      Model_15.t      =   3
      Model_15.r      =   27
      Model_15.b      =   11
      Actual_15.l     =   3
      Actual_15.t     =   23
      Actual_15.r     =   86
      Actual_15.b     =   78
      Picc_15         =   441
      Color_15        =   12632256
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
      C[0]_19         =   0
      C[1]_19         =   0
      Image_19        =   8
      Animator_19     =   0
      Blinker_19      =   0
      list[2]_17      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   60
      C[0]_20         =   0
      C[1]_20         =   0
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
      ClassName_23    =   "CPiccListSquareLED"
      count_23        =   2
      list[2]_23      =   24
      ClassName_24    =   "CCWPicc"
      opts_24         =   62
      Image_24        =   0
      Animator_24     =   0
      Blinker_24      =   0
      Size_24.cx      =   30
      Size_24.cy      =   14
      Model_24.r      =   30
      Model_24.b      =   14
      Actual_24.t     =   20
      Actual_24.r     =   89
      Actual_24.b     =   81
      Picc_24         =   412
      Color_24        =   65380
      Name_24         =   "Divot"
      list[1]_23      =   25
      ClassName_25    =   "CCWPicc"
      opts_25         =   62
      Image_25        =   0
      Animator_25     =   0
      Blinker_25      =   0
      Size_25.cx      =   30
      Size_25.cy      =   14
      Model_25.l      =   3
      Model_25.t      =   3
      Model_25.r      =   27
      Model_25.b      =   11
      Actual_25.l     =   3
      Actual_25.t     =   23
      Actual_25.r     =   86
      Actual_25.b     =   78
      Picc_25         =   441
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
      Style_1         =   17
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
   Begin CWUIControlsLib.CWButton cwbtnStatus 
      Height          =   1215
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _Version        =   393218
      _ExtentX        =   2355
      _ExtentY        =   2143
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
      szText_6        =   "Status 2"
      style_6         =   1076101120
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
      font_10         =   0
      Animator_9      =   0
      Blinker_9       =   0
      list[1]_4       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   12632256
      C[1]_11         =   12632256
      Image_11        =   12
      ClassName_12    =   "CCWPiccListImage"
      opts_12         =   1280
      Rows_12         =   1
      Cols_12         =   1
      F_12            =   12632256
      B_12            =   12632256
      ColorReplaceWith_12=   8421504
      ColorReplace_12 =   8421504
      Tolerance_12    =   2
      UsePiccList_12  =   -1  'True
      PiccList_12     =   13
      ClassName_13    =   "CPiccListSquareLED"
      count_13        =   2
      list[2]_13      =   14
      ClassName_14    =   "CCWPicc"
      opts_14         =   62
      Image_14        =   0
      Animator_14     =   0
      Blinker_14      =   0
      Size_14.cx      =   30
      Size_14.cy      =   14
      Model_14.r      =   30
      Model_14.b      =   14
      Actual_14.t     =   20
      Actual_14.r     =   89
      Actual_14.b     =   81
      Picc_14         =   412
      Color_14        =   12632256
      Name_14         =   "Divot"
      list[1]_13      =   15
      ClassName_15    =   "CCWPicc"
      opts_15         =   62
      Image_15        =   0
      Animator_15     =   0
      Blinker_15      =   0
      Size_15.cx      =   30
      Size_15.cy      =   14
      Model_15.l      =   3
      Model_15.t      =   3
      Model_15.r      =   27
      Model_15.b      =   11
      Actual_15.l     =   3
      Actual_15.t     =   23
      Actual_15.r     =   86
      Actual_15.b     =   78
      Picc_15         =   441
      Color_15        =   12632256
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
      C[0]_19         =   0
      C[1]_19         =   0
      Image_19        =   8
      Animator_19     =   0
      Blinker_19      =   0
      list[2]_17      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   60
      C[0]_20         =   0
      C[1]_20         =   0
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
      ClassName_23    =   "CPiccListSquareLED"
      count_23        =   2
      list[2]_23      =   24
      ClassName_24    =   "CCWPicc"
      opts_24         =   62
      Image_24        =   0
      Animator_24     =   0
      Blinker_24      =   0
      Size_24.cx      =   30
      Size_24.cy      =   14
      Model_24.r      =   30
      Model_24.b      =   14
      Actual_24.t     =   20
      Actual_24.r     =   89
      Actual_24.b     =   81
      Picc_24         =   412
      Color_24        =   65380
      Name_24         =   "Divot"
      list[1]_23      =   25
      ClassName_25    =   "CCWPicc"
      opts_25         =   62
      Image_25        =   0
      Animator_25     =   0
      Blinker_25      =   0
      Size_25.cx      =   30
      Size_25.cy      =   14
      Model_25.l      =   3
      Model_25.t      =   3
      Model_25.r      =   27
      Model_25.b      =   11
      Actual_25.l     =   3
      Actual_25.t     =   23
      Actual_25.r     =   86
      Actual_25.b     =   78
      Picc_25         =   441
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
      Style_1         =   17
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
   Begin CWUIControlsLib.CWButton cwbtnStatus 
      Height          =   1215
      Index           =   4
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
      _Version        =   393218
      _ExtentX        =   2355
      _ExtentY        =   2143
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
      szText_6        =   "Status 4"
      style_6         =   1076101120
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
      font_10         =   0
      Animator_9      =   0
      Blinker_9       =   0
      list[1]_4       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   12632256
      C[1]_11         =   12632256
      Image_11        =   12
      ClassName_12    =   "CCWPiccListImage"
      opts_12         =   1280
      Rows_12         =   1
      Cols_12         =   1
      F_12            =   12632256
      B_12            =   12632256
      ColorReplaceWith_12=   8421504
      ColorReplace_12 =   8421504
      Tolerance_12    =   2
      UsePiccList_12  =   -1  'True
      PiccList_12     =   13
      ClassName_13    =   "CPiccListSquareLED"
      count_13        =   2
      list[2]_13      =   14
      ClassName_14    =   "CCWPicc"
      opts_14         =   62
      Image_14        =   0
      Animator_14     =   0
      Blinker_14      =   0
      Size_14.cx      =   30
      Size_14.cy      =   14
      Model_14.r      =   30
      Model_14.b      =   14
      Actual_14.t     =   20
      Actual_14.r     =   89
      Actual_14.b     =   81
      Picc_14         =   412
      Color_14        =   12632256
      Name_14         =   "Divot"
      list[1]_13      =   15
      ClassName_15    =   "CCWPicc"
      opts_15         =   62
      Image_15        =   0
      Animator_15     =   0
      Blinker_15      =   0
      Size_15.cx      =   30
      Size_15.cy      =   14
      Model_15.l      =   3
      Model_15.t      =   3
      Model_15.r      =   27
      Model_15.b      =   11
      Actual_15.l     =   3
      Actual_15.t     =   23
      Actual_15.r     =   86
      Actual_15.b     =   78
      Picc_15         =   441
      Color_15        =   12632256
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
      C[0]_19         =   0
      C[1]_19         =   0
      Image_19        =   8
      Animator_19     =   0
      Blinker_19      =   0
      list[2]_17      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   60
      C[0]_20         =   0
      C[1]_20         =   0
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
      ClassName_23    =   "CPiccListSquareLED"
      count_23        =   2
      list[2]_23      =   24
      ClassName_24    =   "CCWPicc"
      opts_24         =   62
      Image_24        =   0
      Animator_24     =   0
      Blinker_24      =   0
      Size_24.cx      =   30
      Size_24.cy      =   14
      Model_24.r      =   30
      Model_24.b      =   14
      Actual_24.t     =   20
      Actual_24.r     =   89
      Actual_24.b     =   81
      Picc_24         =   412
      Color_24        =   65380
      Name_24         =   "Divot"
      list[1]_23      =   25
      ClassName_25    =   "CCWPicc"
      opts_25         =   62
      Image_25        =   0
      Animator_25     =   0
      Blinker_25      =   0
      Size_25.cx      =   30
      Size_25.cy      =   14
      Model_25.l      =   3
      Model_25.t      =   3
      Model_25.r      =   27
      Model_25.b      =   11
      Actual_25.l     =   3
      Actual_25.t     =   23
      Actual_25.r     =   86
      Actual_25.b     =   78
      Picc_25         =   441
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
      Style_1         =   17
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
   Begin CWUIControlsLib.CWButton cwbtnStatus 
      Height          =   1215
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
      _Version        =   393218
      _ExtentX        =   2355
      _ExtentY        =   2143
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
      szText_6        =   "Status 3"
      style_6         =   1076101120
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
      font_10         =   0
      Animator_9      =   0
      Blinker_9       =   0
      list[1]_4       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   12632256
      C[1]_11         =   12632256
      Image_11        =   12
      ClassName_12    =   "CCWPiccListImage"
      opts_12         =   1280
      Rows_12         =   1
      Cols_12         =   1
      F_12            =   12632256
      B_12            =   12632256
      ColorReplaceWith_12=   8421504
      ColorReplace_12 =   8421504
      Tolerance_12    =   2
      UsePiccList_12  =   -1  'True
      PiccList_12     =   13
      ClassName_13    =   "CPiccListSquareLED"
      count_13        =   2
      list[2]_13      =   14
      ClassName_14    =   "CCWPicc"
      opts_14         =   62
      Image_14        =   0
      Animator_14     =   0
      Blinker_14      =   0
      Size_14.cx      =   30
      Size_14.cy      =   14
      Model_14.r      =   30
      Model_14.b      =   14
      Actual_14.t     =   20
      Actual_14.r     =   89
      Actual_14.b     =   81
      Picc_14         =   412
      Color_14        =   12632256
      Name_14         =   "Divot"
      list[1]_13      =   15
      ClassName_15    =   "CCWPicc"
      opts_15         =   62
      Image_15        =   0
      Animator_15     =   0
      Blinker_15      =   0
      Size_15.cx      =   30
      Size_15.cy      =   14
      Model_15.l      =   3
      Model_15.t      =   3
      Model_15.r      =   27
      Model_15.b      =   11
      Actual_15.l     =   3
      Actual_15.t     =   23
      Actual_15.r     =   86
      Actual_15.b     =   78
      Picc_15         =   441
      Color_15        =   12632256
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
      C[0]_19         =   0
      C[1]_19         =   0
      Image_19        =   8
      Animator_19     =   0
      Blinker_19      =   0
      list[2]_17      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   60
      C[0]_20         =   0
      C[1]_20         =   0
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
      ClassName_23    =   "CPiccListSquareLED"
      count_23        =   2
      list[2]_23      =   24
      ClassName_24    =   "CCWPicc"
      opts_24         =   62
      Image_24        =   0
      Animator_24     =   0
      Blinker_24      =   0
      Size_24.cx      =   30
      Size_24.cy      =   14
      Model_24.r      =   30
      Model_24.b      =   14
      Actual_24.t     =   20
      Actual_24.r     =   89
      Actual_24.b     =   81
      Picc_24         =   412
      Color_24        =   65380
      Name_24         =   "Divot"
      list[1]_23      =   25
      ClassName_25    =   "CCWPicc"
      opts_25         =   62
      Image_25        =   0
      Animator_25     =   0
      Blinker_25      =   0
      Size_25.cx      =   30
      Size_25.cy      =   14
      Model_25.l      =   3
      Model_25.t      =   3
      Model_25.r      =   27
      Model_25.b      =   11
      Actual_25.l     =   3
      Actual_25.t     =   23
      Actual_25.r     =   86
      Actual_25.b     =   78
      Picc_25         =   441
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
      Style_1         =   17
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
End
Attribute VB_Name = "ctrStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Revision   Date      Initials  Explanation
' 1.0     09-15-2004    SRC     New User Control for Status Display

Option Explicit

Public Property Get StatusCaption(IndicatorNum As Integer) As String
'
'   PURPOSE: Return the Caption of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusCaption = cwbtnStatus(IndicatorNum).Caption

End Property

Public Property Let StatusCaption(IndicatorNum As Integer, NewCaption As String)
'
'   PURPOSE: Set the Caption of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewCaption   = New Caption
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Caption = NewCaption

End Property

Public Property Get StatusFont(IndicatorNum As Integer) As String
'
'   PURPOSE: Return the Font of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusFont = cwbtnStatus(IndicatorNum).Font

End Property

Public Property Let StatusFont(IndicatorNum As Integer, NewFont As String)
'
'   PURPOSE: Set the Font of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewFont   = New Font
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Font = NewFont

End Property

Public Property Get StatusFontBold(IndicatorNum As Integer) As Boolean
'
'   PURPOSE: Return the Font.Bold property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusFontBold = cwbtnStatus(IndicatorNum).Font.Bold

End Property

Public Property Let StatusFontBold(IndicatorNum As Integer, FontBold As Boolean)
'
'   PURPOSE: Set the Font.Bold of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            FontBold     = Whether or not to make the font bold
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Font.Bold = FontBold

End Property

Public Property Get StatusFontSize(IndicatorNum As Integer) As Integer
'
'   PURPOSE: Return the Font.Size property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusFontSize = cwbtnStatus(IndicatorNum).Font.Size

End Property

Public Property Let StatusFontSize(IndicatorNum As Integer, FontSize As Integer)
'
'   PURPOSE: Set the Font.Size of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            FontSize     = New Font Size
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Font.Size = FontSize

End Property

Public Property Get StatusOffColor(IndicatorNum As Integer) As ColorConstants
'
'   PURPOSE: Return the OffColor property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusOffColor = cwbtnStatus(IndicatorNum).OffColor

End Property

Public Property Let StatusOffColor(IndicatorNum As Integer, NewOffColor As ColorConstants)
'
'   PURPOSE: Set the OffColor property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewOffColor   = New Off Color
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).OffColor = NewOffColor

End Property

Public Property Get StatusOffText(IndicatorNum As Integer) As String
'
'   PURPOSE: Return the OffText property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusOffText = cwbtnStatus(IndicatorNum).OffText

End Property

Public Property Let StatusOffText(IndicatorNum As Integer, NewOffText As String)
'
'   PURPOSE: Set the OffText property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewOffText   = New Off Text
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).OffText = NewOffText

End Property

Public Property Get StatusOnColor(IndicatorNum As Integer) As ColorConstants
'
'   PURPOSE: Return the OnColor property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusOnColor = cwbtnStatus(IndicatorNum).OnColor

End Property

Public Property Let StatusOnColor(IndicatorNum As Integer, NewOnColor As ColorConstants)
'
'   PURPOSE: Set the OnColor property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewOnColor   = New On Color
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).OnColor = NewOnColor

End Property

Public Property Get StatusOnText(IndicatorNum As Integer) As String
'
'   PURPOSE: Return the OnText property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusOnText = cwbtnStatus(IndicatorNum).OnText

End Property

Public Property Let StatusOnText(IndicatorNum As Integer, NewOnText As String)
'
'   PURPOSE: Set the OnText property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            NewOnText    = New On Text
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).OnText = NewOnText

End Property

Public Property Get StatusVisible(IndicatorNum As Integer) As Boolean
'
'   PURPOSE: Return the Visible property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusVisible = cwbtnStatus(IndicatorNum).Visible

End Property

Public Property Let StatusVisible(IndicatorNum As Integer, VisibleIndicator As Boolean)
'
'   PURPOSE: Set the Visible property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum     = Selected Indicator
'            VisibleIndicator = Whether or not the indicator is visible
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Visible = VisibleIndicator

End Property

Public Property Get StatusValue(IndicatorNum As Integer) As Boolean
'
'   PURPOSE: Return the Value property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'
' OUTPUT(S):

StatusValue = cwbtnStatus(IndicatorNum).Value

End Property

Public Property Let StatusValue(IndicatorNum As Integer, OnOrOff As Boolean)
'
'   PURPOSE: Set the Value property of the selected Status Indicator
'
'  INPUT(S): IndicatorNum = Selected Indicator
'            OnOrOff      = Whether the indicator is on or off
'
' OUTPUT(S):

cwbtnStatus(IndicatorNum).Value = OnOrOff

End Property


