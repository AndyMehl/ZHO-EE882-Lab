VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "CWUI.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMLX90293 
   Caption         =   "Melexis 90293 Test Utility"
   ClientHeight    =   12690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   12690
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer tmrMLX 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   120
   End
   Begin VB.Frame fraClampLevels 
      Caption         =   "Clamping Levels (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1575
      Left            =   120
      TabIndex        =   83
      Top             =   3240
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampLo 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   84
         Top             =   600
         Width           =   855
         _Version        =   393218
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampHi 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   85
         Top             =   1080
         Width           =   855
         _Version        =   393218
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampLo 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   86
         Top             =   600
         Width           =   855
         _Version        =   393218
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampHi 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   87
         Top             =   1080
         Width           =   855
         _Version        =   393218
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Clamp Low (0-100)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   91
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Clamp High (0-100)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   90
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   2040
         TabIndex        =   89
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   3000
         TabIndex        =   88
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Offset Temp Comp (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2775
      Left            =   8520
      TabIndex        =   80
      Top             =   9840
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWOSYS1 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   152
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS1 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   153
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS2 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   155
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS3 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   156
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS4 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   157
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS5 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   158
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS2 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   159
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS6 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   160
         Top             =   2280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS3 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   161
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS4 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   162
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS5 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   163
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOSYS6 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   164
         Top             =   2280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @160 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   120
         TabIndex        =   169
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @125 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   120
         TabIndex        =   168
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @ 90 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   120
         TabIndex        =   167
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @ 20 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   166
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @-15 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   165
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @-50 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   154
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   3000
         TabIndex        =   82
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   2040
         TabIndex        =   81
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraCommunicationStatus 
      Caption         =   "Communication Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3015
      Left            =   120
      TabIndex        =   72
      Top             =   0
      Width           =   4095
      Begin VB.ComboBox cboComPortNum 
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   76
         Text            =   "4"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboComPortNum 
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   74
         Text            =   "3"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdResetCommunication 
         Caption         =   "Reset Communication Both PTC-04's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   73
         Top             =   2160
         Width           =   3615
      End
      Begin CWUIControlsLib.CWButton cwbtnActive 
         Height          =   375
         Left            =   240
         TabIndex        =   75
         Top             =   1560
         Width           =   375
         _Version        =   393218
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
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
         C[0]_11         =   19230
         C[1]_11         =   19230
         Image_11        =   12
         ClassName_12    =   "CCWPiccListImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         F_12            =   19230
         B_12            =   19230
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
         Color_14        =   19230
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
         Color_15        =   19230
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
      Begin VB.Label lblPTC04 
         Caption         =   "PTC-03 #1 Comm Port Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   79
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblPTC04 
         Caption         =   "PTC-03 #2 Comm Port Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   78
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblbPTC04CommActive 
         Caption         =   "PTC-04 Communication Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   77
         Top             =   1680
         Width           =   3135
      End
   End
   Begin VB.Frame fraReadWrite 
      Caption         =   "Read/Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1335
      Left            =   4320
      TabIndex        =   67
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdReadEEPROM 
         Caption         =   "Read All EEPROM Locations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdWriteEEPROM 
         Caption         =   "Write to all EEPROM Locations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   69
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdClearData 
         Caption         =   "Clear Display Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   68
         Top             =   720
         Width           =   3615
      End
      Begin CWUIControlsLib.CWButton cwbtnStatus 
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   720
         Width           =   3615
         _Version        =   393218
         _ExtentX        =   6376
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.74
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
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Please Wait"
         style_8         =   16777217
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
         szText_10       =   "Ready"
         style_10        =   16777217
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65280
         C[1]_11         =   65280
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   286
         F_12            =   65280
         B_12            =   65280
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
         C[0]_18         =   65535
         C[1]_18         =   65535
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   286
         F_19            =   65535
         B_19            =   65535
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
         mechAction_1    =   3
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
   Begin VB.Frame fraLock 
      Caption         =   "Locking (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   975
      Left            =   120
      TabIndex        =   63
      Top             =   7320
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWMLXLock 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   92
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWMLXLock 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   93
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Memory Lock (0-3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   66
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   3120
         TabIndex        =   65
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   2040
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraOffset 
      Caption         =   "Output Values (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2055
      Left            =   120
      TabIndex        =   51
      Top             =   5040
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWVG 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   52
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOM 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   53
         Top             =   1080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOS 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   54
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWVG 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   55
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOM 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   56
         Top             =   1080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWOS 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   57
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "VG (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   62
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Output Mode (0-7)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   61
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Output Scaling (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   60
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   2040
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraSensitivity 
      Caption         =   "Linear Set Points (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   6975
      Left            =   4320
      TabIndex        =   43
      Top             =   1560
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWY0 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   44
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY1 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   45
         Top             =   960
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY0 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   46
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY1 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   47
         Top             =   960
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY3 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   95
         Top             =   1680
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY3 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   96
         Top             =   1680
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY4 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   97
         Top             =   2040
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY4 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   98
         Top             =   2040
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY5 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   99
         Top             =   2400
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY5 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   100
         Top             =   2400
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY6 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   101
         Top             =   2760
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY7 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   102
         Top             =   3120
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY8 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   103
         Top             =   3480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY9 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   104
         Top             =   3840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY10 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   105
         Top             =   4200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY11 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   106
         Top             =   4560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY12 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   107
         Top             =   4920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY13 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   108
         Top             =   5280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY14 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   109
         Top             =   5640
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY15 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   110
         Top             =   6000
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY16 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   111
         Top             =   6360
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY6 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   112
         Top             =   2760
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY7 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   113
         Top             =   3120
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY8 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   114
         Top             =   3480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY9 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   115
         Top             =   3840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY10 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   116
         Top             =   4200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY11 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   117
         Top             =   4560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY12 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   118
         Top             =   4920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY13 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   119
         Top             =   5280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY14 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   120
         Top             =   5640
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY15 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   121
         Top             =   6000
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY16 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   122
         Top             =   6360
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY2 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   227
         Top             =   1320
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWY2 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   228
         Top             =   1320
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1023
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Y5 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   83
         Left            =   240
         TabIndex        =   226
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y16 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   240
         TabIndex        =   136
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y15 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   135
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y14 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   240
         TabIndex        =   134
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y13 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   240
         TabIndex        =   133
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y12 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   240
         TabIndex        =   132
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y11 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   131
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y10 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   240
         TabIndex        =   130
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y9 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   240
         TabIndex        =   129
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y8 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   240
         TabIndex        =   128
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y7 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   240
         TabIndex        =   127
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y6 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   126
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y4 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   125
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y3 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   124
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y2 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   123
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y1 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   94
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Y0 (-50 to +150)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   2040
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   48
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraTimeGenerator 
      Caption         =   "Sensitivity Temp Comp (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3855
      Left            =   4320
      TabIndex        =   31
      Top             =   8640
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWBPivot 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   32
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnSSYS 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   33
         Top             =   960
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS1 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   34
         Top             =   1440
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWBPivot 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   35
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnSSYS 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   36
         Top             =   960
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS1 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   37
         Top             =   1440
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS2 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   137
         Top             =   1800
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS3 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   138
         Top             =   2160
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS4 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   139
         Top             =   2520
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS5 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   140
         Top             =   2880
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS2 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   141
         Top             =   1800
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS6 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   142
         Top             =   3240
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS3 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   143
         Top             =   2160
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS4 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   144
         Top             =   2520
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS5 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   145
         Top             =   2880
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSSYS6 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   146
         Top             =   3240
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @160 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   240
         TabIndex        =   151
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @125 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   240
         TabIndex        =   150
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @ 90 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   240
         TabIndex        =   149
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @ 20 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   240
         TabIndex        =   148
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @-15 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   147
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "TC Pivot Pt (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Enable SSYS DS (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   41
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Comp @-50 (0-1.000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   40
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraFiltering 
      Caption         =   "Diagnostics (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2295
      Left            =   8520
      TabIndex        =   25
      Top             =   7440
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWDiagOM 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   26
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDiagOM 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   27
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDiagOL 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   171
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbSD 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   172
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbCT 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   173
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbCI 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   174
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDiagOL 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   175
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbSD 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   176
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbCT 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   177
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDbCI 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   178
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Db Increment (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   66
         Left            =   240
         TabIndex        =   181
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Db Threshold (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   65
         Left            =   240
         TabIndex        =   180
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Db Step Down (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   240
         TabIndex        =   179
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Output Level (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   240
         TabIndex        =   170
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Output Mode (0-7)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Other Settings (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   5895
      Left            =   8520
      TabIndex        =   19
      Top             =   1440
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWLPM 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   20
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWLPM 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   21
         Top             =   480
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnS 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   182
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnS 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   183
         Top             =   840
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnP 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   185
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnP 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   186
         Top             =   1200
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWPPol 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   188
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWPPol 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   189
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnPS 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   191
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnPS 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   192
         Top             =   1920
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSPS 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   194
         Top             =   2280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSPS 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   195
         Top             =   2280
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSP 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   197
         Top             =   2640
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWSP 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   198
         Top             =   2640
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWMFC 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   200
         Top             =   3000
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWMFC 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   201
         Top             =   3000
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDF 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   203
         Top             =   3360
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDF 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   204
         Top             =   3360
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnO 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   206
         Top             =   3720
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnO 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   207
         Top             =   3720
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnSP 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   209
         Top             =   4080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnSP 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   210
         Top             =   4080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDAC 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   212
         Top             =   4440
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWDAC 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   213
         Top             =   4440
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnHT 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   215
         Top             =   4800
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWEnHT 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   216
         Top             =   4800
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWHT 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   218
         Top             =   5160
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWHT 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   219
         Top             =   5160
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWIIR 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   221
         Top             =   5520
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWIIR 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   222
         Top             =   5520
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "IIR Filter K (0-65535)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   80
         Left            =   240
         TabIndex        =   223
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Hard Thr (0-65535)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   79
         Left            =   240
         TabIndex        =   220
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable Hard Thr (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   78
         Left            =   240
         TabIndex        =   217
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "DAC Out Sign (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   77
         Left            =   240
         TabIndex        =   214
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable SENT PN (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   76
         Left            =   240
         TabIndex        =   211
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable OSYS (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   75
         Left            =   240
         TabIndex        =   208
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Digital Filter (0-3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   74
         Left            =   240
         TabIndex        =   205
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Mag. Field Comp (0-3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   73
         Left            =   240
         TabIndex        =   202
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Sense Polarity (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   72
         Left            =   240
         TabIndex        =   199
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Sig. Proc. Seq. (0-3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   71
         Left            =   240
         TabIndex        =   196
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable PWM SR (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   70
         Left            =   240
         TabIndex        =   193
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PWM Polarity (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   69
         Left            =   240
         TabIndex        =   190
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable PWM (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   68
         Left            =   240
         TabIndex        =   187
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Enable SENT (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   67
         Left            =   240
         TabIndex        =   184
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Low Power Mode (0-1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer ID (Read/Write)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   11400
      Width           =   4095
      Begin VB.TextBox CWNumEditCustomerID 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   230
         Text            =   "0"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox CWNumEditCustomerID 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   229
         Text            =   "0"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraMLXID 
      Caption         =   "Melexis ID (Read Only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditX 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   127
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditY 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   2
         Top             =   1080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   127
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditWafer 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   31
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditLot 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   2040
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   131071
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditX 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   127
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditY 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   1080
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   127
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditWafer 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   1560
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   31
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditLot 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   2040
         Width           =   615
         _Version        =   393218
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         opts_1          =   131134
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         Mode_1          =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   131071
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Caption         =   "0123456789"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   82
         Left            =   720
         TabIndex        =   225
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SN: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   81
         Left            =   240
         TabIndex        =   224
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Wafer X (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Wafer Y (0-255)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Wafer (0-31)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Lot # (0-262143)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PTC-04 #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin MSCommLib.MSComm SerialPort 
      Index           =   1
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   115200
   End
   Begin MSCommLib.MSComm SerialPort 
      Index           =   2
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      BaudRate        =   115200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsTimeStreamSettings 
         Caption         =   "&Time Stream Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSaveMLXData 
         Caption         =   "&Save MLX Data to File"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMLX90293"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************** Melexis 90293 Programming Interface *****************
'
'   Scott R Calkins
'   CTS Corporation Automotive Products
'   1142 West Beardsley Avenue
'   Elkhart, Indiana    46514
'   (574) 295-3575
'
'Ver      Date      By   Purpose of modification
'1.0.0  03/19/2019  ANM  First release.
'1.1.0  12/12/2019  ANM  Add 90293 read/write per SCN# 6185.
'

Option Explicit

Private mblnKillTimeDone As Boolean
Private mstrRevisionLevel As String

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

Private Sub cmdClearData_Click()
'
'   PURPOSE:  To zero all parameters on the GUI
'
'  INPUT(S):   None
' OUTPUT(S):   None

Dim lintProgrammerNum As Integer

For lintProgrammerNum = 1 To 2

    'Clamping Levels
    CWNumEditClampLo(lintProgrammerNum).Value = 0
    CWNumEditClampHi(lintProgrammerNum).Value = 0

    'Output Values
    CWVG(lintProgrammerNum).Value = 0
    CWOM(lintProgrammerNum).Value = 0
    CWOS(lintProgrammerNum).Value = 0

    'Locking
    CWMLXLock(lintProgrammerNum).Value = 0
    
    'Melexis ID (Read Only)
    CWNumEditX(lintProgrammerNum).Value = 0
    CWNumEditY(lintProgrammerNum).Value = 0
    CWNumEditWafer(lintProgrammerNum).Value = 0
    CWNumEditLot(lintProgrammerNum).Value = 0
    Label1(82).Caption = "0"
    
    'Customer ID
    CWNumEditCustomerID(lintProgrammerNum).Text = "0"
    
    'Linear Set Pts
    CWY0(lintProgrammerNum).Value = 0
    CWY1(lintProgrammerNum).Value = 0
    CWY2(lintProgrammerNum).Value = 0
    CWY3(lintProgrammerNum).Value = 0
    CWY4(lintProgrammerNum).Value = 0
    CWY5(lintProgrammerNum).Value = 0
    CWY6(lintProgrammerNum).Value = 0
    CWY7(lintProgrammerNum).Value = 0
    CWY8(lintProgrammerNum).Value = 0
    CWY9(lintProgrammerNum).Value = 0
    CWY10(lintProgrammerNum).Value = 0
    CWY11(lintProgrammerNum).Value = 0
    CWY12(lintProgrammerNum).Value = 0
    CWY13(lintProgrammerNum).Value = 0
    CWY14(lintProgrammerNum).Value = 0
    CWY15(lintProgrammerNum).Value = 0
    CWY16(lintProgrammerNum).Value = 0

    'Sensitivity TC
    CWBPivot(lintProgrammerNum).Value = 0
    CWEnSSYS(lintProgrammerNum).Value = 0
    CWSSYS1(lintProgrammerNum).Value = 0
    CWSSYS2(lintProgrammerNum).Value = 0
    CWSSYS3(lintProgrammerNum).Value = 0
    CWSSYS4(lintProgrammerNum).Value = 0
    CWSSYS5(lintProgrammerNum).Value = 0
    CWSSYS6(lintProgrammerNum).Value = 0

    'Other Settings
    CWLPM(lintProgrammerNum).Value = 0
    CWEnS(lintProgrammerNum).Value = 0
    CWEnP(lintProgrammerNum).Value = 0
    CWPPol(lintProgrammerNum).Value = 0
    CWEnPS(lintProgrammerNum).Value = 0
    CWSPS(lintProgrammerNum).Value = 0
    CWSP(lintProgrammerNum).Value = 0
    CWMFC(lintProgrammerNum).Value = 0
    CWDF(lintProgrammerNum).Value = 0
    CWEnO(lintProgrammerNum).Value = 0
    CWEnSP(lintProgrammerNum).Value = 0
    CWDAC(lintProgrammerNum).Value = 0
    CWEnHT(lintProgrammerNum).Value = 0
    CWHT(lintProgrammerNum).Value = 0
    CWIIR(lintProgrammerNum).Value = 0

    'Diagnostics
    CWDiagOM(lintProgrammerNum).Value = 0
    CWDiagOL(lintProgrammerNum).Value = 0
    CWDbSD(lintProgrammerNum).Value = 0
    CWDbCT(lintProgrammerNum).Value = 0
    CWDbCI(lintProgrammerNum).Value = 0

    'Offset TC
    CWOSYS1(lintProgrammerNum).Value = 0
    CWOSYS2(lintProgrammerNum).Value = 0
    CWOSYS3(lintProgrammerNum).Value = 0
    CWOSYS4(lintProgrammerNum).Value = 0
    CWOSYS5(lintProgrammerNum).Value = 0
    CWOSYS6(lintProgrammerNum).Value = 0
    
Next lintProgrammerNum

End Sub

Private Sub cmdReadEEPROM_Click()
'
'   PURPOSE:  To call the read EEPROM function on the click of the button
'
'  INPUT(S):   None
' OUTPUT(S):   None

Dim lblnVotingError As Boolean
Dim lintProgrammerNum As Integer
Dim lintMLXID(6) As Integer
Dim lintCustID(3) As Integer
Dim X As Integer
Dim PSFMan As PSF090293AAMLXManager
Dim DevicesCol As ObjectCollection
Dim i As Long
Dim SN As String
Dim d As Double

'Indicate that we are reading from EEprom
cwbtnStatus.Value = True

'Disable buttons on the form
cmdReadEEPROM.Enabled = False
cmdWriteEEPROM.Enabled = False
cmdClearData.Enabled = False

'Check if communication is active
If gblnGoodPTC04Link Then

    Call MyDev(lintDev1).DeviceReplaced
    Call MyDev(lintDev1).ReadFullDevice
    
    Call MyDev(lintDev2).DeviceReplaced
    Call MyDev(lintDev2).ReadFullDevice
            
    lintMLXID(0) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID0)
    lintMLXID(1) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID1)
    lintMLXID(2) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID2)
    lintMLXID(3) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID3)
    lintMLXID(4) = MyDev(lintDev1).GetEEParameterCode(CodeMLXID4)

    gudtMLX90277(1).Read.X = (Int(lintMLXID(0) / BIT5) And &H7) + ((lintMLXID(1) And &H1F) * BIT3)
    gudtMLX90277(1).Read.Y = (Int(lintMLXID(1) / BIT5) And &H7) + ((lintMLXID(2) And &H1F) * BIT3)
    gudtMLX90277(1).Read.Wafer = (lintMLXID(0) And &H1F)
    d = CDbl(lintMLXID(4)) * BIT11
    gudtMLX90277(1).Read.Lot = (Int(lintMLXID(2) / BIT5) And &H7) + (lintMLXID(3) * BIT3) + d

    SN = MLX90293.EncodePartID
    Label1(82).Caption = SN
    
    lintMLXID(0) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID0)
    lintMLXID(1) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID1)
    lintMLXID(2) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID2)
    lintMLXID(3) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID3)
    lintMLXID(4) = MyDev(lintDev2).GetEEParameterCode(CodeMLXID4)

    gudtMLX90277(2).Read.X = (Int(lintMLXID(0) / BIT5) And &H7) + ((lintMLXID(1) And &H1F) * BIT3)
    gudtMLX90277(2).Read.Y = (Int(lintMLXID(1) / BIT5) And &H7) + ((lintMLXID(2) And &H1F) * BIT3)
    gudtMLX90277(2).Read.Wafer = (lintMLXID(0) And &H1F)
    d = CDbl(lintMLXID(4)) * BIT11
    gudtMLX90277(2).Read.Lot = (Int(lintMLXID(2) / BIT5) And &H7) + (lintMLXID(3) * BIT3) + d
    
    'Clamping Levels
    CWNumEditClampLo(1).Value = MyDev(lintDev1).GetEEParameterValue(CodeCLAMPHIGH)
    CWNumEditClampHi(1).Value = MyDev(lintDev1).GetEEParameterValue(CodeCLAMPLOW)
    
    CWNumEditClampLo(2).Value = MyDev(lintDev2).GetEEParameterValue(CodeCLAMPHIGH)
    CWNumEditClampHi(2).Value = MyDev(lintDev2).GetEEParameterValue(CodeCLAMPLOW)

    'Output Values
    CWVG(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeVG)
    CWOM(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeOSNORMMODE)
    CWOS(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeOUTPUTSCALING)

    CWVG(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeVG)
    CWOM(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeOSNORMMODE)
    CWOS(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeOUTPUTSCALING)
    
    'Locking
    CWMLXLock(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeMEMLOCK)
    CWMLXLock(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeMEMLOCK)
    
    'Melexis ID (Read Only)
    CWNumEditX(1).Value = gudtMLX90277(1).Read.X
    CWNumEditY(1).Value = gudtMLX90277(1).Read.Y
    CWNumEditWafer(1).Value = gudtMLX90277(1).Read.Wafer
    CWNumEditLot(1).Value = gudtMLX90277(1).Read.Lot
    
    CWNumEditX(2).Value = gudtMLX90277(2).Read.X
    CWNumEditY(2).Value = gudtMLX90277(2).Read.Y
    CWNumEditWafer(2).Value = gudtMLX90277(2).Read.Wafer
    CWNumEditLot(2).Value = gudtMLX90277(2).Read.Lot
    
    'Customer ID
    glngCUSTID1 = MyDev(lintDev2).GetEEParameterCode(CodeUSERID1)
    glngCUSTID2 = MyDev(lintDev2).GetEEParameterCode(CodeUSERID2)
    CWNumEditCustomerID(2).Text = MLX90293.DecodeCustomerID90293
    
    glngCUSTID1 = MyDev(lintDev1).GetEEParameterCode(CodeUSERID1)
    glngCUSTID2 = MyDev(lintDev1).GetEEParameterCode(CodeUSERID2)
    CWNumEditCustomerID(1).Text = MLX90293.DecodeCustomerID90293
    
    'Linear Set Pts
    CWY0(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY0)
    CWY1(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY1)
    CWY2(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY2)
    CWY3(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY3)
    CWY4(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY4)
    CWY5(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY5)
    CWY6(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY6)
    CWY7(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY7)
    CWY8(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY8)
    CWY9(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY9)
    CWY10(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY10)
    CWY11(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY11)
    CWY12(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY12)
    CWY13(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY13)
    CWY14(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY14)
    CWY15(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY15)
    CWY16(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLNRY16)

    CWY0(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY0)
    CWY1(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY1)
    CWY2(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY2)
    CWY3(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY3)
    CWY4(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY4)
    CWY5(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY5)
    CWY6(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY6)
    CWY7(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY7)
    CWY8(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY8)
    CWY9(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY9)
    CWY10(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY10)
    CWY11(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY11)
    CWY12(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY12)
    CWY13(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY13)
    CWY14(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY14)
    CWY15(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY15)
    CWY16(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLNRY16)
    
    'Sensitivity TC
    CWBPivot(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeBPIVOT)
    CWEnSSYS(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS2XSPAN)
    CWSSYS1(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS1)
    CWSSYS2(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS2)
    CWSSYS3(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS3)
    CWSSYS4(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS5)
    CWSSYS5(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS6)
    CWSSYS6(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS7)
    
    CWBPivot(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeBPIVOT)
    CWEnSSYS(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS2XSPAN)
    CWSSYS1(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS1)
    CWSSYS2(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS2)
    CWSSYS3(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS3)
    CWSSYS4(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS5)
    CWSSYS5(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS6)
    CWSSYS6(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS7)

    'Other Settings
    CWLPM(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeLPMODE)
    CWEnS(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSENT)
    CWEnP(1).Value = MyDev(lintDev1).GetEEParameterCode(CodePWM)
    CWPPol(1).Value = MyDev(lintDev1).GetEEParameterCode(CodePWMPOLARITY)
    CWEnPS(1).Value = MyDev(lintDev1).GetEEParameterCode(CodePWMSLEWRATE)
    CWSPS(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeTREATSEQ)
    CWSP(1).Value = MyDev(lintDev1).GetEEParameterCode(CodePOLARITY)
    CWMFC(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeMAPXYZ)
    CWDF(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeFILTER)
    CWEnO(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSELECT_SENT_OSYS)
    CWEnSP(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSENT_PAUSE_NIBBLEOUT)
    CWDAC(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeDAC_OUTPUT_SIGN)
    CWEnHT(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeENABLEHARDTHRESHOLD)
    CWHT(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeHARDTHRESHOLD)
    CWIIR(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeIIR_K)

    CWLPM(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeLPMODE)
    CWEnS(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSENT)
    CWEnP(2).Value = MyDev(lintDev2).GetEEParameterCode(CodePWM)
    CWPPol(2).Value = MyDev(lintDev2).GetEEParameterCode(CodePWMPOLARITY)
    CWEnPS(2).Value = MyDev(lintDev2).GetEEParameterCode(CodePWMSLEWRATE)
    CWSPS(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeTREATSEQ)
    CWSP(2).Value = MyDev(lintDev2).GetEEParameterCode(CodePOLARITY)
    CWMFC(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeMAPXYZ)
    CWDF(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeFILTER)
    CWEnO(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSELECT_SENT_OSYS)
    CWEnSP(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSENT_PAUSE_NIBBLEOUT)
    CWDAC(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeDAC_OUTPUT_SIGN)
    CWEnHT(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeENABLEHARDTHRESHOLD)
    CWHT(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeHARDTHRESHOLD)
    CWIIR(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeIIR_K)
    
    'Diagnostics
    CWDiagOM(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeOSDIAGMODE)
    CWDiagOL(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeOSDIAGDIAG)
    CWDbSD(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeDIAGDEBOUNCESTEPDOWN)
    CWDbCT(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeDIAGDEBOUNCETHRESH)
    CWDbCI(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeDIAGDEBOUNCESTEPUP)

    CWDiagOM(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeOSDIAGMODE)
    CWDiagOL(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeOSDIAGDIAG)
    CWDbSD(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeDIAGDEBOUNCESTEPDOWN)
    CWDbCT(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeDIAGDEBOUNCETHRESH)
    CWDbCI(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeDIAGDEBOUNCESTEPUP)
    
    'Offset TC
    CWOSYS1(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS1)
    CWOSYS2(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS2)
    CWOSYS3(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS3)
    CWOSYS4(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS5)
    CWOSYS5(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS6)
    CWOSYS6(1).Value = MyDev(lintDev1).GetEEParameterCode(CodeSSYS7)
    
    CWOSYS1(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS1)
    CWOSYS2(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS2)
    CWOSYS3(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS3)
    CWOSYS4(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS5)
    CWOSYS5(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS6)
    CWOSYS6(2).Value = MyDev(lintDev2).GetEEParameterCode(CodeSSYS7)
    
Else
    'Let the user know if communication is not active
    MsgBox "Communication Not Established!  Reset Communication!", vbOKOnly & vbCritical, "Programmer Communication Not Active"
End If

'Indicate that we are done reading from EEprom
cwbtnStatus.Value = False

'Enable buttons on the form
cmdReadEEPROM.Enabled = True
cmdWriteEEPROM.Enabled = True
cmdClearData.Enabled = True

End Sub

Private Sub cmdResetCommunication_Click()
'
'   PURPOSE:   Resets communication to the Melexis programmer selected by Index
'
'  INPUT(S):   None
' OUTPUT(S):   None

Dim lintProgrammerNum As Integer

On Error GoTo BadComboBox

''Assign Comm Port Numbers
'For lintProgrammerNum = 1 To 2
'    gudtPTC04(lintProgrammerNum).CommPortNum = CInt((cboComPortNum(lintProgrammerNum).Text))
'Next lintProgrammerNum

'Establish Communication with the programmer
Call MLX90293.EstablishCommunication

Exit Sub
'Error Trap
BadComboBox:

    MsgBox "Invalid Com Port Number. " _
           & vbCrLf & "Try again.", vbOKOnly + vbCritical, "Error"


End Sub

Private Sub cmdWriteEEPROM_Click()
'
'   PURPOSE:   Writes values from numEdit boxes to EEPROM
'
'  INPUT(S):   None
' OUTPUT(S):   None

Dim lintZeroCnt As Integer
Dim lintResponse As Integer
Dim lblnVotingError As Boolean
Dim lintProgrammerNum As Integer

'Indicate that we are writing to the EEprom
cwbtnStatus.Value = True

'Disable buttons on the form
cmdReadEEPROM.Enabled = False
cmdWriteEEPROM.Enabled = False
cmdClearData.Enabled = False

'Check if Communication is active
If gblnGoodPTC04Link Then

   Call MyDev(lintDev1).SetEEParameterCode(CodeVG, CWVG(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSNORMMODE, CWOM(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOUTPUTSCALING, CWOS(1).Value)
   Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPHIGH, CWNumEditClampLo(1).Value)
   Call MyDev(lintDev1).SetEEParameterValue(CodeCLAMPLOW, CWNumEditClampHi(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeFILTER, CWDF(1).Value)
   'Call MyDev(lintDev1).SetEEParameterCode(CodeMEMLOCK, CWMLXLock(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeENABLEHARDTHRESHOLD, CWEnHT(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeHARDTHRESHOLD, CWHT(1).Value)
   
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingBpivot, CWBPivot(1).Value)
   Call MyDev(lintDev1).Advanced.SetSolverSetting(SolverSettingTreatSeq, CWSPS(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS1, CWOSYS1(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS2, CWOSYS2(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS3, CWOSYS3(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS5, CWOSYS4(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS6, CWOSYS5(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSYS7, CWOSYS6(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS1, CWSSYS1(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS2, CWSSYS2(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS3, CWSSYS3(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS5, CWSSYS4(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS6, CWSSYS5(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS7, CWSSYS6(1).Value)

   Call MLX90293.EncodeCustomerID90293(CWNumEditCustomerID(1).Text)
   Call MyDev(lintDev1).SetEEParameterCode(CodeUSERID1, glngCUSTID1)
   Call MyDev(lintDev1).SetEEParameterCode(CodeUSERID2, glngCUSTID2)
    
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY0, CWY0(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY1, CWY1(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY2, CWY2(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY3, CWY3(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY4, CWY4(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY5, CWY5(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY6, CWY6(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY7, CWY7(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY8, CWY8(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY9, CWY9(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY10, CWY10(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY11, CWY11(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY12, CWY12(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY13, CWY13(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY14, CWY14(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY15, CWY15(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLNRY16, CWY16(1).Value)
    
   Call MyDev(lintDev1).SetEEParameterCode(CodeSSYS2XSPAN, CWEnSSYS(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeLPMODE, CWLPM(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSENT, CWEnS(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodePWM, CWEnP(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodePWMPOLARITY, CWPPol(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodePWMSLEWRATE, CWEnPS(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodePOLARITY, CWSP(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeMAPXYZ, CWMFC(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSELECT_SENT_OSYS, CWEnO(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeSENT_PAUSE_NIBBLEOUT, CWEnSP(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeDAC_OUTPUT_SIGN, CWDAC(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeIIR_K, CWIIR(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSDIAGMODE, CWDiagOM(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeOSDIAGDIAG, CWDiagOL(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeDIAGDEBOUNCESTEPDOWN, CWDbSD(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeDIAGDEBOUNCETHRESH, CWDbCT(1).Value)
   Call MyDev(lintDev1).SetEEParameterCode(CodeDIAGDEBOUNCESTEPUP, CWDbCI(1).Value)
   
   Call MyDev(lintDev1).ProgramDevice
   
   'Dev 2
   Call MyDev(lintDev2).SetEEParameterCode(CodeVG, CWVG(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSNORMMODE, CWOM(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOUTPUTSCALING, CWOS(2).Value)
   Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPHIGH, CWNumEditClampLo(2).Value)
   Call MyDev(lintDev2).SetEEParameterValue(CodeCLAMPLOW, CWNumEditClampHi(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeFILTER, CWDF(2).Value)
   'Call MyDev(lintDev2).SetEEParameterCode(CodeMEMLOCK, CWMLXLock(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeENABLEHARDTHRESHOLD, CWEnHT(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeHARDTHRESHOLD, CWHT(2).Value)
   
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingBpivot, CWBPivot(2).Value)
   Call MyDev(lintDev2).Advanced.SetSolverSetting(SolverSettingTreatSeq, CWSPS(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS1, CWOSYS1(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS2, CWOSYS2(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS3, CWOSYS3(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS5, CWOSYS4(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS6, CWOSYS5(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSYS7, CWOSYS6(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS1, CWSSYS1(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS2, CWSSYS2(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS3, CWSSYS3(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS5, CWSSYS4(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS6, CWSSYS5(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS7, CWSSYS6(2).Value)

   'Call MLX90293.EncodeCustomerID90293(CWNumEditCustomerID(2).Text)
   'Call MyDev(lintDev2).SetEEParameterCode(CodeUSERID1, glngCUSTID1)
   'Call MyDev(lintDev2).SetEEParameterCode(CodeUSERID2, glngCUSTID2)
    
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY0, CWY0(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY1, CWY1(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY2, CWY2(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY3, CWY3(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY4, CWY4(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY5, CWY5(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY6, CWY6(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY7, CWY7(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY8, CWY8(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY9, CWY9(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY10, CWY10(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY11, CWY11(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY12, CWY12(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY13, CWY13(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY14, CWY14(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY15, CWY15(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLNRY16, CWY16(2).Value)
    
   Call MyDev(lintDev2).SetEEParameterCode(CodeSSYS2XSPAN, CWEnSSYS(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeLPMODE, CWLPM(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSENT, CWEnS(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodePWM, CWEnP(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodePWMPOLARITY, CWPPol(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodePWMSLEWRATE, CWEnPS(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodePOLARITY, CWSP(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeMAPXYZ, CWMFC(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSELECT_SENT_OSYS, CWEnO(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeSENT_PAUSE_NIBBLEOUT, CWEnSP(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeDAC_OUTPUT_SIGN, CWDAC(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeIIR_K, CWIIR(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSDIAGMODE, CWDiagOM(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeOSDIAGDIAG, CWDiagOL(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeDIAGDEBOUNCESTEPDOWN, CWDbSD(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeDIAGDEBOUNCETHRESH, CWDbCT(2).Value)
   Call MyDev(lintDev2).SetEEParameterCode(CodeDIAGDEBOUNCESTEPUP, CWDbCI(2).Value)
   
   Call MyDev(lintDev2).ProgramDevice
Else
    'Let the user know if communication is not active
    MsgBox "Communication Not Established!  Reset Communication!", vbOKOnly & vbCritical, "Programmer Communication Not Active"
End If

'Indicate that we are done writing to the EEPROM
cwbtnStatus.Value = False

'Enable buttons on the form
cmdReadEEPROM.Enabled = True
cmdWriteEEPROM.Enabled = True
cmdClearData.Enabled = True

End Sub

Private Sub Form_Load()
'
'   PURPOSE:   Executes when form is loaded, initializes form
'
'  INPUT(S):   None
' OUTPUT(S):   None

Me.top = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE:   Executes on Form_Unload Event
'
'  INPUT(S):   None
' OUTPUT(S):   None

tmrMLX.Enabled = False

End Sub

Private Sub mnuFileExit_Click()
'
'   PURPOSE:  Makes the form invisible
'
'  INPUT(S):   None
' OUTPUT(S):   None

Visible = False

End Sub

Private Sub mnuMLX90277RevisionCx_Click()   'V1.2
'
'   PURPOSE:   Selects the revision level
'
'  INPUT(S):   None
' OUTPUT(S):   None

'mnuMLX90277RevisionCx.Checked = True
'mnuMLX90277RevisionFA.Checked = False
'
''Cx Level Chip
'gstrMLX90277Revision = "Cx"

End Sub

Private Sub mnuMLX90277RevisionFA_Click()   'V1.2
'
'   PURPOSE:   Selects the revision level
'
'  INPUT(S):   None
' OUTPUT(S):   None

'mnuMLX90277RevisionCx.Checked = False
'mnuMLX90277RevisionFA.Checked = True
''FA Level Chip
'gstrMLX90277Revision = "FA"

End Sub

Private Sub mnuToolsSaveMLXData_Click()
'
'   PURPOSE:   Saves MLX data to file.
'
'  INPUT(S):   None
' OUTPUT(S):   None

'Call Solver90277.SaveMLXtoFile

End Sub

Private Sub mnuToolsTimeStreamSettings_Click()
'
'   PURPOSE:   Displays Time Stream Settings Form
'
'  INPUT(S):   None
' OUTPUT(S):   None

'frmTimeSettings.Show vbModal

End Sub

Private Sub tmrKillTime_Timer()
'
'   PURPOSE:   Event triggered when when timer, tmrKillTime, is complete.
'
'  INPUT(S):   None
' OUTPUT(S):   None

mblnKillTimeDone = True

End Sub

Private Sub tmrMLX_Timer()
'
'   PURPOSE:   Update the value of the CommunicationActive Display LED
'
'  INPUT(S):   None
' OUTPUT(S):   None


cwbtnActive.Value = gblnGoodPTC04Link

End Sub
