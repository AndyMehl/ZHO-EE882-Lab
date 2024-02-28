VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMLX90277 
   Caption         =   "Melexis 90277 Test Utility"
   ClientHeight    =   10440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraVotingError 
      Caption         =   "TC Table (Read Only)"
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
      Left            =   9240
      TabIndex        =   119
      Top             =   1920
      Width           =   4095
      Begin VB.CheckBox chkTCTable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   121
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkTCTable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   120
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "TC Table"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   124
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
         Index           =   42
         Left            =   2040
         TabIndex        =   123
         Top             =   240
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
         Index           =   43
         Left            =   3000
         TabIndex        =   122
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
      Height          =   2535
      Left            =   9240
      TabIndex        =   58
      Top             =   7440
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditX 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   59
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
         TabIndex        =   61
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
         TabIndex        =   63
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
         TabIndex        =   65
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
         TabIndex        =   115
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
         TabIndex        =   116
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
         TabIndex        =   117
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
         TabIndex        =   118
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
         TabIndex        =   114
         Top             =   240
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
         Index           =   49
         Left            =   3000
         TabIndex        =   113
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Lot # (0-131071)"
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
         TabIndex        =   66
         Top             =   2040
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
         TabIndex        =   64
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Wafer Y (0-127)"
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
         TabIndex        =   62
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Wafer X (0-127)"
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
         TabIndex        =   60
         Top             =   600
         Width           =   1815
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   47
      Top             =   7440
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditCustomerID 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   48
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
         RangeMax_Val_1  =   16777215
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditCustomerID 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   102
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
         RangeMax_Val_1  =   16777215
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
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
         Left            =   2040
         TabIndex        =   101
         Top             =   240
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
         Index           =   39
         Left            =   3000
         TabIndex        =   100
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ID  (0-16777215)"
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
         TabIndex        =   49
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Output Driver (Read/Write)"
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
      Height          =   1095
      Left            =   9240
      TabIndex        =   44
      Top             =   3240
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditMode 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   45
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditMode 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   99
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
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
         TabIndex        =   98
         Top             =   240
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
         Index           =   45
         Left            =   3000
         TabIndex        =   97
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mode (0-3)"
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
         TabIndex        =   46
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraFiltering 
      Caption         =   "Filtering (Read/Write)"
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   41
      Top             =   6000
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditFilter 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   42
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditFilter 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   240
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
         Index           =   37
         Left            =   3000
         TabIndex        =   94
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filter (0-15)"
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
         TabIndex        =   43
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraTimeGenerator 
      Caption         =   "Time Generator (Read/Write)"
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
      Height          =   2415
      Left            =   9240
      TabIndex        =   40
      Top             =   4680
      Width           =   4095
      Begin VB.CheckBox chkSlowMode 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   112
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkSlowMode 
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   57
         Top             =   1920
         Width           =   255
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditOscillatorAdjust 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   50
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
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditCKDACCH 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   52
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditCKANACH 
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditOscillatorAdjust 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   109
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
         RangeMax_Val_1  =   15
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditCKDACCH 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   110
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditCKANACH 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   111
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
         RangeMax_Val_1  =   3
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
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
         TabIndex        =   108
         Top             =   240
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
         Index           =   47
         Left            =   3000
         TabIndex        =   107
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Slow (Y/N)"
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
         TabIndex        =   56
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cap. Freq. Adj. (0-3)"
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
         TabIndex        =   55
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "DAC Freq. Adj. (0-3)"
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
         TabIndex        =   53
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Oscillator Adj. (0-15)"
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
         TabIndex        =   51
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraSensitivity 
      Caption         =   "Sensitivity (Read/Write)"
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
      Left            =   360
      TabIndex        =   33
      Top             =   7920
      Width           =   4095
      Begin VB.CheckBox chkInvertSlope 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   85
         Top             =   1500
         Width           =   255
      End
      Begin VB.CheckBox chkInvertSlope 
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   39
         Top             =   1500
         Width           =   255
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditRGain 
         Height          =   315
         Index           =   2
         Left            =   3240
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditFGain 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   37
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditRGain 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   83
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditFGain 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   84
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
         TabIndex        =   82
         Top             =   360
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
         Index           =   30
         Left            =   2040
         TabIndex        =   81
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Invert Slope (Y/N)"
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
         TabIndex        =   38
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Fine Gain (0-1023)"
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
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Rough Gain (0-15)"
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
         TabIndex        =   34
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraOffset 
      Caption         =   "Offset Voltage (Read/Write)"
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
      Left            =   360
      TabIndex        =   26
      Top             =   5520
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditOffset 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   27
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditAGND 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   30
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditDrift 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   32
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditOffset 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   78
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditAGND 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   79
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditDrift 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   80
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
         TabIndex        =   77
         Top             =   360
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
         Index           =   28
         Left            =   2040
         TabIndex        =   76
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Offset Drift (0-15)"
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
         TabIndex        =   31
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "AGND (0-1023)"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Offset (0-1023)"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame fraEEPROMFault 
      Caption         =   "Fault Level (Read/Write)"
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   24
      Top             =   8880
      Width           =   4095
      Begin VB.CheckBox chkFaultLevel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   86
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkFaultLevel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   25
         Top             =   600
         Width           =   255
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
         Index           =   40
         Left            =   2040
         TabIndex        =   88
         Top             =   240
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
         Index           =   41
         Left            =   3000
         TabIndex        =   87
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fault Level (Y/N)"
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
         TabIndex        =   67
         Top             =   600
         Width           =   1815
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
      Height          =   1335
      Left            =   4800
      TabIndex        =   21
      Top             =   1920
      Width           =   4095
      Begin VB.CheckBox chkMemLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   106
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkMLXLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   105
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkMemLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkMLXLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   22
         Top             =   600
         Width           =   255
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
         TabIndex        =   104
         Top             =   240
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
         Index           =   33
         Left            =   3000
         TabIndex        =   103
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Memory Lock (Y/N)"
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
         TabIndex        =   69
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Melexis Lock (Y/N)"
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
         TabIndex        =   68
         Top             =   600
         Width           =   1815
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
      Left            =   4800
      TabIndex        =   16
      Top             =   240
      Width           =   8535
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
         Left            =   4560
         TabIndex        =   19
         Top             =   720
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
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   3615
      End
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
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
      Begin CWUIControlsLib.CWButton cwbtnStatus 
         Height          =   375
         Left            =   480
         TabIndex        =   20
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
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   4095
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
         TabIndex        =   71
         Top             =   2160
         Width           =   3615
      End
      Begin VB.ComboBox cboComPortNum 
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   70
         Text            =   "3"
         Top             =   480
         Width           =   855
      End
      Begin CWUIControlsLib.CWButton cwbtnActive 
         Height          =   375
         Left            =   240
         TabIndex        =   15
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
      Begin VB.ComboBox cboComPortNum 
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   13
         Text            =   "4"
         Top             =   1080
         Width           =   855
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
         TabIndex        =   126
         Top             =   1680
         Width           =   3135
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
         TabIndex        =   125
         Top             =   1080
         Width           =   2655
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
         TabIndex        =   14
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Temperature Compensation (Read/Write)"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   3600
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditTC 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   9
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
         RangeMax_Val_1  =   31
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditTC2nd 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   10
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
         RangeMax_Val_1  =   63
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditTCWin 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   11
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
         RangeMax_Val_1  =   7
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditTCWin 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   89
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
         RangeMax_Val_1  =   7
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditTC 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   92
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
         RangeMax_Val_1  =   31
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit CWNumEditTC2nd 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   93
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
         RangeMax_Val_1  =   63
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
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
         TabIndex        =   91
         Top             =   240
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
         Index           =   35
         Left            =   3000
         TabIndex        =   90
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "TC (0-31)"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "TC2nd (0-63)"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "TCWin (0-7)"
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
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
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   4095
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampLo 
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
         Left            =   2280
         TabIndex        =   72
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
      Begin CWUIControlsLib.CWNumEdit CWNumEditClampHi 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   360
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
         Index           =   26
         Left            =   2040
         TabIndex        =   73
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Clamp High (0-1023)"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Clamp Low (0-1023)"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
   End
   Begin MSCommLib.MSComm SerialPort 
      Index           =   1
      Left            =   1080
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   115200
   End
   Begin VB.Timer tmrMLX 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrKillTime 
      Enabled         =   0   'False
      Left            =   480
      Top             =   0
   End
   Begin MSCommLib.MSComm SerialPort 
      Index           =   2
      Left            =   1680
      Top             =   -120
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
      End
      Begin VB.Menu mnuToolsSaveMLXData 
         Caption         =   "&Save MLX Data to File"
      End
   End
   Begin VB.Menu mnuMLX90277Revision 
      Caption         =   "MLX90277 Revision Level"
      Begin VB.Menu mnuMLX90277RevisionCx 
         Caption         =   "Cx"
      End
      Begin VB.Menu mnuMLX90277RevisionFA 
         Caption         =   "FA"
      End
   End
End
Attribute VB_Name = "frmMLX90277"
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
'1.0.0  04/19/2004  SRC  First release per PR12722-A.
'1.1.0  06/30/2004  SRC  Re-released for use with PTC-04 instead of PTC-03
'                        per PR12722-B.  Re-wrote interface routines, added
'                        necessary decode/encode routines, and coded for
'                        simultaneous communication with two programmers.
'1.1.1  09/25/2004  SRC  Updated form for more accurate representation of
'                        Read/Write parameters.  Added current limit for
'                        PTC-04 Programmers.
'1.1.2  01/13/2005  SRC  Minor comment changes in SendCommandGetResponse.
'1.2.0  03/05/2005  SRC  Updated ReadEEPROM and associated variable structure
'                        for new MLX firmware.  Additions add capability to
'                        communicate with MLX90277FA parts, while retaining
'                        communication capabilities with MLX90277Cx parts.
'                        Updated form to include option to communicate with
'                        either chip.  Changes noted as 'V1.2
'1.3.0  05/01/2007  ANM  Added output 2 SN sub 'V1.3
'1.4.0  08/14/2008  ANM  Added GetCurrent '3.3ANM
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

    'EEPROM Fault Level
    chkFaultLevel(lintProgrammerNum).Value = 0

    'Locks
    chkMLXLock(lintProgrammerNum).Value = 0
    chkMemLock(lintProgrammerNum).Value = 0

    'TC Table Check
    chkTCTable(lintProgrammerNum).Value = 0

    'Clamping Levels
    CWNumEditClampLo(lintProgrammerNum).Value = 0
    CWNumEditClampHi(lintProgrammerNum).Value = 0

    'Offset
    CWNumEditOffset(lintProgrammerNum).Value = 0
    CWNumEditAGND(lintProgrammerNum).Value = 0
    CWNumEditDrift(lintProgrammerNum).Value = 0

    'Sensitivity
    CWNumEditRGain(lintProgrammerNum).Value = 0
    CWNumEditFGain(lintProgrammerNum).Value = 0
    chkInvertSlope(lintProgrammerNum).Value = 0

    'TC
    CWNumEditTCWin(lintProgrammerNum).Value = 0
    CWNumEditTC(lintProgrammerNum).Value = 0
    CWNumEditTC2nd(lintProgrammerNum).Value = 0

    'Filtering
    CWNumEditFilter(lintProgrammerNum).Value = 0

    'Output Driver
    CWNumEditMode(lintProgrammerNum).Value = 0

    'Customer ID
    CWNumEditCustomerID(lintProgrammerNum).Value = 0

    'Time Generator (Read Only)
    CWNumEditOscillatorAdjust(lintProgrammerNum).Value = 0
    CWNumEditCKDACCH(lintProgrammerNum).Value = 0
    CWNumEditCKANACH(lintProgrammerNum).Value = 0
    chkSlowMode(lintProgrammerNum).Value = 0

    'Melexis ID (Read Only)
    CWNumEditX(lintProgrammerNum).Value = 0
    CWNumEditY(lintProgrammerNum).Value = 0
    CWNumEditWafer(lintProgrammerNum).Value = 0
    CWNumEditLot(lintProgrammerNum).Value = 0

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

'Indicate that we are reading from EEprom
cwbtnStatus.Value = True

'Disable buttons on the form
cmdReadEEPROM.Enabled = False
cmdWriteEEPROM.Enabled = False
cmdClearData.Enabled = False


'Check if communication is active
If gblnGoodPTC04Link Then

    '*** Read values back from EEprom ***
    Call MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError)

    'If there is a voting error, display a message box
    If lblnVotingError Then MsgBox "EEPROM Voting Failure!  MLX90277 IC Problem!", vbOKOnly & vbCritical, "EEPROM Failure!"

    'Loop through both programmers
    For lintProgrammerNum = 1 To 2

        '*** Make sure that the read variables get transferred into the write variables ***
        Call MLX90277.CopyMLXReadsToMLXWrites(lintProgrammerNum)

        '*** Update Display Variables ***

        'EEPROM Fault Level
        If gudtMLX90277(lintProgrammerNum).Read.FaultLevel Then
            chkFaultLevel(lintProgrammerNum).Value = 1
        Else
            chkFaultLevel(lintProgrammerNum).Value = 0
        End If

        'Melexis Lock
        If gudtMLX90277(lintProgrammerNum).Read.MelexisLock Then
            chkMLXLock(lintProgrammerNum).Value = 1
        Else
            chkMLXLock(lintProgrammerNum).Value = 0
        End If
        'Memory Lock
        If gudtMLX90277(lintProgrammerNum).Read.MemoryLock Then
            chkMemLock(lintProgrammerNum).Value = 1
        Else
            chkMemLock(lintProgrammerNum).Value = 0
        End If

        'Clamping Levels
        CWNumEditClampLo(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.clampLow
        CWNumEditClampHi(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.clampHigh

        'Offset
        CWNumEditOffset(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.offset
        CWNumEditAGND(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.AGND
        CWNumEditDrift(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Drift

        'Sensitivity
        CWNumEditRGain(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.RGain
        CWNumEditFGain(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.FGain
        If gudtMLX90277(lintProgrammerNum).Read.InvertSlope Then
            chkInvertSlope(lintProgrammerNum).Value = 1
        Else
            chkInvertSlope(lintProgrammerNum).Value = 0
        End If

        'TC
        CWNumEditTCWin(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.TCWin
        CWNumEditTC(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.TC
        CWNumEditTC2nd(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.TC2nd

        'Filtering
        CWNumEditFilter(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Filter

        'Customer ID
        CWNumEditCustomerID(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.CustID

        'Output Driver (Read Only)
        CWNumEditMode(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Mode

        'Time Generator (Read Only)
        CWNumEditOscillatorAdjust(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.FCKADJ
        CWNumEditCKDACCH(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.CKDACCH
        CWNumEditCKANACH(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.CKANACH
        If gudtMLX90277(lintProgrammerNum).Read.SlowMode Then
            chkSlowMode(lintProgrammerNum).Value = 1
        Else
            chkSlowMode(lintProgrammerNum).Value = 0
        End If

        'Melexis ID (Read Only)
        CWNumEditX(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.X
        CWNumEditY(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Y
        CWNumEditWafer(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Wafer
        CWNumEditLot(lintProgrammerNum).Value = gudtMLX90277(lintProgrammerNum).Read.Lot

        'TC Table
        If (MLX90277.VerifyMLXCRC(lintProgrammerNum) And MLX90277.VerifyCustomerCRC(lintProgrammerNum)) Then
            chkTCTable(lintProgrammerNum).Value = 1
        Else
            chkTCTable(lintProgrammerNum).Value = 0
        End If
    Next lintProgrammerNum
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

'Assign Comm Port Numbers
For lintProgrammerNum = 1 To 2
    gudtPTC04(lintProgrammerNum).CommPortNum = CInt((cboComPortNum(lintProgrammerNum).Text))
Next lintProgrammerNum

'Establish Communication with the programmer
Call MLX90277.EstablishCommunication

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

    'Loop through both programmers
    For lintProgrammerNum = 1 To 2
        'Make sure we are not going to write zeros to all locations
        If CWNumEditClampLo(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditClampHi(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditOffset(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditAGND(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditDrift(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditFGain(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditTCWin(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditTC(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If
        If CWNumEditTC2nd(lintProgrammerNum).Value = 0 Then
            lintZeroCnt = lintZeroCnt + 1
        End If

        'Display message to prevent clearing EEprom
        If lintZeroCnt > 2 Then

            lintResponse = MsgBox("More than one EEPROM location will be cleared.  Are you sure you want to do this?" _
                                  & vbCrLf & vbCrLf & "Click <OK> to Continue or <Cancel> to Abort Write to EEPROM.", _
                                  vbOKCancel + vbExclamation, "Abort Clearing EEPROM Locations")

            If lintResponse = vbCancel Then
                'Indicate that we are done writing to EEprom
                cwbtnStatus.Value = False
                Exit For
            End If
        End If

        '*** Transfer Display Contents to EEPROM Variables ***

        'Fault Level
        If chkFaultLevel(lintProgrammerNum).Value = 1 Then
            gudtMLX90277(lintProgrammerNum).Write.FaultLevel = True
        Else
            gudtMLX90277(lintProgrammerNum).Write.FaultLevel = False
        End If

        'Melexis Lock
        If chkMLXLock(lintProgrammerNum).Value = 1 Then
            gudtMLX90277(lintProgrammerNum).Write.MelexisLock = True
        Else
            gudtMLX90277(lintProgrammerNum).Write.MelexisLock = False
        End If

        'Memory Lock
        If chkMemLock(lintProgrammerNum).Value = 1 Then
            gudtMLX90277(lintProgrammerNum).Write.MemoryLock = True
        Else
            gudtMLX90277(lintProgrammerNum).Write.MemoryLock = False
        End If

        'Clamping Levels
        gudtMLX90277(lintProgrammerNum).Write.clampLow = CWNumEditClampLo(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.clampHigh = CWNumEditClampHi(lintProgrammerNum).Value

        'Offset
        gudtMLX90277(lintProgrammerNum).Write.offset = CWNumEditOffset(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.AGND = CWNumEditAGND(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.Drift = CWNumEditDrift(lintProgrammerNum).Value

        'Sensitivity
        gudtMLX90277(lintProgrammerNum).Write.RGain = CWNumEditRGain(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.FGain = CWNumEditFGain(lintProgrammerNum).Value
        If chkInvertSlope(lintProgrammerNum).Value = 1 Then
            gudtMLX90277(lintProgrammerNum).Write.InvertSlope = True
        Else
            gudtMLX90277(lintProgrammerNum).Write.InvertSlope = False
        End If

        'TC
        gudtMLX90277(lintProgrammerNum).Write.TCWin = CWNumEditTCWin(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.TC = CWNumEditTC(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.TC2nd = CWNumEditTC2nd(lintProgrammerNum).Value

        'Filtering
        gudtMLX90277(lintProgrammerNum).Write.Filter = CWNumEditFilter(lintProgrammerNum).Value

        'Output Driver
        gudtMLX90277(lintProgrammerNum).Write.Mode = CWNumEditMode(lintProgrammerNum).Value

        'Customer ID
        gudtMLX90277(lintProgrammerNum).Write.CustID = CWNumEditCustomerID(lintProgrammerNum).Value

        'Time Generator
        gudtMLX90277(lintProgrammerNum).Write.FCKADJ = CWNumEditOscillatorAdjust(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.CKANACH = CWNumEditCKANACH(lintProgrammerNum).Value
        gudtMLX90277(lintProgrammerNum).Write.CKDACCH = CWNumEditCKDACCH(lintProgrammerNum).Value
        If frmMLX90277.chkSlowMode(lintProgrammerNum).Value = 1 Then
            gudtMLX90277(lintProgrammerNum).Write.SlowMode = True
        Else
            gudtMLX90277(lintProgrammerNum).Write.SlowMode = False
        End If

        'Melexis ID (Read Only)

        'Encode the EEPROM Write
        Call MLX90277.EncodeEEpromWrite(lintProgrammerNum)
    
    Next lintProgrammerNum

    'Write values to the EEPROM (addressed 0 - 64)
    If Not MLX90277.WriteEEPROMBlockByRows(0, 7) Then
        MsgBox "Error Executing Block Write!  Codes Not Programmed!", vbOKOnly & vbCritical, "Programmer Communication Error"
    End If

    'Read the EEPROM to verify the writes
    Call MLX90277.ReadEEPROM(gstrMLX90277Revision, lblnVotingError)

    'If there is a voting error, display a message box
    If lblnVotingError Then MsgBox "EEPROM Voting Failure!  MLX90277 IC Problem!", vbOKOnly & vbCritical, "EEPROM Failure!"

    'Loop through both programmers
    For lintProgrammerNum = 1 To 2
        'Verify that the reads and writes match
        If Not MLX90277.CompareReadsAndWrites(lintProgrammerNum) Then
            MsgBox "Reads Do Not Match Writes on PTC-04 #" & Format(lintProgrammerNum, "0"), vbOKOnly + vbCritical, "Programming Error!"
        End If

        'TC Table
        If (MLX90277.VerifyMLXCRC(lintProgrammerNum) And MLX90277.VerifyCustomerCRC(lintProgrammerNum)) Then
            chkTCTable(lintProgrammerNum).Value = 1
        Else
            chkTCTable(lintProgrammerNum).Value = 0
        End If
    Next lintProgrammerNum
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

'V1.2  This code should be used if the software is to be stand-alone
'Default to Read the MLX90277FA with the Read Button
'frmMLX90277.mnuMLX90277RevisionCx.Checked = False
'frmMLX90277.mnuMLX90277RevisionFA.Checked = True
'gstrMLX90277Revision = "FA"

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

mnuMLX90277RevisionCx.Checked = True
mnuMLX90277RevisionFA.Checked = False

'Cx Level Chip
gstrMLX90277Revision = "Cx"

End Sub

Private Sub mnuMLX90277RevisionFA_Click()   'V1.2
'
'   PURPOSE:   Selects the revision level
'
'  INPUT(S):   None
' OUTPUT(S):   None

mnuMLX90277RevisionCx.Checked = False
mnuMLX90277RevisionFA.Checked = True
'FA Level Chip
gstrMLX90277Revision = "FA"

End Sub

Private Sub mnuToolsSaveMLXData_Click()
'
'   PURPOSE:   Saves MLX data to file.
'
'  INPUT(S):   None
' OUTPUT(S):   None

Call Solver90277.SaveMLXtoFile

End Sub

Private Sub mnuToolsTimeStreamSettings_Click()
'
'   PURPOSE:   Displays Time Stream Settings Form
'
'  INPUT(S):   None
' OUTPUT(S):   None

frmTimeSettings.Show vbModal

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
