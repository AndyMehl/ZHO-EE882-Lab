VERSION 5.00
Begin VB.UserControl ctrLotSummary 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LockControls    =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   7455
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSummary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   "Error:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   4260
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPartYield 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblSummary 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   "Good:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   "Rejected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2220
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   "Severe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   "Lot Yield:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6300
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSummary 
         Caption         =   " Yield:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "ctrLotSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Revision   Date      Initials  Explanation
' 1.0     09-15-2004    SRC     Based on ctrLotSummary11.  Re-wrote
'                               for use as generic summary control.

Option Explicit

Public Property Get FrameCaption() As String
'
'   PURPOSE: Return the Caption of a Frame
'
'  INPUT(S): None
'
' OUTPUT(S):

FrameCaption = fraSummary.Caption

End Property

Public Property Let FrameCaption(ByVal NewCaption As String)
'
'   PURPOSE: Update the Caption of a Frame
'
'  INPUT(S): NewCaption = New Caption for the Label
'
' OUTPUT(S):

fraSummary.Caption = NewCaption

End Property

Public Property Get LabelCaption(ByVal LabelNum As Integer) As String
'
'   PURPOSE: Return the Caption of a label
'
'  INPUT(S): LabelNum = Which Label to return the caption of
'
' OUTPUT(S):

LabelCaption = lblSummary(LabelNum).Caption

End Property

Public Property Let LabelCaption(ByVal LabelNum As Integer, ByVal NewCaption As String)
'
'   PURPOSE: Update the Caption of a Label
'
'  INPUT(S): LabelNum   = Which Label to update
'            NewCaption = New Caption for the Label
'
' OUTPUT(S):

lblSummary(LabelNum).Caption = NewCaption

End Property

Public Property Let TextBackgroundColor(TextBoxNum As Integer, ByVal NewColor As ColorConstants)
'
'   PURPOSE: Update the Background Color of a Textbox
'
'  INPUT(S): TextBoxNum = Which TextBox to update
'            NewColor   = New Color for the TextBox background
'
' OUTPUT(S):

txtSummary(TextBoxNum).BackColor = NewColor

End Property

Public Property Get TextBoxText(ByVal TextBoxNum As Integer) As String
'
'   PURPOSE: Return the Text in a Textbox
'
'  INPUT(S): TextBoxNum = Which Textbox to return the text in
'
' OUTPUT(S):

TextBoxText = txtSummary(TextBoxNum).Text

End Property

Public Property Let TextBoxText(ByVal TextBoxNum As Integer, ByVal NewText As String)
'
'   PURPOSE: Update the Text in a TextBox
'
'  INPUT(S): TextBoxNumNum = Which TextBox to update
'            NewText       = New Text for the Text Box
'
' OUTPUT(S):

txtSummary(TextBoxNum).Text = NewText

End Property

