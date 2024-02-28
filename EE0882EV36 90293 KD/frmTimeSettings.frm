VERSION 5.00
Begin VB.Form frmTimeSettings 
   Caption         =   "Time Settings"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimeProg 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   21
      Text            =   "0"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtTimeHold 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   20
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtTimePor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   19
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtTimeProg 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Text            =   "0"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtTimeHold 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtTimePor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtTimePuls 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Form"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
   End
   Begin VB.CommandButton cmdLoadTimeSettings 
      Caption         =   "Load New Time Settings"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox txtTimePuls 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tpor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   2760
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Thold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   2760
      TabIndex        =   26
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tprog:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   2760
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   4800
      TabIndex        =   24
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   4800
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   4800
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tpor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Thold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tprog:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   2160
      TabIndex        =   15
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   2160
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   2  'Center
      Caption         =   "PTC-03 #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tpuls:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   2  'Center
      Caption         =   "PTC-03 #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "uS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblTimeSettings 
      Alignment       =   1  'Right Justify
      Caption         =   "Tpuls:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTimeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'
'   PURPOSE: To unload the time settings from the form
'
'  INPUT(S): None.
' OUTPUT(S): None.

Unload Me

End Sub

Private Sub cmdLoadTimeSettings_Click()
'
'   PURPOSE: To load the time settings from the form
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer

On Error GoTo BadTimeSettings

For lintProgrammerNum = 1 To 2
    gudtPTC04(lintProgrammerNum).Tpuls = CInt(txtTimePuls(lintProgrammerNum).Text)
    gudtPTC04(lintProgrammerNum).Tpor = CInt(txtTimePor(lintProgrammerNum).Text)
    gudtPTC04(lintProgrammerNum).Thold = CInt(txtTimeHold(lintProgrammerNum).Text)
    gudtPTC04(lintProgrammerNum).Tprog = CInt(txtTimeProg(lintProgrammerNum).Text)
Next lintProgrammerNum

Exit Sub

BadTimeSettings:

    MsgBox "There was an error loading the time settings.  Please" & vbCrLf & _
           "enter only numbers into the text boxes.", vbOKOnly, "Error Loading From Text Boxes"

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To load the time settings form
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lintProgrammerNum As Integer

For lintProgrammerNum = 1 To 2
    txtTimePuls(lintProgrammerNum).Text = gudtPTC04(lintProgrammerNum).Tpuls
    txtTimePor(lintProgrammerNum).Text = gudtPTC04(lintProgrammerNum).Tpor
    txtTimeHold(lintProgrammerNum).Text = gudtPTC04(lintProgrammerNum).Thold
    txtTimeProg(lintProgrammerNum).Text = gudtPTC04(lintProgrammerNum).Tprog
Next lintProgrammerNum

End Sub
