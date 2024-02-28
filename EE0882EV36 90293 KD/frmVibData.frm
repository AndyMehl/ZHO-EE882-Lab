VERSION 5.00
Begin VB.Form frmVibData 
   Caption         =   "Vibration Exposure Data"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtFrequency 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtNumCycles 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtDuration 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtPlanes 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtProfile 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Frequency:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Cycles:"
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
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Duration:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperature:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Planes:"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Profile:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmVibData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1.5ANM new form

Private Sub cmdClearAll_Click()
'
'   PURPOSE: To clear all info boxes
'
'  INPUT(S):
' OUTPUT(S):

txtProfile.Text = ""
txtTemp.Text = ""
txtDuration.Text = ""
txtPlanes.Text = ""
txtNumCycles.Text = ""
txtFrequency.Text = ""

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To load variables with data entered
'
'  INPUT(S):
' OUTPUT(S):

gudtExposure.Vibration.Profile = txtProfile.Text
gudtExposure.Vibration.Temperature = txtTemp.Text
gudtExposure.Vibration.Duration = txtDuration.Text
gudtExposure.Vibration.Planes = txtPlanes.Text
gudtExposure.Vibration.NumberofCycles = txtNumCycles.Text
gudtExposure.Vibration.Frequency = txtFrequency.Text

'Unload the form
Unload Me

End Sub

Private Sub Form_Load()
'
'   PURPOSE: Load the form
'
'  INPUT(S):
' OUTPUT(S):

'Center window on screen
Me.top = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2

'Fill in boxes with current information
txtProfile.Text = gudtExposure.Vibration.Profile
txtTemp.Text = gudtExposure.Vibration.Temperature
txtDuration.Text = gudtExposure.Vibration.Duration
txtPlanes.Text = gudtExposure.Vibration.Planes
txtNumCycles.Text = gudtExposure.Vibration.NumberofCycles
txtFrequency.Text = gudtExposure.Vibration.Frequency

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: Unload the form
'
'  INPUT(S):
' OUTPUT(S):

'Update the variables
gudtExposure.Vibration.Profile = txtProfile.Text
gudtExposure.Vibration.Temperature = txtTemp.Text
gudtExposure.Vibration.Duration = txtDuration.Text
gudtExposure.Vibration.Planes = txtPlanes.Text
gudtExposure.Vibration.NumberofCycles = txtNumCycles.Text
gudtExposure.Vibration.Frequency = txtFrequency.Text

End Sub
