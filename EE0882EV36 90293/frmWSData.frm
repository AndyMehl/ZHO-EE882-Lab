VERSION 5.00
Begin VB.Form frmWSData 
   Caption         =   "Water Spray Exposure Data"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtCondition 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtDuration 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Condition:"
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
      TabIndex        =   2
      Top             =   960
      Width           =   2055
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
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmWSData"
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

txtDuration.Text = ""
txtCondition.Text = ""

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To load variables with data entered
'
'  INPUT(S):
' OUTPUT(S):

gudtExposure.WaterSpray.Duration = txtDuration.Text
gudtExposure.WaterSpray.Condition = txtCondition.Text

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
txtDuration.Text = gudtExposure.WaterSpray.Duration
txtCondition.Text = gudtExposure.WaterSpray.Condition

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: Unload the form
'
'  INPUT(S):
' OUTPUT(S):

'Update the variables
gudtExposure.WaterSpray.Duration = txtDuration.Text
gudtExposure.WaterSpray.Condition = txtCondition.Text

End Sub

