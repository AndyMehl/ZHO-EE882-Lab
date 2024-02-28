VERSION 5.00
Begin VB.Form frmTSData 
   Caption         =   "Thermal Shock Exposure Data"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtCondition 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtNumCycles 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtHighTempDur 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtHighTemp 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtLowTempDur 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtLowTemp 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "°C"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "°C"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   360
      Width           =   375
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
      Left            =   600
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
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "High Temp Duration:"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "High Temp:"
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
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Low Temp Duration:"
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
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Low Temp:"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmTSData"
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

txtLowTemp.Text = "0"
txtLowTempDur.Text = "0"
txtHighTemp.Text = "0"
txtHighTempDur.Text = "0"
txtNumCycles.Text = ""
txtCondition.Text = ""

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To load variables with data entered
'
'  INPUT(S):
' OUTPUT(S):

gudtExposure.ThermalShock.LowTemp = CSng(txtLowTemp.Text)
gudtExposure.ThermalShock.LowTempTime = CSng(txtLowTempDur.Text)
gudtExposure.ThermalShock.HighTemp = CSng(txtHighTemp.Text)
gudtExposure.ThermalShock.HighTempTime = CSng(txtHighTempDur.Text)
gudtExposure.ThermalShock.NumberofCycles = txtNumCycles.Text
gudtExposure.ThermalShock.Condition = txtCondition.Text

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
txtLowTemp.Text = CStr(gudtExposure.ThermalShock.LowTemp)
txtLowTempDur.Text = CStr(gudtExposure.ThermalShock.LowTempTime)
txtHighTemp.Text = CStr(gudtExposure.ThermalShock.HighTemp)
txtHighTempDur.Text = CStr(gudtExposure.ThermalShock.HighTempTime)
txtNumCycles.Text = gudtExposure.ThermalShock.NumberofCycles
txtCondition.Text = gudtExposure.ThermalShock.Condition

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: Unload the form
'
'  INPUT(S):
' OUTPUT(S):

'Update the variables
gudtExposure.ThermalShock.LowTemp = CSng(txtLowTemp.Text)
gudtExposure.ThermalShock.LowTempTime = CSng(txtLowTempDur.Text)
gudtExposure.ThermalShock.HighTemp = CSng(txtHighTemp.Text)
gudtExposure.ThermalShock.HighTempTime = CSng(txtHighTempDur.Text)
gudtExposure.ThermalShock.NumberofCycles = txtNumCycles.Text
gudtExposure.ThermalShock.Condition = txtCondition.Text

End Sub

Private Sub txtHighTemp_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtHighTemp.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtHighTempDur.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtHighTempDur.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

Private Sub txtHighTempDur_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtHighTempDur.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtLowTemp.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtLowTemp.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

Private Sub txtLowTempDur_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtLowTempDur.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtNumCycles.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtNumCycles.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

Private Sub txtLowTemp_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtLowTemp.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtLowTempDur.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtLowTempDur.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept the Empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub
