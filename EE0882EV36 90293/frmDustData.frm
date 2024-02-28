VERSION 5.00
Begin VB.Form frmDustData 
   Caption         =   "Dust Exposure Data"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtCondition 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txtNumCycles 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtFreq 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtSettleTime 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtStirTime 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtDustAmount 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtDustType 
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
      TabIndex        =   12
      Top             =   3960
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
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label5 
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
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Settle Time:"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Stir Time:  "
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount of Dust:"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Type of Dust:"
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
Attribute VB_Name = "frmDustData"
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

txtDustType.Text = ""
txtDustAmount.Text = "0"
txtStirTime.Text = "0"
txtSettleTime.Text = "0"
txtFreq.Text = ""
txtNumCycles.Text = ""
txtCondition.Text = ""

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To load variables with data entered
'
'  INPUT(S):
' OUTPUT(S):

gudtExposure.Dust.TypeofDust = txtDustType.Text
gudtExposure.Dust.AmountofDust = CSng(txtDustAmount.Text)
gudtExposure.Dust.StirTime = CSng(txtStirTime.Text)
gudtExposure.Dust.SettleTime = CSng(txtSettleTime.Text)
gudtExposure.Dust.Frequency = txtFreq.Text
gudtExposure.Dust.NumberofCycles = txtNumCycles.Text
gudtExposure.Dust.Condition = txtCondition.Text

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
txtDustType.Text = gudtExposure.Dust.TypeofDust
txtDustAmount.Text = CStr(gudtExposure.Dust.AmountofDust)
txtStirTime.Text = CStr(gudtExposure.Dust.StirTime)
txtSettleTime.Text = CStr(gudtExposure.Dust.SettleTime)
txtFreq.Text = gudtExposure.Dust.Frequency
txtNumCycles.Text = gudtExposure.Dust.NumberofCycles
txtCondition.Text = gudtExposure.Dust.Condition

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: Unload the form
'
'  INPUT(S):
' OUTPUT(S):

'Update the variables
gudtExposure.Dust.TypeofDust = txtDustType.Text
gudtExposure.Dust.AmountofDust = CSng(txtDustAmount.Text)
gudtExposure.Dust.StirTime = CSng(txtStirTime.Text)
gudtExposure.Dust.SettleTime = CSng(txtSettleTime.Text)
gudtExposure.Dust.Frequency = txtFreq.Text
gudtExposure.Dust.NumberofCycles = txtNumCycles.Text
gudtExposure.Dust.Condition = txtCondition.Text

End Sub

Private Sub txtDustAmount_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtDustAmount.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtStirTime.SetFocus
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
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtStirTime.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

Private Sub txtStirTime_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtStirTime.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtSettleTime.SetFocus
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
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtSettleTime.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

Private Sub txtSettleTime_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtSettleTime.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtFreq.SetFocus
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
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
Else
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtFreq.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 48 To 57       '0-9
                'Accept the character
            Case Empty
                'Accept empty
            Case Else
                KeyAscii = 0    ' Cancel the character.
                Beep            ' Sound error signal.
        End Select
    End If
End If

End Sub

