VERSION 5.00
Begin VB.Form frmHTHHSData 
   Caption         =   "High Temperature High Humidity Exposure Data"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRelHum 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtCondition 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtDuration 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Width           =   255
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
      Left            =   4320
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Relative Humidity:"
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
      TabIndex        =   9
      Top             =   960
      Width           =   2295
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
      TabIndex        =   4
      Top             =   2160
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmHTHHSData"
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

txtTemp.Text = ""
txtRelHum.Text = "0"
txtDuration.Text = ""
txtCondition.Text = ""

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To load variables with data entered
'
'  INPUT(S):
' OUTPUT(S):

gudtExposure.HTempHHumiditySoak.Temperature = txtTemp.Text
gudtExposure.HTempHHumiditySoak.RelativeHumidity = CSng(txtRelHum.Text)
gudtExposure.HTempHHumiditySoak.Duration = txtDuration.Text
gudtExposure.HTempHHumiditySoak.Condition = txtCondition.Text

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
txtTemp.Text = gudtExposure.HTempHHumiditySoak.Temperature
txtRelHum.Text = CStr(gudtExposure.HTempHHumiditySoak.RelativeHumidity)
txtDuration.Text = gudtExposure.HTempHHumiditySoak.Duration
txtCondition.Text = gudtExposure.HTempHHumiditySoak.Condition

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE: Unload the form
'
'  INPUT(S):
' OUTPUT(S):

'Update the variables
gudtExposure.HTempHHumiditySoak.Temperature = txtTemp.Text
gudtExposure.HTempHHumiditySoak.RelativeHumidity = CSng(txtRelHum.Text)
gudtExposure.HTempHHumiditySoak.Duration = txtDuration.Text
gudtExposure.HTempHHumiditySoak.Condition = txtCondition.Text

End Sub

Private Sub txtRelHum_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtRelHum.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtDuration.SetFocus
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
        txtDuration.SetFocus
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
