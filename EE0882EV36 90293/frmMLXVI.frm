VERSION 5.00
Begin VB.Form frmMLXVI 
   Caption         =   "Melexis VI"
   ClientHeight    =   2475
   ClientLeft      =   7035
   ClientTop       =   4710
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4875
   Begin VB.TextBox txtRS1 
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtRS2 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtC2 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtS2 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Text            =   "5"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "Read Current && Supply"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtC1 
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtS1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Text            =   "5"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Set Supply Here --------->"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "V"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "V"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Output 1             Output 2"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "mA"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "mA"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "V"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "V"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmMLXVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************** Melexis 90277 V/I  *****************
'
'Ver      Date      By   Purpose of modification
'1.0.0  10/02/2009  ANM  First release per SCN# 4391.
'1.0.1  11/21/2019  ANM  Adj for 90293.
'

Private Sub cmdCurrent_Click()
'
'   PURPOSE: To read the MLX current on demand.
'
'  INPUT(S): None.
' OUTPUT(S): None.

Dim lsng5V As Single
Dim lsng5V2 As Single

'If gudtMachine.Enable90288 Then
'    Call MLX90288.GetCurrent
'Else
    'Call MLX90277.GetCurrent
    'Set float then get bytes from float
    lsng5V = CSng(frmMLXVI.txtS1.Text)
    lsng5V2 = CSng(frmMLXVI.txtS2.Text)
    Call MyDev(lintDev1).Advanced.SetVdd(lsng5V)
    Call MyDev(lintDev2).Advanced.SetVdd(lsng5V2)

    gudtReading(0).mlxCurrent = MyDev(lintDev1).GetIdd
    gudtReading(1).mlxCurrent = MyDev(lintDev2).GetIdd
    gudtReading(0).mlxSupply = MyDev(lintDev1).GetVdd
    gudtReading(1).mlxSupply = MyDev(lintDev2).GetVdd
'End If

txtC1.Text = Format(gudtReading(0).mlxCurrent, "#0.00")
txtC2.Text = Format(gudtReading(1).mlxCurrent, "#0.00")
txtRS1.Text = Format(gudtReading(0).mlxSupply, "#0.00")
txtRS2.Text = Format(gudtReading(1).mlxSupply, "#0.00")

End Sub

'Private Sub cmdSuppy_Click()
''
''   PURPOSE: To change the MLX supply voltage on demand.
''
''  INPUT(S): None.
'' OUTPUT(S): None.
'
'Call Supply
'
'End Sub

'Private Sub Supply()
''
''   PURPOSE: To change the MLX supply voltage on demand.
''
''  INPUT(S): None.
'' OUTPUT(S): None.
'
'Dim lstrWrite(1 To 2) As String
'Dim lstrResponse(1 To 2) As String
'Dim MLXbyte(3) As Byte
'Dim MLXbyte2(3) As Byte
'Dim PPSStr(2) As String
'Dim lsng5V As Single
'Dim lsng5V2 As Single
'Dim X As Integer
'Dim lblnError As Boolean
'
'On Error GoTo SetErr
'
'lblnError = False
'
''Set float then get bytes from float
'lsng5V = CSng(frmMLXVI.txtS1.Text)
'lsng5V2 = CSng(frmMLXVI.txtS2.Text)
'CopyMemory MLXbyte(0), lsng5V, 4
'CopyMemory MLXbyte2(0), lsng5V2, 4
'
''Convert bytes to string
'PPSStr(0) = Chr$(MLXbyte(0)) & Chr$(MLXbyte(1)) & Chr$(MLXbyte(2)) & Chr$(MLXbyte(3))
'PPSStr(1) = Chr$(MLXbyte2(0)) & Chr$(MLXbyte2(1)) & Chr$(MLXbyte2(2)) & Chr$(MLXbyte2(3))
'
''Create the string to write to the programmer
'For X = 1 To 2
'    lstrWrite(X) = Chr$(PTC04CommandType.ptcSetPPS) & Chr$(0) & PPSStr(X - 1)
'Next X
'
''Send the string
''If gudtMachine.Enable90288 Then
''    If MLX90288.SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
''        'The response from the programmer should be the command
''        lblnError = Not (lstrResponse(1) = Chr$(PTC04CommandType.ptcSetPPS))
''    End If
''Else
'    If MLX90277.SendCommandGetResponse(lstrWrite(), lstrResponse()) Then
'        'The response from the programmer should be the command
'        lblnError = Not (lstrResponse(1) = Chr$(PTC04CommandType.ptcSetPPS))
'    End If
''End If
'
'If Not lblnError Then Exit Sub
'
'SetErr:
'    MsgBox "Error setting MLX Supply!", vbOKOnly, "MLX Error"
'
'End Sub

Private Sub Form_Load()
'
'   PURPOSE:   Executes when form is loaded, initializes form
'
'  INPUT(S):   None
' OUTPUT(S):   None

Me.top = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2
txtS1.Text = 5
txtS2.Text = 5

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'   PURPOSE:   Executes when form is closed, resets supply to 5V
'
'  INPUT(S):   None
' OUTPUT(S):   None

txtS1.Text = 5
txtS2.Text = 5

'Call Supply
Call MyDev(lintDev1).Advanced.SetVdd(5)
Call MyDev(lintDev2).Advanced.SetVdd(5)

End Sub

Private Sub txtS1_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtS1.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        txtS2.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 46             '.
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
        txtS2.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 46             '.
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

Private Sub txtS2_Change()
'
'   PURPOSE: To check the validity of the input character
'
'  INPUT(S): none
' OUTPUT(S): none

If txtS2.Text = "" Then
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdSuppy.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 43             '+
                'Accept the character
            Case 45             '-
                'Accept the character
            Case 46             '.
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
        cmdSuppy.SetFocus
    Else
        'Accept appropriate characters
        Select Case KeyAscii
            Case 8              'backspace
                'Accept the character
            Case 46             '.
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
