VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2100
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Height          =   264
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   468
      Width           =   2028
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1572
      TabIndex        =   3
      Top             =   972
      Width           =   816
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   972
      Width           =   816
   End
   Begin VB.Label Label1 
      Caption         =   "Enter supervisory password:"
      Height          =   228
      Left            =   360
      TabIndex        =   0
      Top             =   252
      Width           =   2280
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

gblnAdministrator = False
Unload Me

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

If UCase(txtPassword.Text) = "CTS" Then
    gblnAdministrator = True
    Unload Me
Else
    Beep
    MsgBox "Incorrect Password.  Please try again"
    gblnAdministrator = False
End If

End Sub


