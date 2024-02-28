VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H00800000&
   Caption         =   "Program & Test System Start-Up"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Please click OK to setup the scanner program or click Exit to end the program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   563
      TabIndex        =   3
      Top             =   2280
      Width           =   10335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Welcome to the 705 Series Programmer/Scanner"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1500
      Left            =   1043
      TabIndex        =   0
      Top             =   360
      Width           =   9375
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'
'   PURPOSE: To exit start up form
'
'  INPUT(S): none
' OUTPUT(S): none

Unload Me               'Unload form
End                     'End Program

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To continue with program
'
'  INPUT(S): none
' OUTPUT(S): none

Unload Me

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To load form in center of screen
'
'  INPUT(S): none
' OUTPUT(S): none

Me.top = 0.15 * Screen.Height
Me.left = (Screen.Width - Me.Width) / 2

End Sub

