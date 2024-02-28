VERSION 5.00
Begin VB.Form frmLotType 
   Caption         =   "Mode Selection"
   ClientHeight    =   945
   ClientLeft      =   6675
   ClientTop       =   5340
   ClientWidth     =   3075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   3075
   Begin VB.ComboBox cboLotFileType 
      Height          =   315
      ItemData        =   "frmLotType.frx":0000
      Left            =   240
      List            =   "frmLotType.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select Mode of Scanner:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLotType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboLotFileType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        gstrLotType = cboLotFileType.Text
        'frmMain.lblLotType.Caption = cboLotFileType.Text
        Unload Me
    End If
End Sub

Private Sub Form_Load()

Call AMAD705_2.PopulateLotTypesList

End Sub
