VERSION 5.00
Begin VB.Form frmMultiScan 
   Caption         =   "Multi-Scan"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumScans 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "10"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Scans (2 - 1000):"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Scans Done:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblCS 
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmMultiScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.7ANM new form

Option Explicit
Private mblnStop As Boolean

Private Sub cmdStart_Click()

Dim lintNum, X As Integer

'Verify number of cycles
lintNum = CInt(txtNumScans.Text)
If (lintNum < 2) Or (lintNum > 1000) Then
    MsgBox "Number of scans must be 2 - 1000!", vbOKOnly, "Incorrect Number of Scans"
    Exit Sub
End If

cmdStart.Enabled = False
cmdStop.Enabled = True

'Cycle thru number of scans
For X = 1 To lintNum
    'Enable stop
    If mblnStop = True Then Exit For

    'Call the executive that handles Scanning
    If (gintAnomaly = 0) Then Call Pedal.RunTest

    'Set graph variable
    If gblnGraphEnable Then gblnGraphsLoaded = True

    'Save the results data if called for
    If ((gintAnomaly = 0) And gblnSaveScanResultsToFile) Then Call TestLab.Save705TLScanResultsToFile

    'Move to the Load Location
    Call MoveToLoadLocation

    'Diplay scan number
    lblCS.Caption = CStr(X)
Next X

'Finish up
mblnStop = False
cmdStart.Enabled = True
cmdStop.Enabled = False

Unload frmMultiScan

End Sub

Private Sub cmdStop_Click()
mblnStop = True
End Sub

Private Sub Form_Load()
cmdStop.Enabled = False
cmdStart.Enabled = True
mblnStop = False
lblCS.Caption = ""
End Sub
