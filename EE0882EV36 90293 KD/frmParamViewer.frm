VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmParamViewer 
   Caption         =   "Parameter Viewer"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   16113
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   3840
      TabIndex        =   0
      Top             =   9360
      Width           =   948
   End
End
Attribute VB_Name = "frmParamViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

'Unload this form
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):
    
Dim lintRow As Integer
Dim lintColumn As Integer
Dim lvntCurrentY As Variant
Dim lstrName As String '3.6ANM \/\/
Dim lstrPN As String
Dim THandle As Long
Dim iret As Long       '3.6ANM /\/\

On Error GoTo ErrPrint:

lintPageHeight = Printer.Height
lsngTwipsPerPage = lintPageHeight - ((0.5 / 11) * (lintPageHeight)) 'take margins into consideration
lsngRowsPerPage = lsngTwipsPerPage / 200
lintCounter = 0

If Int(lsngRowsPerPage) > lsngRowsPerPage Then    'lsngRowsPerPage comes back as a decimal
    lsngRowsPerPage = Int(lsngRowsPerPage) - 2    'number, so these lines of code
Else                                              'change this decimal into an integer
    lsngRowsPerPage = Int(lsngRowsPerPage) - 1    'and rounds the integer down to the closest
End If                                            'low integer.

For lintRow = 0 To MSHFlexGrid1.Rows - 1

    If lintRow <> 0 And lintRow Mod lsngRowsPerPage = 0 Then   'Determines if the page
        lvntCurrentY = 200                                     'being printed is still printing.
        Printer.EndDoc                                         'If the page is done, the code reinitiates for the next page.
        lintCounter = 1                                        'Preset to if the page is on the same page still.
    Else
        lintCounter = lintCounter + 1
        lvntCurrentY = lintCounter * 200    ' Allows the printer not to overwrite the first line of every other page after the first page.
    End If

    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 100: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn)
    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 4000: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn + 1)
    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 5500: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn + 2)
    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 7000: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn + 3)
    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 8500: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn + 4)
    Printer.CurrentY = lvntCurrentY: Printer.CurrentX = 10000: Printer.Print MSHFlexGrid1.TextMatrix(lintRow, lintColumn + 5)

Next lintRow

Printer.EndDoc

'3.6ANM \/\/
If frmMain.mnuFunctionAutoSavePDFs.Checked = True Then
    BlockInput True
    Call frmDAQIO.KillTime(2000)
    
    THandle = FindWindowPartial(PDFWINDOW, "*")
    If THandle = 0 Then
        Call frmDAQIO.KillTime(2000)
        THandle = FindWindowPartial("*page*", "*")
    End If
    iret = BringWindowToTop(THandle)
    
    If THandle <> 0 Then
        lstrPN = left(frmMain.cboParameterFileName.Text, Len(frmMain.cboParameterFileName.Text) - 4)
        lstrName = lstrPN & "v" & CStr(gudtMachine.parameterRev) & Format(Now, " MM-DD-YY HHMMSSAMPM") & ".pdf"
        
        'Check if lot name folder exists, if not create it
        If Not gfsoFileSystemObject.FolderExists(PDFPATH & lstrPN) Then
            gfsoFileSystemObject.CreateFolder (PDFPATH & lstrPN)
        End If
    
        SendKeys "^(s)", True
        'SendKeys "{tab}", True
        'SendKeys "{tab}", True
        'SendKeys "{tab}", True
        Call frmDAQIO.KillTime(200)
        SendKeys PDFPATH & lstrPN & "\" & lstrName, True
        Call frmDAQIO.KillTime(200)
        SendKeys "{enter}", True
        Call frmDAQIO.KillTime(200)
        SendKeys "{esc}", True
        'SendKeys lstrName, True
        'SendKeys "{enter}", True
        'SendKeys "{tab}", True
        'SendKeys "{tab}", True
        'SendKeys "{tab}", True
        'SendKeys "{enter}", True
    End If
    
    BlockInput False
End If
'3.6ANM /\/\

Exit Sub

ErrPrint:
    BlockInput False '3.6ANM
    
End Sub
