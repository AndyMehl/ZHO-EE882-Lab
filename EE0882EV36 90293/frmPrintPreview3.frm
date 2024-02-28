VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrintPreview3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCommands 
      Height          =   500
      Left            =   0
      TabIndex        =   0
      Top             =   10560
      Width           =   8500
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   350
         Left            =   2000
         TabIndex        =   2
         Top             =   125
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   5000
         TabIndex        =   1
         Top             =   125
         Width           =   1500
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   375
      Index           =   3
      Left            =   500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2330
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   375
      Index           =   0
      Left            =   500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   500
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   750
      Index           =   2
      Left            =   500
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1590
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   375
      Index           =   4
      Left            =   500
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2690
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   7365
      Index           =   5
      Left            =   495
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3045
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   12991
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   750
      Index           =   1
      Left            =   500
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   850
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPrintPreview3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1.6ANM new form

Private Const SCALEFACTOR = 4   'Constant for resizing images for better resolution

Private Sub cmdCancel_Click()
'
'   PURPOSE: To cancel printing and unload the form
'
'  INPUT(S): none
' OUTPUT(S): none

Unload frmPrintPreview3

End Sub

Private Sub cmdPrint_Click()
'
'   PURPOSE: To send the current display to the printer.
'
'  INPUT(S): none
' OUTPUT(S): none
Dim i As Integer

'Make the grids invisible while printing
For i = 0 To 5
    flexGrid(i).Visible = False
Next i

'Print the display
Call PrintDisplay

'Make the grids visible again
For i = 0 To 5
    flexGrid(i).Visible = True
Next i

End Sub

Public Sub DisplayData(GridNum As Integer)
'
'   PURPOSE: To populate the print preview form with printable data.
'
'  INPUT(S): none
' OUTPUT(S): none

Dim i As Long, j As Long, k As Long

Dim lintControlHeight As Single
Dim lintControlWidth As Single
Dim lsngVerticalScaler As Single
Dim lsngHorizontalScaler As Single

'Fill the system information grid
flexGrid(0).Clear
flexGrid(0).Rows = 1
flexGrid(0).Cols = 1
flexGrid(0).ColWidth(0) = 7500
flexGrid(0).RowHeight(0) = 375
flexGrid(0).ColAlignment(0) = flexAlignCenterCenter
flexGrid(0).Font.Bold = True
flexGrid(0).TextMatrix(0, 0) = "CTS Corporation   Elkhart, IN, USA"
flexGrid(0).GridLines = flexGridNone
flexGrid(0).GridColor = vbWhite
flexGrid(0).BorderStyle = flexBorderNone

'Fill the customer information grid
flexGrid(1).Clear
flexGrid(1).Rows = 2
flexGrid(1).Cols = 4
flexGrid(1).RowHeight(0) = 375
flexGrid(1).RowHeight(1) = 375
flexGrid(1).ColWidth(0) = 1900: flexGrid(1).ColWidth(1) = 1850
flexGrid(1).ColWidth(2) = 1850: flexGrid(1).ColWidth(3) = 1900
flexGrid(1).TextMatrix(0, 0) = "Customer"
flexGrid(1).TextMatrix(1, 0) = gstrCustomerName                     '1.8ANM
flexGrid(1).TextMatrix(0, 1) = "Customer Part #"
flexGrid(1).TextMatrix(1, 1) = gudtMachine.CustomerPartNum          '1.8ANM
flexGrid(1).TextMatrix(0, 2) = "CTS Part #"
flexGrid(1).TextMatrix(1, 2) = gstrCTSPartNum                       '1.8ANM
flexGrid(1).TextMatrix(0, 3) = "Part Name"
flexGrid(1).TextMatrix(1, 3) = gstrPartName                         '1.8ANM
flexGrid(1).GridLines = flexGridNone
flexGrid(1).GridColor = vbWhite
flexGrid(1).BorderStyle = flexBorderNone

flexGrid(1).Row = 0
For k = 0 To flexGrid(1).Cols - 1
    flexGrid(1).Col = k
    flexGrid(1).CellAlignment = flexAlignCenterCenter
Next k

flexGrid(1).Row = 1
For k = 0 To flexGrid(1).Cols - 1
    flexGrid(1).Col = k
    flexGrid(1).CellAlignment = flexAlignCenterCenter
Next k

'Fill the user information 1 grid
flexGrid(2).Clear
flexGrid(2).Rows = 2
flexGrid(2).Cols = 7
flexGrid(2).RowHeight(0) = 375
flexGrid(2).RowHeight(1) = 375
flexGrid(2).ColWidth(0) = 1000: flexGrid(2).ColWidth(1) = 1000
flexGrid(2).ColWidth(2) = 1400: flexGrid(2).ColWidth(3) = 1100
flexGrid(2).ColWidth(4) = 1000: flexGrid(2).ColWidth(5) = 1000
flexGrid(2).ColWidth(6) = 1000
flexGrid(2).TextMatrix(0, 0) = "Operator"
flexGrid(2).TextMatrix(1, 0) = frmMain.ctrSetupInfo1.Operator
flexGrid(2).TextMatrix(0, 1) = "Series"
flexGrid(2).TextMatrix(1, 1) = frmMain.ctrSetupInfo1.Series
flexGrid(2).TextMatrix(0, 2) = "Serial #"
flexGrid(2).TextMatrix(1, 2) = frmMain.ctrSetupInfo1.PartNum
flexGrid(2).TextMatrix(0, 3) = "Date Code"
flexGrid(2).TextMatrix(1, 3) = frmMain.ctrSetupInfo1.DateCode
flexGrid(2).TextMatrix(0, 4) = "TestLog #"
flexGrid(2).TextMatrix(1, 4) = frmMain.ctrSetupInfo1.TLNum
flexGrid(2).TextMatrix(0, 5) = "Temperature"
flexGrid(2).TextMatrix(1, 5) = frmMain.ctrSetupInfo1.Temperature
flexGrid(2).TextMatrix(0, 6) = "Sample #"
flexGrid(2).TextMatrix(1, 6) = frmMain.ctrSetupInfo1.Sample
flexGrid(2).GridLines = flexGridNone
flexGrid(2).GridColor = vbWhite
flexGrid(2).BorderStyle = flexBorderNone

flexGrid(2).Row = 0
For k = 0 To flexGrid(2).Cols - 1
    flexGrid(2).Col = k
    flexGrid(2).CellAlignment = flexAlignCenterCenter
Next k

flexGrid(2).Row = 1
For k = 0 To flexGrid(2).Cols - 1
    flexGrid(2).Col = k
    flexGrid(2).CellAlignment = flexAlignCenterCenter
Next k

'Fill the setup information grid
flexGrid(3).Clear
flexGrid(3).Rows = 1
flexGrid(3).Cols = 3
flexGrid(3).RowHeight(0) = 375
flexGrid(3).ColWidth(0) = 2505: flexGrid(3).ColWidth(1) = 2640: flexGrid(3).ColWidth(2) = 2355
flexGrid(3).ColAlignment(0) = flexAlignLeftCenter
flexGrid(3).TextMatrix(0, 0) = Format$(Now, "mmmm d, yyyy   h:mm:ss AM/PM")
flexGrid(3).ColAlignment(1) = flexAlignLeftCenter
flexGrid(3).TextMatrix(0, 1) = "Parameter File: " & frmMain.cboParameterFileName
flexGrid(3).ColAlignment(2) = flexAlignLeftCenter
flexGrid(3).TextMatrix(0, 2) = "Lot File: " & frmMain.cboLotFile
flexGrid(3).GridLines = flexGridNone
flexGrid(3).GridColor = vbWhite
flexGrid(3).BorderStyle = flexBorderNone

'Fill the user information 2 grid
flexGrid(4).Clear
flexGrid(4).Rows = 1
flexGrid(4).Cols = 1
flexGrid(4).RowHeight(0) = 375
flexGrid(4).ColWidth(0) = 7500:
flexGrid(4).TextMatrix(0, 0) = "Comment: " & frmMain.ctrSetupInfo1.Comment
flexGrid(4).GridLines = flexGridNone
flexGrid(4).GridColor = vbWhite
flexGrid(4).BorderStyle = flexBorderNone

'Fill the Results/Stats grid
flexGrid(5).Clear
flexGrid(5).Rows = frmMain.ctrResultsTabs1.NumberOfRows(GridNum)
flexGrid(5).Cols = frmMain.ctrResultsTabs1.NumberOfColumns(GridNum)

'Determine the grid height (on the control)
For i = 0 To flexGrid(5).Rows - 1
    lintControlHeight = lintControlHeight + frmMain.ctrResultsTabs1.RowHeight(GridNum, i)
Next i
'Determine the grid width (on the control)
For i = 0 To flexGrid(5).Cols - 1
    lintControlWidth = lintControlWidth + frmMain.ctrResultsTabs1.ColumnSpacing(GridNum, CInt(i))
Next i

'Create scalers for vertical and horizontal scaling
lsngVerticalScaler = flexGrid(5).Height / lintControlHeight
lsngHorizontalScaler = flexGrid(5).Width / lintControlWidth

For i = 0 To flexGrid(5).Rows - 1       'Size the rows
    flexGrid(5).RowHeight(i) = (frmMain.ctrResultsTabs1.RowHeight(GridNum, i) * lsngVerticalScaler)
Next i
For i = 0 To flexGrid(5).Cols - 1       'Size the columns
    flexGrid(5).ColWidth(i) = Fix(frmMain.ctrResultsTabs1.ColumnSpacing(GridNum, CInt(i)) * lsngHorizontalScaler)
Next i

'Put the data in the grid
For i = 0 To flexGrid(5).Rows - 1       'Loop through all the rows
    For j = 0 To flexGrid(5).Cols - 1   'Loop through all the columns
        flexGrid(5).Row = i
        flexGrid(5).Col = j
        flexGrid(5).CellBackColor = &H80000005
        'From the results control:
        flexGrid(5).Text = frmMain.ctrResultsTabs1.Data(GridNum, i, j)
        flexGrid(5).CellAlignment = frmMain.ctrResultsTabs1.TextAlignment(GridNum, i, j)
        flexGrid(5).CellFontName = frmMain.ctrResultsTabs1.CellFont(GridNum, i, j)
        'Scale the font size horizontally
        flexGrid(5).CellFontSize = frmMain.ctrResultsTabs1.CellFontSize(GridNum, i, j) * lsngHorizontalScaler
    Next j
Next i

End Sub

Public Sub PrintDisplay()
'
'   PURPOSE: To print the contents of the display
'
'  INPUT(S): none
' OUTPUT(S): none

Dim lsngVerticalScaler As Single
Dim lsngHorizontalScaler As Single
Dim lsngVerticalScalerGrid As Single
Dim lsngHorizontalScalerGrid As Single
Dim i As Integer
Dim j As Long, k As Long
Dim llngHeight(5) As Long
Dim llngWidth(5) As Long
Dim llngLeft As Long
Dim llngTop As Long
Dim llngRight As Long
Dim llngBottom As Long
Dim lstrName As String '3.6ANM \/\/
Dim lstrPN As String
Dim THandle As Long
Dim iret As Long       '3.6ANM /\/\

On Error GoTo Exit_Sub

'Resize the entire grid by SCALEFACTOR for better resolution
For i = 0 To 5
    flexGrid(i).Height = flexGrid(i).Height * SCALEFACTOR
    flexGrid(i).Width = flexGrid(i).Width * SCALEFACTOR
    flexGrid(i).top = flexGrid(i).top * SCALEFACTOR
    flexGrid(i).left = flexGrid(i).left * SCALEFACTOR
    flexGrid(i).GridLines = flexGridNone    'Turn off the gridlines  (prints better)
    For j = 0 To flexGrid(i).Rows - 1
        flexGrid(i).RowHeight(j) = flexGrid(i).RowHeight(j) * SCALEFACTOR
        flexGrid(i).Row = j
        For k = 0 To flexGrid(i).Cols - 1
            flexGrid(i).Col = k
            flexGrid(i).CellFontSize = flexGrid(i).CellFontSize * SCALEFACTOR
        Next k
    Next j
    For j = 0 To flexGrid(i).Cols - 1
        flexGrid(i).ColWidth(j) = flexGrid(i).ColWidth(j) * SCALEFACTOR
    Next j
Next i
    
lsngHorizontalScaler = ((Printer.ScaleWidth / 8500) / SCALEFACTOR)
lsngVerticalScaler = ((Printer.ScaleHeight / 11000) / SCALEFACTOR)
    
For i = 0 To 5

    'Print the flexgrids
    Printer.PaintPicture flexGrid(i).Picture, flexGrid(i).left * lsngVerticalScaler, flexGrid(i).top * lsngVerticalScaler, flexGrid(i).Width * lsngHorizontalScaler, flexGrid(i).Height * lsngVerticalScaler

    'Add Row sizes
    For j = 0 To flexGrid(i).Rows - 1
        llngHeight(i) = llngHeight(i) + flexGrid(i).RowHeight(j)
    Next j
    'Add Column sizes
    For k = 0 To flexGrid(i).Cols - 1
        llngWidth(i) = llngWidth(i) + flexGrid(i).ColWidth(k)
    Next k

    'Define the scaler values each time to compensate for slightly mis-dimensioned grids
    lsngHorizontalScalerGrid = lsngHorizontalScaler * (flexGrid(i).Width / llngWidth(i))
    lsngVerticalScalerGrid = lsngVerticalScaler * (flexGrid(i).Height / llngHeight(i))

    'Print the gridlines
    If i = 5 Then
        For j = 0 To flexGrid(i).Rows - 1
            flexGrid(i).Row = j
            For k = 0 To flexGrid(i).Cols - 1
                flexGrid(i).Col = k
                'Define the current cells borders
                llngLeft = (flexGrid(i).left * lsngHorizontalScaler) + (flexGrid(i).CellLeft * lsngHorizontalScalerGrid)
                llngTop = (flexGrid(i).top * lsngVerticalScaler) + (flexGrid(i).CellTop * lsngVerticalScalerGrid)
                llngRight = (flexGrid(i).left * lsngHorizontalScaler) + ((flexGrid(i).CellLeft + flexGrid(i).CellWidth) * lsngHorizontalScalerGrid)
                llngBottom = (flexGrid(i).top * lsngVerticalScaler) + ((flexGrid(i).CellTop + flexGrid(i).CellHeight) * lsngVerticalScalerGrid)
                Printer.Line (llngLeft, llngTop)-(llngRight, llngBottom), &H80000008, B
            Next k
        Next j
    End If
Next i

'Let the printer know the document is done
Printer.EndDoc

'3.6ANM \/\/
If frmMain.mnuFunctionAutoSavePDFs.Checked = True And Not gblnSkipPDF Then
    BlockInput True
    Call frmDAQIO.KillTime(2000)
    
    THandle = FindWindowPartial(PDFWINDOW, "*")
    If THandle = 0 Then
        Call frmDAQIO.KillTime(2000)
        THandle = FindWindowPartial("*page*", "*")
    End If
    iret = BringWindowToTop(THandle)
    
    If THandle <> 0 Then
        lstrPN = gstrLotName
        lstrName = gstrSerialNumber & gstrType & "Results " & Format(Now, "MM-DD-YY HHMMSSAMPM") & ".pdf"
        
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

'Shrink the grid back down
For i = 0 To 5
    flexGrid(i).Height = flexGrid(i).Height / SCALEFACTOR
    flexGrid(i).Width = flexGrid(i).Width / SCALEFACTOR
    flexGrid(i).top = flexGrid(i).top / SCALEFACTOR
    flexGrid(i).left = flexGrid(i).left / SCALEFACTOR
    flexGrid(i).GridLines = flexGridFlat
    For j = 0 To flexGrid(i).Rows - 1
        flexGrid(i).RowHeight(j) = flexGrid(i).RowHeight(j) / SCALEFACTOR
        flexGrid(i).Row = j
        For k = 0 To flexGrid(i).Cols - 1
            flexGrid(i).Col = k
            flexGrid(i).CellFontSize = flexGrid(i).CellFontSize / SCALEFACTOR
        Next k
    Next j
    For j = 0 To flexGrid(i).Cols - 1
        flexGrid(i).ColWidth(j) = flexGrid(i).ColWidth(j) / SCALEFACTOR
    Next j
Next i

Exit Sub

Exit_Sub:
    BlockInput False '3.6ANM
    If Err.number Then MsgBox Err.Description, vbOKOnly, "Printer Error"

End Sub
