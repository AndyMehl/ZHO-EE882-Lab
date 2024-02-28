VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{860FFAFB-5AAA-11D2-81EB-006008A2E49D}#1.0#0"; "Pesgo32a.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctrResultsTabs 
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13935
   ScaleHeight     =   8310
   ScaleWidth      =   13935
   Begin TabDlg.SSTab tabResults 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ctrResultsTabs.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flexData(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "ctrResultsTabs.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flexData(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "ctrResultsTabs.ctx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "tlbGraphs"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "picGraph"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "ctrResultsTabs.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "flexData(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "ctrResultsTabs.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "flexData(3)"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox picGraph 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   7215
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   7155
         ScaleWidth      =   13635
         TabIndex        =   3
         Top             =   960
         Width           =   13695
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   1995
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4419
            _ExtentY        =   3519
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":008C
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   1995
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   2520
            Visible         =   0   'False
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":0E83
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   4800
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":1C7A
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   1995
            Index           =   3
            Left            =   2760
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":2A71
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   4
            Left            =   2760
            TabIndex        =   8
            Top             =   2500
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":3868
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   5
            Left            =   2760
            TabIndex        =   9
            Top             =   4800
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":465F
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   6
            Left            =   5400
            TabIndex        =   10
            Top             =   250
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":5456
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   7
            Left            =   5400
            TabIndex        =   11
            Top             =   2500
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":624D
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   8
            Left            =   5400
            TabIndex        =   12
            Top             =   4800
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":7044
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   9
            Left            =   8040
            TabIndex        =   13
            Top             =   250
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":7E3B
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   1995
            Index           =   10
            Left            =   8040
            TabIndex        =   14
            Top             =   2520
            Visible         =   0   'False
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":8C32
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   11
            Left            =   8040
            TabIndex        =   15
            Top             =   4800
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":9A29
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   12
            Left            =   10800
            TabIndex        =   16
            Top             =   250
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":A820
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   13
            Left            =   10800
            TabIndex        =   17
            Top             =   2500
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":B617
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   2000
            Index           =   14
            Left            =   10800
            TabIndex        =   18
            Top             =   4800
            Visible         =   0   'False
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":C40E
         End
         Begin PESGOALib.Pesgoa pesGraph1 
            Height          =   1995
            Index           =   15
            Left            =   11040
            TabIndex        =   19
            Top             =   5280
            Visible         =   0   'False
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   3528
            _StockProps     =   96
            _AllProps       =   "ctrResultsTabs.ctx":D205
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flexData 
         Height          =   7695
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   100
         Cols            =   100
         Enabled         =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid flexData 
         Height          =   7695
         Index           =   1
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   100
         Cols            =   100
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.Toolbar tlbGraphs 
         Height          =   600
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   1058
         ButtonWidth     =   2196
         ButtonHeight    =   953
         Wrappable       =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Graph"
               Object.ToolTipText     =   "Select Graph"
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clear All Graphs"
               Object.ToolTipText     =   "Clear All Graphs"
            EndProperty
         EndProperty
         Begin VB.CommandButton cmdDefGraph 
            Caption         =   "Select Default Graphs"
            Height          =   375
            Left            =   3120
            TabIndex        =   27
            Top             =   120
            Width           =   2175
         End
         Begin VB.Frame fraOrientation 
            Caption         =   "Printing Orientation"
            Height          =   495
            Left            =   9000
            TabIndex        =   22
            Top             =   0
            Width           =   2775
            Begin VB.OptionButton optOrientation 
               Caption         =   "Portrait"
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   24
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optOrientation 
               Caption         =   "LandScape"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdPrintGraphs 
            Caption         =   "Print Visible Graphs"
            Height          =   375
            Left            =   12000
            TabIndex        =   21
            Top             =   120
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flexData 
         Height          =   7695
         Index           =   2
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   100
         Cols            =   100
         Enabled         =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid flexData 
         Height          =   7695
         Index           =   3
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   100
         Cols            =   100
         Enabled         =   -1  'True
      End
   End
End
Attribute VB_Name = "ctrResultsTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'  This control is designed to give the flexibility to write data to the
'  screen with any 5th Generation Software.  Methods & Properties should be
'  added as necessary to allow flexible functionality.  The following routines
'  and properties are externally available:
'
'     Subs :
'           ClearData(GridNumber As Integer, StartColumn As Long, StopColumn As Long)
'           ClearTotalGrid(GridNumber As Integer)
'           ExtractDataEvenXIntervals(GraphArray() As Variant)
'           ExtractDataXAndY(GraphArray())
'           PrintAllGraphsInWindow()
'
'     Properties
'
'     Get :
'           BoldText(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Boolean
'           CellColor(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As ColorConstants
'           CellFont(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As String
'           CellFontSize(GridNumber As Integer, Row As Long, Column As Long) As Integer
'           CellWordWrap(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Boolean
'           ColumnSpacing(GridNumber As Integer, ColumnNum As Integer) As Long
'           Data(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Variant
'           NumberOfColumns(GridNumber As Integer) As Long
'           NumberOfRows(GridNumber As Integer) As Long
'           RowHeight(GridNumber As Integer, RowNum As Long) As Long
'           TabName(TabNum As Integer) As String
'           TextAlignment(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As AlignmentSettings'

'     Let :
'           ActiveTab(ByVal TabNumber As Integer)
'           BoldColumn(GridNumber As Integer, ColumnNum As Long, ByVal Bold As Boolean)
'           BoldRow(GridNumber As Integer, RowNum As Long, ByVal Bold As Boolean)
'           BoldText(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Bold As Boolean)
'           CellColor(GridNumber As Integer, Row As Long, Column As Long, ByVal BackColor As ColorConstants)
'           CellFont(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal FontName As String)
'           CellFontSize(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal FontSize As Integer)
'           CellWordWrap(GridNumber As Integer, RowNum As Long, ColumnNum As Long, WordWrapCell As Boolean)
'           ColumnAlignment(GridNumber As Integer, ColumnNum As Long, ByVal Alignment As AlignmentSettings)
'           ColumnSpacing(GridNumber As Integer, ColumnNum As Integer, SizeInTwips As Long)
'           Data(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Value As Variant)
'           NumberOfColumns(GridNumber As Integer, NumColumns As Long)
'           NumberOfRows(GridNumber As Integer, NumRows As Long)
'           RowAlignment(GridNumber As Integer, RowNum As Long, ByVal Alignment As AlignmentSettings)
'           RowHeight(GridNumber As Integer, RowNum As Long, RowHeightInTwips As Long)
'           TabName(TabNum As Integer, ByVal Name As String)
'           TextAlignment(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Alignment As AlignmentSettings)
'           TotalCellFont(GridNumber As Integer, ByVal FontName As String)
'           TotalCellFontSize(GridNumber As Integer, ByVal FontSize As Integer)
'           TotalRowHeight(GridNumber As Integer, HeightInTwips As Long)
'           TotalWordWrap(GridNumber As Integer, WordWrap As Boolean)
'
'  This control was written by Scott R. Calkins of CTS Automotive Elkhart in July 2004 for use within
'  5th Generation software projects.  Based largely on ctrResGM by Tad L. Miller.
'  REVISION  INIT    DATE      DESCRIPTION
'  1.0.0     SRC  07/08/2004   Initial Release of Software for use with 5th Generation Software.
'  1.0.1     SRC  09/15/2004   Replaced instances of TabPro control with built in VB Tabs.
'  1.1.0     ANM  02/03/2006   Updated to include select default graph button per PR 11801-K
'

'Constants for graph location within the control
Private Const GRAPHLEFT = 10
Private Const GRAPHTOP = 100
Private Const TOTALGRAPHHEIGHT = 7000
Private Const TOTALGRAPHWIDTH = 13600

'Printer Orientation.  Default = 2, Landscape
Private mintOrientation As Integer

Public Property Let ActiveTab(ByVal TabNumber As Integer)
'
'   PURPOSE: Selects which Tab is on top (Active)
'
'  INPUT(S): TabNumber = Tab to make Active
'
' OUTPUT(S): None

'Verify that there is a Tab(TabNum)
If TabNumber >= 0 And TabNumber <= tabResults.Tabs - 1 Then
    tabResults.Tab = TabNumber
End If

End Property

Public Property Let BoldColumn(GridNumber As Integer, ColumnNum As Long, ByVal Bold As Boolean)
'
'   PURPOSE: To set whether or not the cells in a column have a bold font.
'
'  INPUT(S): GridNumber = Selected Grid
'            ColumnNum  = Selected Column
'            Bold       = Whether or not to Bold the selected Column
'
' OUTPUT(S): None

Dim llngRowNum As Long

'Only proceed if the selected column is valid
If ColumnNum <= flexData(GridNumber).Cols Then
    'Set the active column
    flexData(GridNumber).Col = ColumnNum
    'Set the Bold property of the cells in the column
    For llngRowNum = 0 To flexData(GridNumber).Rows - 1
        flexData(GridNumber).Row = llngRowNum
        flexData(GridNumber).CellFontBold = Bold
    Next llngRowNum
End If

End Property

Public Property Let BoldRow(GridNumber As Integer, RowNum As Long, ByVal Bold As Boolean)
'
'   PURPOSE: To set whether or not the cells in a row have a bold font.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            Bold       = Whether or not to Bold the selected Row
'
' OUTPUT(S): None

Dim llngColumnNum As Long

'Only proceed if the selected column is valid
If RowNum <= flexData(GridNumber).Rows Then
    'Set the active row
    flexData(GridNumber).Row = RowNum
    'Set the Bold property of the cells in the row
    For llngColumnNum = 0 To flexData(GridNumber).Cols - 1
        flexData(GridNumber).Col = llngColumnNum
        flexData(GridNumber).CellFontBold = Bold
    Next llngColumnNum
End If

End Property

Public Property Get BoldText(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Boolean
'
'   PURPOSE: To return whether or not the selected cell has a bold font.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Return the Bold property
BoldText = flexData(GridNumber).CellFontBold
    
End Property

Public Property Let BoldText(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Bold As Boolean)
'
'   PURPOSE: To set whether or not the selected cell has a bold font.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'            Bold       = Whether or not to Bold the selected cell
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Set the Bold property
flexData(GridNumber).CellFontBold = Bold

End Property

Public Property Get CellColor(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As ColorConstants
'
'   PURPOSE: To get the background color of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Return the Cell Color
CellColor = flexData(GridNumber).CellBackColor

End Property

Public Property Let CellColor(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal BackColor As ColorConstants)
'
'   PURPOSE: To set the background color of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'            BackColor  = New Cell Back Color
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

flexData(GridNumber).CellBackColor = BackColor

End Property

Public Property Get CellFont(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As String
'
'   PURPOSE: To get the cell font name of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Return the Cell Font Name
CellFont = flexData(GridNumber).CellFontName

End Property

Public Property Let CellFont(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal FontName As String)
'
'   PURPOSE: To set the cell font name of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'            FontName   = Cell Font Name
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Set the Cell Font Name
flexData(GridNumber).CellFontName = FontName

End Property

Public Property Get CellFontSize(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Integer
'
'   PURPOSE: To get the cell font size of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Return the Cell Font Size
CellFontSize = flexData(GridNumber).CellFontSize

End Property

Public Property Let CellFontSize(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal FontSize As Integer)
'
'   PURPOSE: To set the cell font size of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'            FontSize   = Cell Font Size
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Set the Cell Font Size
flexData(GridNumber).CellFontSize = FontSize

End Property

Public Property Get CellWordWrap(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Boolean
'
'   PURPOSE: To get the cell word wrap property of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

CellWordWrap = flexData(GridNumber).WordWrap

End Property

Public Property Let CellWordWrap(GridNumber As Integer, RowNum As Long, ColumnNum As Long, WordWrapCell As Boolean)
'
'   PURPOSE: To set the cell word wrap property of the selected cell.
'
'  INPUT(S): GridNumber   = Selected Grid
'            RowNum       = Selected Row
'            ColumnNum    = Selected Column
'            WordWrapCell = Whether or not to word wrap the selected cell
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

flexData(GridNumber).WordWrap = WordWrapCell

End Property

Public Sub ClearData(GridNumber As Integer, StartColumn As Long, StopColumn As Long)
'
'   PURPOSE: To clear the selected columns in the selected grid.
'
'  INPUT(S): GridNumber     [Integer]       -which grid to modify
'            StartColumn    [Long]          -Start clearing
'            StopColumn     [Long]          -Stop clearing
'
' OUTPUT(S): None

Dim llngRowNum As Long
Dim llngColumnNum As Long

For llngRowNum = 1 To flexData(GridNumber).Rows - 1
    For llngColumnNum = StartColumn To StopColumn
        Data(GridNumber, llngRowNum, llngColumnNum) = ""
        CellColor(GridNumber, llngRowNum, llngColumnNum) = vbWhite
    Next llngColumnNum
Next llngRowNum

End Sub

Public Sub ClearTotalGrid(GridNumber As Integer)
'
'   PURPOSE: To clear the selected grid
'
'  INPUT(S): GridNumber = Which grid to clear
'
' OUTPUT(S): None

flexData(GridNumber).Clear

End Sub

Public Property Let ColumnAlignment(GridNumber As Integer, ColumnNum As Long, ByVal Alignment As AlignmentSettings)
'
'   PURPOSE: To get the cell alignment property for the selected column.
'
'  INPUT(S): GridNumber = Selected Grid
'            ColumnNum  = Selected Column
'            Alignment  = New Alignment setting
'
' OUTPUT(S): None

Dim llngRowNum As Long

'Only proceed if the selected column is valid
If ColumnNum <= flexData(GridNumber).Cols Then
    'Set the active column
    flexData(GridNumber).Col = ColumnNum
    For llngRowNum = 0 To flexData(GridNumber).Rows - 1
        flexData(GridNumber).Row = llngRowNum
        flexData(GridNumber).CellAlignment = Alignment
    Next llngRowNum
End If

End Property

Public Property Get ColumnSpacing(GridNumber As Integer, ColumnNum As Integer) As Long
'
'   PURPOSE: To return the column size on the selected Grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

ColumnSpacing = flexData(GridNumber).ColWidth(ColumnNum)

End Property

Public Property Let ColumnSpacing(GridNumber As Integer, ColumnNum As Integer, SizeInTwips As Long)
'
'   PURPOSE: To set the column size on the selected Grid.
'
'  INPUT(S): GridNumber  = Selected Grid
'            ColumnNum   = Selected Column
'            SizeInTwips = New Column Width in Twips
'
' OUTPUT(S): None

flexData(GridNumber).ColWidth(ColumnNum) = SizeInTwips

End Property

Public Property Get Data(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As Variant
'
'   PURPOSE: To get the text in a selected grid location
'
'  INPUT(S): GridNumber = Selected Grid
'            ColumnNum  = Selected Column
'            RowNum     = Selected Row
'
' OUTPUT(S): None

'Return the text at the selected location
Data = flexData(GridNumber).TextMatrix(RowNum, ColumnNum)

End Property

Public Property Let Data(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Value As Variant)
'
'   PURPOSE: To set the text in a selected grid location
'
'  INPUT(S): GridNumber = Selected Grid
'            ColumnNum  = Selected Column
'            RowNum     = Selected Row
'
' OUTPUT(S): None

'Set the text at the selected location
flexData(GridNumber).TextMatrix(RowNum, ColumnNum) = Value

End Property

Public Sub ExtractDataEvenXIntervals(GraphArray() As Variant)
'
'   PURPOSE: Extract the Data array into meaningful data
'
'  INPUT(S): GraphArray()
'
'            GraphArray(0,0) = Main Title    (e.g. Output #1)
'            GraphArray(0,1) = System        (Series inforamtion)
'            GraphArray(0,2) = x-axis label  (Caption for x-axis)
'            GraphArray(0,3) = y-axis label  (Caption for y-axis)
'            GraphArray(0,4) = x-start       (Where the x-axis starts on the graph)
'            GraphArray(0,5) = x-stop        (Where the x-axis stops on the graph)
'            GraphArray(0,6) = y-high        (Highest value on the y-axis to graph)
'            GraphArray(0,7) = y-low         (Lowest value on the y-axis to graph)
'            GraphArray(0,8) = evaluate start(start of data evaluation)
'            GraphArray(0,9) = evaluate stop (stop of data evaluation)
'            GraphArray(0,10) = increment    (x increment value of graph)
'            GraphArray(0,11) = Sub Title    (e.g. Voltage Gradient)
'            GraphArray(0,12) = part #       (Current part number)
'            GraphArray(0,13) = Number of Data Graphs to graph
'            GraphArray(0,14) = Graph first output name
'            GraphArray(0,15) = Graph second output name
'
'            GraphArray(1,0...end) = y-axis data array
'            GraphArray(2,0...end) = high limit data array
'            GraphArray(3,0...end) = low limit data array

'            GraphArray(4,0) = Main Title    (e.g. Output #1)
'            GraphArray(4,1) = Sub Title     (e.g. Single Point Linearity)
'            GraphArray(4,2) = x-axis label  (Caption for x-axis)
'            GraphArray(4,3) = y-axis label  (Caption for y-axis)
'            GraphArray(4,4) = x-start       (Where the x-axis starts on the graph)
'            GraphArray(4,5) = x-stop        (Where the x-axis stops on the graph)
'            GraphArray(4,6) = y-high        (Highest value on the y-axis to graph)
'            GraphArray(4,7) = y-low         (Lowest value on the y-axis to graph)
'            GraphArray(4,8) = evaluate start(start of data evaluation)
'            GraphArray(4,9) = evaluate stop (stop of data evaluation)
'            GraphArray(4,10) = increment    (x increment value of graph)
'            GraphArray(4,11) = Series info  (Series information)
'            GraphArray(4,12) = part #       (Current part number)
'            GraphArray(4,13) = Number of Data Graphs to graph
'            GraphArray(4,14) = Graph first output name
'            GraphArray(4,15) = Graph second output name
'
'            GraphArray(5,0...end) = y-axis data array
'            GraphArray(6,0...end) = high limit data array
'            GraphArray(7,0...end) = low limit data array
'
' OUTPUT(S): None...

Dim lintGraphNum As Integer
Dim lintPointer As Integer
Dim lintNumberOfSubsets As Integer
Dim llngNumberOfDataPoints As Long
Dim lsngDataY() As Single
Dim lsngDataX() As Single
Dim lintCurrentDataPoint As Integer
Dim lintCurrentSubset As Integer

'Initialize the Pointer
lintPointer = 0

For lintGraphNum = 0 To 99
    'If there's no data in the first position, we're done graphing
    If GraphArray(lintPointer, 0) = 0 Then Exit For
    'Set the properties for the current graph
    Call SetGraphProperties(lintGraphNum)
    'Add a button to the graph toolbar for the current graph
    tlbGraphs.Buttons(1).ButtonMenus.Add
    If pesGraph1(lintGraphNum).Visible = True Then
        'Add * to Text to show graph is selected
        tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text = "*" & GraphArray(lintPointer, 0) & " " & GraphArray(lintPointer, 11)
    Else
        'Show the Text of the non-selected graph
        tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text = GraphArray(lintPointer, 0) & " " & GraphArray(lintPointer, 11)
    End If

    'Set the Main Title
    pesGraph1(lintGraphNum).MainTitle = GraphArray(lintPointer, 11) & "  (" & GraphArray(lintPointer, 0) & ")"
    'Show the date/time as the subtitle
    pesGraph1(lintGraphNum).subTitle = Now & "                              " & GraphArray(lintPointer, 1) & "                                   " & "Part # " & GraphArray(lintPointer, 12)
    'Display the X-axis label
    pesGraph1(lintGraphNum).XAxisLabel = GraphArray(lintPointer, 2)
    'Display the Y-axis label
    pesGraph1(lintGraphNum).YAxisLabel = GraphArray(lintPointer, 3)

    'Number of data lines (subsets) per graph
    lintNumberOfSubsets = GraphArray(lintPointer, 13) + 2
    If lintNumberOfSubsets <= 2 Then
        'Graph three lines on the graph(yData, HighLimit, LowLimit)
        pesGraph1(lintGraphNum).Subsets = 3
        lintNumberOfSubsets = 3
    Else
        'Graph the number of outputs specified by user
        pesGraph1(lintGraphNum).Subsets = lintNumberOfSubsets
    End If

    'Dimension data arrays to the necessary bounds:
    'Bounds of the second dimension of the incoming array * number of subsets
    ReDim lsngDataX(LBound(GraphArray, 2) To UBound(GraphArray, 2) * lintNumberOfSubsets)
    ReDim lsngDataY(LBound(GraphArray, 2) To UBound(GraphArray, 2) * lintNumberOfSubsets)

    'Set up the X and Y limits
    pesGraph1(lintGraphNum).RYAxisComparisonSubsets = 0
    pesGraph1(lintGraphNum).ManualMinX = GraphArray(lintPointer, 4)
    pesGraph1(lintGraphNum).ManualMaxX = GraphArray(lintPointer, 5)
    pesGraph1(lintGraphNum).ManualMaxY = GraphArray(lintPointer, 6)
    pesGraph1(lintGraphNum).ManualMinY = GraphArray(lintPointer, 7)

    llngNumberOfDataPoints = (GraphArray(lintPointer, 9) - GraphArray(lintPointer, 8)) / GraphArray(lintPointer, 10)
    pesGraph1(lintGraphNum).Points = llngNumberOfDataPoints
    
    'Extract the Y data and calculate X data
    For lintCurrentDataPoint = 0 To llngNumberOfDataPoints - 1
        lsngDataY(lintCurrentDataPoint) = GraphArray(lintPointer + 1, lintCurrentDataPoint)
        lsngDataX(lintCurrentDataPoint) = (GraphArray(lintPointer, 10) * lintCurrentDataPoint) + GraphArray(lintPointer, 8)
        lsngDataY(lintCurrentDataPoint + llngNumberOfDataPoints) = GraphArray(lintPointer + 2, lintCurrentDataPoint)
        lsngDataX(lintCurrentDataPoint + llngNumberOfDataPoints) = (GraphArray(lintPointer, 10) * lintCurrentDataPoint) + GraphArray(lintPointer, 8)
        lsngDataY(lintCurrentDataPoint + llngNumberOfDataPoints + llngNumberOfDataPoints) = GraphArray(lintPointer + 3, lintCurrentDataPoint)
        lsngDataX(lintCurrentDataPoint + llngNumberOfDataPoints + llngNumberOfDataPoints) = (GraphArray(lintPointer, 10) * lintCurrentDataPoint) + GraphArray(lintPointer, 8)
    Next lintCurrentDataPoint
    
    'Set the line colors for High/Low/Data Subset 1
    pesGraph1(lintGraphNum).SubsetColors(0) = QBColor(1)           'Output - blue line
    pesGraph1(lintGraphNum).SubsetLineTypes(0) = PELT_MEDIUMSOLID  'Output - thin line
    pesGraph1(lintGraphNum).SubsetColors(1) = vbRed                'High limit - red line
    pesGraph1(lintGraphNum).SubsetLineTypes(1) = PELT_THICKSOLID   'High limit - thick line
    pesGraph1(lintGraphNum).SubsetColors(2) = vbRed                'Low limit - red line
    pesGraph1(lintGraphNum).SubsetLineTypes(2) = PELT_THICKSOLID   'Low limit - thick line

    'Fill multiple graphs
    If lintNumberOfSubsets >= 4 Then
        For lintCurrentSubset = 4 To lintNumberOfSubsets
            For lintCurrentDataPoint = 0 To llngNumberOfDataPoints - 1
                lsngDataY(lintCurrentDataPoint + (lintCurrentSubset - 1) * llngNumberOfDataPoints) = GraphArray(lintPointer + 1, lintCurrentDataPoint + ((lintCurrentSubset - 3) * llngNumberOfDataPoints))
                lsngDataX(lintCurrentDataPoint + (lintCurrentSubset - 1) * llngNumberOfDataPoints) = (GraphArray(lintPointer, 10) * lintCurrentDataPoint) + GraphArray(lintPointer, 8)
            Next lintCurrentDataPoint
            pesGraph1(lintGraphNum).SubsetsToLegend(0) = 0             'Show first subset
            pesGraph1(lintGraphNum).SubsetsToLegend(lintCurrentSubset - 3) = lintCurrentSubset - 1
            pesGraph1(lintGraphNum).SubsetLineTypes(lintCurrentSubset - 1) = PELT_MEDIUMSOLID
            pesGraph1(lintGraphNum).SubsetColors(lintCurrentSubset - 1) = QBColor(lintCurrentSubset - 2)
            pesGraph1(lintGraphNum).SubsetLabels(0) = GraphArray(lintPointer, 14)
            pesGraph1(lintGraphNum).SubsetLabels(lintCurrentSubset - 1) = GraphArray(lintPointer, lintCurrentSubset + 11)
        Next lintCurrentSubset
    Else
        pesGraph1(lintGraphNum).SubsetsToLegend(0) = 3                 'Do not display subset legends
    End If

    'Load the X & Y data arrays into the graph object:
    Call PEvset(pesGraph1(lintGraphNum), PEP_faYDATA, lsngDataY(0), llngNumberOfDataPoints * lintNumberOfSubsets)
    Call PEvset(pesGraph1(lintGraphNum), PEP_faXDATA, lsngDataX(0), llngNumberOfDataPoints * lintNumberOfSubsets)
    
    'Draw the image and show it:
    Call PEreinitialize(pesGraph1(lintGraphNum))
    Call PEresetimage(pesGraph1(lintGraphNum), 0, 0)
    'Set cursor = arrow
    MousePointer = vbArrow

    'Set up the cursors
    pesGraph1(lintGraphNum).CursorMode = PECM_DATACROSS
    pesGraph1(lintGraphNum).MouseCursorControl = True
    pesGraph1(lintGraphNum).CursorPromptStyle = PECPS_XYVALUES
    'Redraw
    pesGraph1(lintGraphNum).PEactions = 3

    'Increment the pointer for the next graph
    lintPointer = lintPointer + 4
Next lintGraphNum

End Sub

Public Sub ExtractDataXAndY(GraphArray())
'
'   PURPOSE:
'
'  INPUT(S):
'
' OUTPUT(S): None...

Dim lintGraphNum As Integer
Dim lintPointer As Integer
Dim lintNumberOfSubsets As Integer
Dim llngNumberOfDataPoints As Long
Dim lsngDataY() As Single
Dim lsngDataX() As Single
Dim lintCurrentDataPoint As Integer
Dim lintCurrentSubset As Integer

lintPointer = 0

For lintGraphNum = 0 To 99
    'If there's no data in the first position, we're done graphing
    If GraphArray(lintPointer, 0) = 0 Then Exit For
    'Set the properties for the current graph
    Call SetGraphProperties(lintGraphNum)
    'Add a button to the graph toolbar for the current graph
    tlbGraphs.Buttons(1).ButtonMenus.Add
    If pesGraph1(lintGraphNum).Visible = True Then
        'Add * to Text to show graph is selected
        tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text = "*" & GraphArray(lintPointer, 0) & " " & GraphArray(lintPointer, 1)
    Else
        'Show the Text of the non-selected graph
        tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text = GraphArray(lintPointer, 0) & " " & GraphArray(lintPointer, 1)
    End If

    'Set the Main Title
    pesGraph1(lintGraphNum).MainTitle = GraphArray(lintPointer, 11) & "  (" & GraphArray(lintPointer, 0) & ")"
    'Show the date/time as the subtitle
    pesGraph1(lintGraphNum).subTitle = Now & "                              " & GraphArray(lintPointer, 1) & "                                   " & "Part # " & GraphArray(lintPointer, 12)
    'Display the X-axis label
    pesGraph1(lintGraphNum).XAxisLabel = GraphArray(lintPointer, 2)
    'Display the Y-axis label
    pesGraph1(lintGraphNum).YAxisLabel = GraphArray(lintPointer, 3)

    'Number of data lines (subsets) per graph
    lintNumberOfSubsets = GraphArray(lintPointer, 13) + 2
    If lintNumberOfSubsets <= 2 Then
        'Graph three lines on the graph(yData, HighLimit, LowLimit)
        pesGraph1(lintGraphNum).Subsets = 3
        lintNumberOfSubsets = 3
    Else
        'Graph the number of outputs specified by user
        pesGraph1(lintGraphNum).Subsets = lintNumberOfSubsets
    End If

    'Dimension data arrays to the necessary bounds:
    'Bounds of the second dimension of the incoming array * number of subsets
    ReDim lsngDataX(LBound(GraphArray, 2) To UBound(GraphArray, 2) * lintNumberOfSubsets)
    ReDim lsngDataY(LBound(GraphArray, 2) To UBound(GraphArray, 2) * lintNumberOfSubsets)

    'Set up the X and Y limits
    pesGraph1(lintGraphNum).RYAxisComparisonSubsets = 0
    pesGraph1(lintGraphNum).ManualMinX = GraphArray(lintPointer, 4)
    pesGraph1(lintGraphNum).ManualMaxX = GraphArray(lintPointer, 5)
    pesGraph1(lintGraphNum).ManualMaxY = GraphArray(lintPointer, 6)
    pesGraph1(lintGraphNum).ManualMinY = GraphArray(lintPointer, 7)
    
    'Calculate the number of data points
    llngNumberOfDataPoints = (GraphArray(lintPointer, 9) - GraphArray(lintPointer, 8)) / GraphArray(lintPointer, 10)
    pesGraph1(lintGraphNum).Points = llngNumberOfDataPoints
    
    'Extract X and Y Data
    For lintCurrentDataPoint = 0 To llngNumberOfDataPoints - 1
        lsngDataX(lintCurrentDataPoint) = GraphArray(lintPointer + 1, lintCurrentDataPoint)
        lsngDataY(lintCurrentDataPoint) = GraphArray(lintPointer + 2, lintCurrentDataPoint)
        lsngDataX(lintCurrentDataPoint + llngNumberOfDataPoints) = GraphArray(lintPointer + 1, lintCurrentDataPoint)
        lsngDataY(lintCurrentDataPoint + llngNumberOfDataPoints) = GraphArray(lintPointer + 3, lintCurrentDataPoint)
        lsngDataX(lintCurrentDataPoint + llngNumberOfDataPoints + llngNumberOfDataPoints) = GraphArray(lintPointer + 1, lintCurrentDataPoint)
        lsngDataY(lintCurrentDataPoint + llngNumberOfDataPoints + llngNumberOfDataPoints) = GraphArray(lintPointer + 4, lintCurrentDataPoint)
    Next lintCurrentDataPoint

    'Set the line colors for High/Low/Data Subset 1
    pesGraph1(lintGraphNum).SubsetColors(0) = QBColor(1)           'Output - blue line
    pesGraph1(lintGraphNum).SubsetLineTypes(0) = PELT_MEDIUMSOLID  'Output - thin line
    pesGraph1(lintGraphNum).SubsetColors(1) = vbRed                'High limit - red line
    pesGraph1(lintGraphNum).SubsetLineTypes(1) = PELT_THICKSOLID   'High limit - thick line
    pesGraph1(lintGraphNum).SubsetColors(2) = vbRed                'Low limit - red line
    pesGraph1(lintGraphNum).SubsetLineTypes(2) = PELT_THICKSOLID   'Low limit - thick line

    'Fill multiple graphs
    If lintNumberOfSubsets >= 4 Then
        For lintCurrentSubset = 4 To lintNumberOfSubsets
            For lintCurrentDataPoint = 0 To llngNumberOfDataPoints - 1
                lsngDataX(lintCurrentDataPoint + (lintCurrentSubset - 1) * llngNumberOfDataPoints) = GraphArray(lintPointer + 1, lintCurrentDataPoint + ((lintCurrentSubset - 3) * llngNumberOfDataPoints))
                lsngDataY(lintCurrentDataPoint + (lintCurrentSubset - 1) * llngNumberOfDataPoints) = GraphArray(lintPointer + 2, lintCurrentDataPoint + ((lintCurrentSubset - 3) * llngNumberOfDataPoints))
            Next lintCurrentDataPoint
            pesGraph1(lintGraphNum).SubsetsToLegend(0) = 0             'Show first subset
            pesGraph1(lintGraphNum).SubsetsToLegend(lintCurrentSubset - 3) = lintCurrentSubset - 1
            pesGraph1(lintGraphNum).SubsetLineTypes(lintCurrentSubset - 1) = PELT_MEDIUMSOLID
            pesGraph1(lintGraphNum).SubsetColors(lintCurrentSubset - 1) = QBColor(lintCurrentSubset - 2)
            pesGraph1(lintGraphNum).SubsetLabels(0) = GraphArray(lintPointer, 14)
            pesGraph1(lintGraphNum).SubsetLabels(lintCurrentSubset - 1) = GraphArray(lintPointer, lintCurrentSubset + 11)
        Next lintCurrentSubset
    Else
        pesGraph1(lintGraphNum).SubsetsToLegend(0) = 3                 'Do not display subset legends
    End If

    'Load the X & Y data arrays into the graph object:
    Call PEvset(pesGraph1(lintGraphNum), PEP_faYDATA, lsngDataY(0), llngNumberOfDataPoints * lintNumberOfSubsets)
    Call PEvset(pesGraph1(lintGraphNum), PEP_faXDATA, lsngDataX(0), llngNumberOfDataPoints * lintNumberOfSubsets)

    'Draw the image and show it:
    Call PEreinitialize(pesGraph1(lintGraphNum))
    Call PEresetimage(pesGraph1(lintGraphNum), 0, 0)
    'Set cursor = arrow
    MousePointer = vbArrow

    'Set up the cursors
    pesGraph1(lintGraphNum).CursorMode = PECM_DATACROSS
    pesGraph1(lintGraphNum).MouseCursorControl = True
    pesGraph1(lintGraphNum).CursorPromptStyle = PECPS_XYVALUES
    'Redraw
    pesGraph1(lintGraphNum).PEactions = 3

    'Increment the pointer for the next graph
    lintPointer = lintPointer + 5
Next lintGraphNum

End Sub

Public Property Get NumberOfColumns(GridNumber As Integer) As Long
'
'   PURPOSE: To return the number of columns in the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'
' OUTPUT(S): None

'Return the new number of columns
NumberOfColumns = flexData(GridNumber).Cols

End Property

Public Property Let NumberOfColumns(GridNumber As Integer, NumColumns As Long)
'
'   PURPOSE: To set the number of columns in the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            NumColumns = New Number of Columns in Grid
'
' OUTPUT(S): None

'Set the new number of columns
flexData(GridNumber).Cols = NumColumns

End Property

Public Property Get NumberOfRows(GridNumber As Integer) As Long
'
'   PURPOSE: To return the number of rows in the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'
' OUTPUT(S): None

'Return the new number of rows
NumberOfRows = flexData(GridNumber).Rows

End Property

Public Property Let NumberOfRows(GridNumber As Integer, NumRows As Long)
'
'   PURPOSE: To set the number of rows in the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            NumRows    = New Number of Rows in Grid
'
' OUTPUT(S): None

'Set the new number of rows
flexData(GridNumber).Rows = NumRows

End Property

Public Sub PrintAllGraphsInWindow()
'
'   PURPOSE: Print Graphs which apprear in graph window
'
'  INPUT(S): None
'
' OUTPUT(S): None
           
Dim lintGraphNum As Integer
Dim lintVisibleGraphNum As Integer
Dim lintNumberOfGraphs As Integer
Dim lintVisibleGraphs(0 To 15) As Integer
Dim lintLeft As Integer
Dim lintTop As Integer
Dim lintWidth As Integer
Dim lintHeight As Integer
Dim llngMeta As Long
Dim llngOldMapMode As Long
Dim pt As POINTSTRUCT
Dim lstrName As String '3.6ANM \/\/
Dim lstrPN As String
Dim THandle As Long
Dim iret As Long       '3.6ANM /\/\

On Error GoTo ErrGraph:

'Set printer quality to 200 dpi to speed up printing
Printer.PrintQuality = 200

'Set the Printer Orientation
If (mintOrientation = 1) Or (mintOrientation = 2) Then
    Printer.Orientation = mintOrientation
Else
    Printer.Orientation = 2     'Default before mintOrientation is set (Lanscape)
End If

UserControl.MousePointer = vbHourglass
Printer.Print                               'Send printer something so VB knows to start page
Printer.ScaleMode = 3                       'Set size of page
llngOldMapMode = SetMapMode(Printer.hDC, 8) 'Set mapping mode MM_ANSIOTROPIC
llngOldMapMode = SetMapMode(Printer.hDC, 8) 'Set viewport org and extents

lintGraphNum = 0
'Count to see how many graphs are visible and keep track of which ones
Do
    If pesGraph1(lintGraphNum).Visible = True Then
        lintNumberOfGraphs = lintNumberOfGraphs + 1
        lintVisibleGraphs(lintVisibleGraphNum) = lintGraphNum
        lintVisibleGraphNum = lintVisibleGraphNum + 1
    End If
    lintGraphNum = lintGraphNum + 1
Loop Until lintGraphNum >= 16  'Maximum Number of Graphs = 16
lintVisibleGraphNum = 0

Select Case lintNumberOfGraphs

    Case 0
        MsgBox "I can't read your mind.  If you would like to print a graph you must select one from the Pull Down Tool Bar...", vbInformation, "What are you thinking?"
    Case 1  'send 1 graph per page
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
    Case 2  'send 2 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight / 2
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
    Case 3  'send 3 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight / 3
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight / (3 / 2)
        lintWidth = Printer.ScaleWidth: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
    Case 4  'send 4 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight / 2
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0 + Printer.ScaleWidth / 2: lintTop = 0
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #4
        lintLeft = 0 + Printer.ScaleWidth / 2: lintTop = 0 + Printer.ScaleHeight / 2
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 2
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        
    Case 5, 6  'send 5 or 6 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight * 1 / 3
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0: lintTop = 0 + Printer.ScaleHeight * 2 / 3
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #4
        lintLeft = 0 + Printer.ScaleWidth / 2: lintTop = 0
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #5
        lintLeft = 0 + Printer.ScaleWidth / 2: lintTop = 0 + Printer.ScaleHeight * 1 / 3
        lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        
        If lintNumberOfGraphs = 6 Then
            'graph #6
            lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
            lintLeft = 0 + Printer.ScaleWidth / 2: lintTop = 0 + Printer.ScaleHeight * 2 / 3
            lintWidth = Printer.ScaleWidth / 2: lintHeight = Printer.ScaleHeight / 3
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
        End If
                
    Case 7, 8, 9 'send 7, 8, or 9 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + 1 * Printer.ScaleHeight / 3
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0: lintTop = 0 + 2 * Printer.ScaleHeight / 3
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #4
        lintLeft = 0 + Printer.ScaleWidth / 3: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #5
        lintLeft = 0 + Printer.ScaleWidth / 3: lintTop = 0 + 1 * Printer.ScaleHeight / 3
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #6
        lintLeft = 0 + Printer.ScaleWidth / 3: lintTop = 0 + 2 * Printer.ScaleHeight / 3
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #7
        lintLeft = 0 + 2 * Printer.ScaleWidth / 3: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
                
        'graph #8
        If lintNumberOfGraphs > 7 Then
            lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
            lintLeft = 0 + 2 * Printer.ScaleWidth / 3: lintTop = 0 + 1 * Printer.ScaleHeight / 3
            lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
        End If
        
        'graph #9
        If lintNumberOfGraphs > 8 Then
            lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
            lintLeft = 0 + 2 * Printer.ScaleWidth / 3: lintTop = 0 + 2 * Printer.ScaleHeight / 3
            lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 3
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
        End If
            
    Case 10, 11, 12 'send 10, 11, or 12 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + 1 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0: lintTop = 0 + 2 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #4
        lintLeft = 0: lintTop = 0 + 3 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #5
        lintLeft = 1 * Printer.ScaleWidth / 3: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #6
        lintLeft = 1 * Printer.ScaleWidth / 3: lintTop = 0 + 1 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #7
        lintLeft = 1 * Printer.ScaleWidth / 3: lintTop = 0 + 2 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #8
        lintLeft = 1 * Printer.ScaleWidth / 3: lintTop = 0 + 3 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #9
        lintLeft = 2 * Printer.ScaleWidth / 3: lintTop = 0
        lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #10
        If lintNumberOfGraphs > 9 Then
            lintLeft = 2 * Printer.ScaleWidth / 3: lintTop = 0 + 1 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
        End If
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #11
        If lintNumberOfGraphs > 10 Then
            lintLeft = 2 * Printer.ScaleWidth / 3: lintTop = 0 + 2 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
            lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        End If
        'graph #12
        If lintNumberOfGraphs > 11 Then
            lintLeft = 2 * Printer.ScaleWidth / 3: lintTop = 0 + 3 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 3: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
        End If
        
    Case 13, 14, 15, 16 'send 13, 14, 15, or 16 graphs per page
        'graph #1
        lintLeft = 0: lintTop = 0
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #2
        lintLeft = 0: lintTop = 0 + 1 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #3
        lintLeft = 0: lintTop = 0 + 2 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #4
        lintLeft = 0: lintTop = 0 + 3 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        
        'graph #5
        lintLeft = Printer.ScaleWidth / 4: lintTop = 0
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #6
        lintLeft = Printer.ScaleWidth / 4: lintTop = 0 + 1 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #7
        lintLeft = Printer.ScaleWidth / 4: lintTop = 0 + 2 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #8
        lintLeft = Printer.ScaleWidth / 4: lintTop = 0 + 3 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        
        'graph #9
        lintLeft = 2 * Printer.ScaleWidth / 4: lintTop = 0
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #10
        lintLeft = 2 * Printer.ScaleWidth / 4: lintTop = 0 + 1 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #11
        lintLeft = 2 * Printer.ScaleWidth / 4: lintTop = 0 + 2 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        'graph #12
        lintLeft = 2 * Printer.ScaleWidth / 4: lintTop = 0 + 3 * Printer.ScaleHeight / 4
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
                        
        'graph #13
        lintLeft = 3 * Printer.ScaleWidth / 4: lintTop = 0
        lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
        Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
        Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
        Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
        llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
        Call PlayMetaFile(Printer.hDC, llngMeta)
        lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        'graph #14
        If lintNumberOfGraphs > 13 Then
            lintLeft = 3 * Printer.ScaleWidth / 4: lintTop = 0 + 1 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
            lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        End If
        'graph #15
        If lintNumberOfGraphs > 14 Then
            lintLeft = 3 * Printer.ScaleWidth / 4: lintTop = 0 + 2 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
            lintVisibleGraphNum = lintVisibleGraphNum + 1   'increment graph
        End If
        'graph #16
        If lintNumberOfGraphs > 15 Then
            lintLeft = 3 * Printer.ScaleWidth / 4: lintTop = 0 + 3 * Printer.ScaleHeight / 4
            lintWidth = Printer.ScaleWidth / 4: lintHeight = Printer.ScaleHeight / 4
            Call SetViewportOrgEx(Printer.hDC, lintLeft, lintTop, pt)
            Call SetViewportExtEx(Printer.hDC, lintWidth, lintHeight, pt)
            Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), lintWidth, lintHeight)
            llngMeta = PEgetmeta(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)))
            Call PlayMetaFile(Printer.hDC, llngMeta)
            lintVisibleGraphNum = lintVisibleGraphNum + 1       'increment graph
        End If

End Select

'Let the printer know the document is done
Printer.EndDoc

'3.6ANM \/\/
If frmMain.mnuFunctionAutoSavePDFs.Checked = True Then
    BlockInput True
    Call frmDAQIO.KillTime(3000)
    
    THandle = FindWindowPartial(PDFWINDOW, "*")
    If THandle = 0 Then
        Call frmDAQIO.KillTime(3000)
        THandle = FindWindowPartial("*page*", "*")
    End If
    iret = BringWindowToTop(THandle)
    
    If THandle <> 0 Then
        lstrPN = gstrLotName
        lstrName = gstrSerialNumber & gstrType & Format(Now, "MM-DD-YY HHMMSSAMPM") & ".pdf"
        
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

'Reset mapping mode
Call SetMapMode(Printer.hDC, llngOldMapMode)

'Reset image to current aspect ratio
Call PEresetimage(pesGraph1(lintVisibleGraphs(lintVisibleGraphNum)), 0, 0)

'Reset the MousePointer
UserControl.MousePointer = vbNormal

Exit Sub

ErrGraph:
    BlockInput False '3.6ANM
End Sub

Public Property Let RowAlignment(GridNumber As Integer, RowNum As Long, ByVal Alignment As AlignmentSettings)
'
'   PURPOSE: To get the cell alignment property for the selected row.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum  = Selected Column
'            Alignment  = New Alignment setting
'
' OUTPUT(S): None

Dim llngColumnNum As Long

'Only proceed if the selected row is valid
If RowNum <= flexData(GridNumber).Rows Then
    'Set the active row
    flexData(GridNumber).Row = RowNum
    For llngColumnNum = 0 To flexData(GridNumber).Cols - 1
        flexData(GridNumber).Col = llngColumnNum
        flexData(GridNumber).CellAlignment = Alignment
    Next llngColumnNum
End If

End Property

Public Property Get RowHeight(GridNumber As Integer, RowNum As Long) As Long
'
'   PURPOSE: To get the height of the selected row.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'
' OUTPUT(S): None

'Set the RowHeight property
RowHeight = flexData(GridNumber).RowHeight(RowNum)

End Property

Public Property Let RowHeight(GridNumber As Integer, RowNum As Long, RowHeightInTwips As Long)
'
'   PURPOSE: To set the height of the selected row.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            RowHeightInTwips = New Row height, in Twips
'
' OUTPUT(S): None

'Set the row height
flexData(GridNumber).RowHeight(RowNum) = RowHeightInTwips

End Property

Private Sub SetGraphProperties(GraphNum)
'
'   PURPOSE: Set Graph Properties
'
'  INPUT(S): GraphNum = Graph Number to set properties for
'
' OUTPUT(S): None

Call PEreset(pesGraph1(GraphNum))

'Setup the non-default graph object parameters:
pesGraph1(GraphNum).AllowAnnotationControl = True               'Enables items to be added to pop-up menu
pesGraph1(GraphNum).AllowDataHotSpots = True                    'Enables data hot spots
pesGraph1(GraphNum).AllowGraphHotSpots = True                   'Enables graph Hot Spots
pesGraph1(GraphNum).AllowVertLineAnnotHotSpots = True           'Enables vert. line annotation Hot-Spots
pesGraph1(GraphNum).AllowZooming = PEAZ_HORZANDVERT             'Enables horz. & vert. zooming
pesGraph1(GraphNum).CursorPromptStyle = PECPS_XYVALUES          'Display x & y values of cursor
pesGraph1(GraphNum).CursorPromptTracking = True                 'Enables tracking of mouse to cursor
pesGraph1(GraphNum).DefOrientation = PEDO_LANDSCAPE             'Sets orientation of printout
pesGraph1(GraphNum).FontSize = PEFS_SMALL                       'Sets font size
pesGraph1(GraphNum).GraphForeColor = &H40&                      'Sets foreground color
pesGraph1(GraphNum).GridLineControl = PEGLC_YAXIS               'Sets horizontal grid lines only
pesGraph1(GraphNum).HotSpotSize = PEHSS_LARGE                   'Sets the size of the hot spot locator
pesGraph1(GraphNum).LabelBold = True                            'Sets all labels in BOLD
pesGraph1(GraphNum).MainTitleFont = "Arial"                     'Sets font used in main title
pesGraph1(GraphNum).ManualScaleControlX = PEMSC_MINMAX          'Enables user control of x-axis scale
pesGraph1(GraphNum).ManualScaleControlY = PEMSC_MINMAX          'Enables user control of y-axis scale
pesGraph1(GraphNum).NullDataValueY = -9999                      'Data equal to this won't get plotted
pesGraph1(GraphNum).NullDataValueX = -9999                      'Data equal to this won't get plotted
pesGraph1(GraphNum).PointSize = PEPS_SMALL                      'Sets size of plot points
pesGraph1(GraphNum).PrepareImages = True                        'Prepare images in memory, not on screen
pesGraph1(GraphNum).RYAxisScaleControl = PEAC_NORMAL            'Set right y-axis grid scale to linear
pesGraph1(GraphNum).ScrollingHorzZoom = True                    'Enables horz. panning after zoom
pesGraph1(GraphNum).ShadowColor = &HFFFFFF                      'Same as background color, so...no shadow
pesGraph1(GraphNum).ShowVertLineAnnotations = True              'Always show vertical line annotations
pesGraph1(GraphNum).SubTitleFont = "Arial"                      'Sets font used in sub title
pesGraph1(GraphNum).XAxisScaleControl = PEAC_NORMAL             'Set x-axis grid scale to linear
pesGraph1(GraphNum).YAxisScaleControl = PEAC_NORMAL             'Set y-axis grid scale to linear
pesGraph1(GraphNum).ZoomInterfaceOnly = PEZIO_NORMAL            'Enables built-in zooming
pesGraph1(GraphNum).VertLineAnnotationType(0) = PELT_THINSOLID  'Sets vertical annotation line type
pesGraph1(GraphNum).VertLineAnnotationColor(0) = vbRed          'Sets line annotation color (red)

End Sub

Public Property Get TabName(TabNum As Integer) As String
'
'   PURPOSE: To return the Caption of the selected Tab
'
'  INPUT(S): TabNum  = Selected Tab
'
' OUTPUT(S): None

'Verify that there is a Tab(TabNum)
If TabNum >= 0 And TabNum <= tabResults.Tabs - 1 Then
    'Return the Caption
    TabName = tabResults.TabCaption(TabNum)
End If

End Property

Public Property Let TabName(TabNum As Integer, ByVal Name As String)
'
'   PURPOSE: To set the Caption of the selected Tab
'
'  INPUT(S): TabNum  = Selected Tab
'            Name    = New Caption for the selected Tab
'
' OUTPUT(S): None

'Verify that there is a Tab(TabNum)
If TabNum >= 0 And TabNum <= tabResults.Tabs - 1 Then
     'Set the new Caption
    tabResults.TabCaption(TabNum) = Name
End If

End Property

Public Property Get TextAlignment(GridNumber As Integer, RowNum As Long, ColumnNum As Long) As AlignmentSettings
'
'   PURPOSE: To return the Alignment of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Return the alignment of the selected cell
TextAlignment = flexData(GridNumber).CellAlignment
    
End Property

Public Property Let TextAlignment(GridNumber As Integer, RowNum As Long, ColumnNum As Long, ByVal Alignment As AlignmentSettings)
'
'   PURPOSE: To set the Alignment of the selected cell.
'
'  INPUT(S): GridNumber = Selected Grid
'            RowNum     = Selected Row
'            ColumnNum  = Selected Column
'
' OUTPUT(S): None

'Select the Row & Column
flexData(GridNumber).Row = RowNum
flexData(GridNumber).Col = ColumnNum

'Set the alignment of the selected cell
flexData(GridNumber).CellAlignment = Alignment
    
End Property

Public Property Let TotalCellFont(GridNumber As Integer, ByVal FontName As String)
'
'   PURPOSE: To set the cell font of all cells on the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            FontName   = New Cell Font
'
' OUTPUT(S): None

Dim llngRowNum As Long
Dim llngColumnNum As Long

'Loop through all rows
For llngRowNum = 0 To flexData(GridNumber).Rows - 1
    'Loop through all columns
    For llngColumnNum = 0 To flexData(GridNumber).Cols - 1
        flexData(GridNumber).Row = llngRowNum
        flexData(GridNumber).Col = llngColumnNum
        flexData(GridNumber).CellFontName = FontName
    Next llngColumnNum
Next llngRowNum

End Property

Public Property Let TotalCellFontSize(GridNumber As Integer, ByVal FontSize As Integer)
'
'   PURPOSE: To set the cell font of all cells on the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            FontSize   = New Cell Font Size
'
' OUTPUT(S): None

Dim llngRowNum As Long
Dim llngColumnNum As Long

'Loop through all rows
For llngRowNum = 0 To flexData(GridNumber).Rows - 1
    'Loop through all columns
    For llngColumnNum = 0 To flexData(GridNumber).Cols - 1
        flexData(GridNumber).Row = llngRowNum
        flexData(GridNumber).Col = llngColumnNum
        flexData(GridNumber).CellFontSize = FontSize
    Next llngColumnNum
Next llngRowNum

End Property

Public Property Let TotalRowHeight(GridNumber As Integer, HeightInTwips As Long)
'
'   PURPOSE: To set the row height of all rows for the selected grid.
'
'  INPUT(S): GridNumber    = Selected Grid
'            HeightInTwips = New Row Height, in twips.
'
' OUTPUT(S): None

Dim llngRowNum As Long

'Loop through all rows setting the height
For llngRowNum = 0 To NumberOfRows(GridNumber) - 1
    flexData(GridNumber).RowHeight(llngRowNum) = HeightInTwips
Next llngRowNum

End Property

Public Property Let TotalWordWrap(GridNumber As Integer, WordWrap As Boolean)
'
'   PURPOSE: To set the WordWrap Property for the selected grid.
'
'  INPUT(S): GridNumber = Selected Grid
'            WordWrap   = Whether or not to WordWrap
'
' OUTPUT(S): None

Dim llngRowNum As Long
Dim llngColumnNum As Long

'Loop through all rows
For llngRowNum = 0 To flexData(GridNumber).Rows - 1
    'Loop through all columns
    For llngColumnNum = 0 To flexData(GridNumber).Cols - 1
        flexData(GridNumber).Row = llngRowNum
        flexData(GridNumber).Col = llngColumnNum
        flexData(GridNumber).WordWrap = WordWrap
    Next llngColumnNum
Next llngRowNum

End Property

Private Sub cmdDefGraph_Click()
'
'   PURPOSE: Allow user to quickly select Default Graphs
'
'  INPUT(S): None
'
' OUTPUT(S): None

Dim lintTextLength As Integer
Dim X As Integer

If gblnGraphEnable And gblnGraphsLoaded Then
    'Clear all graphs
    For X = 0 To 15
        If pesGraph1(X).Visible = True Then
            'If it's already visible, make it invisible
            pesGraph1(X).Visible = False
            lintTextLength = Len(tlbGraphs.Buttons(1).ButtonMenus(X + 1).Text)
            'Remove * to Text to show graph is not selected
            tlbGraphs.Buttons(1).ButtonMenus(X + 1).Text = Mid(tlbGraphs.Buttons(1).ButtonMenus(X + 1).Text, 2, lintTextLength)
        End If
    Next X
    
    'Select default graphs
    'Make graphs visible
    pesGraph1(1).Visible = True
    pesGraph1(6).Visible = True
    pesGraph1(10).Visible = True
    pesGraph1(12).Visible = True
        
    'Add * to Text to show graph is selected
    tlbGraphs.Buttons(1).ButtonMenus(2).Text = "*" & tlbGraphs.Buttons(1).ButtonMenus(2).Text
    tlbGraphs.Buttons(1).ButtonMenus(7).Text = "*" & tlbGraphs.Buttons(1).ButtonMenus(7).Text
    tlbGraphs.Buttons(1).ButtonMenus(11).Text = "*" & tlbGraphs.Buttons(1).ButtonMenus(11).Text
    tlbGraphs.Buttons(1).ButtonMenus(13).Text = "*" & tlbGraphs.Buttons(1).ButtonMenus(13).Text
    
    'Center the four graphs
    'Graph 1
    pesGraph1(1).left = GRAPHLEFT
    pesGraph1(1).Width = TOTALGRAPHWIDTH / 2
    pesGraph1(1).top = GRAPHTOP
    pesGraph1(1).Height = TOTALGRAPHHEIGHT / 2
    'Graph 2
    pesGraph1(6).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
    pesGraph1(6).Width = TOTALGRAPHWIDTH / 2
    pesGraph1(6).top = GRAPHTOP
    pesGraph1(6).Height = TOTALGRAPHHEIGHT / 2
    'Graph 3
    pesGraph1(10).left = GRAPHLEFT
    pesGraph1(10).Width = TOTALGRAPHWIDTH / 2
    pesGraph1(10).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 2)
    pesGraph1(10).Height = TOTALGRAPHHEIGHT / 2
    'Graph 4
    pesGraph1(12).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
    pesGraph1(12).Width = TOTALGRAPHWIDTH / 2
    pesGraph1(12).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 2)
    pesGraph1(12).Height = TOTALGRAPHHEIGHT / 2
Else
    MsgBox "Graphs not enabled or loaded", vbOKOnly, "Error displaying graphs"
End If

End Sub

'*** Events ***

Private Sub cmdPrintGraphs_Click()
'
'   PURPOSE: Allow user to Print All Visible Graphs
'
'  INPUT(S): None
'
' OUTPUT(S): None

gstrType = " Graph " '3.6ANM
Call PrintAllGraphsInWindow
    
End Sub

Private Sub optOrientation_Click(Index As Integer)
'
'   PURPOSE: Select the printing orientation (Landscape or Portrait)
'
'  INPUT(S): Index = 0 = Landscape
'                    1 = Portrait
'
' OUTPUT(S): None

Select Case Index
    Case 0
        mintOrientation = vbPRORLandscape
    Case 1
        mintOrientation = vbPRORPortrait
End Select

End Sub

Private Sub tlbGraphs_ButtonClick(ByVal Button As MSComctlLib.Button)
'
'   PURPOSE: All Graphs
'
'  INPUT(S): None
'
' OUTPUT(S): None

Dim lintGraphNum As Integer
Dim lintTextLength As Integer

If Button = "Clear All Graphs" Then
    For lintGraphNum = 0 To 15
        If pesGraph1(lintGraphNum).Visible = True Then
            pesGraph1(lintGraphNum).Visible = False
            lintTextLength = Len(tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text)
            'Remove * to Text to show graph is not selected
            tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text = Mid(tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum + 1).Text, 2, lintTextLength)
        End If
    Next lintGraphNum
End If

End Sub

Private Sub tlbGraphs_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'
'   PURPOSE: Update Graph Viewing based on the number of graphs selected
'
'  INPUT(S): ButtonMenu = Button pressed in Menu
'
' OUTPUT(S): None

Dim lintVisibleGraphNum As Integer
Dim lintGraphNum As Integer
Dim lintNumberOfGraphs As Integer
Dim lintVisibleGraphs(0 To 15) As Integer
Dim lintTextLength As Integer

'Loop checking for which button was pressed
For lintGraphNum = 1 To 16  'Maximum Number of Graphs = 16
    'See which button was pressed
    If tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text = ButtonMenu Then
        If pesGraph1(lintGraphNum - 1).Visible = True Then
            'If it's already visible, make it invisible
            pesGraph1(lintGraphNum - 1).Visible = False
            lintTextLength = Len(tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text)
            'Remove * to Text to show graph is not selected
            tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text = Mid(tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text, 2, lintTextLength)
        Else
            'If it's invisible, make it visible
            pesGraph1(lintGraphNum - 1).Visible = True
            'Add * to Text to show graph is selected
            tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text = "*" & tlbGraphs.Buttons(1).ButtonMenus(lintGraphNum).Text
        End If
        'Once we've found which button was pressed, exit the for loop
        Exit For
    End If
Next lintGraphNum

'Count to see how many graphs are visible and keep track of which ones.
For lintGraphNum = 0 To 15     'Maximum Number of Graphs = 16
    If pesGraph1(lintGraphNum).Visible = True Then
        lintNumberOfGraphs = lintNumberOfGraphs + 1
        'Track which Graphs are visible
        lintVisibleGraphs(lintVisibleGraphNum) = lintGraphNum
        lintVisibleGraphNum = lintVisibleGraphNum + 1
    End If
Next lintGraphNum

lintVisibleGraphNum = 0
Select Case lintNumberOfGraphs

    Case 1  'One Graph
        pesGraph1(lintVisibleGraphs(0)).left = GRAPHLEFT
        pesGraph1(lintVisibleGraphs(0)).Width = TOTALGRAPHWIDTH
        pesGraph1(lintVisibleGraphs(0)).top = GRAPHTOP
        pesGraph1(lintVisibleGraphs(0)).Height = TOTALGRAPHHEIGHT
    Case 2  'Two Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT
    Case 3  'Three Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP + (2 * (TOTALGRAPHHEIGHT / 3))
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 3
    Case 4 'Four Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 2
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 2
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 2)
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 2
        'Graph 4
        pesGraph1((lintVisibleGraphs(3))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(3))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(3))).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 2)
        pesGraph1((lintVisibleGraphs(3))).Height = TOTALGRAPHHEIGHT / 2
    Case 5, 6 'Five Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP + (2 * (TOTALGRAPHHEIGHT / 3))
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 4
        pesGraph1((lintVisibleGraphs(3))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(3))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(3))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(3))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 5
        pesGraph1((lintVisibleGraphs(4))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(4))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(4))).top = GRAPHTOP + (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(4))).Height = TOTALGRAPHHEIGHT / 3
        'If the Fifth Graph was the last Graph, Exit Sub
        If lintNumberOfGraphs = 5 Then Exit Sub
        'Graph 6
        pesGraph1((lintVisibleGraphs(5))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 2)
        pesGraph1((lintVisibleGraphs(5))).Width = TOTALGRAPHWIDTH / 2
        pesGraph1((lintVisibleGraphs(5))).top = GRAPHTOP + (2 * (TOTALGRAPHHEIGHT / 3))
        pesGraph1((lintVisibleGraphs(5))).Height = TOTALGRAPHHEIGHT / 3
    Case 7, 8, 9    'Seven, Eight, or Nine Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 4
        pesGraph1((lintVisibleGraphs(3))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(3))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(3))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(3))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 5
        pesGraph1((lintVisibleGraphs(4))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(4))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(4))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(4))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 6
        pesGraph1((lintVisibleGraphs(5))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(5))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(5))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(5))).Height = TOTALGRAPHHEIGHT / 3
        'Graph 7
        pesGraph1((lintVisibleGraphs(6))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(6))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(6))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(6))).Height = TOTALGRAPHHEIGHT / 3
        'If the Seventh Graph was the last Graph, Exit Sub
        If lintNumberOfGraphs = 7 Then Exit Sub
        'Graph 8
        pesGraph1((lintVisibleGraphs(7))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(7))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(7))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(7))).Height = TOTALGRAPHHEIGHT / 3
        'If the Eighth Graph was the last Graph, Exit Sub
        If lintNumberOfGraphs = 8 Then Exit Sub
        'Graph 9
        pesGraph1((lintVisibleGraphs(8))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(8))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(8))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 3)
        pesGraph1((lintVisibleGraphs(8))).Height = TOTALGRAPHHEIGHT / 3
    Case 10, 11, 12 'Ten, Eleven, or Twelve Graphs
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 4
        pesGraph1((lintVisibleGraphs(3))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(3))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(3))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(3))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 5
        pesGraph1((lintVisibleGraphs(4))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(4))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(4))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(4))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 6
        pesGraph1((lintVisibleGraphs(5))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(5))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(5))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(5))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 7
        pesGraph1((lintVisibleGraphs(6))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(6))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(6))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(6))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 8
        pesGraph1((lintVisibleGraphs(7))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(7))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(7))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(7))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 9
        pesGraph1((lintVisibleGraphs(8))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(8))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(8))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(8))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 10
        pesGraph1((lintVisibleGraphs(9))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(9))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(9))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(9))).Height = TOTALGRAPHHEIGHT / 4
        'If the Tenth graph was the last Graph, Exit Sub
        If lintNumberOfGraphs = 10 Then Exit Sub
        'Graph 11
        pesGraph1((lintVisibleGraphs(10))).left = GRAPHLEFT + 1 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(10))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(10))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(10))).Height = TOTALGRAPHHEIGHT / 4
        'If the Eleventh graph was the Last Graph, Exit Sub
        If lintNumberOfGraphs = 11 Then Exit Sub
        'Graph 12
        pesGraph1((lintVisibleGraphs(11))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 3)
        pesGraph1((lintVisibleGraphs(11))).Width = TOTALGRAPHWIDTH / 3
        pesGraph1((lintVisibleGraphs(11))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(11))).Height = TOTALGRAPHHEIGHT / 4
    Case 13, 14, 15, 16
        'Graph 1
        pesGraph1((lintVisibleGraphs(0))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(0))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(0))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(0))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 2
        pesGraph1((lintVisibleGraphs(1))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(1))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(1))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(1))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 3
        pesGraph1((lintVisibleGraphs(2))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(2))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(2))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(2))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 4
        pesGraph1((lintVisibleGraphs(3))).left = GRAPHLEFT + 3 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(3))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(3))).top = GRAPHTOP
        pesGraph1((lintVisibleGraphs(3))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 5
        pesGraph1((lintVisibleGraphs(4))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(4))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(4))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(4))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 6
        pesGraph1((lintVisibleGraphs(5))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(5))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(5))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(5))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 7
        pesGraph1((lintVisibleGraphs(6))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(6))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(6))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(6))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 8
        pesGraph1((lintVisibleGraphs(7))).left = GRAPHLEFT + 3 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(7))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(7))).top = GRAPHTOP + 1 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(7))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 9
        pesGraph1((lintVisibleGraphs(8))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(8))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(8))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(8))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 10
        pesGraph1((lintVisibleGraphs(9))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(9))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(9))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(9))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 11
        pesGraph1((lintVisibleGraphs(10))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(10))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(10))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(10))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 12
        pesGraph1((lintVisibleGraphs(11))).left = GRAPHLEFT + 3 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(11))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(11))).top = GRAPHTOP + 2 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(11))).Height = TOTALGRAPHHEIGHT / 4
        'Graph 13
        pesGraph1((lintVisibleGraphs(12))).left = GRAPHLEFT
        pesGraph1((lintVisibleGraphs(12))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(12))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(12))).Height = TOTALGRAPHHEIGHT / 4
        'If the Thirteenth graph was the Last Graph, Exit Sub
        If lintNumberOfGraphs = 13 Then Exit Sub
        'Graph 14
        pesGraph1((lintVisibleGraphs(13))).left = GRAPHLEFT + (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(13))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(13))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(13))).Height = TOTALGRAPHHEIGHT / 4
        'If the Fourteenth graph was the Last Graph, Exit Sub
        If lintNumberOfGraphs = 14 Then Exit Sub
        'Graph 15
        pesGraph1((lintVisibleGraphs(14))).left = GRAPHLEFT + 2 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(14))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(14))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(14))).Height = TOTALGRAPHHEIGHT / 4
        'If the Fifteenth graph was the Last Graph, Exit Sub
        If lintNumberOfGraphs = 15 Then Exit Sub
        'Graph 16
        pesGraph1((lintVisibleGraphs(15))).left = GRAPHLEFT + 3 * (TOTALGRAPHWIDTH / 4)
        pesGraph1((lintVisibleGraphs(15))).Width = TOTALGRAPHWIDTH / 4
        pesGraph1((lintVisibleGraphs(15))).top = GRAPHTOP + 3 * (TOTALGRAPHHEIGHT / 4)
        pesGraph1((lintVisibleGraphs(15))).Height = TOTALGRAPHHEIGHT / 4
End Select

End Sub

