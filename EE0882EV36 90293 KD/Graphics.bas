Attribute VB_Name = "Graphics"
Option Explicit

Global DemoDir As String
Global HelpDirAndName As String

'///////////////////////////////////////////////////////////////////
'// 32 BIT PEGRP32A API FUNCTIONS AND DEFINES //
'// For VB /32 bit                            //
'// Copyright (c) 1996-1999 Gigasoft, Inc.    //
'///////////////////////////////////////////////
'
'
' ** NOTE, this is a partial list of constants.
' If the constant you need can not be found here,
' look in PEGRPAPI.H which is the true header file
' for the DLL.
'

Global Const PESTA_CENTER = 0
Global Const PESTA_LEFT = 1
Global Const PESTA_RIGHT = 2

Global Const PEDO_DRIVERDEFAULT = 0
Global Const PEDO_LANDSCAPE = 1
Global Const PEDO_PORTRAIT = 2

Global Const PEVS_COLOR = 0
Global Const PEVS_MONO = 1
Global Const PEVS_MONOWITHSYMBOLS = 2

Global Const PEDP_NODECIMALS = 0
Global Const PEDP_ONEDECIMAL = 1
Global Const PEDP_TWODECIMALS = 2
Global Const PEDP_THREEDECIMALS = 3

Global Const PECPS_NONE = 0
Global Const PECPS_XVALUE = 1
Global Const PECPS_YVALUE = 2
Global Const PECPS_XYVALUES = 3

Global Const PEZIO_NORMAL = 0
Global Const PEZIO_RECT = 1
Global Const PEZIO_LINE = 2

Global Const PEFS_LARGE = 0
Global Const PEFS_MEDIUM = 1
Global Const PEFS_SMALL = 2

Global Const PEVB_NONE = 0
Global Const PEVB_TOP = 1
Global Const PEVB_BOTTOM = 2
Global Const PEVB_TOPANDBOTTOM = 3

Global Const PEAC_AUTO = 0
Global Const PEAC_NORMAL = 1
Global Const PEAC_LOG = 2

Global Const PEGLC_BOTH = 0
Global Const PEGLC_YAXIS = 1
Global Const PEGLC_XAXIS = 2
Global Const PEGLC_NONE = 3

Global Const PEAS_SUMPP = 51
Global Const PEAS_MINAP = 1
Global Const PEAS_MINPP = 52
Global Const PEAS_MAXAP = 2
Global Const PEAS_MAXPP = 53
Global Const PEAS_AVGAP = 3
Global Const PEAS_AVGPP = 54
Global Const PEAS_P1SDAP = 4
Global Const PEAS_P1SDPP = 55
Global Const PEAS_P2SDAP = 5
Global Const PEAS_P2SDPP = 56
Global Const PEAS_P3SDAP = 6
Global Const PEAS_P3SDPP = 57
Global Const PEAS_M1SDAP = 7
Global Const PEAS_M1SDPP = 58
Global Const PEAS_M2SDAP = 8
Global Const PEAS_M2SDPP = 59
Global Const PEAS_M3SDAP = 9
Global Const PEAS_M3SDPP = 60
Global Const PEAS_PARETO_ASC = 90
Global Const PEAS_PARETO_DEC = 91

Global Const PEPTGI_FIRSTPOINTS = 0
Global Const PEPTGI_LASTPOINTS = 1

Global Const PEPTGV_SEQUENTIAL = 0
Global Const PEPTGV_RANDOM = 1

Global Const PEGPT_GRAPH = 0
Global Const PEGPT_TABLE = 1
Global Const PEGPT_BOTH = 2

Global Const PETW_GRAPHED = 0
Global Const PETW_ALLSUBSETS = 1

Global Const PEDLT_PERCENTAGE = 0
Global Const PEDLT_VALUE = 1

Global Const PEMSC_NONE = 0
Global Const PEMSC_MIN = 1
Global Const PEMSC_MAX = 2
Global Const PEMSC_MINMAX = 3

Global Const PEHS_NONE = 0
Global Const PEHS_SUBSET = 1
Global Const PEHS_POINT = 2
Global Const PEHS_GRAPH = 3
Global Const PEHS_TABLE = 4
Global Const PEHS_DATAPOINT = 5
Global Const PEHS_ANNOTATION = 6
Global Const PEHS_XAXISANNOTATION = 7
Global Const PEHS_YAXISANNOTATION = 8
Global Const PEHS_HORZLINEANNOTATION = 9
Global Const PEHS_VERTLINEANNOTATION = 10
Global Const PEHS_MAINTITLE = 11
Global Const PEHS_SUBTITLE = 12
Global Const PEHS_MULTISUBTITLE = 13
Global Const PEHS_MULTIBOTTOMTITLE = 14
Global Const PEHS_YAXISLABEL = 15
Global Const PEHS_XAXISLABEL = 16
Global Const PEHS_YAXIS = 17
Global Const PEHS_XAXIS = 18

Global Const PESPM_NONE = 0
Global Const PESPM_HIGHLOWBAR = 1
Global Const PESPM_HIGHLOWLINE = 2
Global Const PESPM_HIGHLOWCLOSE = 3
Global Const PESPM_OPENHIGHLOWCLOSE = 4
Global Const PESPM_BOXPLOT = 5

Global Const PELT_THINSOLID = 0
Global Const PELT_DASH = 1
Global Const PELT_DOT = 2
Global Const PELT_DASHDOT = 3
Global Const PELT_DASHDOTDOT = 4
Global Const PELT_MEDIUMSOLID = 5
Global Const PELT_THICKSOLID = 6
Global Const PELAT_GRIDTICK = 7
Global Const PELAT_GRIDLINE = 8
Global Const PELT_MEDIUMTHINSOLID = 9
Global Const PELT_MEDIUMTHICKSOLID = 10
Global Const PELT_EXTRATHICKSOLID = 11

Global Const PEPS_SMALL = 0
Global Const PEPS_MEDIUM = 1
Global Const PEPS_LARGE = 2
Global Const PEPS_MICRO = 3

Global Const PEPT_PLUS = 0
Global Const PEPT_CROSS = 1
Global Const PEPT_DOT = 2
Global Const PEPT_DOTSOLID = 3
Global Const PEPT_SQUARE = 4
Global Const PEPT_SQUARESOLID = 5
Global Const PEPT_DIAMOND = 6
Global Const PEPT_DIAMONDSOLID = 7
Global Const PEPT_UPTRIANGLE = 8
Global Const PEPT_UPTRIANGLESOLID = 9
Global Const PEPT_DOWNTRIANGLE = 10
Global Const PEPT_DOWNTRIANGLESOLID = 11

Global Const PEADL_NONE = 0
Global Const PEADL_DATAVALUES = 1
Global Const PEADL_POINTLABELS = 2
Global Const PEADL_DATAPOINTLABELS = 3

Global Const PEAZ_NONE = 0
Global Const PEAZ_HORIZONTAL = 1
Global Const PEAZ_VERTICAL = 2
Global Const PEAZ_HORZANDVERT = 3

Global Const PEBFD_2ND = 0
Global Const PEBFD_3RD = 1
Global Const PEBFD_4TH = 2

Global Const PEBS_SMALL = 0
Global Const PEBS_MEDIUM = 1
Global Const PEBS_LARGE = 2

Global Const PECG_COARSE = 0
Global Const PECG_MEDIUM = 1
Global Const PECG_FINE = 2

Global Const PEAE_NONE = 0
Global Const PEAE_ALLSUBSETS = 1
Global Const PEAE_INDSUBSETS = 2

Global Const PECM_NOCURSOR = 0
Global Const PECM_POINT = 1
Global Const PECM_DATACROSS = 2
Global Const PECM_DATASQUARE = 3
Global Const PECM_FLOATINGY = 4
Global Const PECM_FLOATINGXY = 5

Global Const PEGAT_NOSYMBOL = 0
Global Const PEGAT_PLUS = 1
Global Const PEGAT_CROSS = 2
Global Const PEGAT_DOT = 3
Global Const PEGAT_DOTSOLID = 4
Global Const PEGAT_SQUARE = 5
Global Const PEGAT_SQUARESOLID = 6
Global Const PEGAT_DIAMOND = 7
Global Const PEGAT_DIAMONDSOLID = 8
Global Const PEGAT_UPTRIANGLE = 9
Global Const PEGAT_UPTRIANGLESOLID = 10
Global Const PEGAT_DOWNTRIANGLE = 11
Global Const PEGAT_DOWNTRIANGLESOLID = 12
Global Const PEGAT_SMALLPLUS = 13
Global Const PEGAT_SMALLCROSS = 14
Global Const PEGAT_SMALLDOT = 15
Global Const PEGAT_SMALLDOTSOLID = 16
Global Const PEGAT_SMALLSQUARE = 17
Global Const PEGAT_SMALLSQUARESOLID = 18
Global Const PEGAT_SMALLDIAMOND = 19
Global Const PEGAT_SMALLDIAMONDSOLID = 20
Global Const PEGAT_SMALLUPTRIANGLE = 21
Global Const PEGAT_SMALLUPTRIANGLESOLID = 22
Global Const PEGAT_SMALLDOWNTRIANGLE = 23
Global Const PEGAT_SMALLDOWNTRIANGLESOLID = 24
Global Const PEGAT_LARGEPLUS = 25
Global Const PEGAT_LARGECROSS = 26
Global Const PEGAT_LARGEDOT = 27
Global Const PEGAT_LARGEDOTSOLID = 28
Global Const PEGAT_LARGESQUARE = 29
Global Const PEGAT_LARGESQUARESOLID = 30
Global Const PEGAT_LARGEDIAMOND = 31
Global Const PEGAT_LARGEDIAMONDSOLID = 32
Global Const PEGAT_LARGEUPTRIANGLE = 33
Global Const PEGAT_LARGEUPTRIANGLESOLID = 34
Global Const PEGAT_LARGEDOWNTRIANGLE = 35
Global Const PEGAT_LARGEDOWNTRIANGLESOLID = 36

Global Const PEGAT_POINTER = 37

Global Const PEGAT_THINSOLIDLINE = 38
Global Const PEGAT_DASHLINE = 39
Global Const PEGAT_DOTLINE = 40
Global Const PEGAT_DASHDOTLINE = 41
Global Const PEGAT_DASHDOTDOTLINE = 42
Global Const PEGAT_MEDIUMSOLIDLINE = 43
Global Const PEGAT_THICKSOLIDLINE = 44
Global Const PEGAT_LINECONTINUE = 45
                                     
Global Const PEGAT_TOPLEFT = 46
Global Const PEGAT_BOTTOMRIGHT = 47

Global Const PEGAT_RECT_THIN = 48
Global Const PEGAT_RECT_DASH = 49
Global Const PEGAT_RECT_DOT = 50
Global Const PEGAT_RECT_DASHDOT = 51
Global Const PEGAT_RECT_DASHDOTDOT = 52
Global Const PEGAT_RECT_MEDIUM = 53
Global Const PEGAT_RECT_THICK = 54
Global Const PEGAT_RECT_FILL = 55

Global Const PEGAT_ROUNDRECT_THIN = 56
Global Const PEGAT_ROUNDRECT_DASH = 57
Global Const PEGAT_ROUNDRECT_DOT = 58
Global Const PEGAT_ROUNDRECT_DASHDOT = 59
Global Const PEGAT_ROUNDRECT_DASHDOTDOT = 60
Global Const PEGAT_ROUNDRECT_MEDIUM = 61
Global Const PEGAT_ROUNDRECT_THICK = 62
Global Const PEGAT_ROUNDRECT_FILL = 63

Global Const PEGAT_ELLIPSE_THIN = 64
Global Const PEGAT_ELLIPSE_DASH = 65
Global Const PEGAT_ELLIPSE_DOT = 66
Global Const PEGAT_ELLIPSE_DASHDOT = 67
Global Const PEGAT_ELLIPSE_DASHDOTDOT = 68
Global Const PEGAT_ELLIPSE_MEDIUM = 69
Global Const PEGAT_ELLIPSE_THICK = 70
Global Const PEGAT_ELLIPSE_FILL = 71

Global Const PEGAT_DASH = 72
Global Const PEGAT_PIXEL = 73

Global Const PEGAT_STARTPOLY = 74
Global Const PEGAT_ADDPOLYPOINT = 75
Global Const PEGAT_ENDPOLYGON = 76
Global Const PEGAT_ENDPOLYLINE_THIN = 77
Global Const PEGAT_ENDPOLYLINE_MEDIUM = 78
Global Const PEGAT_ENDPOLYLINE_THICK = 79
Global Const PEGAT_ENDPOLYLINE_DASH = 80
Global Const PEGAT_ENDPOLYLINE_DOT = 81
Global Const PEGAT_ENDPOLYLINE_DASHDOT = 82
Global Const PEGAT_ENDPOLYLINE_DASHDOTDOT = 83

Global Const PEGAT_STARTTEXT = 84
Global Const PEGAT_ADDTEXT = 85
Global Const PEGAT_PARAGRAPH = 86

Global Const PEDTM_NONE = 0
Global Const PEDTM_VB = 1
Global Const PEDTM_DELPHI = 2

Global Const PESA_ALL = 0
Global Const PESA_AXISLABELS = 1
Global Const PESA_GRIDNUMBERS = 2
Global Const PESA_NONE = 3
Global Const PESA_LABELONLY = 4
Global Const PESA_EMPTY = 5

Global Const PELL_TOP = 0
Global Const PELL_BOTTOM = 1
Global Const PELL_LEFT = 2
Global Const PELL_RIGHT = 3

Global Const PESS_NONE = 0
Global Const PESS_FINANCIAL = 1

Global Const PEHSS_SMALL = 0
Global Const PEHSS_MEDIUM = 1
Global Const PEHSS_LARGE = 2

Global Const PEDS_NONE = 0
Global Const PEDS_SHADOWS = 1
Global Const PEDS_3D = 2

Global Const PEGP_NO_GROUPING = 0
Global Const PEGP_1_PERCENT = 1
Global Const PEGP_2_PERCENT = 2
Global Const PEGP_3_PERCENT = 3
Global Const PEGP_4_PERCENT = 4
Global Const PEGP_5_PERCENT = 5

Global Const PESF_TWENTYFOURPERCENT = 0
Global Const PESF_TWENTYTWOPERCENT = 1
Global Const PESF_TWENTYPERCENT = 2
Global Const PESF_EIGHTEENPRECENT = 3
Global Const PESF_SIXTEENPERCENT = 4
Global Const PESF_FOURTEENPERCENT = 5
Global Const PESF_TWELVEPERCENT = 6
Global Const PESF_TENPERCENT = 7
Global Const PESF_EIGHTPERCENT = 8
Global Const PESF_SIXPERCENT = 9

Global Const PESC_POLAR = 0
Global Const PESC_SMITH = 1
Global Const PESC_ROSE = 2

Global Const PESPG_NONE = 0
Global Const PESPG_HORZ = 1
Global Const PESPG_HORZANDVERT = 2

Global Const PEGS_THIN = 0
Global Const PEGS_THICK = 1
Global Const PEGS_DOT = 2
Global Const PEGS_DASH = 3

Global Const PEFVP_AUTO = 0
Global Const PEFVP_VERT = 1
Global Const PEFVP_HORZ = 2
Global Const PEFVP_SLANTED = 3

Global Const PEMAS_NONE = 0
Global Const PEMAS_THIN = 1
Global Const PEMAS_MEDIUM = 2
Global Const PEMAS_THICK = 3
Global Const PEMAS_THICKPLUSTICK = 4

Global Const PERI_INCBY15 = 0
Global Const PERI_INCBY10 = 1
Global Const PERI_INCBY5 = 2
Global Const PERI_INCBY2 = 3
Global Const PERI_INCBY1 = 4
Global Const PERI_DECBY1 = 5
Global Const PERI_DECBY2 = 6
Global Const PERI_DECBY5 = 7
Global Const PERI_DECBY10 = 8
Global Const PERI_DECBY15 = 9

Global Const PERD_WIREFRAME = 0
Global Const PERD_PLOTTINGMETHOD = 1
Global Const PERD_FULLDETAIL = 2

Global Const PESBB_WHILEROTATING = 0
Global Const PESBB_ALWAYS = 1
Global Const PESBB_NEVER = 2

Global Const PEPM_SURFACEPOLYGONS = 1
Global Const PEPM_3DBAR = 2
Global Const PEPM_POLYGONDATA = 3
Global Const PEPM_SCATTER = 4

Global Const PESC_NONE = 0
Global Const PESC_TOPLINES = 1
Global Const PESC_BOTTOMLINES = 2
Global Const PESC_TOPCOLORS = 3
Global Const PESC_BOTTOMCOLORS = 4

Global Const PESS_WHITESHADING = 0
Global Const PESS_COLORSHADING = 1


Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Type POINTSTRUCT
    X As Long
    Y As Long
End Type

Type HOTSPOTDATA
    HotSpotL As Long
    HotSpotT As Long
    HotSpotR As Long
    HotSpotB As Long
    nHotSpotType As Long
    n1 As Long
    n2 As Long
End Type

Type POINT3D
    X As Single
    Y As Single
    z As Single
End Type

Type POLYGONDATA
    Vertices(0 To 3) As POINT3D
    NumberOfVertices As Long
    PolyColor As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type

Type GLOBALPROPERTIES
    nObjectType As Long
    szMainTitle As String * 48
    szSubTitle As String * 48
    nSubsets As Long
    nPoints As Long
    bMonoWithSymbols As Long
    nDefOrientation As Long
    nPrepareImages As Long
    b3dDialogs As Long
    bDataShadows As Long
    bAllowCustomization As Long
    bAllowExporting As Long
    bAllowMaximization As Long
    bAllowPopup As Long
    nPageWidth As Long
    nPageHeight As Long
    rectLogicalLoc As Rect
    bCustom As Long
    nViewingStyle As Long
    nCViewingStyle As Long
    dwMonoDeskColor As Long
    dwMonoTextColor As Long
    dwMonoShadowColor As Long
    dwMonoGraphForeColor As Long
    dwMonoGraphBackColor As Long
    dwMonoTableForeColor As Long
    dwMonoTableBackColor As Long
    dwCMonoDeskColor As Long
    dwCMonoTextColor As Long
    dwCMonoShadowColor As Long
    dwCMonoGraphForeColor As Long
    dwCMonoGraphBackColor As Long
    dwCMonoTableForeColor As Long
    dwCMonoTableBackColor As Long
    dwDeskColor As Long
    dwTextColor As Long
    dwShadowColor As Long
    dwGraphForeColor As Long
    dwGraphBackColor As Long
    dwTableForeColor As Long
    dwTableBackColor As Long
    dwCDeskColor As Long
    dwCTextColor As Long
    dwCShadowColor As Long
    dwCGraphForeColor As Long
    dwCGraphBackColor As Long
    dwCTableForeColor As Long
    dwCTableBackColor As Long
    nDataPrecision As Long
    nCDataPrecision As Long
    nFontSize As Long
    nCFontSize As Long
    szMainTitleFont As String * 48
    bMainTitleBold As Long
    bMainTitleItalic As Long
    bMainTitleUnderline As Long
    szCMainTitleFont As String * 48
    bCMainTitleBold As Long
    bCMainTitleItalic As Long
    bCMainTitleUnderline As Long
    szSubTitleFont As String * 48
    bSubTitleBold As Long
    bSubTitleItalic As Long
    bSubTitleUnderline As Long
    szCSubTitleFont As String * 48
    bCSubTitleBold As Long
    bCSubTitleItalic As Long
    bCSubTitleUnderline As Long
    szLabelFont As String * 48
    bLabelBold As Long
    bLabelItalic As Long
    bLabelUnderline As Long
    szCLabelFont As String * 48
    bCLabelBold As Long
    bCLabelItalic As Long
    bCLabelUnderline As Long
    szTableFont As String * 48
    szCTableFont As String * 48
    bAllowSubsetHotSpots As Long
    bAllowPointHotSpots As Long
End Type

'Graph Plotting Methods for PEGO32.OCX
Global Const GPM_LINE = 0
Global Const GPM_BAR = 1
Global Const GPM_POINT = 2
Global Const GPM_AREA = 3
Global Const GPM_AREASTACKED = 4
Global Const GPM_AREASTACKEDPERCENT = 5
Global Const GPM_BARSTACKED = 6
Global Const GPM_BARSTACKEDPERCENT = 7
Global Const GPM_POINTSPLUSBFL = 8
Global Const GPM_POINTSPLUSBFLGRAPHED = 9
Global Const GPM_POINTPLUSBFC = 10
Global Const GPM_POINTPLUSBFCGRAPHED = 11
Global Const GPM_POINTPLUSSPLINE = 12
Global Const GPM_SPLINE = 13
Global Const GPM_HISTOGRAM = 14
Global Const GPM_SPECIFICPLOTMODE = 15
Global Const GPM_POINTSPLUSLINE = 16
Global Const GPM_HORIZONTALBAR = 17
Global Const GPM_HORZSTACKEDBAR = 18
Global Const GPM_HORZSTACKEDBARPERCENT = 19
Global Const GPM_STEP = 20
Global Const GPM_RIBBON = 21

Global Const GPMII_LINE = 0
Global Const GPMII_POINT = 1
Global Const GPMII_POINTSPLUSBFL = 2
Global Const GPMII_POINTSPLUSBFLGRAPHED = 3
Global Const GPMII_POINTPLUSBFC = 4
Global Const GPMII_POINTPLUSBFCGRAPHED = 5
Global Const GPMII_POINTPLUSSPLINE = 6
Global Const GPMII_SPLINE = 7
Global Const GPMII_POINTSPLUSLINE = 8
Global Const GPMII_STEP = 9
Global Const GPMII_RIBBON = 10

'Scientific Graph Plotting Methods for PESGO32.OCX
Global Const SGPM_LINE = 0
Global Const SGPM_POINT = 1
Global Const SGPM_STICK = 2
Global Const SGPM_POINTSPLUSBFL = 3
Global Const SGPM_POINTSPLUSBFC = 4
Global Const SGPM_POINTSPLUSSPLINE = 5
Global Const SGPM_SPLINE = 6
Global Const SGPM_BUBBLE = 7
Global Const SGPM_POINTSPLUSLINE = 8
Global Const SGPM_AREA = 9
Global Const SGPM_BAR = 10
Global Const SGPM_SPECIFICPLOTMODE = 11
Global Const SGPM_STEP = 12
Global Const SGPM_RIBBON = 13
Global Const SGPM_CONTOURLINES = 14
Global Const SGPM_CONTOURCOLORS = 15

Global Const SGPMII_LINE = 0
Global Const SGPMII_POINT = 1
Global Const SGPMII_POINTSPLUSBFL = 2
Global Const SGPMII_POINTSPLUSBFC = 3
Global Const SGPMII_POINTSPLUSSPLINE = 4
Global Const SGPMII_SPLINE = 5
Global Const SGPMII_POINTSPLUSLINE = 6
Global Const SGPMII_RIBBON = 7

'Polar/Smith Plotting Methods for PEPSO32.OCX
Global Const PGPM_LINE = 0
Global Const PGPM_POINT = 1
Global Const PGPM_POINTSPLUSLINE = 2

Global Const PGPMII_LINE = 0
Global Const PGPMII_POINT = 1
Global Const PGPMII_POINTSPLUSLINE = 2

'3D Object Comparison Plotting Method
Global Const TDPM_0 = 0
Global Const TDPM_1 = 1
Global Const TDPM_2 = 2
Global Const TDPM_3 = 3
Global Const TDPM_4 = 4
Global Const TDPM_5 = 5

Global Const TDPMII_POINT = 0
Global Const TDPMII_LINE = 1
Global Const TDPMII_POINTSPLUSLINE = 2

'// MOST COMMON CONSTANTS USED WITH VB //'
'// IF OTHERS ARE NEEDED, FIND THEM IN PEGRPAPI.TXT //'
Global Const PEP_nSUBSETS = 2115
Global Const PEP_nPOINTS = 2120
Global Const PEP_szaSUBSETLABELS = 2125
Global Const PEP_szaPOINTLABELS = 2130
Global Const PEP_faXDATA = 2135
Global Const PEP_faYDATA = 2140
Global Const PEP_bCUSTOM = 2225
Global Const PEP_faAPPENDYDATA = 3276
Global Const PEP_szaAPPENDPOINTLABELDATA = 3277
Global Const PEP_faAPPENDYDATAII = 3024
Global Const PEP_faAPPENDXDATA = 3658
Global Const PEP_faAPPENDXDATAII = 3659

Global Const PEP_ptLASTMOUSEMOVE = 2637
Global Const PEP_rectGRAPH = 3049
Global Const PEP_rectAXIS = 3051
Global Const PEP_structHOTSPOTDATA = 2610
Global Const PEP_structPOLYDATA = 4014

Global Const PEP_faVERTLINEANNOTATION = 3218
Global Const PEP_szaVERTLINEANNOTATIONTEXT = 3219
Global Const PEP_naVERTLINEANNOTATIONTYPE = 3221
Global Const PEP_dwaVERTLINEANNOTATIONCOLOR = 3222

'////// API FUNCTIONS //////'
Declare Function PEsetglobal Lib "PEGRP32A.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long
Declare Function PEgetglobal Lib "PEGRP32A.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long
Declare Function PEvset Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, lpvData As Any, ByVal nItems&) As Long
Declare Function PEvget Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, lpvDest As Any) As Long
Declare Function PEvsetcell Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvData As Any) As Long
Declare Function PEvgetcell Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvDest As Any) As Long
Declare Function PEszset Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
Declare Function PEszget Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
Declare Function PEnset Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
Declare Function PEnget Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&) As Long
Declare Function PElset Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
Declare Function PElget Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nProperty&) As Long
Declare Function PEcreate Lib "PEGRP32A.DLL" (ByVal nObjectType&, ByVal dwStyle&, lpRect As Rect, ByVal hParent&, ByVal nID&) As Long
Declare Function PEdestroy Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PEload Lib "PEGRP32A.DLL" (ByVal hObject&, lphGlbl As Any) As Long
Declare Function PEstore Lib "PEGRP32A.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
Declare Function PEloadpartial Lib "PEGRP32A.DLL" (ByVal hObject&, lphGlbl As Any) As Long
Declare Function PEstorepartial Lib "PEGRP32A.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
Declare Function PEgetmeta Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PEresetimage Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nLength&, ByVal nHeight&) As Long
Declare Function PElaunchcustomize Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PElaunchexport Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PElaunchmaximize Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PElaunchtextexport Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal bToFile&, ByVal lpszFilename$) As Long
Declare Function PElaunchprintdialog Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal bFullPage&, lpPoint As POINTSTRUCT) As Long
Declare Function PElaunchcolordialog Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PElaunchfontdialog Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PElaunchpopupmenu Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEreinitialize Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PEreinitializecustoms Lib "PEGRP32A.DLL" (ByVal hObject&) As Long
Declare Function PEgethelpcontext Lib "PEGRP32A.DLL" (ByVal HWND&) As Long
Declare Function PEcopymetatoclipboard Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEcopymetatofile Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
Declare Function PEcopybitmaptoclipboard Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEcopybitmaptofile Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
Declare Function PEcopyoletoclipboard Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEprintgraph Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal hDC&, ByVal nWidth&, ByVal nHeight&, ByVal nOrient&) As Long
Declare Function PEconvpixeltograph Lib "PEGRP32A.DLL" (ByVal hObject&, ByRef nAxis&, ByRef nX&, ByRef nY&, ByRef fX#, ByRef fY#, ByVal bRight&, ByVal bTop&, ByVal bVV&) As Long
Declare Function PEreset Lib "PEGRP32A.DLL" (ByVal hObject&) As Long

Declare Function PEgethotspot Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nX&, ByVal nY&) As Long
Declare Function PEvsetEx Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal property&, ByVal nStartingCell&, ByVal nCellCount&, lpData As Any, lpMemSetValue As Any) As Long
Declare Function PEvgetEx Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal property&, ByVal nStartingCell&, ByVal nCellCount&, lpData As Any) As Long
Declare Function PEpartialresetimage Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal nStartPoint&, ByVal nPointsToAdd&) As Long
Declare Function PEsavetofile Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal lpFileName$) As Long
Declare Function PEloadfromfile Lib "PEGRP32A.DLL" (ByVal hObject&, ByVal lpFileName$) As Long
Declare Function PEcreatefromfile Lib "PEGRP32A.DLL" (ByVal lpFileName$, ByVal hParent&, lpRect As Rect, ByVal nID&) As Long
Declare Function PEcopyjpegtoclipboard Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEcopyjpegtofile Lib "PEGRP32A.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long


'**** Graph output to Printer (Metafile)
Declare Function SetMapMode Lib "GDI32" (ByVal hDC&, ByVal Mode&) As Long
Declare Function SetViewportExtEx Lib "GDI32" (ByVal hDC&, ByVal X&, ByVal Y&, lpPoint As Any) As Long
Declare Function SetViewportOrgEx Lib "GDI32" (ByVal hDC&, ByVal X&, ByVal Y&, lpPoint As Any) As Long
Declare Function PlayMetaFile Lib "GDI32" (ByVal hDC&, ByVal hMeta&) As Long
