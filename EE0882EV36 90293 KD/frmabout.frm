VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3165
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5745
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2181.088
   ScaleMode       =   0  'User
   ScaleWidth      =   5392.023
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmabout.frx":0000
      ScaleHeight     =   336.791
      ScaleMode       =   0  'User
      ScaleWidth      =   336.791
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2196
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Caption         =   "n/a"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1050
      TabIndex        =   7
      Top             =   240
      Width           =   4051
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright © 2006  CTS Corporation  (Elkhart, IN)"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   1290
      Width           =   4050
   End
   Begin VB.Label lblDeveloper 
      Caption         =   "Program developed by CTS Electronics Engineering"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.47
      X2              =   5309.43
      Y1              =   1385.146
      Y2              =   1385.146
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   101.364
      X2              =   5312.246
      Y1              =   1404.441
      Y2              =   1404.441
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   270
      Width           =   4050
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning:  This computer program is considered a trade secret of CTS Corporation and is protected by all applicable laws."
      ForeColor       =   &H00000000&
      Height          =   828
      Left            =   252
      TabIndex        =   3
      Top             =   2196
      Width           =   3816
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1       'set default lower array bound to 1 instead of 0

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
'
'   PURPOSE: To display system information
'
'  INPUT(S): none
' OUTPUT(S): none

Call StartSysInfo

End Sub

Private Sub cmdOK_Click()
'
'   PURPOSE: To simulate an ok via a button.  By pressing this button the form
'            will unload.
'
'  INPUT(S): none
' OUTPUT(S): none

Unload Me

End Sub

Private Sub Form_Load()
'
'   PURPOSE: To Load the form with the appropriate version number and title
'
'  INPUT(S): none
' OUTPUT(S): none

Me.Caption = "About " & App.Title
lblTitle.Caption = frmMain.Caption

End Sub

Public Sub StartSysInfo()
'
'   PURPOSE: To display system information
'
'  INPUT(S): none
' OUTPUT(S): none

Dim rc As Long
Dim SysInfoPath As String

On Error GoTo SysInfoErr

' Try To Get System Info Program Path\Name From Registry...
If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
' Try To Get System Info Program Path Only From Registry...
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Validate Existance Of Known 32 Bit File Version
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        
    ' Error - File Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
' Error - Registry Entry Can Not Be Found...
Else
    GoTo SysInfoErr
End If

Call Shell(SysInfoPath, vbNormalFocus)

Exit Sub

SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'
'   PURPOSE:
'
'  INPUT(S):
' OUTPUT(S):

Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
'------------------------------------------------------------
' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
'------------------------------------------------------------
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

tmpVal = String$(1024, 0)                             ' Allocate Variable Space
KeyValSize = 1024                                       ' Mark Variable Size

'------------------------------------------------------------
' Retrieve Registry Key Value...
'------------------------------------------------------------
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                     KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                    
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
    tmpVal = left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
Else                                                    ' WinNT Does NOT Null Terminate String...
    tmpVal = left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
End If
'------------------------------------------------------------
' Determine Key Value Type For Conversion...
'------------------------------------------------------------
Select Case KeyValType                                  ' Search Data Types...
Case REG_SZ                                             ' String Registry Key Data Type
    KeyVal = tmpVal                                     ' Copy String Value
Case REG_DWORD                                          ' Double Word Registry Key Data Type
    For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
    Next
    KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
End Select

GetKeyValue = True                                      ' Return Success
rc = RegCloseKey(hKey)                                  ' Close Registry Key
Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

