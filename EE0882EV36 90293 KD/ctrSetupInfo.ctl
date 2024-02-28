VERSION 5.00
Begin VB.UserControl ctrSetupInfo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   ScaleHeight     =   945
   ScaleWidth      =   10575
   Begin VB.Frame frmSetupInformation 
      Caption         =   "Setup Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtTLNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtSample 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTemperature 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtDateCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         MaxLength       =   8
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtSeries 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtOperator 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "TL #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblSample 
         Caption         =   "Sample #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1540
         TabIndex        =   16
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lbTemperature 
         Caption         =   "Temperature:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   15
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblShift 
         Caption         =   "Date Code:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   150
         Width           =   975
      End
      Begin VB.Label lblPart 
         Caption         =   "Serial #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lblSeries 
         Caption         =   "Series:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lblOperator 
         Caption         =   "Operator:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.PictureBox picPreviewResults 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11100
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   11160
   End
End
Attribute VB_Name = "ctrSetupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Revision   Date      Initials  Explanation
' 1.1     02-03-2006    ANM     Modified for TestLab items
' 1.2     07-24-2008    ANM     Restricted items in some boxes
'

Public Property Get Operator() As String
'Get Operator Initials
Operator = txtOperator.Text
End Property

Public Property Let Operator(ByVal NewText As String)
'Set Operator Initials
txtOperator.Text = NewText
End Property

Public Property Get Series() As String
'Get Series
Series = txtSeries.Text
End Property

Public Property Let Series(ByVal NewText As String)
'Set Series
txtSeries.Text = NewText
End Property

Public Property Get DateCode() As String
'Get Date Code
DateCode = txtDateCode.Text
End Property

Public Property Let DateCode(ByVal NewText As String)
'Set DateCode
txtDateCode.Text = NewText
End Property

Public Property Get PartNum() As String
'Get Part Number
PartNum = txtPart.Text
End Property

Public Property Let PartNum(ByVal NewText As String)
'Set Part Number
txtPart.Text = NewText
End Property

Public Property Get Temperature() As String
'Get Temperature
Temperature = txtTemperature.Text
End Property

Public Property Let Temperature(ByVal NewText As String)
'Set Temperature
txtTemperature.Text = NewText
End Property

Public Property Get Sample() As String
'Get Sample
Sample = txtSample.Text
End Property

Public Property Let Sample(ByVal NewText As String)
'Set Sample
txtSample.Text = NewText
End Property

Public Property Get TLNum() As String
'Get TLNum
TLNum = txtTLNum.Text
End Property

Public Property Let TLNum(ByVal NewText As String)
'Set TLNum
txtTLNum.Text = NewText
End Property

Public Property Get Comment() As String
'Get Comment
Comment = txtComment.Text
End Property

Public Property Let Comment(ByVal NewText As String)
'Set Comment
txtComment.Text = NewText
End Property

Private Sub txtComment_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To restrict certain characters from the comment field.
'            This subroutine will be initiated by a key pressed event.
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none
'3.4aANM new sub

'Accept only letters, numbers, & appropriate characters
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 32 To 33       'space and !
        'Accept the character
    Case 35 To 41       '# $ % & ` ( )
        'Accept the character
    Case 43             '+
        'Accept the character
    Case 45             '-
        'Accept the character
    Case 46             '.
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 64 To 90       '@ and A-Z (upper case)
        'Accept the character
    Case 94 To 95       '^ and underscore
        'Accept the character
    Case 97 To 122      'a-z (lower case)
        'Accept the character
    Case 189, 188, 181, 177, 61, 247, 152, 176, 183, 178
        'Accept the character
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select

End Sub

Private Sub txtOperator_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To restrict certain characters from the operator field.
'            This subroutine will be initiated by a key pressed event.
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none
'3.4aANM new sub

'Accept only letters, numbers, & appropriate characters
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 65 To 90       'A-Z (upper case)
        'Accept the character
    Case 97 To 122      'a-z (lower case)
        'Accept the character
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select

End Sub

Private Sub txtSeries_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To restrict certain characters from the series field.
'            This subroutine will be initiated by a key pressed event.
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none
'3.4aANM new sub

'Accept only letters, numbers, & appropriate characters
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 32             'space
        'Accept the character
    Case 43             '+
        'Accept the character
    Case 45             '-
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 65 To 90       'A-Z (upper case)
        'Accept the character
    Case 95             'underscore
        'Accept the character
    Case 97 To 122      'a-z (lower case)
        'Accept the character
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select

End Sub

Private Sub txtTemperature_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To restrict certain characters from the temp field.
'            This subroutine will be initiated by a key pressed event.
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none
'3.4aANM new sub

'Accept only letters, numbers, & appropriate characters
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 32 To 33       'space and !
        'Accept the character
    Case 35 To 41       '# $ % & ` ( )
        'Accept the character
    Case 43             '+
        'Accept the character
    Case 45             '-
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 64 To 90       '@ and A-Z (upper case)
        'Accept the character
    Case 94 To 95       '^ and underscore
        'Accept the character
    Case 97 To 122      'a-z (lower case)
        'Accept the character
    Case 176            '°
        'Accept the character
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select

End Sub

Private Sub txtTLNum_KeyPress(KeyAscii As Integer)
'
'   PURPOSE: To restrict certain characters from the TL# field.
'            This subroutine will be initiated by a key pressed event.
'
'  INPUT(S): KeyAscii = ascii representation of the key pressed
' OUTPUT(S): none
'3.4aANM new sub

'Accept only letters, numbers, & appropriate characters
Select Case KeyAscii
    Case 3, 8, 22       'copy, backspace & paste
        'Accept the character
    Case 48 To 57       '0-9
        'Accept the character
    Case 65 To 90       'A-Z (upper case)
        'Accept the character
    Case 95             'underscore
        'Accept the character
    Case 97 To 122      'a-z (lower case)
        'Accept the character
    Case Else
        KeyAscii = 0    ' Cancel the character.
        Beep            ' Sound error signal.
End Select

End Sub

