VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChat 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "usuario - Conversación"
   ClientHeight    =   6180
   ClientLeft      =   975
   ClientTop       =   1050
   ClientWidth     =   8850
   Icon            =   "talka.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8850
   Begin RichTextLib.RichTextBox bandeja 
      Height          =   2775
      Left            =   300
      TabIndex        =   1
      Top             =   1500
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"talka.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox writer 
      Height          =   735
      Left            =   300
      TabIndex        =   0
      Top             =   4850
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      MaxLength       =   300
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"talka.frx":060C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00814D3C&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   5700
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Algo de texto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00814D3C&
      Height          =   195
      Left            =   795
      TabIndex        =   3
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      ForeColor       =   &H00814D3C&
      Height          =   195
      Left            =   350
      TabIndex        =   2
      Top             =   1170
      Width           =   375
   End
   Begin VB.Image Image21 
      Height          =   5100
      Left            =   0
      Picture         =   "talka.frx":068E
      Stretch         =   -1  'True
      Top             =   870
      Width           =   75
   End
   Begin VB.Image Image20 
      Height          =   870
      Left            =   0
      Picture         =   "talka.frx":06E6
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image19 
      Height          =   870
      Left            =   135
      Picture         =   "talka.frx":0C1A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4950
   End
   Begin VB.Image Image18 
      Height          =   870
      Left            =   5080
      Picture         =   "talka.frx":0DD8
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image17 
      Height          =   315
      Left            =   5920
      Picture         =   "talka.frx":14CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image Image16 
      Height          =   315
      Left            =   8695
      Picture         =   "talka.frx":1597
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image15 
      Height          =   5205
      Left            =   8775
      Picture         =   "talka.frx":17FC
      Stretch         =   -1  'True
      Top             =   315
      Width           =   75
   End
   Begin VB.Image Image14 
      Height          =   75
      Left            =   225
      Picture         =   "talka.frx":1854
      Stretch         =   -1  'True
      Top             =   6105
      Width           =   7935
   End
   Begin VB.Image Image13 
      Height          =   210
      Left            =   0
      Picture         =   "talka.frx":18AC
      Top             =   5965
      Width           =   225
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   6000
      Picture         =   "talka.frx":1A09
      Top             =   4890
      Width           =   705
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   345
      Left            =   330
      Picture         =   "talka.frx":2963
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   6390
   End
   Begin VB.Image Image9 
      Height          =   345
      Left            =   6720
      Picture         =   "talka.frx":2A5B
      Top             =   5640
      Width           =   90
   End
   Begin VB.Image Image8 
      Height          =   345
      Left            =   6720
      Picture         =   "talka.frx":2B3D
      Top             =   4455
      Width           =   90
   End
   Begin VB.Image Image7 
      Height          =   345
      Left            =   330
      Picture         =   "talka.frx":2C25
      Stretch         =   -1  'True
      Top             =   4455
      Width           =   6390
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   330
      Picture         =   "talka.frx":2D1D
      Stretch         =   -1  'True
      Top             =   1095
      Width           =   6390
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   6720
      Picture         =   "talka.frx":2D67
      Top             =   1095
      Width           =   90
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   240
      Picture         =   "talka.frx":2DBC
      Top             =   1095
      Width           =   90
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   240
      Picture         =   "talka.frx":2E11
      Top             =   4450
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   240
      Picture         =   "talka.frx":2EF8
      Top             =   5640
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   8160
      Picture         =   "talka.frx":2FDC
      Top             =   5520
      Width           =   690
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A06B4F&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   240
      Top             =   4800
      Width           =   6570
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00A06B4F&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Top             =   1440
      Width           =   6570
   End
   Begin VB.Image Image12 
      Height          =   4860
      Left            =   1515
      Picture         =   "talka.frx":3633
      Top             =   1320
      Width           =   7275
   End
   Begin VB.Image Image22 
      Height          =   6180
      Left            =   0
      Picture         =   "talka.frx":8435
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8805
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Dim Clip As String
Private Const WM_USER = &H400
Private Const EM_AUTOURLDETECT = &H45B
Private Type POINTAPI
x As Long
y As Long
End Type
Private Const RGN_COPY = 5
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Private Const TVM_SETBKCOLOR = 4381&
Private Const EM_CHARFROMPOS& = &HD7

Private Sub bandeja_Change()
bandeja.SelStart = Len(bandeja.Text)
End Sub

Private Sub Form_Activate()
Label2.Caption = Trim(Mid(Me.Caption, 1, Len(Me.Caption) - 15))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub writer_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim buffer As String
If KeyAscii = 13 Then
buffer = Trim(writer.Text)
buffer = Replace(buffer, vbCrLf, "")
If Trim(buffer) = "" Then
writer.Text = ""
KeyAscii = 0
Exit Sub
End If
If Trim(writer.Text) = "" Then Exit Sub
bandeja.SelStart = Len(bandeja.Text)
bandeja.SelFontName = "Verdana"
bandeja.SelBold = False
bandeja.SelItalic = False
bandeja.SelStrikeThru = False
bandeja.SelUnderline = False
bandeja.SelColor = RGB(150, 150, 150)
bandeja.SelFontSize = 10
bandeja.SelText = Replace(frmMain.Label5, Chr(12), vbCrLf) & " dice: " & vbCrLf
bandeja.SelFontSize = 9 'Opciones de usuario
bandeja.SelFontName = "Verdana" 'Opciones de usuario
bandeja.SelBold = True 'Opciones de usuario
bandeja.SelItalic = False 'Opciones de usuario
bandeja.SelStrikeThru = False 'Opciones de usuario
bandeja.SelUnderline = False 'Opciones de usuario
bandeja.SelColor = RGB(30, 45, 150) 'Opciones de usuario
bandeja.SelText = writer.Text & vbCrLf
bandeja.SelStart = Len(bandeja.Text)
bandeja.SelFontSize = 10
CheckSmileys Len(writer.Text) + 64
frmMain.sock(iSockets).SendData "IMDAT" & "SND" & frmMain.Label5.Caption & "RCE" & Label2.Caption & "MSG" & writer.Text
writer.Text = ""
KeyAscii = 0
End If
End Sub

Public Sub CheckSmileys(Length As Integer)
'ESTA FUNCIÓN GENERA UNA GRAVE RALENTIZACIÓN DEL PROCESAMIENTO DE TEXTO
On Error Resume Next
Dim a As Integer
Dim bClipHasImage As Boolean
Dim found As Integer
Dim scase As String
Clip = ""
'On Error GoTo cont:
On Error Resume Next
If Trim(Clipboard.GetText) <> "" Then Clip = Clipboard.GetText
cont:
found = 0
bandeja.Locked = False
For a = (Len(bandeja.Text) - (Length + 1)) To Len(bandeja.Text)
bandeja.SelStart = a - 1
bandeja.SelLength = 2
Clipboard.Clear
scase = Mid(bandeja.Text, a, 2)
Select Case scase
Case ":)"
Clipboard.SetData frmImages.imgIcon(56).Picture
Case ":o"
Clipboard.SetData frmImages.imgIcon(52).Picture
Case ":O"
Clipboard.SetData frmImages.imgIcon(52).Picture
Case ";)"
Clipboard.SetData frmImages.imgIcon(67).Picture
Case ":s"
Clipboard.SetData frmImages.imgIcon(37).Picture
Case ":S"
Clipboard.SetData frmImages.imgIcon(37).Picture
Case ":["
Clipboard.SetData frmImages.imgIcon(29).Picture
Case ":d"
Clipboard.SetData frmImages.imgIcon(61).Picture
Case ":D"
Clipboard.SetData frmImages.imgIcon(61).Picture
Case ":p"
Clipboard.SetData frmImages.imgIcon(64).Picture
Case ":P"
Clipboard.SetData frmImages.imgIcon(64).Picture
Case ":("
Clipboard.SetData frmImages.imgIcon(58).Picture
Case ":|"
Clipboard.SetData frmImages.imgIcon(65).Picture
Case ":$"
Clipboard.SetData frmImages.imgIcon(55).Picture
Case ":@"
Clipboard.SetData frmImages.imgIcon(28).Picture
End Select
If Str(Clipboard.GetData) <> 0 Then SendMessage bandeja.hWnd, WM_PASTE, 0, 0
Clipboard.Clear
scase = Mid(bandeja.Text, a, 3)
Select Case scase
Case ":-)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(56).Picture
Case ":-O"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(52).Picture
Case ":-o"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(52).Picture
Case ";-)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(67).Picture
Case ":-S"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(37).Picture
Case ":-s"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(37).Picture
Case ":'("
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(38).Picture
Case "(H)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(59).Picture
Case "(h)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(59).Picture
Case "(A)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(27).Picture
Case "(a)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(27).Picture
Case ":-#"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(0).Picture
Case "8-|"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(25).Picture
Case ":-*"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(4).Picture
Case ":^)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(21).Picture
Case "|-)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(26).Picture
Case "(Y)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(63).Picture
Case "(y)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(63).Picture
Case "(B)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(30).Picture
Case "(b)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(30).Picture
Case "(X)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(43).Picture
Case "(x)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(43).Picture
Case "({)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(68).Picture
Case ":-["
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(29).Picture
Case "(L)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(45).Picture
Case "(l)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(45).Picture
Case "(K)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(46).Picture
Case "(k)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(46).Picture
Case "(F)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(57).Picture
Case "(f)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(57).Picture
Case "(P)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(33).Picture
Case "(p)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(33).Picture
Case "(@)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(34).Picture
Case "(T)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(53).Picture
Case "(t)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(53).Picture
Case "(8)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(51).Picture
Case "(*)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(60).Picture
Case "(O)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(35).Picture
Case "(o)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(35).Picture
Case ":-D"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(61).Picture
Case ":-d"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(61).Picture
Case ":-P"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(64).Picture
Case ":-p"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(64).Picture
Case ":-("
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(58).Picture
Case ":-|"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(65).Picture
Case ":-$"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(55).Picture
Case ":-@"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(28).Picture
Case "(6)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(39).Picture
Case "8o|"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(1).Picture
Case "8O|"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(1).Picture
Case LCase("^o)")
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(3).Picture
Case "+o("
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(6).Picture
Case "*-)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(22).Picture
Case "8-)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(2).Picture
Case "(C)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(36).Picture
Case "(c)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(36).Picture
Case "(N)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(62).Picture
Case "(n)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(62).Picture
Case "(D)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(48).Picture
Case "(d)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(48).Picture
Case "(Z)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(69).Picture
Case "(z)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(69).Picture
Case "(})"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(44).Picture
Case "(^)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(32).Picture
Case "(U)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(31).Picture
Case "(u)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(31).Picture
Case "(G)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(54).Picture
Case "(g)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(54).Picture
Case "(W)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(66).Picture
Case "(w)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(66).Picture
Case "(~)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(42).Picture
Case "(&)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(42).Picture
Case "(I)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(47).Picture
Case "(i)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(47).Picture
Case "(S)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(50).Picture
Case "(s)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(50).Picture
Case "(E)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(41).Picture
Case "(e)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(41).Picture
Case "(M)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(49).Picture
Case "(m)"
bandeja.SelLength = 3
Clipboard.SetData frmImages.imgIcon(49).Picture
End Select
scase = Mid(bandeja.Text, a, 4)
Select Case scase
Case "<:o)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(24).Picture
Case "(sn)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(7).Picture
Case "(SN)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(7).Picture
Case "(Sn)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(7).Picture
Case "(sN)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(7).Picture
Case "(pl)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(8).Picture
Case "(PL)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(8).Picture
Case "(Pl)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(8).Picture
Case "(pL)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(8).Picture
Case "(pi)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(10).Picture
Case "(PI)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(10).Picture
Case "(Pi)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(10).Picture
Case "(pI)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(10).Picture
Case "(au)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(12).Picture
Case "(AU)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(12).Picture
Case "(Au)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(12).Picture
Case "(aU)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(12).Picture
Case "(um)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(14).Picture
Case "(UM)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(14).Picture
Case "(Um)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(14).Picture
Case "(uM)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(14).Picture
Case "(co)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(16).Picture
Case "(CO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(16).Picture
Case "(Co)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(16).Picture
Case "(cO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(16).Picture
Case "(st)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(18).Picture
Case "(ST)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(18).Picture
Case "(St)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(18).Picture
Case "(sT)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(18).Picture
Case "(mo)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(19).Picture
Case "(MO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(19).Picture
Case "(Mo)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(19).Picture
Case "(mO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(19).Picture
Case "(ba)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(20).Picture
Case "(BA)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(20).Picture
Case "(Ba)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(20).Picture
Case "(bA)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(20).Picture
Case "(||)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(9).Picture
Case "(so)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(11).Picture
Case "(SO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(11).Picture
Case "(So)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(11).Picture
Case "(sO)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(11).Picture
Case "(ap)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(13).Picture
Case "(AP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(13).Picture
Case "(Ap)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(13).Picture
Case "(aP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(13).Picture
Case "(ip)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(15).Picture
Case "(IP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(15).Picture
Case "(Ip)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(15).Picture
Case "(iP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(15).Picture
Case "(mp)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(17).Picture
Case "(MP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(17).Picture
Case "(mP)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(17).Picture
Case "(Mp)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(17).Picture
Case "(li)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(23).Picture
Case "(LI)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(23).Picture
Case "(Li)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(23).Picture
Case "(lI)"
bandeja.SelLength = 4
Clipboard.SetData frmImages.imgIcon(23).Picture
End Select
If Str(Clipboard.GetData) <> 0 Then SendMessage bandeja.hWnd, WM_PASTE, 0, 0
Clipboard.Clear
Next
If Trim(Clip) <> "" Then Clipboard.SetText Clip
bandeja.Locked = True
End Sub
