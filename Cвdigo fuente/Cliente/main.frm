VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FAECE6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IntraMessenger"
   ClientHeight    =   5955
   ClientLeft      =   810
   ClientTop       =   1500
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4365
   Begin VB.Frame fraControl 
      BackColor       =   &H8000000E&
      Height          =   1335
      Left            =   0
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   32
         Left            =   1680
         Top             =   840
      End
      Begin VB.Timer tmrIconica 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   600
         Top             =   840
      End
      Begin VB.Timer tmrTemporizador 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   120
         Top             =   840
      End
      Begin VB.Timer tmrAnim 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2640
         Top             =   840
      End
      Begin VB.Timer tmrFueras 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2160
         Top             =   840
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtClave 
         Height          =   315
         Left            =   2040
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtMyalias 
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPuerto 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtServidor 
         Height          =   315
         Left            =   600
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox estoyOn 
         Height          =   315
         Left            =   1080
         TabIndex        =   21
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTimeout 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSWinsockLib.Winsock sock 
         Index           =   0
         Left            =   1080
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ImageList imagenes 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":20AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":245E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2810
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView contactos 
      Height          =   4485
      Left            =   0
      TabIndex        =   12
      Top             =   650
      Visible         =   0   'False
      Width           =   4390
      _ExtentX        =   7752
      _ExtentY        =   7911
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   3
      Scroll          =   0   'False
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAECE6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   1450
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAECE6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   1450
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -9000
      Picture         =   "main.frx":36D8
      ScaleHeight     =   540
      ScaleWidth      =   23040
      TabIndex        =   5
      Top             =   1550
      Visible         =   0   'False
      Width           =   23040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar sesión"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   730
      MouseIcon       =   "main.frx":11B86
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "main.frx":11CD8
      ScaleHeight     =   540
      ScaleWidth      =   23040
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   23040
   End
   Begin VB.Image Image25 
      Height          =   675
      Left            =   0
      Picture         =   "main.frx":20186
      Top             =   80
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image24 
      Height          =   675
      Left            =   0
      Picture         =   "main.frx":207E5
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image23 
      Height          =   675
      Left            =   0
      Picture         =   "main.frx":20DFD
      Top             =   80
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image22 
      Height          =   270
      Left            =   560
      Picture         =   "main.frx":2147D
      Top             =   280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image21 
      Height          =   270
      Left            =   560
      Picture         =   "main.frx":218FC
      Top             =   280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image20 
      Height          =   270
      Left            =   560
      Picture         =   "main.frx":21D76
      Top             =   280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image19 
      Height          =   675
      Left            =   0
      Picture         =   "main.frx":221EC
      Top             =   80
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mi estado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   920
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image18 
      Height          =   240
      Left            =   1560
      Picture         =   "main.frx":22835
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image17 
      Height          =   240
      Left            =   120
      Picture         =   "main.frx":22BBF
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Left            =   360
      Picture         =   "main.frx":22F49
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   240
      Left            =   600
      Picture         =   "main.frx":232D3
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Algo de texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   920
      MouseIcon       =   "main.frx":2365D
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   330
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image Image14 
      Height          =   240
      Left            =   4080
      Picture         =   "main.frx":237AF
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image13 
      Height          =   240
      Left            =   3720
      Picture         =   "main.frx":23D39
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image12 
      Height          =   240
      Left            =   3360
      Picture         =   "main.frx":242C3
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image11 
      Height          =   240
      Left            =   3000
      Picture         =   "main.frx":2484D
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   240
      Left            =   2640
      Picture         =   "main.frx":24DD7
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   240
      Left            =   2280
      Picture         =   "main.frx":25361
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   240
      Left            =   1920
      Picture         =   "main.frx":258EB
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Haga clic para iniciar una sesión:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1095
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   1200
      Picture         =   "main.frx":25E75
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   840
      Picture         =   "main.frx":263FF
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando sesión..."
      Height          =   210
      Left            =   1537
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image Image4 
      Height          =   1365
      Left            =   0
      Picture         =   "main.frx":26989
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para iniciar sesión con una cuenta diferente, haga clic aquí."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00814D3C&
      Height          =   405
      Left            =   1000
      MouseIcon       =   "main.frx":27198
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3165
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00814D3C&
      Height          =   210
      Left            =   735
      TabIndex        =   1
      Top             =   1500
      Width           =   60
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   1065
      Left            =   710
      Picture         =   "main.frx":272EA
      Stretch         =   -1  'True
      Top             =   1370
      Width           =   3930
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1065
      Left            =   0
      Picture         =   "main.frx":27AAC
      Top             =   1365
      Width           =   705
   End
   Begin VB.Image Image7 
      Enabled         =   0   'False
      Height          =   3510
      Left            =   0
      Picture         =   "main.frx":28D30
      Stretch         =   -1  'True
      Top             =   2445
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   3510
      Left            =   0
      Picture         =   "main.frx":2B4E1
      Stretch         =   -1  'True
      Top             =   2445
      Width           =   4575
   End
   Begin VB.Menu file 
      Caption         =   "Archivo"
      Begin VB.Menu m1 
         Caption         =   "Iniciar sesión como"
      End
      Begin VB.Menu m01 
         Caption         =   "Cancelar inicio de sesión"
         Visible         =   0   'False
      End
      Begin VB.Menu m2 
         Caption         =   "Iniciar sesión..."
      End
      Begin VB.Menu m3 
         Caption         =   "Cerrar sesión"
         Enabled         =   0   'False
      End
      Begin VB.Menu se1 
         Caption         =   "-"
      End
      Begin VB.Menu m4 
         Caption         =   "Mi estado"
         Begin VB.Menu m5 
            Caption         =   "En línea"
         End
         Begin VB.Menu m9 
            Caption         =   "No disponible"
         End
         Begin VB.Menu m7 
            Caption         =   "Vuelvo enseguida"
         End
         Begin VB.Menu m8 
            Caption         =   "Ausente"
         End
         Begin VB.Menu m7a 
            Caption         =   "Al teléfono"
         End
         Begin VB.Menu m6 
            Caption         =   "Salí a comer"
         End
         Begin VB.Menu m10 
            Caption         =   "Sin conexión"
         End
      End
      Begin VB.Menu se2 
         Caption         =   "-"
      End
      Begin VB.Menu m11 
         Caption         =   "Enviar un archivo o una foto..."
         Enabled         =   0   'False
      End
      Begin VB.Menu m12 
         Caption         =   "Abrir archivos recibidos"
      End
      Begin VB.Menu se3 
         Caption         =   "-"
      End
      Begin VB.Menu m14 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu m15 
      Caption         =   "Contactos"
      Begin VB.Menu m16 
         Caption         =   "Agregar un contacto..."
         Enabled         =   0   'False
      End
      Begin VB.Menu m17 
         Caption         =   "Buscar un contacto"
         Enabled         =   0   'False
         Begin VB.Menu ds 
            Caption         =   "Menu no definido"
         End
      End
      Begin VB.Menu m18 
         Caption         =   "Ir a la libreta de direcciones"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu m19 
      Caption         =   "Acciones"
      Begin VB.Menu m20 
         Caption         =   "Enviar un mensaje instantáneo"
         Enabled         =   0   'False
      End
      Begin VB.Menu m21 
         Caption         =   "Enviar un archivo o una foto"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu m22 
      Caption         =   "Herramientas"
      Begin VB.Menu m23 
         Caption         =   "Siempre visible"
      End
      Begin VB.Menu se4 
         Caption         =   "-"
      End
      Begin VB.Menu m24 
         Caption         =   "Opciones..."
      End
   End
   Begin VB.Menu m25 
      Caption         =   "Ayuda"
      Begin VB.Menu m26 
         Caption         =   "Temas de Ayuda"
      End
      Begin VB.Menu m27 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu floar 
      Caption         =   "Flotante"
      Visible         =   0   'False
      Begin VB.Menu mirrorm1 
         Caption         =   "Iniciar sesión como"
      End
      Begin VB.Menu mirrorm01 
         Caption         =   "Cancelar inicio de sesión"
         Visible         =   0   'False
      End
      Begin VB.Menu mirrorm2 
         Caption         =   "Iniciar sesión..."
      End
      Begin VB.Menu mirrorm3 
         Caption         =   "Cerrar sesión"
         Enabled         =   0   'False
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu openswin 
         Caption         =   "Abrir ventana principal"
      End
      Begin VB.Menu closer 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu states 
      Caption         =   "Estados"
      Visible         =   0   'False
      Begin VB.Menu s1 
         Caption         =   "En línea"
      End
      Begin VB.Menu s6 
         Caption         =   "No disponible"
      End
      Begin VB.Menu s3 
         Caption         =   "Vuelvo enseguida"
      End
      Begin VB.Menu s5 
         Caption         =   "Ausente"
      End
      Begin VB.Menu s4 
         Caption         =   "Al teléfono"
      End
      Begin VB.Menu s2 
         Caption         =   "Salí a comer"
      End
      Begin VB.Menu s7 
         Caption         =   "Sin conexión"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Dim iSockets As Integer
Private Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Dim breaker As Integer
Dim Salto As Integer
Dim Contador As Integer
Dim TiempoEspera As Integer
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Dim nid As NOTIFYICONDATA
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_LOOP = &H8
Private Const SND_NODEFAULT = &H2
Dim frmChat(65536) As New frmChat
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub tmrAnim_Timer()
If Salto = 0 Then
Contador = Contador + 1
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image9.Picture
nid.szTip = "IntraMessenger - Sesión iniciada" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 1
ElseIf Salto = 1 Then
nid.hIcon = Image10.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 2
ElseIf Salto = 2 Then
nid.hIcon = Image11.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 3
ElseIf Salto = 3 Then
nid.hIcon = Image12.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 4
ElseIf Salto = 4 Then
nid.hIcon = Image13.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 5
ElseIf Salto = 5 Then
nid.hIcon = Image14.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 6
ElseIf Salto = 6 Then
nid.hIcon = Image8.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 7
ElseIf Salto = 7 Then
nid.hIcon = Image8.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 8
ElseIf Salto = 8 Then
nid.hIcon = Image8.Picture
Shell_NotifyIcon NIM_MODIFY, nid
Salto = 0
If Contador >= 8 Then
If m5.Checked = True Then
nid.hIcon = Image8.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m9.Checked = True Then
nid.hIcon = Image15.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m7.Checked = True Then
nid.hIcon = Image17.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m8.Checked = True Then
nid.hIcon = Image17.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m7a.Checked = True Then
nid.hIcon = Image15.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m6.Checked = True Then
nid.hIcon = Image17.Picture
Shell_NotifyIcon NIM_MODIFY, nid
ElseIf m10.Checked = True Then
nid.hIcon = Image16.Picture
Shell_NotifyIcon NIM_MODIFY, nid
End If
tmrAnim.Enabled = False
End If
End If
End Sub

Private Sub closer_Click()
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posX", frmMain.Top
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posY", frmMain.Left
If Trim(txtClave.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Drivers")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Drivers", "Commands", txtClave.Text
End If
If Trim(txtUsuario.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "LastUser", txtUsuario.Text
End If
Shell_NotifyIcon NIM_DELETE, nid
DoEvents
If estoyOn.Text = "1" Then
Call m10_Click
End If
End
End Sub

Private Sub Command1_Click()
If Trim(txtUsuario.Text) = "" Or Trim(txtClave.Text) = "" Then
frmLogin.Show vbModal, Me
Else
On Error Resume Next
Command1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = True
Timer1.Enabled = True
Picture1.Visible = True
Picture2.Visible = True
Image7.Visible = True
Frame1.Visible = True
Frame2.Visible = True
m2.Enabled = False
m1.Visible = False
mirrorm1.Visible = False
mirrorm2.Enabled = False
m01.Visible = True
mirrorm01.Visible = True
breaker = 0
tmrIconica.Enabled = True
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image6.Picture
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
tmrTemporizador.Enabled = True
Load sock(iSockets)
sock(iSockets).RemotePort = txtPuerto.Text
sock(iSockets).RemoteHost = txtServidor.Text
sock(iSockets).Connect
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Command1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = True
Timer1.Enabled = True
Picture1.Visible = True
Image7.Visible = True
Picture2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
m1.Visible = False
m2.Enabled = False
m01.Visible = True
mirrorm01.Visible = True
mirrorm2.Enabled = False
mirrorm1.Visible = False
breaker = 0
tmrIconica.Enabled = True
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image6.Picture
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
tmrTemporizador.Enabled = True
tmrTemporizador.Enabled = True
Load sock(iSockets)
sock(iSockets).RemotePort = txtPuerto.Text
sock(iSockets).RemoteHost = txtServidor.Text
sock(iSockets).Connect
End Sub

Private Sub contactos_DblClick()
'MÓDULO PARA CAMBIAR IMAGEN DEL NODO PRINCIPAL
If contactos.Nodes.Count = 0 Then
MsgBox "Existe un problema para cargar la lista de contactos. Por favor reinicie la aplicación.", vbInformation, "Información"
Exit Sub
End If
If contactos.SelectedItem.Key = "AMI" Or contactos.SelectedItem.Key = "FAM" Or contactos.SelectedItem.Key = "COM" Or contactos.SelectedItem.Key = "OTR" Then
If contactos.SelectedItem.Expanded = False Then
contactos.SelectedItem.Image = 6
ElseIf contactos.SelectedItem.Expanded = True Then
contactos.SelectedItem.Image = 7
End If
End If
'MÓDULO PARA INICIAR CONVERSACIÓN CON EL USUARIO SELECCIONADO
On Error Resume Next
Dim indice As Integer
indice = contactos.SelectedItem.Index
If contactos.SelectedItem.Image = 6 Or contactos.SelectedItem.Image = 7 Then Exit Sub
If contactos.SelectedItem.Image = 4 Then
MsgBox "No es posible enviar un mensaje a un usuario desconectado.", vbInformation, "Información"
Else
Load frmChat(indice)
frmChat(indice).Show
frmChat(indice).Caption = Trim(Mid(contactos.Nodes.Item(indice).Text, 1, InStrRev(contactos.Nodes(indice).Text, "(") - 1)) & " - Conversación"
frmChat(indice).WindowState = 0
End If
End Sub

Private Sub contactos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = False
End Sub

Private Sub Form_Activate()
m1.Caption = "Iniciar sesión como " & Label1.Caption
mirrorm1.Caption = "Iniciar sesión como " & Label1.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next
If AnotherInstance() Then End
Picture1.Enabled = False
Picture2.Enabled = False
Picture2.Left = 14040
Picture2.Top = Picture1.Top
frmMain.Top = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posX")
frmMain.Left = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posY")
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
txtClave.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Drivers", "Commands")
txtUsuario.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "LastUser")
txtTimeout.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "txtTimeout")
If txtUsuario.Text = "Error" Then
txtUsuario.Text = ""
End If
If txtClave.Text = "Error" Then
txtClave.Text = ""
End If
If txtTimeout.Text = "Error" Then
txtTimeout.Text = "5"
End If
Label1.Caption = txtUsuario.Text
reload
recargar
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim Msg As Long
Dim sFilter As String
Msg = x / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDBLCLK
If frmAbout.Visible = True And Me.Visible = False Then
AppActivate "Acerca de IntraMessenger"
Exit Sub
ElseIf frmAbout.Visible = True And Me.Visible = True Then
AppActivate "Acerca de IntraMessenger"
Exit Sub
End If
Me.Show
Me.WindowState = 0
AppActivate "IntraMessenger"
Case WM_RBUTTONUP
If frmAbout.Visible = True Then
AppActivate "Acerca de IntraMessenger"
PopupMenu menuprincipal, x
Exit Sub
End If
If Me.Visible = True Then
Me.WindowState = 0
AppActivate "IntraMessenger"
End If
PopupMenu floar, x, , , openswin
Case WM_RBUTTONDBLCLK
Case WM_LBUTTONDOWN
If frmAbout.Visible = True Then
AppActivate "Acerca de IntraMessenger"
Exit Sub
End If
If Me.Visible = True Then
Me.WindowState = 0
AppActivate "IntraMessenger"
End If
Case WM_RBUTTONDOWN
If Me.Visible = True And frmAbout.Visible = False Then
Me.WindowState = 0
AppActivate "IntraMessenger"
End If
Case WM_LBUTTONUP
End Select
End Sub

Private Sub Form_Resize()
Image1.Top = 1370
Image1.Left = 0
Image2.Width = frmMain.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posX", frmMain.Top
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger", "posY", frmMain.Left
Cancel = 1
If Me.Visible = True Then
TitleToTray Me
Me.Hide
End If
End Sub

Private Sub tmrFueras_Timer()
TiempoEspera = TiempoEspera + 1
If TiempoEspera = Val(txtTimeout.Text) Then
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = True
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = True
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "IDL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Ausente)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Ausente" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
TiempoEspera = 0
tmrFueras.Enabled = False
End If
End Sub

Private Sub tmrIconica_Timer()
If breaker = 0 Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = frmMain.Icon
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
breaker = 1
Else
nid.hIcon = Image6.Picture
Shell_NotifyIcon NIM_MODIFY, nid
breaker = 0
End If
End Sub

Private Sub Image19_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = False
End Sub

Private Sub Image20_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = True
End Sub

Private Sub Image21_Click()
Image22.Visible = True
PopupMenu states, x, Image22.Left, Image22.Top + Image22.Height + 35
Image22.Visible = False
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = False
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = False
End Sub

Private Sub Label3_Click()
frmLogin.Show vbModal, Me
End Sub

Private Sub Label5_Change()
'Label6.Top = Label5.Top + 70
Label6.Top = Label7.Top
If Label5.Width >= 4100 Then
Label5.Width = 4100
End If
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
End Sub

Private Sub Label5_Click()
Call Image21_Click
End Sub

Private Sub Label6_Change()
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image21.Visible = False
End Sub

Private Sub m01_Click()
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
mirrorm3.Enabled = False
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
End Sub

Private Sub m1_Click()
If Trim(txtUsuario.Text) = "" Or Trim(txtClave.Text) = "" Then
frmLogin.Show vbModal, Me
Else
On Error Resume Next
Command1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = True
Timer1.Enabled = True
Picture1.Visible = True
Image7.Visible = True
Picture2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
m2.Enabled = False
m01.Visible = True
m1.Visible = False
mirrorm2.Enabled = False
mirrorm01.Visible = True
mirrorm1.Visible = False
breaker = 0
tmrIconica.Enabled = True
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image6.Picture
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
tmrTemporizador.Enabled = True
Load sock(iSockets)
sock(iSockets).RemotePort = txtPuerto.Text
sock(iSockets).RemoteHost = txtServidor.Text
sock(iSockets).Connect
End If
End Sub

Private Sub m10_Click()
If m10.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = True
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = True
Image19.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = True
sock(iSockets).SendData "STATE" & "OFF" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Sin conexión)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image16.Picture
nid.szTip = "IntraMessenger - Sin conexión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub m14_Click()
Unload Me
End Sub

Private Sub m16_Click()
frmNewContact.Show vbModal, Me
End Sub

Private Sub m2_Click()
frmLogin.Show vbModal, Me
End Sub

Private Sub m24_Click()
frmOpciones.Show vbModal, Me
End Sub

Private Sub m27_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub m3_click()
'MÓDULO EXPERIMENTAL (CIERRA LAS VENTANAS AL CERRAR SESIÓN)
Dim i
For i = 1 To contactos.Nodes.Count
If contactos.Nodes.Item(i).Image = 1 Or contactos.Nodes.Item(i).Image = 2 Or contactos.Nodes.Item(i).Image = 3 Or contactos.Nodes.Item(i).Image = 5 Then
If frmChat(i).Visible = True Then
Unload frmChat(i)
End If
End If
Next i
DoEvents
'MÓDULO EXPERIMENTAL
sock(iSockets).SendData "STATE" & "KUS" & "$$$" & txtUsuario.Text
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
Label5.Visible = False
Label7.Visible = False
Label6.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
contactos.Visible = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
tmrFueras.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
DoEvents
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
Label1.Caption = txtUsuario.Text
m1.Caption = "Iniciar sesión como " & txtUsuario.Text
mirrorm1.Caption = "Iniciar sesión como " & txtUsuario.Text
End Sub

Private Sub m5_Click()
If m5.Checked = True Then Exit Sub
s1.Checked = True
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = True
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = True
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "ONL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(En línea)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image8.Picture
nid.szTip = "IntraMessenger - Sesión iniciada" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
If tmrFueras.Enabled = False Then
TiempoEspera = 0
tmrFueras.Enabled = True
End If
End Sub

Private Sub m6_Click()
If m6.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = True
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = True
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "LUN" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Salí a comer)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Salí a comer" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub m7_Click()
If m7.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = True
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = True
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "AWA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Vuelvo enseguida)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Vuelvo enseguida" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub m7a_Click()
If m7a.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = True
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = True
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = False
Image24.Visible = True
Image25.Visible = False
sock(iSockets).SendData "STATE" & "ONP" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Al teléfono)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - Al teléfono" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub m8_Click()
If m8.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = True
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = True
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "IDL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Ausente)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Ausente" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub m9_Click()
If m9.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = True
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = True
m10.Checked = False
Image19.Visible = False
Image23.Visible = False
Image24.Visible = True
Image25.Visible = False
sock(iSockets).SendData "STATE" & "NOA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(No disponible)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - No disponible" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub mirrorm3_click()
'MÓDULO EXPERIMENTAL (CIERRA LAS VENTANAS AL CERRAR SESIÓN)
Dim i
For i = 1 To contactos.Nodes.Count
If contactos.Nodes.Item(i).Image = 1 Or contactos.Nodes.Item(i).Image = 2 Or contactos.Nodes.Item(i).Image = 3 Or contactos.Nodes.Item(i).Image = 5 Then
If frmChat(i).Visible = True Then
Unload frmChat(i)
End If
End If
Next i
DoEvents
'MÓDULO EXPERIMENTAL
sock(iSockets).SendData "STATE" & "KUS" & "$$$" & txtUsuario.Text
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
If frmPopup.Visible = True Then
Unload frmPopup
End If
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
contactos.Visible = False
tmrFueras.Enabled = False
Label5.Visible = False
Label7.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Label6.Visible = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
DoEvents
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
m1.Caption = "Iniciar sesión como " & txtUsuario.Text
Label1.Caption = txtUsuario.Text
mirrorm1.Caption = "Iniciar sesión como " & txtUsuario.Text
End Sub

Private Sub m23_Click()
If m23.Checked = True Then
SetWinOnTop = SetWindowPos(frmMain.hWnd, HWND_NOTOPMOST, frmMain.Left, frmMain.Top, frmMain.Width, frmMain.Height, SWP_NOMOVE Or SWP_NOSIZE)
m23.Checked = False
Exit Sub
End If
If m23.Checked = False Then
SetWinOnTop = SetWindowPos(frmMain.hWnd, HWND_TOPMOST, frmMain.Left, frmMain.Top, frmMain.Width, frmMain.Height, SWP_NOMOVE Or SWP_NOSIZE)
m23.Checked = True
Exit Sub
End If
End Sub

Private Sub mirrorm01_Click()
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
mirrorm3.Enabled = False
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
End Sub

Private Sub mirrorm1_Click()
If Trim(txtUsuario.Text) = "" Or Trim(txtClave.Text) = "" Then
frmLogin.Show vbModal, Me
Else
On Error Resume Next
Command1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = True
Timer1.Enabled = True
Picture1.Visible = True
Image7.Visible = True
Picture2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
m2.Enabled = False
m01.Visible = True
m1.Visible = False
mirrorm2.Enabled = False
mirrorm01.Visible = True
mirrorm1.Visible = False
breaker = 0
tmrIconica.Enabled = True
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image6.Picture
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
tmrTemporizador.Enabled = True
Load sock(iSockets)
sock(iSockets).RemotePort = txtPuerto.Text
sock(iSockets).RemoteHost = txtServidor.Text
sock(iSockets).Connect
End If
End Sub

Private Sub mirrorm2_Click()
frmLogin.Show vbModal, Me
End Sub

Private Sub openswin_Click()
frmMain.Visible = True
End Sub

Private Sub s1_Click()
If s1.Checked = True Then Exit Sub
s1.Checked = True
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = True
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = True
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "ONL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(En línea)"
If Label5.Width > 1500 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image8.Picture
nid.szTip = "IntraMessenger - Sesión iniciada" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
If tmrFueras.Enabled = False Then
TiempoEspera = 0
tmrFueras.Enabled = True
End If
End Sub

Private Sub s2_Click()
If s2.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = True
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = True
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "LUN" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Salí a comer)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Salí a comer" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub s3_Click()
If s3.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = True
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = True
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "AWA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Vuelvo enseguida)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Vuelvo enseguida" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub s4_Click()
If s4.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = True
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = True
m8.Checked = False
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = False
Image24.Visible = True
Image25.Visible = False
sock(iSockets).SendData "STATE" & "ONP" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Al teléfono)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - Al teléfono" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub s5_Click()
If s5.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = True
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = True
m9.Checked = False
m10.Checked = False
Image19.Visible = False
Image23.Visible = True
Image24.Visible = False
Image25.Visible = False
sock(iSockets).SendData "STATE" & "IDL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Ausente)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Ausente" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub s6_Click()
If s6.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = True
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = True
m10.Checked = False
Image19.Visible = False
Image23.Visible = False
Image24.Visible = True
Image25.Visible = False
sock(iSockets).SendData "STATE" & "NOA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(No disponible)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - No disponible" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub s7_Click()
If s7.Checked = True Then Exit Sub
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = True
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = True
Image19.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = True
sock(iSockets).SendData "STATE" & "OFF" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Sin conexión)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
If tmrAnim.Enabled = False Then
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image16.Picture
nid.szTip = "IntraMessenger - Sin conexión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End If
TiempoEspera = 0
tmrFueras.Enabled = False
End Sub

Private Sub sock_Close(Index As Integer)
On Error Resume Next
sock(iSockets).Close
iSockets = iSockets - 1
If iSockets = -1 Then
iSockets = 0
End If
'MÓDULO EXPERIMENTAL (CIERRA LAS VENTANAS AL CERRAR SESIÓN)
Dim i
For i = 1 To contactos.Nodes.Count
If contactos.Nodes.Item(i).Image = 1 Or contactos.Nodes.Item(i).Image = 2 Or contactos.Nodes.Item(i).Image = 3 Or contactos.Nodes.Item(i).Image = 5 Then
If frmChat(i).Visible = True Then
Unload frmChat(i)
End If
End If
Next i
DoEvents
'MÓDULO EXPERIMENTAL
Command1.Visible = True
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
mirrorm3.Enabled = False
tmrIconica.Enabled = False
tmrAnim.Enabled = False
tmrAnim.Enabled = False
Label5.Visible = False
Label7.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Label6.Visible = False
contactos.Visible = False
tmrFueras.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
m1.Caption = "Iniciar sesión como " & txtUsuario.Text
Label1.Caption = txtUsuario.Text
mirrorm1.Caption = "Iniciar sesión como " & txtUsuario.Text
Contador = 0
MsgBox "El servicio se ha suspendido porque el servidor ha cerrado el acceso."
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
mirrorm3.Enabled = False
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
End Sub

Private Sub sock_Connect(Index As Integer)
sock(iSockets).SendData "LOGIN" & "$$$" & txtUsuario.Text & "&&&" & txtClave.Text
tmrTemporizador.Enabled = False
estoyOn.Text = 1
End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error Resume Next
If Me.WindowState = vbMinimized Then
Me.WindowState = vbNormal
End If
Dim strData As String
sock(iSockets).GetData strData, vbString
'Comportamiento cuando el servidor acepta la solicitud
Dim inicializacion As String
inicializacion = Mid(strData, 1, 5)
Select Case inicializacion
Case "ALLOW"
sock(iSockets).SendData "LISTA" & "$$$" & txtUsuario.Text
contactos.Nodes.Clear
Command1.Visible = False
Image20.Visible = True
Image19.Visible = True
Image1.Visible = False
Image2.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = True
Picture2.Visible = True
Frame1.Visible = False
Frame2.Visible = False
m3.Enabled = True
m4.Enabled = True
m16.Enabled = True
m5.Checked = True
s1.Checked = True
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
contactos.Visible = True
tmrFueras.Enabled = True
m8.Checked = False
m9.Checked = False
m10.Checked = False
mirrorm3.Enabled = True
m2.Visible = True
m1.Enabled = False
m2.Enabled = False
m01.Visible = False
m1.Visible = True
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = False
mirrorm01.Visible = False
tmrIconica.Enabled = False
tmrAnim.Enabled = True
Label5.Visible = True
Label7.Visible = True
Label6.Visible = True
estoyOn.Text = 1
'Módulo de envío del estado actual (En línea por defecto)
sock(iSockets).SendData "STATE" & "ONL" & "$$$" & txtUsuario.Text
'Comportamiento cuando el servidor rechaza la solicitud
Case "ERROR"
contactos.Nodes.Clear
Command1.Visible = True
Image20.Visible = False
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
MsgBox "Error al frmLogin sesión: usuario o contraseña incorrectos."
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
'Comportamiento cuando el servidor detecta doble inicio de sesión
Case "ALREA"
contactos.Nodes.Clear
Command1.Visible = True
Image20.Visible = False
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
MsgBox "No puede iniciar sesión dos veces al mismo tiempo."
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm3.Enabled = False
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
End Select
'Aquí van TODAS las instrucciones de recibo de datos
If estoyOn.Text = 1 Then
'RECEPCIÓN PRIMARIA DE DATOS (ALIAS E IMAGEN A MOSTRAR)
Dim datosEntrada As String
datosEntrada = strData
datosEntrada = Mid(datosEntrada, 6, 5)
Select Case datosEntrada
Case "ALIAS"
txtMyalias.Text = Mid(strData, InStr(strData, "%%%") + 3)
Label5.Caption = txtMyalias.Text
Label6.Caption = "(En línea)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
Case "IMAGE"
End Select
'RECEPCIÓN SECUNDARIA DE DATOS (LISTA DE CONTACTOS, EN LÍNEA, DESCONECTADOS, ETC...)
Dim datosSecundarios As String
datosSecundarios = strData
datosSecundarios = Mid(datosSecundarios, 1, 5)
Select Case datosSecundarios
Case "LISTA"
contactos.ImageList = imagenes
Dim i As Integer
Dim Lista As String
Dim AgregarNuevo  As Variant
Lista = Mid(strData, 7)
Debug.Print Lista
contactos.Nodes.Clear
contactos.Nodes.Add(, , "AMI", "Usuarios (0/0)", 7).Bold = True
contactos.Nodes.Item(1).ForeColor = &H814D3C
'Módulo experimental
'###################
Dim EstadoActual As String
For i = 0 To UBound(Split(Lista, "_")) - 1
AgregarNuevo = Split(Lista, "_")
EstadoActual = Trim(Right(AgregarNuevo(i), 3))
AgregarNuevo(i) = Mid(AgregarNuevo(i), 1, Len(AgregarNuevo(i)) - 3)
'contactos.Nodes(1).Text = "Usuarios (" & Globalnline & "/" & contactos.Nodes.Item(1).Children & ")"
Select Case EstadoActual
Case "OFF" 'Mostrar el icono cuando el usuario está offline
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Sin conexión)", 4
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Sin conexión)", 4
End If
Case "ONL" 'Mostrar el icono cuando el estado es online
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (En línea)", 1
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (En línea)", 1
End If
Case "NOA" 'Mostrar el icono cuando el estado es no disponible
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (No disponible)", 3
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (No disponible)", 3
End If
Case "IDL" 'Mostrar el icono cuando el estado es ausente
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Ausente)", 2
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Ausente)", 2
End If
Case "AWA" 'Mostrar el icono cuando el estado es vuelvo enseguida
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Vuelvo enseguida)", 2
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Vuelvo enseguida)", 2
End If
Case "LUN" 'Mostrar el icono cuando el estado es salí a comer
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Salí a comer)", 2
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Salí a comer)", 2
End If
Case "ONP" 'Mostrar el icono cuando el estado es al teléfono
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Al teléfono)", 3
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Al teléfono)", 3
End If
Case "BLO" 'Mostrar el icono cuando el estado es sin admisión
If i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Sin admisión)", 5
ElseIf Not i = 0 Then
contactos.Nodes.Add "AMI", tvwChild, , AgregarNuevo(i) & " (Sin admisión)", 5
End If
End Select
Next i
'VARIAS
contactos.Nodes(1).Expanded = True
sock(iSockets).SendData "STATE" & "ONL" & "$$$" & txtUsuario.Text
On Local Error Resume Next
'ELIMINAR EL NOMBRE DE USUARIO DE LA LISTA
Dim k
For k = 2 To contactos.Nodes.Count
If Trim(Mid(contactos.Nodes(k).Text, 1, Len(Label5.Caption))) = Label5.Caption Then
contactos.Nodes.Remove (k)
End If
Next k
'Módulo experimental (Cambia el nick del usuario, sin alterar su estado)
'###################
Case "ALIAS"
txtMyalias.Text = Mid(strData, InStr(strData, "%%%") + 3)
Label5.Caption = txtMyalias.Text
Dim EstadoReal As String
Dim EstadoEnviar As String
If m5.Checked = True Then
EstadoReal = "(" & m5.Caption & ")"
EstadoEnviar = "ONL"
ElseIf m6.Checked = True Then
EstadoReal = "(" & m6.Caption & ")"
EstadoEnviar = "LUN"
ElseIf m7.Checked = True Then
EstadoReal = "(" & m7.Caption & ")"
EstadoEnviar = "AWA"
ElseIf m7a.Checked = True Then
EstadoReal = "(" & m7a.Caption & ")"
EstadoEnviar = "ONP"
ElseIf m8.Checked = True Then
EstadoReal = "(" & m8.Caption & ")"
EstadoEnviar = "IDL"
ElseIf m9.Checked = True Then
EstadoReal = "(" & m9.Caption & ")"
EstadoEnviar = "NOA"
ElseIf m10.Checked = True Then
EstadoReal = "(" & m10.Caption & ")"
EstadoEnviar = "OFF"
End If
sock(iSockets).SendData "STATE" & EstadoEnviar & "$$$" & txtUsuario.Text
Label6.Caption = EstadoReal
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
Case "USREG"
If frmNewContact.Visible = True Then
Call frmNewContact.satisfactorio
End If
Case "EXIST"
If frmNewContact.Visible = True Then
Call frmNewContact.existe
End If
Case "UPDAT" '(Cambia el estado del usuario, sin afectar al nick)
Dim ParseAlias As String
Dim ParseEstado As String
ParseAlias = Mid(strData, InStr(strData, "$$$") + 3)
ParseAlias = Mid(ParseAlias, 1, Len(ParseAlias) - 3)
ParseEstado = Right(strData, 3)
Dim nodos
For nodos = 1 To contactos.Nodes.Count
If Trim(Mid(contactos.Nodes(nodos).Text, 1, InStrRev(contactos.Nodes(nodos).Text, "(") - 1)) = Trim(ParseAlias) Then
Select Case ParseEstado
Case "OFF"
contactos.Nodes(nodos).Image = 4
contactos.Nodes(nodos).Text = ParseAlias & " (Sin conexión)"
Case "ONL"
contactos.Nodes(nodos).Image = 1
contactos.Nodes(nodos).Text = ParseAlias & " (En línea)"
If m9.Checked = True Or m10.Checked = True Then
Else
If frmOpciones.Visible = False And frmAbout.Visible = False And frmNewContact.Visible = False Then
If Text1.Text = "true" Then
frmPopup.Show
End If
If Text3.Text = "true" Then
sndPlaySound Text6.Text, SND_ASYNC Or SND_NODEFAULT
End If
Contador = 0
tmrAnim.Enabled = True
End If
End If
Case "LUN"
contactos.Nodes(nodos).Image = 2
contactos.Nodes(nodos).Text = ParseAlias & " (Salí a comer)"
Case "AWA"
contactos.Nodes(nodos).Image = 2
contactos.Nodes(nodos).Text = ParseAlias & " (Vuelvo enseguida)"
Case "NOA"
contactos.Nodes(nodos).Image = 3
contactos.Nodes(nodos).Text = ParseAlias & " (No disponible)"
Case "IDL"
contactos.Nodes(nodos).Image = 2
contactos.Nodes(nodos).Text = ParseAlias & " (Ausente)"
Case "ONP"
contactos.Nodes(nodos).Image = 3
contactos.Nodes(nodos).Text = ParseAlias & " (Al teléfono)"
End Select
End If
Next nodos
'TAREAS VARIAS (Cuenta y devuelve el número de usuarios conectados)
contactos.Nodes(1).Text = "Usuarios (0/" & contactos.Nodes(1).Children & ")"
Dim RX
For RX = 1 To contactos.Nodes.Count
Dim GlobalOline As Integer
If contactos.Nodes.Item(RX).Image = 1 Or contactos.Nodes.Item(RX).Image = 2 Or contactos.Nodes.Item(RX).Image = 3 Or contactos.Nodes.Item(RX).Image = 5 Then
GlobalOline = GlobalOline + 1
contactos.Nodes(1).Text = "Usuarios (" & GlobalOline & "/" & contactos.Nodes(1).Children & ")"
End If
Next RX
Case "UPALI" 'Cambio de alias de algún usuario
'On Local Error Resume Next
Dim BuscaEste As String
Dim CambiaPorEste As String
Dim EstadoUpdater As String
Dim erIndex As Integer
BuscaEste = Trim(Mid(strData, InStr(strData, "OLD") + 3, InStr(strData, "NEW") - 9))
CambiaPorEste = Trim(Mid(strData, InStr(strData, "NEW") + 3))
Dim look
For look = 1 To contactos.Nodes.Count
If Trim(Mid(contactos.Nodes(look).Text, 1, Len(BuscaEste))) = BuscaEste Then
EstadoUpdater = Trim(Mid(contactos.Nodes(look).Text, InStrRev(contactos.Nodes(look).Text, "(") - 1, InStrRev(contactos.Nodes(look).Text, ")") + 1))
contactos.Nodes(look).Text = CambiaPorEste & " " & EstadoUpdater
erIndex = contactos.Nodes(look).Index
frmChat(erIndex).Caption = CambiaPorEste & " - Conversación"
frmChat(erIndex).Label2.Caption = CambiaPorEste
End If
Next look
Case "IMDAT" 'Recepción de texto en una conversación
Dim QuienEnvia As String
Dim CuerpoMSG As String
QuienEnvia = strData
QuienEnvia = Mid(QuienEnvia, InStr(QuienEnvia, "SND") + 3, InStr(QuienEnvia, "MSG") - 9)
CuerpoMSG = strData
CuerpoMSG = Mid(CuerpoMSG, InStr(CuerpoMSG, "MSG") + 3)
Dim buscaForm
Dim indexUser
For buscaForm = 1 To contactos.Nodes.Count
If QuienEnvia = Trim(Mid(contactos.Nodes(buscaForm).Text, 1, Len(QuienEnvia))) Then
indexUser = contactos.Nodes(buscaForm).Index
If frmChat(indexUser).Visible = True Then
Else
If frmAbout.Visible = True Then
Unload frmAbout
End If
If frmNewContact.Visible = True Then
Unload frmNewContact
End If
If frmOpciones.Visible = True Then
Unload frmOpciones
End If
If Text2.Text = "true" Then
frmPopup.Show
End If
Load frmChat(indexUser)
frmChat(indexUser).WindowState = 1
frmChat(indexUser).Show
If Text3.Text = "true" Then
sndPlaySound Text4.Text, SND_ASYNC Or SND_NODEFAULT
End If
Contador = 0
tmrAnim.Enabled = True
End If
End If
Next buscaForm
On Error Resume Next
'TRATAMIENT0 DE TEXTO - REVISADO -
frmChat(indexUser).bandeja.Visible = False
frmChat(indexUser).bandeja.SelStart = Len(frmChat(indexUser).bandeja.Text)
frmChat(indexUser).bandeja.SelStrikeThru = False
frmChat(indexUser).bandeja.SelColor = RGB(150, 150, 150)
frmChat(indexUser).bandeja.SelFontSize = 10
frmChat(indexUser).bandeja.SelText = Replace(QuienEnvia, Chr(12), vbCrLf) & " dice: " & vbCrLf
frmChat(indexUser).bandeja.SelStart = Len(frmChat(indexUser).bandeja.Text)
frmChat(indexUser).bandeja.SelStrikeThru = False
frmChat(indexUser).bandeja.SelText = CuerpoMSG & vbCrLf
frmChat(indexUser).bandeja.SelStart = Len(frmChat(indexUser).bandeja.Text)
frmChat(indexUser).bandeja.SelStrikeThru = False
frmChat(indexUser).CheckSmileys Len(CuerpoMSG) + 64
frmChat(indexUser).bandeja.Visible = True
frmChat(indexUser).Caption = QuienEnvia & " - Conversación"
frmChat(indexUser).Label3.Caption = "Último mensaje recibido el " & Format(Date) & " a las " & Format(Time)
'TRATAMIENT0 DE TEXTO - REVISADO -
DoEvents
Case "IMWRI" 'Indica si un usuario está escribiendo un mensaje (Este módulo aún NO está en funcionamiento)
Dim buscaForms
Dim indexUsers
Dim QuienSend As String
QuienSend = strData
QuienSend = Mid(QuienSend, InStr(QuienSend, "SND") + 3)
For buscaForms = 1 To contactos.Nodes.Count
If QuienSend = Trim(Mid(contactos.Nodes(buscaForms).Text, 1, Len(QuienSend))) Then
indexUsers = contactos.Nodes(buscaForms).Index
End If
Next buscaForms
QuienSend = strData
QuienSend = Mid(QuienSend, InStr(QuienSend, "SND") + 3)
If Trim(Mid(strData, 1, 5)) = "IMWRI" Then
DoEvents
frmChat(indexUsers).Label3.Caption = QuienSend & " está escribiendo un mensaje." & CuerpoMSG
Else
DoEvents
frmChat(indexUsers).Label3.Caption = QuienSend & " NO está escribiendo un mensaje." & CuerpoMSG
End If
End Select
End If
End Sub

Private Sub st1_Click()
s1.Checked = True
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = True
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
sock(iSockets).SendData "STATE" & "ONL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(En línea)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image8.Picture
nid.szTip = "IntraMessenger - Sesión iniciada" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st2_Click()
s1.Checked = False
s2.Checked = True
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = True
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
sock(iSockets).SendData "STATE" & "LUN" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Salí a comer)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Salí a comer" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st3_Click()
s1.Checked = False
s2.Checked = False
s3.Checked = True
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = True
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
sock(iSockets).SendData "STATE" & "AWA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Vuelvo enseguida)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Vuelvo enseguida" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st4_Click()
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = True
s5.Checked = False
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = True
m8.Checked = False
m9.Checked = False
m10.Checked = False
sock(iSockets).SendData "STATE" & "ONP" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Al teléfono)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image17.Picture
nid.szTip = "IntraMessenger - Al teléfono" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st5_Click()
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = True
s6.Checked = False
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = True
m9.Checked = False
m10.Checked = False
sock(iSockets).SendData "STATE" & "IDL" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Ausente)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - Ausente" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st6_Click()
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = True
s7.Checked = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = True
m10.Checked = False
sock(iSockets).SendData "STATE" & "NOA" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(No disponible)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image15.Picture
nid.szTip = "IntraMessenger - No disponible" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub st7_Click()
s1.Checked = False
s2.Checked = False
s3.Checked = False
s4.Checked = False
s5.Checked = False
s6.Checked = False
s7.Checked = True
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = True
sock(iSockets).SendData "STATE" & "OFF" & "$$$" & txtUsuario.Text
Label5.Caption = txtMyalias.Text
Label6.Caption = "(Sin conexión)"
If Label5.Width >= 2000 And Label5.Width <= 3165 Then
Label6.Left = Label5.Width - Label6.Width + Label5.Left - 35
Else
Label6.Left = Label7.Left + Label7.Width + 55
End If
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image16.Picture
nid.szTip = "IntraMessenger - Sin conexión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub tmrTemporizador_Timer()
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
mirrorm3.Enabled = False
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
Dim permuta
permuta = MsgBox("No se ha podido iniciar sesión, compruebe los parámetros y la conexión.", vbInformation + vbRetryCancel, "IntraMessenger")
If permuta = vbRetry Then
Call llamada
ElseIf permuta = vbCancel Then
Command1.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = False
Timer1.Enabled = False
Picture1.Visible = False
Image7.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
m01.Visible = False
m1.Visible = True
m2.Visible = True
m1.Enabled = True
m2.Enabled = True
m3.Enabled = False
If frmPopup.Visible = True Then
Unload frmPopup
End If
mirrorm3.Enabled = False
m4.Enabled = False
m16.Enabled = False
m5.Checked = False
m6.Checked = False
m7.Checked = False
m7a.Checked = False
m8.Checked = False
m9.Checked = False
m10.Checked = False
Contador = 0
mirrorm01.Visible = False
mirrorm1.Visible = True
mirrorm2.Visible = True
mirrorm1.Enabled = True
mirrorm2.Enabled = True
tmrIconica.Enabled = False
tmrAnim.Enabled = False
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image5.Picture
nid.szTip = "IntraMessenger - No ha iniciado sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
sock(iSockets).Close
tmrTemporizador.Enabled = False
estoyOn.Text = 0
End If
End Sub

Private Sub Timer1_Timer()
Picture1.Left = Picture1.Left - 720
Picture2.Left = Picture2.Left - 720
If Picture1.Left <= -23040 Then
Picture1.Left = Picture2.Left + Picture2.Width
End If
If Picture2.Left <= -23040 Then
Picture2.Left = Picture1.Left + Picture1.Width
End If
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Public Sub llamada()
On Error Resume Next
Command1.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = True
Timer1.Enabled = True
Picture1.Visible = True
Image7.Visible = True
Picture2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
m1.Visible = False
m2.Enabled = False
m01.Visible = True
mirrorm01.Visible = True
mirrorm2.Enabled = False
mirrorm1.Visible = False
breaker = 0
tmrIconica.Enabled = True
nid.cbSize = Len(nid)
nid.hWnd = frmMain.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Image6.Picture
nid.szTip = "IntraMessenger - Iniciando sesión" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
tmrTemporizador.Enabled = True
tmrTemporizador.Enabled = True
Load sock(iSockets)
sock(iSockets).RemotePort = txtPuerto.Text
sock(iSockets).RemoteHost = txtServidor.Text
sock(iSockets).Connect
End Sub

Public Sub reload()
txtPuerto.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "port")
txtServidor.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "server")
End Sub

Private Function AnotherInstance() As Integer
Dim AppTitle$
If App.PrevInstance Then
AppTitle$ = App.Title
App.Title = "IntraMessenger"
AppActivate AppTitle$
AnotherInstance = True
Else
AnotherInstance = False
End If
End Function

Public Sub recargar()
txtTimeout.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "txtTimeout")
Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsContacts")
Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsMessages")
Text3.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Sounds")
Text4.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoMsg")
Text5.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoNot")
Text6.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "UsrOnl")
End Sub
