VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSonidos 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especificar sonidos"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   1620
   ClientWidth     =   6255
   Icon            =   "sonidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command8 
      Caption         =   ">>"
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   460
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">>"
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   820
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>"
      Height          =   315
      Left            =   4440
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1190
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Examinar"
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   820
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Examinar"
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   460
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Examinar"
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Top             =   1190
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00EEF3F4&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00EEF3F4&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00EEF3F4&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo mensaje:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo aviso:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto en línea:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar sonidos de notificación"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmSonidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_LOOP = &H8
Private Const SND_NODEFAULT = &H2

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoMsg", Text2.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "UsrOnl", Text1.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoNot", Text3.Text
frmMain.recargar
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo Exits
With dialogo
.DialogTitle = "Abrir"
.DefaultExt = "wav"
.Filter = "Sonidos (*.wav)|*.wav|"
.Flags = cdlOFNFileMustExist
.ShowOpen
End With
Text3.Text = dialogo.FileName
Exits:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo Exits
With dialogo
.DialogTitle = "Abrir"
.DefaultExt = "wav"
.Filter = "Sonidos (*.wav)|*.wav|"
.Flags = cdlOFNFileMustExist
.ShowOpen
End With
Text1.Text = dialogo.FileName
Exits:
Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo Exits
With dialogo
.DialogTitle = "Abrir"
.DefaultExt = "wav"
.Filter = "Sonidos (*.wav)|*.wav|"
.Flags = cdlOFNFileMustExist
.ShowOpen
End With
Text2.Text = dialogo.FileName
Exits:
Exit Sub
End Sub

Private Sub Command6_Click()
sndPlaySound Text3.Text, SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Command7_Click()
sndPlaySound Text2.Text, SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Command8_Click()
sndPlaySound Text1.Text, SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Form_Load()
Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoMsg")
Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "UsrOnl")
Text3.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "NvoNot")
If Trim(Text2.Text) = "" Or Text2.Text = "Error" Then
Text2.Text = App.Path & "\type.wav"
End If
If Trim(Text1.Text) = "" Or Text1.Text = "Error" Then
Text1.Text = App.Path & "\online.wav"
End If
If Trim(Text3.Text) = "" Or Text3.Text = "Error" Then
Text3.Text = App.Path & "\newalert.wav"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound vbNullString, SND_ASYNC
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Text1_GotFocus()
Command8.Default = True
End Sub

Private Sub Text1_LostFocus()
Command8.Default = False
End Sub

Private Sub Text2_GotFocus()
Command7.Default = True
End Sub

Private Sub Text2_LostFocus()
Command7.Default = False
End Sub

Private Sub Text3_GotFocus()
Command6.Default = True
End Sub

Private Sub Text3_LostFocus()
Command6.Default = False
End Sub
