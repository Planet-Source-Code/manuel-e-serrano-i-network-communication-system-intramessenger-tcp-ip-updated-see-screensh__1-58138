VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWizard 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asistente para IntraNET Messenger Services"
   ClientHeight    =   3255
   ClientLeft      =   5895
   ClientTop       =   2565
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "asistente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tester 
      Left            =   480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   0
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox otraclave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox clave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox usuario 
      Height          =   315
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña:"
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario:"
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"asistente.frx":000C
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Suscribase ahora para tener acceso a los servicios de mensajería instantánea de IntraNET Messenger Service."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Dim cerrar As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo solveErr
If Trim(txtUsuario.Text) = "" Then
MsgBox "Debe indicar un nombre de usuario para continuar.", vbInformation, "Información"
txtUsuario.SetFocus
Exit Sub
End If
If Trim(txtClave.Text) = "" Then
MsgBox "Debe escribir la contraseña para continuar.", vbInformation, "Información"
txtClave.SetFocus
Exit Sub
End If
If Trim(otraclave.Text) = "" Then
MsgBox "Debe escribir la confirmación de la contraseña para continuar.", vbInformation, "Información"
otraclave.SetFocus
Exit Sub
End If
If txtClave.Text <> otraclave.Text Then
MsgBox "Las contraseñas no coinciden, por favor intente de nuevo.", vbExclamation, "Error en contraseñas"
txtClave.SelStart = 0
txtClave.SelLength = Len(txtClave.Text)
txtClave.SetFocus
Exit Sub
End If
socket.SendData "NUEVO" & "$$$" & LCase(txtUsuario.Text) & "&&&" & txtClave.Text
solveErr:
If Err.Number = 40006 Then
MsgBox "Imposible registrar usuario, puede que no haya conexión con el servidor.", vbExclamation, "Error fatal"
Exit Sub
End If
End Sub

Private Sub Form_Load()
DoEvents
socket.RemotePort = frmMain.txtPuerto.Text
socket.RemoteHost = frmMain.txtServidor.Text
socket.Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cerrar = 0 Then
Dim advertencia
advertencia = MsgBox("El proceso de suscripción no ha finalizado. ¿Desea cerrar el aistente?", vbQuestion + vbYesNo, "Pregunta")
If advertencia = vbYes Then
cerrar = 1
Cancel = 0
Unload Me
Else
cerrar = 0
Cancel = 1
Exit Sub
End If
End If
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
socket.GetData strData, vbString
Dim resultado As String
resultado = Mid(strData, 1, 5)
Select Case resultado
Case "USREG"
Dim permuta
permuta = MsgBox("El usuario " & txtUsuario.Text & " ha sido registrado satisfactoriamente." & vbCrLf & "¿Desea frmLogin sesión ahora?", vbInformation + vbYesNo, "Información")
If permuta = vbYes Then
frmMain.txtUsuario.Text = txtUsuario.Text
frmMain.txtClave.Text = txtClave.Text
cerrar = 1
Unload Me
Unload frmLogin
Call frmMain.llamada
Unload Me
Else
cerrar = 1
Unload Me
End If
Case "ERROR"
MsgBox "Error al registrar nuevo usuario, puede estar usando un nombre ya reservado.", vbExclamation, "Error al registrar"
Exit Sub
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Function PortInUse(ByVal PortNumber As Integer) As Boolean
Dim bAns As Boolean
On Error Resume Next
tester.LocalPort = frmMain.txtPuerto.Text
tester.Listen
bAns = Err.Number = 10048
tester.Close
PortInUse = bAns
End Function

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_.1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
Beep
KeyAscii = 0
End If
End Sub
