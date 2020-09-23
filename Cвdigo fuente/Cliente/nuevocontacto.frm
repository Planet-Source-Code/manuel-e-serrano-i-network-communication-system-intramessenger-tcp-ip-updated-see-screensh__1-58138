VERSION 5.00
Begin VB.Form frmNewContact 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar nuevo contacto"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "nuevocontacto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox usuario 
      Height          =   315
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   0
      Top             =   690
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar un nuevo contacto a su libreta de direcciones."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5640
   End
End
Attribute VB_Name = "frmNewContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Trim(txtUsuario.Text) = "" Then
MsgBox "Debe indicar un nombre de usuario para continuar.", vbInformation, "Información"
txtUsuario.SetFocus
Exit Sub
End If
frmMain.sock(iSockets).SendData "AGREG" & "$$$" & frmMain.txtUsuario.Text & ">=<" & txtUsuario.Text
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_.1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
Beep
KeyAscii = 0
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Public Sub satisfactorio()
Dim pregunta
pregunta = MsgBox("El usuario " & txtUsuario.Text & " ha sido agregado a su libreta de direcciones." & vbCrLf & "¿Desea agregar un nuevo contacto?", vbInformation + vbYesNo, "Información")
If pregunta = vbYes Then
txtUsuario.Text = ""
txtUsuario.SetFocus
Else
frmMain.sock(iSockets).SendData "LISTA" & "$$$" & frmMain.txtUsuario.Text
Unload Me
End If
End Sub

Public Sub existe()
MsgBox "El contacto ya existe en su libreta de direcciones.", vbInformation, "Información"
Exit Sub
End Sub
