VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio de mensajería"
   ClientHeight    =   2895
   ClientLeft      =   8820
   ClientTop       =   3030
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "iniciar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registrar un nuevo usuario"
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
      Height          =   210
      Left            =   240
      MouseIcon       =   "iniciar.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2400
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicie sesión con su dirección asignada para poder recibir y enviar mensajes, transferir archivos etc."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Private Sub Command1_Click()
If Trim(Text1.Text) = "" Then
MsgBox "Debe escribir el nombre de txtUsuario.", vbInformation, "Información"
Text1.SetFocus
Exit Sub
End If
If Trim(Text2.Text) = "" Then
MsgBox "Debe escribir la contraseña para continuar.", vbInformation, "Información"
Text2.SetFocus
Exit Sub
End If
frmMain.Label1.Caption = Text1.Text
frmMain.txtUsuario.Text = LCase(Text1.Text)
frmMain.txtClave.Text = Text2.Text
frmMain.llamada
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Label4_Click()
frmWizard.Show vbModal, Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_.1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
Beep
KeyAscii = 0
End If
End Sub
