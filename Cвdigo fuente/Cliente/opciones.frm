VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00EEF3F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   4335
   ClientLeft      =   8655
   ClientTop       =   3030
   ClientWidth     =   7215
   Icon            =   "opciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5400
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox buffert3 
      Height          =   285
      Left            =   2760
      TabIndex        =   56
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1560
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2040
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2520
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3000
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3480
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3960
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4440
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4920
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox t5 
      Height          =   285
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   40
      Top             =   2490
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox t4 
      Height          =   285
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   39
      Top             =   2130
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Co6 
      Caption         =   "Cambiar..."
      Height          =   375
      Left            =   5880
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox c10 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Reproducir un sonido cuando los contactos inicien sesión o envíen un mensaje"
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   2595
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CheckBox c9 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Mostrar avisos al recibir un mensaje"
      Height          =   375
      Left            =   2040
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox c8 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Mostrar avisos cuando se conecten los contactos"
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CheckBox c7 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Abrir la ventana principal de IntraMessenger al iniciar este programa"
      Height          =   375
      Left            =   2040
      TabIndex        =   32
      Top             =   1035
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CheckBox c6 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Ejecutar IntraMessegner automáticamente cuando abro sesión en Windows"
      Height          =   375
      Left            =   2040
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CheckBox c5 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Guardar automáticamente un historial de mis conversaciones"
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Co5 
      Caption         =   "Cambiar..."
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2475
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox t3 
      BackColor       =   &H00EEF3F4&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CheckBox c4 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Mostrar iconos gestuales en mensajes instantáneos"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Co4 
      Caption         =   "Cambiar fuente..."
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox t2 
      Height          =   285
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   15
      Text            =   "5"
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox c3 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Mostrarme ""Ausente"" si estoy inactivo "
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox c2 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Visualizar la imagen para mostrar de los demás en conversaciones de mensajes instantáneos"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton co3 
      Caption         =   "Cambiar imagen..."
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox c1 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Visualizar mi imagen para mostrar y permitir a los demás la vean"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox t1 
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
      Begin VB.Image Image7 
         Height          =   480
         Left            =   360
         Picture         =   "opciones.frx":000C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Conexión"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
      Begin VB.Image Image6 
         Height          =   480
         Left            =   360
         Picture         =   "opciones.frx":08D6
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
      Begin VB.Image Image5 
         Height          =   480
         Left            =   360
         Picture         =   "opciones.frx":11A0
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensajes"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   360
         Picture         =   "opciones.frx":1A6A
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label l18 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre servidor:"
      Height          =   255
      Left            =   1320
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label l17 
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto:"
      Height          =   255
      Left            =   1320
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label l16 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualmente NO está conectado a IntraNET Messenger Service."
      Height          =   255
      Left            =   1320
      TabIndex        =   44
      Top             =   1680
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label l15 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración actual"
      Height          =   255
      Left            =   1320
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label l14 
      BackColor       =   &H00EEF3F4&
      BackStyle       =   0  'Transparent
      Caption         =   $"opciones.frx":2334
      Height          =   495
      Left            =   2040
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image i10 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":23C2
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l13 
      BackStyle       =   0  'Transparent
      Caption         =   "Preferencias de conexión"
      Height          =   255
      Left            =   1320
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image i9 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":2C8C
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l12 
      BackStyle       =   0  'Transparent
      Caption         =   "Avisos"
      Height          =   255
      Left            =   1320
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image i8 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":3556
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l11 
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar sesión"
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image i7 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":3E20
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l10 
      BackStyle       =   0  'Transparent
      Caption         =   "Historial de mensajes"
      Height          =   255
      Left            =   1320
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label l9 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Colocar los archivos recibidos de otros usuarios en esta carpeta:"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image i6 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":46EA
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transferencia de archivos"
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label l7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar la fuente y la dirección de mis mensajes instantáneos"
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image i4 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":4FB4
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l6 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del mensaje"
      Height          =   255
      Left            =   1320
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image i5 
      Height          =   285
      Left            =   5880
      Picture         =   "opciones.frx":587E
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label l5 
      BackStyle       =   0  'Transparent
      Caption         =   "minutos"
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i3 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":5E36
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mi estado"
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i2 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":6700
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mi imagen para mostrar"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image i1 
      Height          =   480
      Left            =   1320
      Picture         =   "opciones.frx":6FCA
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label l2 
      BackColor       =   &H00EEF3F4&
      Caption         =   "Escriba su nombre tal y como desea que lo vean los demás usuarios."
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mi nombre para mostrar"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000014&
      BorderColor     =   &H80000014&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000014&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000014&
      BorderColor     =   &H80000014&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000014&
      BorderColor     =   &H80000014&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Dim auto As String
Dim presionado1 As Integer
Dim presionado2 As Integer
Dim presionado3 As Integer
Dim presionado4 As Integer
Dim verMiImagen As String
Dim verImagen As String
Dim Ausencia As String
Dim VerIconos As String
Dim MiCarpeta As String
Dim GuardarHistorial As String
Dim EjecutarIntraMessenger As String
Dim ModoSilencioso As String
Dim MostrarAvisos As String
Dim AvisarMensajes As String
Dim txtPuerto As String
Dim txtServidor As String
Dim TiempoEspera As String
Dim Sonidos As String
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Type BrowseInfo
hwndOwner      As Long
pIDLRoot       As Long
pszDisplayName As Long
lpszTitle      As Long
ulFlags        As Long
lpfnCallback   As Long
lParam         As Long
iImage         As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Sub c1_Click()
If c1.Value = 1 Then
If frmMain.estoyOn.Text = "1" Then
co3.Enabled = True
Else
co3.Enabled = False
End If
Else
co3.Enabled = False
End If
End Sub

Private Sub c10_Click()
If c10.Value = vbChecked Then
Co6.Enabled = True
Else
Co6.Enabled = False
End If
End Sub

Private Sub c3_Click()
If c3.Value = 0 Then
t2.Enabled = False
t2.BackColor = &HEEF3F4
Else
t2.Enabled = True
t2.BackColor = &H80000005
End If
End Sub

Private Sub Co5_Click()
buffert3.Text = t3.Text
t3.Text = OpenDirectoryTV(frmOpciones, "Seleccione otra carpeta para guardar los archivos recibidos.")
If Trim(t3.Text) = "" Then
t3.Text = buffert3.Text
End If
End Sub

Public Function OpenDirectoryTV(odtvOwner As Form, Optional odtvTitle As String) As String
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
szTitle = odtvTitle
With tBrowseInfo
.hwndOwner = odtvOwner.hWnd
.lpszTitle = lstrcat(szTitle, "")
.ulFlags = BIF_RETURNONLYFSDIRS
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
sBuffer = Space(MAX_PATH)
SHGetPathFromIDList lpIDList, sBuffer
sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
OpenDirectoryTV = sBuffer
End If
End Function

Private Sub Co6_Click()
frmSonidos.Show vbModal, Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim porDefecto As String
auto = frmMain.txtUsuario.Text
If Trim(auto) = "" Then
auto = "No especificado"
End If
Dim j
For j = 1 To frmMain.contactos.Nodes.Count
If LCase(Trim(Mid(frmMain.contactos.Nodes(j).Text, 1, InStr(frmMain.contactos.Nodes(j).Text, "(") - 1))) = LCase(t1.Text) Then
MsgBox "No puede especificar un nombre utilizado por otro usuario", vbInformation, "Información"
Exit Sub
End If
Next j
If Trim(t1.Text) = "" Then
t1.Text = auto
End If
porDefecto = App.Path & "\" & "Mis archivos recibidos"
'Guardar todas las configuraciones
If c6.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "RunStartup", "true"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "IntraMessenger", App.Path & "\" & App.EXEName
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "RunStartup", "false"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "IntraMessenger", " "
End If
If c7.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SilentMode", "false"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SilentMode", "true"
End If
If c8.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsContacts", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsContacts", "false"
End If
If c9.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsMessages", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsMessages", "false"
End If
If c1.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowMyImage", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowMyImage", "false"
End If
If c2.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowImages", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowImages", "false"
End If
If c3.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Absent", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Absent", "false"
End If
If c4.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowIcons", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowIcons", "false"
End If
If c5.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SaveHistory", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SaveHistory", "false"
End If
If c10.Value = vbChecked Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Sounds", "true"
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Sounds", "false"
End If
'''''''''''''''''''''''''''
If Trim(t3.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "path", t3.Text
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "path", porDefecto
End If
If Trim(t2.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "txtTimeout", t2.Text
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "txtTimeout", "5"
End If
If Trim(t4.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "port", t4.Text
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "port", "No configurado"
End If
If Trim(t5.Text) <> "" Then
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "server", t5.Text
Else
CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "server", "No configurado"
End If
frmMain.reload
frmMain.recargar
If frmMain.estoyOn.Text = "1" Then
frmMain.sock(iSockets).SendData "ALIAS" & "$$$" & frmMain.txtUsuario.Text & "%%%" & t1.Text
frmMain.reload
frmMain.recargar
End If
Unload Me
End Sub

Private Sub tabs_Click(PreviousTab As Integer)
On Error Resume Next
If Tabs.Tab = 0 Then
Text1.SetFocus
End If
If Tabs.Tab = 3 Then
Text4.SetFocus
End If
End Sub

Private Sub Form_Load()
'Leer configuraciones del registro
verMiImagen = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowMyImage")
verImagen = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowImages")
Ausencia = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Absent")
VerIconos = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "ShowIcons")
MiCarpeta = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "path")
GuardarHistorial = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SaveHistory")
EjecutarIntraMessenger = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "RunStartup")
ModoSilencioso = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "SilentMode")
MostrarAvisos = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsContacts")
AvisarMensajes = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "AlertsMessages")
txtPuerto = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "port")
txtServidor = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "server")
TiempoEspera = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "txtTimeout")
Sonidos = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\IntraMessenger\Options", "Sounds")
Text1.Text = verMiImagen
Text2.Text = verImagen
Text3.Text = Ausencia
Text4.Text = VerIconos
t3.Text = MiCarpeta
Text6.Text = GuardarHistorial
Text7.Text = EjecutarIntraMessenger
Text8.Text = ModoSilencioso
Text9.Text = MostrarAvisos
Text10.Text = AvisarMensajes
t4.Text = txtPuerto
t5.Text = txtServidor
t2.Text = TiempoEspera
Text5.Text = Sonidos
'Aplicar datos obtenidos en los objetos
If Text1.Text = "true" Then
c1.Value = 1
Else
c1.Value = 0
End If
If Text2.Text = "true" Then
c2.Value = 1
Else
c2.Value = 0
End If
If Text3.Text = "true" Then
c3.Value = 1
t2.Enabled = True
t2.BackColor = &H80000005
Else
c3.Value = 0
t2.Enabled = False
t2.BackColor = &HEEF3F4
End If
If Text4.Text = "true" Then
c4.Value = 1
Else
c4.Value = 0
End If
If Text6.Text = "true" Then
c5.Value = 1
Else
c5.Value = 0
End If
If Text7.Text = "true" Then
c6.Value = 1
Else
c6.Value = 0
End If
If Text8.Text = "true" Then
c7.Value = 0
Else
c7.Value = 1
End If
If Text9.Text = "true" Then
c8.Value = 1
Else
c8.Value = 0
End If
If Text10.Text = "true" Then
c9.Value = 1
Else
c9.Value = 0
End If
If Text5.Text = "true" Then
c10.Value = 1
Else
c10.Value = 0
End If
'Preparar formulario para mostrar al usuario
presionado1 = 1
presionado2 = 0
presionado3 = 0
presionado4 = 0
Frame1.BackColor = &H80000013
l1.Visible = True
l2.Visible = True
l3.Visible = True
l4.Visible = True
l5.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
co3.Visible = True
t1.Visible = True
t2.Visible = True
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
''''''''''''''''''''
auto = frmMain.txtUsuario.Text
If Trim(auto) = "" Then
auto = "No especificado"
End If
l1.Visible = True
l2.Visible = True
l3.Visible = True
l4.Visible = True
l5.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
co3.Visible = True
t1.Visible = True
t2.Visible = True
t1.Text = frmMain.txtMyalias.Text
If Trim(t1.Text) = "" Then
t1.Text = auto
End If
If frmMain.estoyOn.Text = "0" Then
t1.Text = ""
t1.Enabled = False
t1.BackColor = &HEEF3F4
l2.Enabled = False
co3.Enabled = False
c5.Enabled = False
l16.Caption = "Actualmente NO está conectado a IntraNET Messenger Service."
Else
l16.Caption = "Actualmente está conectado al servicio IntraNET Messenger Service."
End If
If t4.Text = "Error" Then
t4.Text = "No configurado"
End If
If t2.Text = "Er" Then
t2.Text = "5"
End If
If t5.Text = "Error" Then
t5.Text = "No configurado"
End If
If t3.Text = "Error" Then
t3.Text = App.Path & "\" & "Mis archivos recibidos"
End If
If c10.Value = vbChecked Then
Co6.Enabled = True
Else
Co6.Enabled = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Frame1_Click()
presionado1 = 1
presionado2 = 0
presionado3 = 0
presionado4 = 0
Frame1.BackColor = &H80000013
l1.Visible = True
l2.Visible = True
l3.Visible = True
l4.Visible = True
l5.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
co3.Visible = True
t1.Visible = True
t2.Visible = True
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado1 = 0 Then
Frame1.BackColor = &HEEF3F4
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Frame2_Click()
presionado1 = 0
presionado2 = 1
presionado3 = 0
presionado4 = 0
Frame2.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = True
l7.Visible = True
l8.Visible = True
l9.Visible = True
l10.Visible = True
i4.Visible = True
i5.Visible = True
i6.Visible = True
i7.Visible = True
c4.Visible = True
c5.Visible = True
Co4.Visible = True
Co5.Visible = True
t3.Visible = True
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado2 = 0 Then
Frame2.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Frame3_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 1
presionado4 = 0
Frame3.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = True
i8.Visible = True
i9.Visible = True
l11.Visible = True
l12.Visible = True
c6.Visible = True
c7.Visible = True
c8.Visible = True
c9.Visible = True
c10.Visible = True
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado3 = 0 Then
Frame3.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Frame4_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 0
presionado4 = 1
Frame4.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = True
l14.Visible = True
l15.Visible = True
l16.Visible = True
l17.Visible = True
l18.Visible = True
i10.Visible = True
t4.Visible = True
t5.Visible = True
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado4 = 0 Then
Frame4.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
End Sub

Private Sub Image4_Click()
presionado1 = 1
presionado2 = 0
presionado3 = 0
presionado4 = 0
Frame1.BackColor = &H80000013
l1.Visible = True
l2.Visible = True
l3.Visible = True
l4.Visible = True
l5.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
co3.Visible = True
t1.Visible = True
t2.Visible = True
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado1 = 0 Then
Frame1.BackColor = &HEEF3F4
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Image5_Click()
presionado1 = 0
presionado2 = 1
presionado3 = 0
presionado4 = 0
Frame2.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = True
l7.Visible = True
l8.Visible = True
l9.Visible = True
l10.Visible = True
i4.Visible = True
i5.Visible = True
i6.Visible = True
i7.Visible = True
c4.Visible = True
c5.Visible = True
Co4.Visible = True
Co5.Visible = True
t3.Visible = True
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado2 = 0 Then
Frame2.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Image6_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 1
presionado4 = 0
Frame3.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = True
i8.Visible = True
i9.Visible = True
l11.Visible = True
l12.Visible = True
c6.Visible = True
c7.Visible = True
c8.Visible = True
c9.Visible = True
c10.Visible = True
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado3 = 0 Then
Frame3.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Image7_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 0
presionado4 = 1
Frame4.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = True
l14.Visible = True
l15.Visible = True
l16.Visible = True
l17.Visible = True
l18.Visible = True
i10.Visible = True
t4.Visible = True
t5.Visible = True
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado4 = 0 Then
Frame4.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
End Sub

Private Sub Label6_Click()
presionado1 = 1
presionado2 = 0
presionado3 = 0
presionado4 = 0
Frame1.BackColor = &H80000013
l1.Visible = True
l2.Visible = True
l3.Visible = True
l4.Visible = True
l5.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
c1.Visible = True
c2.Visible = True
c3.Visible = True
co3.Visible = True
t1.Visible = True
t2.Visible = True
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado1 = 0 Then
Frame1.BackColor = &HEEF3F4
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Label7_Click()
presionado1 = 0
presionado2 = 1
presionado3 = 0
presionado4 = 0
Frame2.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = True
l7.Visible = True
l8.Visible = True
l9.Visible = True
l10.Visible = True
i4.Visible = True
i5.Visible = True
i6.Visible = True
i7.Visible = True
c4.Visible = True
c5.Visible = True
Co4.Visible = True
Co5.Visible = True
t3.Visible = True
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado2 = 0 Then
Frame2.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Label8_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 1
presionado4 = 0
Frame3.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = True
i8.Visible = True
i9.Visible = True
l11.Visible = True
l12.Visible = True
c6.Visible = True
c7.Visible = True
c8.Visible = True
c9.Visible = True
c10.Visible = True
'''''''''''''''
l13.Visible = False
l14.Visible = False
l15.Visible = False
l16.Visible = False
l17.Visible = False
l18.Visible = False
i10.Visible = False
t4.Visible = False
t5.Visible = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado3 = 0 Then
Frame3.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado4 = 0 Then
Frame4.BackColor = &H80000014
End If
End Sub

Private Sub Label9_Click()
presionado1 = 0
presionado2 = 0
presionado3 = 0
presionado4 = 1
Frame4.BackColor = &H80000013
l1.Visible = False
l2.Visible = False
l3.Visible = False
l4.Visible = False
l5.Visible = False
i1.Visible = False
i2.Visible = False
i3.Visible = False
c1.Visible = False
c2.Visible = False
c3.Visible = False
co3.Visible = False
t1.Visible = False
t2.Visible = False
'''''''''''''''''
l6.Visible = False
l7.Visible = False
l8.Visible = False
l9.Visible = False
l10.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
c4.Visible = False
c5.Visible = False
Co4.Visible = False
Co5.Visible = False
t3.Visible = False
'''''''''''''''
Co6.Visible = False
i8.Visible = False
i9.Visible = False
l11.Visible = False
l12.Visible = False
c6.Visible = False
c7.Visible = False
c8.Visible = False
c9.Visible = False
c10.Visible = False
'''''''''''''''
l13.Visible = True
l14.Visible = True
l15.Visible = True
l16.Visible = True
l17.Visible = True
l18.Visible = True
i10.Visible = True
t4.Visible = True
t5.Visible = True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If presionado4 = 0 Then
Frame4.BackColor = &HEEF3F4
End If
If presionado1 = 0 Then
Frame1.BackColor = &H80000014
End If
If presionado2 = 0 Then
Frame2.BackColor = &H80000014
End If
If presionado3 = 0 Then
Frame3.BackColor = &H80000014
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
Beep
KeyAscii = 0
End If
End Sub

Private Sub t3_GotFocus()
Co5.Default = True
End Sub

Private Sub t3_LostFocus()
Co5.Default = False
End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)
If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
Beep
KeyAscii = 0
End If
End Sub
