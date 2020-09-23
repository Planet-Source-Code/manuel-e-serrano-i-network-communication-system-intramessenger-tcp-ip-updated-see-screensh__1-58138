VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor IntraMessenger"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton updater 
      Caption         =   "Actualizar"
      Height          =   330
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3630
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detener"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   3360
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   4800
      TabIndex        =   5
      Top             =   3840
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   130
      TabIndex        =   4
      Top             =   3840
      Width           =   45
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim elalias As String
Dim agregado As Integer
Dim sendalias As String
Dim YesMatch As Integer
Dim iSockets As Integer
Dim GlobalAlias As String
Dim GlobalEstado As String
Dim TotalEnviados As Long
Dim TotalRecibidos As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Command1_Click()
On Error Resume Next
Load sock(Index)
sock(Index).LocalPort = 25
sock(Index).Listen
Command1.Enabled = False
Command2.Enabled = True
Command3.Visible = True
YesMatch = 0
Label1.Caption = "Puerto servidor: " & sock(0).LocalPort & vbCrLf & "Nombre servidor: " & UCase(sock(0).LocalHostName)
Label2.Caption = "Bytes totales recibidos: " & sock(0).BytesReceived & vbCrLf & "Bytes totales enviados: " & TotalEnviados
Call TodosOFF
End Sub

Sub TodosOFF()
'Copyright by Dj-Wincha
'Establece a todos los usuarios en OFF al iniciar el servidor (Módulo experimental)
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
Dim Inix
For Inix = 1 To rs.RecordCount
If rs!Estado <> "OFF" Then
rs.Edit
rs!Estado = "OFF"
rs.Update
End If
rs.MoveNext
Next Inix
Wend
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim nombreCerrar As String
If iSockets >= 1 And List1.ListCount > 0 Then
Dim advertencia
advertencia = MsgBox("Precaución, todos los usuarios serán desconectados del servicio. ¿Desea continuar?", vbExclamation + vbYesNo, "Confirmación")
If advertencia = vbYes Then
TotalEnviados = 0
TotalRecibidos = 0
Dim x
For x = 0 To List1.ListCount
nombreCerrar = Trim(Mid(List1.List(x), 1, InStr(List1.List(x), " - ")))
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If nombreCerrar = rs!IDUsuario Then
rs.Edit
rs!Estado = "OFF"
rs.Update
rs.MoveLast
rs.MoveNext
List1.RemoveItem List1.List(x)
Else
rs.MoveNext
End If
Wend
rs.Close
Next x
Dim i
For i = 0 To iSockets
sock(i).Close
Unload sock(i)
Next i
Command1.Enabled = True
Command2.Enabled = False
Label1.Caption = ""
Label2.Caption = ""
TotalEnviados = 0
TotalRecibidos = 0
YesMatch = 0
List1.Clear
Else
Exit Sub
End If
Else
Dim a
For a = 0 To iSockets
sock(a).Close
Unload sock(a)
Next a
Command1.Enabled = True
Command2.Enabled = False
Label1.Caption = ""
Label2.Caption = ""
TotalEnviados = 0
TotalRecibidos = 0
YesMatch = 0
End If
End Sub

Private Sub Command3_Click()
Label1.Caption = "Puerto servidor: " & sock(0).LocalPort & vbCrLf & "Nombre servidor: " & UCase(sock(0).LocalHostName)
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "Ya existe otra instancia de este programa en ejecución.", vbExclamation, "Error"
Unload Me
Exit Sub
End If
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\msgdata.dat.mdb", False, False, "MS Access;PWD=G8h13k7mh")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim nombreCerrar As String
If iSockets >= 1 And List1.ListCount > 0 Then
Dim advertencia
advertencia = MsgBox("Precaución, todos los usuarios serán desconectados del servicio. ¿Desea continuar?", vbExclamation + vbYesNo, "Confirmación")
If advertencia = vbYes Then
Dim x
For x = 0 To List1.ListCount
nombreCerrar = Trim(Mid(List1.List(x), 1, InStr(List1.List(x), " - ")))
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If nombreCerrar = rs!IDUsuario Then
rs.Edit
rs!Estado = "OFF"
rs.Update
rs.MoveLast
rs.MoveNext
List1.RemoveItem List1.List(x)
Else
rs.MoveNext
End If
Wend
rs.Close
Next x
Cancel = 0
Unload Me
Else
Cancel = 1
Exit Sub
End If
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
Dim indix As Integer
indix = Trim(Mid(List1.Text, InStr(List1.Text, "} ") + 1))
Set rs = db.OpenRecordset("Usuarios")
sock(indix).SendData "LISTA" & Chr(10)
While Not rs.EOF
sock(indix).SendData rs!Alias & rs!Estado & "_"
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub sock_Close(Index As Integer)
On Error Resume Next
Dim i
Dim nombreCerrar As String
For i = 0 To List1.ListCount
If Trim(Mid(List1.List(i), InStr(List1.List(i), "("), 6)) = "(" & sock(Index).SocketHandle & ")" Then
nombreCerrar = Trim(Mid(List1.List(i), 1, InStr(List1.List(i), " - ")))
'NUEVO MÓDULO
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If nombreCerrar = rs!IDUsuario Then
rs.Edit
rs!Estado = "OFF"
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
'NUEVO MÓDULO
List1.RemoveItem (i)
End If
Next i
DoEvents
sock(Index).Close
Unload sock(Index)
iSockets = iSockets - 1
If iSockets = -1 Then
iSockets = 0
End If
YesMatch = 0
'NUEVO MÓDULO
Dim x
Dim indix As Integer
For x = 0 To List1.ListCount
indix = Trim(Mid(List1.List(x), InStr(List1.List(x), "} ") + 1))
DoEvents
sock(indix).SendData "UPDAT" & "$$$" & nombreCerrar & "OFF"
Next x
End Sub

Private Sub sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
iSockets = iSockets + 1
Load sock(iSockets)
sock(iSockets).LocalPort = 25
sock(iSockets).Accept requestID
YesMatch = 0
End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
TotalRecibidos = Val(TotalRecibidos) + Val(bytesTotal)
Label2.Caption = "Bytes totales recibidos: " & TotalRecibidos & vbCrLf & "Bytes totales enviados: " & TotalEnviados
'#####################################
'CONSTANTE DE RECEPCIÓN DE DATOS
Dim strData As String
sock(Index).GetData strData, vbString
Dim textoGenerico As String
textoGenerico = Mid(strData, 1, 5)
Select Case textoGenerico
'#####################################
'AUTENTIFICACIÓN
Case "LOGIN"
Dim clave As String
Dim usuario As String
Dim revision As String
Dim currentState As String
usuario = strData
usuario = Mid(usuario, InStr(usuario, "$$$") + 3, InStr(usuario, "&&&") - 9)
clave = strData
clave = Mid(clave, InStr(clave, "&&&") + 3)
Dim v
For v = 0 To List1.ListCount
revision = Trim(Mid(List1.List(v), 1, InStr(List1.List(v), " - ")))
currentState = Trim(Mid(List1.List(v), InStr(List1.List(v), "{"), InStr(List1.List(v), "}")))
If Mid(currentState, 1, InStr(currentState, "}")) <> "{Sin conexión}" Then
If List1.ListCount > 0 Then
If revision = usuario Then
sock(Index).SendData "ALREA"
End If
End If
End If
Next v
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuario = rs!IDUsuario And clave = rs!Contraseña Then
elalias = rs!Alias
List1.AddItem usuario & " - " & sock(Index).RemoteHostIP & " (" & sock(Index).SocketHandle & ")" & " {En línea} " & sock(Index).Index
sock(Index).SendData "ALLOW"
sock(Index).SendData "ALIAS" & "%%%" & elalias
YesMatch = 1
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Clone
If YesMatch = 0 Then
sock(Index).SendData "ERROR"
End If
'#####################################
'CAMBIO DE NOMBRE A MOSTRAR
Case "ALIAS"
Dim user As String
Dim newalias As String
user = strData
user = Mid(user, InStr(user, "$$$") + 3, InStr(user, "%%%") - 9)
newalias = strData
newalias = Mid(newalias, InStr(newalias, "%%%") + 3)
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If user = rs!IDUsuario Then
OldAlias = rs!Alias
rs.Edit
rs!Alias = newalias
rs.Update
sock(Index).SendData "ALIAS" & "%%%" & newalias
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
Dim a
For a = 0 To rs.RecordCount
sock(Index).SendData rs!Alias & rs!Estado & "_"
rs.MoveNext
Next a
rs.Close
DoEvents
Dim UserNUM
Dim elIndex As Integer
For UserNUM = 0 To List1.ListCount
DoEvents
elIndex = Trim(Mid(List1.List(UserNUM), InStr(List1.List(UserNUM), "} ") + 1))
sock(elIndex).SendData "UPALI" & "OLD" & OldAlias & "NEW" & newalias
Next UserNUM
'#####################################
'CAMBIO DE ESTADO DEL CLIENTE
Case "STATE"
Dim usuarioID As String
usuarioID = strData
usuarioID = Mid(usuarioID, InStr(usuarioID, "$$$") + 3)
Dim Estados As String
Estados = Mid(strData, 6, 3)
Dim EstadosUsuario As String
Select Case Estados
'####################################
Case "OFF" 'Sin conexión
EstadosUsuario = "Sin conexión"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "OFF"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "NOA" 'No disponible
EstadosUsuario = "No disponible"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "NOA"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "ONL" 'En línea
EstadosUsuario = "En línea"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "ONL"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "AWA" 'Vuelvo enseguida
EstadosUsuario = "Vuelvo enseguida"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "AWA"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "LUN" 'Salí a comer
EstadosUsuario = "Salí a comer"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "LUN"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "ONP" 'Al teléfono
EstadosUsuario = "Al teléfono"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "ONP"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "IDL" 'Ausente
EstadosUsuario = "Ausente"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "IDL"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
Case "KUS" 'Sin conexión (Logout)
Dim kill
For kill = 0 To List1.ListCount
If usuarioID = Trim(Mid(List1.List(kill), 1, InStr(List1.List(kill), "-") - 1)) Then
List1.RemoveItem (kill)
End If
Next kill
EstadosUsuario = "Sin conexión"
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If usuarioID = rs!IDUsuario Then
rs.Edit
rs!Estado = "OFF"
GlobalAlias = rs!Alias
GlobalEstado = rs!Estado
rs.Update
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
DoEvents
Call updater_Click
'####################################
End Select
Dim i
For i = 0 To List1.ListCount
If Trim(Mid(List1.List(i), 1, InStr(List1.List(i), " - "))) = usuarioID Then
List1.RemoveItem (i)
List1.AddItem usuarioID & " - " & sock(Index).RemoteHostIP & " (" & sock(Index).SocketHandle & ")" & " {" & EstadosUsuario & "} " & sock(Index).Index
End If
Next i
Exit Sub
Set rs = db.OpenRecordset(usuarioID)
Dim x
For x = 0 To rs.RecordCount
MsgBox rs!NombreContacto
rs.MoveNext
Next x
rs.Close
'#####################################
'ENVÍO DE LISTA DE CONTACTOS DE CLIENTE Y STATUS DE LOS MISMOS
Case "LISTA"
Dim cliente As String
cliente = strData
cliente = Mid(cliente, InStr(cliente, "$$$") + 3)
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
sock(Index).SendData "LISTA" & Chr(10)
Dim c
For c = 0 To rs.RecordCount
sock(Index).SendData rs!Alias & rs!Estado & "_"
rs.MoveNext
Next c
Wend
rs.Close
'#####################################
'REGISTRAR UN NUEVO USUARIO EN EL SISTEMA
Case "NUEVO"
Dim encontrado As Integer
Dim newCliente As String
Dim newContraseña As String
newCliente = strData
newCliente = Mid(newCliente, InStr(newCliente, "$$$") + 3, InStr(newCliente, "&&&") - 9)
newContraseña = strData
newContraseña = Mid(newContraseña, InStr(newContraseña, "&&&") + 3)
Set rs = db.OpenRecordset("lista")
While Not rs.EOF
If newCliente = rs!Usuarios Then
sock(Index).SendData "ERROR"
encontrado = 1
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
If encontrado = 0 Then
db.Execute "CREATE TABLE [" & newCliente & "] (NombreContacto Text);"
Set rs = db.OpenRecordset("lista")
rs.AddNew
rs!Usuarios = newCliente
rs.Update
rs.Close
Set rs = db.OpenRecordset("Usuarios")
Dim NuevoID As Integer
If rs.RecordCount = 0 Then
NuevoID = "1"
Else
rs.MoveLast
NuevoID = rs!IDData + 1
End If
rs.AddNew
rs!IDData = NuevoID
rs!IDUsuario = newCliente
rs!Contraseña = newContraseña
rs!Alias = newCliente
rs!Online = "false"
rs.Update
rs.Close
sock(Index).SendData "USREG"
End If
'#####################################
'AGREGA/MODIFICA/ELIMINA CONTACTOS DE UN CLIENTE PARTICULAR
'Caso: Agregar un cliente nuevo a la lista de contactos
Case "AGREG"
Dim client As String
Dim positivo As Integer
Dim elcontacto As String
client = strData
client = Mid(client, InStr(client, "$$$") + 3, InStr(client, ">=<") - 9)
elcontacto = strData
elcontacto = Mid(elcontacto, InStr(elcontacto, ">=<") + 3)
Set rs = db.OpenRecordset(client)
While Not rs.EOF
If elcontacto = rs!NombreContacto Then
sock(Index).SendData "EXIST"
positivo = 1
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Close
If positivo = 0 Then
Set rs = db.OpenRecordset(client)
rs.AddNew
rs!NombreContacto = elcontacto
sock(Index).SendData "USREG"
rs.Update
rs.Close
End If
'#####################################
'MÓDULO DE ENVÍO DE MENSAJES (VERSIÓN BETA 1.0)
Case "IMDAT"
'Variables relativas a quien envía, recibe y el cuerpo del mensaje
Dim QuienEnvia As String
Dim QuienRecibe As String
Dim NombrePila As String
Dim CuerpoMSG As String
QuienEnvia = strData
QuienEnvia = Mid(QuienEnvia, InStr(QuienEnvia, "SND") + 3, InStr(QuienEnvia, "RCE") - 9)
NombrePila = strData
NombrePila = Mid(NombrePila, InStr(NombrePila, "RCE") + 3)
NombrePila = Mid(NombrePila, 1, InStr(NombrePila, "MSG") - 1)
CuerpoMSG = strData
CuerpoMSG = Mid(CuerpoMSG, InStr(CuerpoMSG, "MSG") + 3)
DoEvents
'Variables relativas a quien envía, recibe y el cuerpo del mensaje
'Buscar el nombre de pila en la base de datos y establecer
'el nombre del usuario real (nombre de suscripción)
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
If NombrePila = rs!Alias Then
QuienRecibe = rs!IDUsuario
rs.MoveLast
rs.MoveNext
Else
rs.MoveNext
End If
Wend
rs.Clone
DoEvents
Dim buscar
For buscar = 0 To List1.ListCount
If Trim(Mid(List1.List(buscar), 1, InStr(List1.List(buscar), " - "))) = QuienRecibe Then
Dim indix As Integer
indix = Trim(Mid(List1.List(buscar), InStr(List1.List(buscar), "} ") + 1))
DoEvents
sock(indix).SendData "IMDAT" & "SND" & QuienEnvia & "MSG" & CuerpoMSG
End If
Next buscar
End Select
End Sub

Private Sub sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
YesMatch = 0
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub sock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
TotalEnviados = Val(TotalEnviados) + Val(bytesSent)
Label2.Caption = "Bytes recibidos: " & TotalRecibidos & vbCrLf & "Bytes totales enviados: " & TotalEnviados
End Sub

Private Sub updater_Click()
'On Error Resume Next
Dim x
Dim indix As Integer
For x = 0 To List1.ListCount
indix = Trim(Mid(List1.List(x), InStr(List1.List(x), "} ") + 1))
DoEvents
sock(indix).SendData "UPDAT" & "$$$" & GlobalAlias & GlobalEstado
DoEvents
Next x
End Sub

Public Sub backup()
'################################################
'#ESTE MÓDULO ES SÓLAMENTE A MODO DE REFERENCIA,#
'#NO CUMPLE NINGUNA FUNCIÓN EN ESPECÍFICO.      #
'################################################
On Error Resume Next
Dim x
Dim indix As Integer
For x = 0 To List1.ListCount - 1
indix = Trim(Mid(List1.List(x), InStr(List1.List(x), "} ") + 1))
DoEvents
Set rs = db.OpenRecordset("Usuarios")
sock(indix).SendData "LISTA" & Chr(10)
While Not rs.EOF
sock(indix).SendData rs!Alias & rs!Estado & "_"
rs.MoveNext
Wend
rs.Close
Next x
End Sub

Public Sub upgradelist()
'On Error Resume Next
Dim x
Dim indix As Integer
For x = 0 To List1.ListCount
indix = Trim(Mid(List1.List(x), InStr(List1.List(x), "} ") + 1))
DoEvents
Set rs = db.OpenRecordset("Usuarios")
While Not rs.EOF
sock(indix).SendData "LISTA" & Chr(10)
Dim c
For c = 0 To rs.RecordCount
sock(indix).SendData rs!Alias & rs!Estado & "_"
rs.MoveNext
Next c
Wend
rs.Close
DoEvents
Next x
End Sub
