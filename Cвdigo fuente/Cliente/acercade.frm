VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FAECE6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de IntraMessenger"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "acercade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   32
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAECE6&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   338
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         Picture         =   "acercade.frx":000C
         ScaleHeight     =   540
         ScaleWidth      =   23040
         TabIndex        =   3
         Top             =   60
         Width           =   23040
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   240
         Picture         =   "acercade.frx":E4BA
         ScaleHeight     =   540
         ScaleWidth      =   23040
         TabIndex        =   4
         Top             =   240
         Width           =   23040
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 Manuel E. Serrano I. -- Venezuela --"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IntraMessenger"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Picture2.Top = Picture1.Top
Picture2.Left = Picture1.Left + Picture1.Width
Label2.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
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
