VERSION 5.00
Begin VB.Form frmPopup 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ControlBox      =   0   'False
   Icon            =   "popup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "popup.frx":000C
   ScaleHeight     =   90
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IntraMessenger"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'*CLIENTE INTRAMESSENGER (BETA 1.1)  *
'*IDEA ORIGINAL: MANUEL E. SERRANO I.*
'*************************************
Dim DirectionIsUp As Boolean
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
DoEvents
Me.Height = 0
DirectionIsUp = True
Me.Top = Screen.Height - 450
Me.Left = Screen.Width - (Me.Width + 200)
SetWinOnTop = SetWindowPos(frmPopup.hWnd, HWND_TOPMOST, frmPopup.Left, frmPopup.Top, frmPopup.Width, frmPopup.Height, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Timer1_Timer()
On Error GoTo solveErr
Timer1.Interval = 11
If DirectionIsUp = True Then
Me.Top = Me.Top - 55
If Me.Top <= Screen.Height - 450 Then
Me.Height = Me.Height + 55
End If
If Me.Top <= Screen.Height - 2190 Then
Me.Top = Screen.Height - 2190
Me.Height = 1740
Timer1.Interval = 4000
DirectionIsUp = False
End If
Else
Me.Height = Me.Height - 55
Me.Top = Me.Top + 55
If Me.Top >= Screen.Height + 50 Then
Timer1.Enabled = False
Unload Me
End If
End If
solveErr:
If Err.Number = 380 Then
DoEvents
Me.Visible = False
Unload Me
End If
End Sub
