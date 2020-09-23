Attribute VB_Name = "modWindows"
Option Explicit
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const WM_USER = &H400
Public rctFrom As RECT
Public rctTo As RECT
Public lngTrayHand As Long
Public lngStartMenuHand As Long
Public lngChildHand As Long
Public strClass As String * 255
Public lngClassNameLen As Long
Public lngRetVal As Long

Public Function TitleToTray(frm As Form)
lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
Do
lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
If InStr(1, strClass, "TrayNotifyWnd") Then
lngTrayHand = lngChildHand
Exit Do
End If
lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop
lngRetVal = GetWindowRect(frm.hWnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hWnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
End Function

Public Function TrayToTitle(frm As Form)
lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
Do
lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
If InStr(1, strClass, "TrayNotifyWnd") Then
lngTrayHand = lngChildHand
Exit Do
End If
lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop
lngRetVal = GetWindowRect(frm.hWnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)
End Function
