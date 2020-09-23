Attribute VB_Name = "registros"
Option Explicit
Public Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal mKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegEnumValueType Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal null1 As Long, ByVal null2 As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegEnumValueString Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As String, lpcbData As Long) As Long
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERR_MORE_DATA = 234
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_ALL_ACCESS = &H3F
Public Const KEY_CREATE_SUBKEY = &H4&
Public Const KEY_ENUMERATE_SUBKEY = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_CREATE_LINK = &H20
Public Const READ_CONTROL = &H20000
Public Const WRITE_OWNER = &H80000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUBKEY Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUBKEY
Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_DWORD_LITTLE_ENDIAN = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_LINK = 6
Public Const REG_MULTI_SZ = 7
Public Const REG_RESOURCE_LIST = 8
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Global mMainKey As Long
Global mIndent As Integer
Global mStopFlag As Boolean
Global mAccumText As String
Global mResult As Long
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const WRITE_DAC = &H40000
Const SYNCHRONIZE = &H100000
Const KEY_EXECUTE = KEY_READ
Const REG_OPTION_NON_VOLATILE = 0
Const ERROR_NONE = 0
Const ERROR_INVALID_PARAMETERS = 87
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Const DisplayErrorMsg = False
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const MAX_PATH = 260
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Public Const LVIF_STATE = &H8
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVIS_STATEIMAGEMASK As Long = &HF000
Public Type LV_ITEM
mask         As Long
iItem        As Long
iSubItem     As Long
state        As Long
stateMask    As Long
pszText      As String
cchTextMax   As Long
iImage       As Long
lParam       As Long
iIndent      As Long
End Type
Global ErrorString As String
Global f, bChecked, bValid As Boolean
Global iIndex As Long
Global sDirectory As String
Dim hkey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

Public Function ReplaceString(sTarget As String, sSearchString As String, sReplaceString As String) As String
Dim sTemp As String
Dim lLength As Long
lLength = Len(sTarget)
sTemp = Replace(sTarget, sSearchString, sReplaceString, 1, lLength, vbTextCompare)
ReplaceString = sTemp
End Function

Public Function GetMainKey(ByVal inName As String) As Long
GetMainKey = 0
If inName = "HKEY_CLASSES_ROOT" Then
GetMainKey = HKEY_CLASSES_ROOT
ElseIf inName = "HKEY_CURRENT_USER" Then
GetMainKey = HKEY_CURRENT_USER
ElseIf inName = "HKEY_LOCAL_MACHINE" Then
GetMainKey = HKEY_LOCAL_MACHINE
ElseIf inName = "HKEY_USERS" Then
GetMainKey = HKEY_USERS
ElseIf inName = "HKEY_PERFORMANCE_DATA" Then
GetMainKey = HKEY_PERFORMANCE_DATA
ElseIf inName = "HKEY_CURRENT_CONFIG" Then
GetMainKey = HKEY_CURRENT_CONFIG
ElseIf inName = "HKEY_DYN_DATA" Then
GetMainKey = HKEY_CURRENT_CONFIG
End If
End Function

Public Function GetMainKeyString(ByVal MainKey As Long) As String
Select Case MainKey
Case HKEY_CLASSES_ROOT
GetMainKeyString = "HKEY_CLASSES_ROOT"
Case HKEY_CURRENT_USER
GetMainKeyString = "HKEY_CURRENT_USER"
Case HKEY_LOCAL_MACHINE
GetMainKeyString = "HKEY_LOCAL_MACHINE"
Case HKEY_USERS
GetMainKeyString = "HKEY_USERS"
Case HKEY_PERFORMANCE_DATA
GetMainKeyString = "HKEY_PERFORMANCE_DATA"
Case HKEY_CURRENT_CONFIG
GetMainKeyString = "HKEY_CURRENT_CONFIG"
Case HKEY_DYN_DATA
GetMainKeyString = "HKEY_CURRENT_CONFIG"
End Select
End Function

Public Function DoEnumSubKeys(ByVal inMainKey As Long, ByVal inSubKey As String)
Dim mKey As Long
Dim colSubKeys As Collection
Dim mBuffer As String * 256
Dim mBufSize As Long
Dim mClassBuffer As String * 256
Dim mClassBufSize As Long
Dim typLastWriteTime As FILETIME
Dim mSubKeyName As String
Dim mSubKeyValue As String
Dim mValType As Long
Dim mIndex As Integer
Dim i As Integer
Set colSubKeys = New Collection
If RegOpenKeyEx(inMainKey, inSubKey, 0&, KEY_ALL_ACCESS, mKey) <> 0 Then
ErrorString = ErrorString & GetMainKeyString(inMainKey) & "\" & inSubKey & vbNewLine
Exit Function
End If
ListEntryValues inMainKey, inSubKey
mIndex = 0
Do
mClassBuffer = ""
mClassBufSize = 0
mBufSize = 256
mSubKeyName = Space$(mBufSize)
mResult = RegEnumKeyEx(mKey, mIndex, mSubKeyName, mBufSize, 0, mClassBuffer, _
mClassBufSize, typLastWriteTime)
If mResult <> ERR_MORE_DATA And mResult <> 0 Then
Exit Do
End If
mSubKeyName = Left$(mSubKeyName, InStr(mSubKeyName, Chr$(0)) - 1)
colSubKeys.Add mSubKeyName
mIndex = mIndex + 1
Loop
RegCloseKey mKey
If mIndex > 0 Then
mIndent = mIndent + 3
Else
mIndent = 3
End If
For i = 1 To colSubKeys.Count
DoEvents
If mStopFlag Then
Exit For
End If
DoEnumSubKeys inMainKey, inSubKey & "\" & colSubKeys(i)
Next i
End Function

Public Function ListEntryValues(ByVal inMainKey As Long, ByVal inSubKey As String)
Dim mKey As Long
Dim mIndex As Long
Dim mReturnEntry, lKey As String
Dim mEntry As String
Dim mEntryLength As Long
Dim mDataType As Long
Dim arrDataByte(1 To 1024) As Byte
Dim mDataByteLength As Long
Dim mDataByteValue, sTemp As String
Dim i As Integer
Dim j
Dim y
Dim tmp As String, mChr As String
Dim mAlignPos, mAlignPos0, mAlignPos1
mResult = RegOpenKeyEx(inMainKey, inSubKey, 0&, KEY_ALL_ACCESS, mKey)
If mResult <> 0 Then
ErrorString = ErrorString & GetMainKeyString(inMainKey) & "\" & inSubKey & vbNewLine
Exit Function
End If
tmp = ""
lKey = "[" & GetMainKeyString(inMainKey) & "\" & inSubKey & "]"
mAccumText = mAccumText & lKey & tmp & vbNewLine
mAlignPos0 = 5
mAlignPos1 = mAlignPos0 + mIndent + Len(tmp)
mIndex = 0
Do
If mIndex = 0 Then
mAlignPos = mAlignPos0
Else
mAlignPos = mAlignPos1
End If
mEntryLength = 1024
mDataByteLength = 1024
mEntry = Space$(mEntryLength)
mResult = RegEnumValue(mKey, mIndex, mEntry, mEntryLength, 0, _
mDataType, arrDataByte(1), mDataByteLength)
If mResult <> 0 Then
Exit Do
End If
mEntry = Left$(mEntry, mEntryLength)
If mEntry = "" And mDataByteLength > 0 Then
mEntry = Chr(64)
Else
mEntry = Chr(34) & mEntry & Chr(34)
End If
mDataByteValue = ""
Select Case mDataType
Case REG_SZ
Dim sTempString As String
For i = 1 To mDataByteLength - 1
mDataByteValue = mDataByteValue & Chr$(arrDataByte(i))
Next i
sTempString = mDataByteValue
mDataByteValue = ReplaceString(sTempString, "\", "\\")
mAccumText = mAccumText & mEntry & "=" & Chr(34) & mDataByteValue & Chr(34) & vbCrLf
Case REG_DWORD
tmp = ""
For i = 4 To 1 Step -1
mChr = Hex(Asc(Chr(arrDataByte(i))))
If Len(mChr) = 1 Then
mChr = "0" + mChr
End If
tmp = tmp & mChr
Next
y = 0
For i = 1 To Len(tmp) Step 2
mChr = Mid(tmp, i, 2)
mChr = ConvNum(mChr, 16, 10)
j = Val(mChr)
Select Case i
Case 1
j = j * &H1000000
Case 2
j = j * &H10000
Case 3
j = j * &H100
End Select
y = y + j
Next
mDataByteValue = CStr(y)
sTemp = "dword:" & tmp
mAccumText = mAccumText & mEntry & "=" & sTemp & vbCrLf
Case REG_BINARY
For i = 1 To mDataByteLength
mChr = Hex$(Asc(Chr(arrDataByte(i))))
If Len(mChr) = 1 Then
mChr = "0" & mChr
End If
mDataByteValue = mDataByteValue & mChr & " "
Next
sTemp = "hex:" & mDataByteValue
mAccumText = mAccumText & mEntry & "=" & sTemp & vbCrLf
Case Else
For i = 1 To mDataByteLength - 1
mDataByteValue = mDataByteValue & Chr$(arrDataByte(i))
Next i
mAccumText = mAccumText & mEntry & "=" & Chr(34) & mDataByteValue & Chr(34) & vbCrLf
End Select
mIndex = mIndex + 1
Loop
mAccumText = mAccumText & vbCrLf
End Function

Public Function ConvNum(inNum As String, inFrom As Integer, inTo As Integer) As String
Dim mChr As String
Dim i As Integer, j As Integer, k As Integer
Dim mTotal As Double, mMod As Double, tmp As Double
ConvNum = ""
If inNum = "" Then
Exit Function
ElseIf inFrom < 2 Or inFrom > 36 Then
Exit Function
ElseIf inTo < 1 Or inTo > 36 Then
Exit Function
End If
inNum = UCase$(inNum)
k = Len(inNum)
For i = 1 To Len(inNum)
k = k - 1
mChr = Mid$(inNum, i, 1)
j = 0
If Asc(mChr) > 64 And Asc(mChr) < 91 Then
j = Asc(mChr) - 55
End If
If j = 0 Then
If Asc(mChr) < 48 Or Asc(mChr) > 57 Then
Exit Function
End If
j = Val(mChr)
End If
If j < 0 Or j > inFrom - 1 Then
Exit Function
End If
mTotal = mTotal + j * (inFrom ^ k)
Next i
Do While mTotal > 0
tmp = CDbl(inTo)
mMod = mTotal - (Int(mTotal / tmp) * tmp)
mTotal = (mTotal - mMod) / inTo
If mMod >= 10 Then
mChr = Chr$(mMod + 55)
Else
mChr = Right$(Str$(mMod), Len(Str$(mMod)) - 1)
End If
ConvNum = mChr & ConvNum
Loop
If ConvNum = "" Then
ConvNum = "0"
End If
End Function

Public Function DeleteRegistryValue(ByVal hkey As Long, ByVal KeyName As String, _
ByVal ValueName As String) As Boolean
Dim handle As Long
Dim sRoot As String
If RegOpenKeyEx(hkey, KeyName, 0, KEY_WRITE, handle) Then
ErrorString = ErrorString & GetMainKeyString(hkey) & "\" & KeyName & vbNewLine
Exit Function
End If
DeleteRegistryValue = (RegDeleteValue(handle, ValueName) = 0)
If DeleteRegistryValue = False Then
ErrorString = ErrorString & GetMainKeyString(hkey) & "\" & KeyName & vbNewLine
End If
RegCloseKey handle
End Function
Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
Dim lRetVal As Long
Dim hkey As Long
lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hkey)
lRetVal = RegDeleteValue(hkey, sValueName)
RegCloseKey (hkey)
End Function

Public Function SetValueEx(ByVal hkey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
Case REG_SZ
sValue = vValue
SetValueEx = RegSetValueExString(hkey, sValueName, 0&, lType, sValue, Len(sValue))
Case REG_DWORD
lValue = vValue
SetValueEx = RegSetValueExLong(hkey, sValueName, 0&, lType, lValue, 4)
End Select
End Function

Public Function CheckRegistryValue(ByVal hkey As Long, ByVal KeyName As String, _
ByVal ValueName As String) As Boolean
Dim handle As Long
If RegOpenKeyEx(hkey, KeyName, 0, KEY_WRITE, handle) Then
CheckRegistryValue = False
Exit Function
Else
CheckRegistryValue = True
End If
RegCloseKey handle
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String
On Error GoTo QueryValueExError
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> ERROR_NONE Then Error 5
Select Case lType
Case REG_SZ:
sValue = String(cch, 0)
lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
If lrc = ERROR_NONE Then
vValue = Left$(sValue, cch)
Else
vValue = Empty
End If
Case REG_DWORD:
lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
If lrc = ERROR_NONE Then vValue = lValue
Case Else
lrc = -1
End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function

Function CreateKey(SubKey As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegCreateKey(MainKeyHandle, SubKey, hkey)
If rtn = ERROR_SUCCESS Then
rtn = RegCloseKey(hkey)
End If
End If
End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
Dim hNewKey As Long
Dim lRetVal As Long
lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
RegCloseKey (hNewKey)
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
Dim lRetVal As Long
Dim hkey As Long
lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hkey)
lRetVal = SetValueEx(hkey, sValueName, lValueType, vValueSetting)
RegCloseKey (hkey)
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
Dim lRetVal As Long
Dim hkey As Long
Dim vValue As Variant
lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hkey)
lRetVal = QueryValueEx(hkey, sValueName, vValue)
QueryValue = vValue
RegCloseKey (hkey)
End Function

Public Function CheckRegistryKey(ByVal hkey As Long, ByVal KeyName As String) As _
Boolean
Dim handle As Long
If RegOpenKeyEx(hkey, KeyName, 0, KEY_READ, handle) = 0 Then
CheckRegistryKey = True
RegCloseKey handle
End If
End Function

Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
ByVal strRegKeyPath As String) As Boolean
Dim m_lngRetVal As Long
Dim lngKeyHandle As Long
lngKeyHandle = 0
m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
If lngKeyHandle = 0 Then
regDoes_Key_Exist = False
Else
regDoes_Key_Exist = True
End If
m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function

Public Function GetMainKeyHandle(MainKeyName As String) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Select Case MainKeyName
Case "HKEY_CLASSES_ROOT"
GetMainKeyHandle = HKEY_CLASSES_ROOT
Case "HKEY_CURRENT_USER"
GetMainKeyHandle = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
GetMainKeyHandle = HKEY_LOCAL_MACHINE
Case "HKEY_USERS"
GetMainKeyHandle = HKEY_USERS
Case "HKEY_PERFORMANCE_DATA"
GetMainKeyHandle = HKEY_PERFORMANCE_DATA
Case "HKEY_CURRENT_CONFIG"
GetMainKeyHandle = HKEY_CURRENT_CONFIG
Case "HKEY_DYN_DATA"
GetMainKeyHandle = HKEY_DYN_DATA
End Select
End Function

Public Function ErrorMsg(lErrorCode As Long) As String
Dim GetErrorMsg As String
Select Case lErrorCode
Case 1009, 1015
GetErrorMsg = "The Registry Database is corrupt!"
Case 2, 1010
GetErrorMsg = "Bad Key Name"
Case 1011
GetErrorMsg = "Can't Open Key"
Case 4, 1012
GetErrorMsg = "Can't Read Key"
Case 5
GetErrorMsg = "Access to this key is denied"
Case 1013
GetErrorMsg = "Can't Write Key"
Case 8, 14
GetErrorMsg = "Out of memory"
Case 87
GetErrorMsg = "Invalid Parameter"
Case 234
GetErrorMsg = "There is more data than the buffer has been allocated to hold."
Case Else
GetErrorMsg = "Undefined Error Code:  " & Str(lErrorCode)
End Select
ErrorMsg = GetErrorMsg
End Function

Public Function GetStringValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hkey)
If rtn = ERROR_SUCCESS Then
sBuffer = Space(255)
lBufferSize = Len(sBuffer)
rtn = RegQueryValueEx(hkey, Entry, 0, REG_SZ, sBuffer, lBufferSize)
If rtn = ERROR_SUCCESS Then
rtn = RegCloseKey(hkey)
sBuffer = Trim(sBuffer)
GetStringValue = Left(sBuffer, Len(sBuffer) - 1)
Else
GetStringValue = "Error"
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn), , "Asistente de negocios"
End If
End If
Else
GetStringValue = "Error"
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn), , "Asistente de negocios"
End If
End If
End If
End Function

Public Function ParseKey(KeyName As String, Keyhandle As Long)
rtn = InStr(KeyName, "\")
If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then
MsgBox "Formato incorrecto:" + Chr(10) + Chr(10) + KeyName, , "Asistente de negocios"
Exit Function
ElseIf rtn = 0 Then
Keyhandle = GetMainKeyHandle(KeyName)
KeyName = ""
Else
Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1))
KeyName = Right(KeyName, Len(KeyName) - rtn)
End If
End Function

Public Function SetStringValue(SubKey As String, Entry As String, Value As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hkey)
If rtn = ERROR_SUCCESS Then
rtn = RegSetValueEx(hkey, Entry, 0, REG_SZ, ByVal Value, Len(Value))
If Not rtn = ERROR_SUCCESS Then
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn), , "Asistente de negocios"
End If
End If
rtn = RegCloseKey(hkey)
Else
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn), , "Asistente de negocios"
End If
End If
End If
End Function

Public Function GetKeyInfo(ByVal Section As Long, ByVal key_name As String, ByVal indent As Integer) As String
Dim subkeys As Collection
Dim subkey_values As Collection
Dim subkey_num As Integer
Dim subkey_name As String
Dim subkey_value As String
Dim Length As Long
Dim hkey As Long
Dim txt As String
Dim subkey_txt As String
Set subkeys = New Collection
Set subkey_values = New Collection
If Right(key_name, 1) = "\" Then key_name = Left(key_name, Len(key_name) - 1)
If RegOpenKeyEx(Section, key_name, 0&, KEY_ALL_ACCESS, hkey) <> ERROR_SUCCESS Then
Exit Function
End If
subkey_num = 0
Do
Length = 256
subkey_name = Space(Length)
If RegEnumKey(hkey, subkey_num, subkey_name, Length) <> ERROR_SUCCESS Then Exit Do
subkey_num = subkey_num + 1
subkey_name = Left(subkey_name, InStr(subkey_name, Chr(0)) - 1)
subkeys.Add subkey_name
Length = 256
subkey_value = Space(Length)
If RegQueryValue(hkey, subkey_name, subkey_value, Length) <> ERROR_SUCCESS Then
subkey_values.Add "Error"
Else
subkey_value = Left(subkey_value, Length - 1)
subkey_values.Add subkey_value
End If
Loop
If RegCloseKey(hkey) <> ERROR_SUCCESS Then
MsgBox "Error closing key.", vbExclamation, "Asistente de negocios"
End If
For subkey_num = 1 To subkeys.Count
subkey_txt = GetKeyInfo(Section, key_name & "\" & subkeys(subkey_num), indent + 2)
txt = txt & Space(indent) & subkeys(subkey_num) & ": " & subkey_values(subkey_num) & vbCrLf & subkey_txt
Next subkey_num
GetKeyInfo = txt
End Function

Public Function DeleteSubkeys(ByVal Section As Long, ByVal key_name As String)
Dim hkey As Long
Dim subkeys As Collection
Dim subkey_num As Long
Dim Length As Long
Dim subkey_name As String
Dim sRoot As String
If RegOpenKeyEx(Section, key_name, 0&, KEY_ALL_ACCESS, hkey) <> ERROR_SUCCESS Then
ErrorString = ErrorString & GetMainKeyString(Section) & "\" & key_name & vbNewLine
Exit Function
End If
Set subkeys = New Collection
subkey_num = 0
Do
Length = 256
subkey_name = Space(Length)
If RegEnumKey(hkey, subkey_num, subkey_name, Length) <> ERROR_SUCCESS Then Exit Do
subkey_num = subkey_num + 1
subkey_name = Left(subkey_name, InStr(subkey_name, Chr(0)) - 1)
subkeys.Add subkey_name
Loop
For subkey_num = 1 To subkeys.Count
DeleteSubkeys Section, key_name & "\" & subkeys(subkey_num)
RegDeleteKey hkey, subkeys(subkey_num)
Next subkey_num
RegCloseKey hkey
End Function

Public Function DeleteKey(ByVal Section As Long, ByVal key_name As String)
Dim pos As Integer
Dim parent_key_name As String
Dim parent_hKey As Long
If Right(key_name, 1) = "\" Then key_name = Left(key_name, Len(key_name) - 1)
DeleteSubkeys Section, key_name
pos = InStrRev(key_name, "\")
If pos = 0 Then
RegDeleteKey Section, key_name
Else
parent_key_name = Left(key_name, pos - 1)
key_name = Mid(key_name, pos + 1)
If RegOpenKeyEx(Section, parent_key_name, 0&, KEY_ALL_ACCESS, parent_hKey) <> ERROR_SUCCESS Then
Else
RegDeleteKey parent_hKey, key_name
RegCloseKey parent_hKey
End If
End If
End Function

Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, _
ByVal Key As String) As String
Dim lResult As Long
Dim lKeyValue As Long
Dim lDataTypeValue As Long
Dim lValueLength As Long
Dim sValue As String
Dim td As Double
Dim TStr2, TStr1  As String
Dim i As Integer
Dim lngDataPntr, lngType As Long
Dim strValue As String
Dim sResult As String
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sResult = Space(2048)
lValueLength = Len(sValue)
lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
Select Case lDataTypeValue
Case REG_SZ:
sResult = "String"
Case REG_DWORD:
sResult = "DWord"
Case REG_BINARY:
sResult = "Binary"
Case REG_NONE:
sResult = "String"
Case REG_EXPAND_SZ:
sResult = "String"
Case REG_DWORD_LITTLE_ENDIAN:
sResult = "String"
Case REG_DWORD_BIG_ENDIAN:
sResult = "String"
Case REG_LINK:
sResult = "String"
Case REG_MULTI_SZ:
sResult = "Binary"
Case REG_RESOURCE_LIST:
sResult = "String"
Case REG_FULL_RESOURCE_DESCRIPTOR:
sResult = "String"
Case REG_RESOURCE_REQUIREMENTS_LIST:
sResult = "String"
Case Else:
sResult = "Other"
End Select
Else
sResult = "Other"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sResult
End Function
