VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim HKeySet As Long
Dim DType As Long
Dim KPath As String
Dim Handle As Long
Dim Success As Long
Dim ValueN As String
Dim SData As String
Dim Bdata() As Byte
Dim DData As Long
Dim Buffsize As Long
Dim sec As SECURITY_ATTRIBUTES

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Const KEY_ALL_ACCESS = &HF003F

Enum hKey
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_USER = &H80000001
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_USERS = &H80000003
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_DYN_DATA = &H80000006

End Enum

Enum lpType
 REG_SZ = 1
 REG_BINARY = 3
 REG_DWORD = 4
End Enum


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Event onSuccess()

Public Event onFailed()


Public Property Get hKey() As hKey
    hKey = HKeySet
End Property

Public Property Let hKey(ByVal Key As hKey)
    HKeySet = Key
End Property

Public Property Get DataType() As lpType
    DataType = DType
End Property

Public Property Let DataType(ByVal Dat As lpType)
    DType = Dat
End Property

Public Property Get SubKey() As String
 SubKey = KPath
End Property

Public Property Let SubKey(ByVal KP As String)
 KPath = KP
End Property

Public Property Get isSuccess() As Boolean
 If Success = 0 Then
  isSuccess = True
 Else
  isSuccess = False
 End If
 Success = 1
End Property

Public Property Let isSuccess(ByVal worked As Boolean)
 MsgBox "isSuccess Property is Read Only"
End Property

Public Property Get ValueName() As String
    ValueName = ValueN
End Property

Public Property Let ValueName(ByVal val As String)
    ValueN = val
End Property

Public Property Get Data() As Variant
Select Case DType
 Case 1
  Data = SData
 Case 3
  Data = Bdata
 Case 4
  Data = DData
End Select
End Property

Public Property Let Data(ByVal Dat1 As Variant)
 Select Case DType
 Case 1
  SData = Dat1 & vbNullChar
 Case 3
  Buffsize = Len(Dat1) / 2
  ReDim Bdata(Buffsize) As Byte
  Bdata = Dat1
 Case 4
  Buffsize = 4
  DData = Dat1
End Select
End Property

Public Sub DeleteValue()
Success = 0
 Success = RegOpenKeyEx(HKeySet, KPath, 0, KEY_ALL_ACCESS, Handle)
 If Success <> 0 Then GoTo Failed
  Success = RegDeleteValue(Handle, ValueN)
 If Success <> 0 Then GoTo Failed
  Success = RegCloseKey(Handle)
 If Success <> 0 Then GoTo Failed
 RaiseEvent onSuccess
Exit Sub
Failed:
 Success = RegCloseKey(Handle)
RaiseEvent onFailed
End Sub

Public Sub DeleteKey()
Success = RegDeleteKey(HKeySet, KPath)
 If Success <> 0 Then GoTo Failed
 RaiseEvent onSuccess
Exit Sub
Failed:
 Success = 1
RaiseEvent onFailed
End Sub

Public Sub CreateKey()
Dim Neworused As Long

sec.nLength = Len(secattr)
sec.lpSecurityDescriptor = 0
sec.bInheritHandle = True

Success = RegCreateKeyEx(HKeySet, KPath, 0, "", 0, KEY_ALL_ACCESS, sec, Handle, Neworused)

 If Success <> 0 Then GoTo Failed
Success = RegCloseKey(Handle)
 If Success <> 0 Then GoTo Failed
 RaiseEvent onSuccess
 Exit Sub
Failed:
 Success = RegCloseKey(Handle)
 Success = 1
 RaiseEvent onFailed
End Sub

Public Sub SaveSetting()
Dim Neworused As Long

sec.nLength = Len(secattr)
sec.lpSecurityDescriptor = 0
sec.bInheritHandle = True

Success = RegCreateKeyEx(HKeySet, KPath, 0, "", 0, KEY_ALL_ACCESS, sec, Handle, Neworused)


 If Success <> 0 Then GoTo Failed

 Select Case DType
 Case 1

 Success = RegSetValueEx(Handle, ValueN, 0, DType, ByVal SData, Len(SData) + 1)
  If Success <> 0 Then GoTo Failed
 Case 3
  Success = RegSetValueEx(Handle, ValueN, 0, DType, Bdata(0), UBound(Bdata()))
  If Success <> 0 Then GoTo Failed
 Case 4
  Success = RegSetValueEx(Handle, ValueN, 0, DType, DData, 4)
  If Success <> 0 Then GoTo Failed

Success = RegCloseKey(Handle)
 If Success <> 0 Then GoTo Failed

End Select
 RaiseEvent onSuccess
 Exit Sub
Failed:
 Success = RegCloseKey(Handle)
 Success = 1
 RaiseEvent onFailed
End Sub


Public Sub GetSetting()
  Dim buff As Long
  Success = RegOpenKeyEx(HKeySet, KPath, 0, KEY_ALL_ACCESS, Handle)
  If Success <> 0 Then GoTo Failed
 Select Case DType
 Case 1
  Success = RegQueryValueEx(Handle, ValueN, 0, DType, "", buff)
  SData = Space(buff)
  Success = RegQueryValueEx(Handle, ValueN, 0, DType, ByVal SData, buff)
  SData = Left$(SData, buff)
  If Success <> 0 Then GoTo Failed
 Case 3

 ReDim Bdata(255)
   Success = RegQueryValueEx(Handle, ValueN, 0, DType, Bdata(0), buff)
  ReDim Bdata(buff)
  Success = RegQueryValueEx(Handle, ValueN, 0, DType, Bdata(0), buff)
  If Success <> 0 Then GoTo Failed
 Case 4
 buff = 4
  Success = RegQueryValueEx(Handle, ValueN, 0, DType, DData, buff)
  If Success <> 0 Then GoTo Failed

Success = RegCloseKey(Handle)
 If Success <> 0 Then GoTo Failed

End Select
 RaiseEvent onSuccess
 Exit Sub
Failed:
 Success = RegCloseKey(Handle)
 Success = 1
 RaiseEvent onFailed
End Sub
