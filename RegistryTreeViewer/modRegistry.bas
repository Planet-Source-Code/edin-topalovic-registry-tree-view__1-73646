Attribute VB_Name = "mR"
Option Explicit
Public Enum KeyRoot
  [HKEY_CLASSES_ROOT] = &H80000000
  [HKEY_CURRENT_CONFIG] = &H80000005
  [HKEY_CURRENT_USER] = &H80000001
  [HKEY_LOCAL_MACHINE] = &H80000002
  [HKEY_USERS] = &H80000003 '
End Enum
Public Enum KeyType
  [REG_BINARY] = 3
  [REG_DWORD] = 4
  [REG_SZ] = 1
End Enum


Private Const HKEY_PERFORMANCE_DATA = &H80000004

Const HKEY_DYN_DATA = &H80000006

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


 Const KEY_ALL_ACCESS = &HF003F
 Const KEY_ENUMERATE_SUB_KEYS = &H8
 Const KEY_READ = &H20019
 Const KEY_WRITE = &H20006
 Const KEY_QUERY_VALUE = &H1 '

 Const REG_FORCE_RESTORE As Long = 8&
 Const TOKEN_QUERY As Long = &H8&
 Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
 Const SE_PRIVILEGE_ENABLED As Long = &H2
 Const SE_RESTORE_NAME = "SeRestorePrivilege"
 Const SE_BACKUP_NAME = "SeBackupPrivilege"

Const REG_OPTION_NON_VOLATILE = 0
Const REG_OPTION_VOLATILE = 1
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Private Type LUID
  lowpart As Long
  highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbdata As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hkey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hkey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbdata As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbdata As Long) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public NewMessage, NewMessage1
Public cnt, cnt1 As Long
Public knm, knm1 As String
Public Message

Public Function ExportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long

  
  If EnablePrivilege(SE_BACKUP_NAME) = False Then
    ExportRegKey = False
    Exit Function
  End If
  
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hkey)
  If ReturnValue <> 0 Then
    
    ExportRegKey = False
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  
  If Dir(FileName) <> "" Then Kill FileName
  
  ReturnValue = RegSaveKey(hkey, FileName, ByVal 0&)
  If ReturnValue = 0 Then
    
    ExportRegKey = True
  Else
    
    ExportRegKey = False
  End If
  
  ReturnValue = RegCloseKey(hkey)
End Function
Public Function ImportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long

  
  If EnablePrivilege(SE_RESTORE_NAME) = False Then
    ImportRegKey = False
    Exit Function
  End If
  
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hkey)
  If ReturnValue <> 0 Then
    
    ImportRegKey = False
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  
  ReturnValue = RegRestoreKey(hkey, FileName, REG_FORCE_RESTORE)
  If ReturnValue = 0 Then
    
    ImportRegKey = True
  Else
    
    ImportRegKey = False
  End If
  
  ReturnValue = RegCloseKey(hkey)
End Function
Public Function ReadRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String, Optional NoKeyFoundValue As String = "") As String
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long

  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_READ, hkey)
  If ReturnValue <> 0 Then
    ReadRegKey = NoKeyFoundValue
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  ReadRegKey = GetSubKeyValue(hkey, SubKey)
  ReturnValue = RegCloseKey(hkey)
End Function
Public Function WriteRegKey(KeyType As KeyType, KeyRoot As KeyRoot, KeyPath As String, SubKey As String, SubKeyValue As String) As Boolean
  
  On Error Resume Next
  Dim hkey As Long
  Dim SecurityAttribute As SECURITY_ATTRIBUTES
  Dim NewKey As Long
  Dim ReturnValue As Long

  
  SecurityAttribute.nLength = Len(SecurityAttribute)
  SecurityAttribute.lpSecurityDescriptor = 0
  SecurityAttribute.bInheritHandle = True

  
  ReturnValue = RegCreateKeyEx(KeyRoot, KeyPath, 0, "", 0, KEY_WRITE, SecurityAttribute, hkey, NewKey)
  If ReturnValue <> 0 Then
    
    WriteRegKey = False
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If

  
  Select Case KeyType
    Case REG_SZ
      ReturnValue = RegSetValueEx(hkey, SubKey, 0, KeyType, ByVal SubKeyValue, Len(SubKeyValue))
    Case REG_DWORD
      ReturnValue = RegSetValueEx(hkey, SubKey, 0, KeyType, CLng(SubKeyValue), 4)
    Case REG_BINARY
      ReturnValue = RegSetValueEx(hkey, SubKey, 0, KeyType, CByte(SubKeyValue), 4)
  End Select

  If ReturnValue = 0 Then
    
    WriteRegKey = True
  Else
    
    WriteRegKey = False
  End If

  
  ReturnValue = RegCloseKey(hkey)
End Function

Public Function EnumerateRegKeys(KeyRoot As KeyRoot, KeyPath As String) As String
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim ClassNameBuffer As String
  Dim ClassNameBufferSize As Long
  Dim LastWrite As FILETIME

  
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ENUMERATE_SUB_KEYS, hkey)
  If ReturnValue <> 0 Then
    
    EnumerateRegKeys = ""
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  Counter = 0
  '
  frmRegistryView.sb.Panels(1).Text = "Reading..."
  Do Until ReturnValue <> 0
    frmRegistryView.MousePointer = 11
    MyBuffer = Space(255)
    ClassNameBuffer = Space(255)
    MyBufferSize = 255
    ClassNameBufferSize = 255
    ReturnValue = RegEnumKeyEx(hkey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
      EnumerateRegKeys = EnumerateRegKeys & MyBuffer & Chr(0)
    End If
    Counter = Counter + 1
    frmRegistryView.sb.Panels(1).Text = "Reading... (" & Counter & " key(s))"
    If (Counter Mod 100) = 0 Then
        DoEvents
    End If
  Loop
  frmRegistryView.sb.Panels(1).Text = ""
  frmRegistryView.MousePointer = 0
  If EnumerateRegKeys <> "" Then EnumerateRegKeys = Left$(EnumerateRegKeys, Len(EnumerateRegKeys) - 1)
  
  ReturnValue = RegCloseKey(hkey)
End Function
Public Function EnumerateRegKeyValues(KeyRoot As KeyRoot, KeyPath As String) As String
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim KeyType As KeyType

  
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_QUERY_VALUE, hkey)
  
  If ReturnValue <> 0 Then
    EnumerateRegKeyValues = ""
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  Counter = 0
  
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    MyBufferSize = 255
    ReturnValue = RegEnumValue(hkey, Counter, MyBuffer, MyBufferSize, 0, KeyType, ByVal 0&, ByVal 0&) 'ByteData(0), ByteDataSize)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      EnumerateRegKeyValues = EnumerateRegKeyValues & MyBuffer & Chr(1)
      EnumerateRegKeyValues = EnumerateRegKeyValues & GetSubKeyValue(hkey, MyBuffer) & Chr(0)
    End If
    Counter = Counter + 1
  Loop
  
  If EnumerateRegKeyValues <> "" Then EnumerateRegKeyValues = Left$(EnumerateRegKeyValues, Len(EnumerateRegKeyValues) - 1)
  
  ReturnValue = RegCloseKey(hkey)
End Function
Public Function DeleteRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String) As Boolean
  On Error Resume Next
  Dim ReturnValue As Long

  
  ReturnValue = RegDeleteKey(KeyRoot, KeyPath & "\" & SubKey)
  If ReturnValue = 0 Then
   
    DeleteRegKey = True
  Else
    
    DeleteRegKey = False
  End If
End Function
Public Function DeleteRegKeyValue(KeyRoot As KeyRoot, KeyPath As String, Optional SubKey As String = "") As Boolean
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long

  
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ALL_ACCESS, hkey)
  If ReturnValue <> 0 Then
    '
    DeleteRegKeyValue = False
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  
  If SubKey = "" Then SubKey = KeyPath
  
  ReturnValue = RegDeleteValue(hkey, SubKey)
  If ReturnValue = 0 Then
    
    DeleteRegKeyValue = True
  Else
    
    DeleteRegKeyValue = False
  End If
  
  ReturnValue = RegCloseKey(hkey)
End Function
Private Function GetSubKeyValue(ByVal hkey As Long, ByVal SubKey As String) As String
  
  On Error Resume Next
  Dim ReturnValue As Long
  Dim KeyType As KeyType
  Dim MyBuffer As String
  Dim MyBufferSize As Long

  
  ReturnValue = RegQueryValueEx(hkey, SubKey, 0, KeyType, ByVal 0, MyBufferSize)
  If ReturnValue = 0 Then
    
    Select Case KeyType
      Case REG_SZ
        
        MyBuffer = String(MyBufferSize, Chr$(0))
        
        ReturnValue = RegQueryValueEx(hkey, SubKey, 0, 0, ByVal MyBuffer, MyBufferSize)
        If ReturnValue = 0 Then
          
          GetSubKeyValue = Left$(MyBuffer, InStr(1, MyBuffer, Chr$(0)) - 1)
        End If
      Case Else
        Dim MyNewBuffer As Long
        
        ReturnValue = RegQueryValueEx(hkey, SubKey, 0, 0, MyNewBuffer, MyBufferSize)
        If ReturnValue = 0 Then
          GetSubKeyValue = MyNewBuffer
        End If
    End Select
  End If
End Function
Private Function EnablePrivilege(seName As String) As Boolean
  On Error Resume Next
  Dim p_lngRtn As Long
  Dim p_lngToken As Long
  Dim p_lngBufferLen As Long
  Dim p_typLUID As LUID
  Dim p_typTokenPriv As TOKEN_PRIVILEGES
  Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

  p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
  If p_lngRtn = 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  If Err.LastDllError <> 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
  If p_lngRtn = 0 Then
    EnablePrivilege = False
    Exit Function
  End If
  p_typTokenPriv.PrivilegeCount = 1
  p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
  p_typTokenPriv.Privileges.pLuid = p_typLUID
  EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

Public Function EnumerateRegKeys1(KeyRoot As KeyRoot, KeyPath As String) As String
  
  On Error Resume Next
  Dim hkey As Long
  Dim ReturnValue As Long
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim ClassNameBuffer As String
  Dim ClassNameBufferSize As Long
  Dim LastWrite As FILETIME

  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ENUMERATE_SUB_KEYS, hkey)
  If ReturnValue <> 0 Then

    EnumerateRegKeys1 = ""
    ReturnValue = RegCloseKey(hkey)
    Exit Function
  End If
  Counter = 0
 
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    ClassNameBuffer = Space(255)
    MyBufferSize = 255
    ClassNameBufferSize = 255
    ReturnValue = RegEnumKeyEx(hkey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
      EnumerateRegKeys1 = MyBuffer
      Exit Function
    End If
    Counter = Counter + 1
    If (Counter Mod 100) = 0 Then
        DoEvents
    End If
  Loop

  If EnumerateRegKeys1 <> "" Then EnumerateRegKeys1 = Left$(EnumerateRegKeys1, Len(EnumerateRegKeys1) - 1)
  
  ReturnValue = RegCloseKey(hkey)
End Function


Public Function SaveValue(ByVal hive As String, ByVal key As String, ByVal valuename As String, ByVal value As Variant, ByVal datatype As String) As Boolean
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim ss As String
Dim s As String
Dim ii As Long
Dim f As Long
hkey = hkeyf(hive)

    Dim a As SECURITY_ATTRIBUTES
    a.lpSecurityDescriptor = 0
    a.bInheritHandle = True
    a.nLength = Len(a)
i = RegCreateKeyEx(hkey, key, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, a, h, k)
If i = ERROR_SUCCESS Then
value = Trim(value)
Select Case LCase(datatype)
Case "string"
    j = RegSetValueEx(h, valuename, 0&, REG_SZ, value, Len(value))
Case "binary"
    For ii = 1 To Len(value) Step 2
         ss = ""
         ss = Mid(value, ii, 2)
         ss = "&H" & ss
         s = s & Chr(Val(CDec(ss)))
    Next
    value = s
    j = RegSetValueExString(h, valuename, 0&, REG_BINARY, value, Len(value))
Case "dword"     'For dword value
    j = RegSetValueExLong(h, valuename, 0&, REG_DWORD, value, 4)
End Select
SaveValue = Message(j)
Else
SaveValue = False
End If
RegCloseKey h
End Function

Public Function hkeyf(ByVal hive As String) As Long
Dim hkey As Long
If hive = "HKEY_CLASSES_ROOT" Then
hkey = HKEY_CLASSES_ROOT
ElseIf hive = "HKEY_CURRENT_USER" Then
hkey = HKEY_CURRENT_USER
ElseIf hive = "HKEY_LOCAL_MACHINE" Then
hkey = HKEY_LOCAL_MACHINE
ElseIf hive = "HKEY_USERS" Then
hkey = HKEY_USERS
ElseIf hive = "HKEY_PERFORMANCE_DATA" Then
hkey = HKEY_PERFORMANCE_DATA
ElseIf hive = "HKEY_CURRENT_CONFIG" Then
hkey = HKEY_CURRENT_CONFIG
ElseIf hive = "HKEY_DYN_DATA" Then
hkey = HKEY_DYN_DATA
End If
hkeyf = hkey
End Function
