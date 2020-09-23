Attribute VB_Name = "modControlPanel"
' This Modul read and write Registrykeys.
'---------------------------------------------------------------
'- API-Declarationen for Registrysettings
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

'---------------------------------------------------------------
'- API-Constant for Registry...
'---------------------------------------------------------------
' Registrytype...
Const REG_SZ = 1 ' Null-terminate Unicode-Zeichenfolge
Const REG_EXPAND_SZ = 2 ' Null-terminate Unicode-Zeichenfolge
Const REG_BINARY = 3&
Const REG_DWORD = 4 ' 32-Bit-Number

' Create Registrykeys-Typs...
Const REG_OPTION_NON_VOLATILE = 0 ' Key is exists on Systemstart

' Registrykeys Securityoptions...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Registry-types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' Feedback...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- Accessattrib of Registry...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Boolean
End Type

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
'-----------------------------------------------------------
'Example - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-----------------------------------------------------------
Private Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
Dim rc As Long ' Feedback-Code
Dim hKey As Long ' Nummer for Registrykey
Dim hDepth As Long '
Dim lpAttr As SECURITY_ATTRIBUTES ' Security for Registry

lpAttr.nLength = 50 ' Set the Securitysattrib of Default Settings...
lpAttr.lpSecurityDescriptor = 0 ' ...
lpAttr.bInheritHandle = True ' ...

'------------------------------------------------------------
'- Create or Open Registrierykey...
'------------------------------------------------------------
rc = RegCreateKeyEx(KeyRoot, KeyName, _
0, REG_SZ, _
REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
hKey, hDepth) ' //KeyRoot//KeyName create/open

If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError ' Errorhandler...

'------------------------------------------------------------
'- Create/Open Key...
'------------------------------------------------------------
If (SubKeyValue = "") Then SubKeyValue = " " ' RegSetValueEx() needs a Space...

' Create/Open Key
rc = RegSetValueEx(hKey, SubKeyName, _
0, REG_SZ, _
SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError ' Errorhandler
'------------------------------------------------------------
'- Close Registry...
'------------------------------------------------------------
rc = RegCloseKey(hKey) ' Close Key

UpdateKey = True ' Feedback Registry entry as boolean
Exit Function ' End
CreateKeyError:
UpdateKey = False ' Errorcodes
rc = RegCloseKey(hKey) ' Try to close the key
End Function

'------------------------------------------------------------
'Example - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'------------------------------------------------------------
Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
Dim i As Long ' Counter
Dim rc As Long ' Feedback-Code
Dim hKey As Long ' Nummer for the Open Key
Dim hDepth As Long '
Dim sKeyVal As String
Dim lKeyValType As Long ' Datetyp of one Registrykey
Dim tmpVal As String ' Temp
Dim KeyValSize As Long ' Size of Registryvar's

' Registrykey under this tree {HKEY_LOCAL_MACHINE...} open
'------------------------------------------------------------
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registrierykey

If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Errorhandler...

tmpVal = String$(1024, 0) ' Save Space for Var's
KeyValSize = 1024 ' Declare Size of Var's

'------------------------------------------------------------
' Call Registrykeyvalue...
'------------------------------------------------------------
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
lKeyValType, tmpVal, KeyValSize) ' Create/Open Value

If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Errorhandler

tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

'------------------------------------------------------------
' Setting Up Keytyp...
'------------------------------------------------------------
Select Case lKeyValType ' searching Datatyp...
Case REG_SZ, REG_EXPAND_SZ ' Character sequence for type of registration key data
sKeyVal = tmpVal ' Copy Character sequence
Case REG_DWORD ' type of registration key DWORD
For i = Len(tmpVal) To 1 Step -1 ' Each bit convert
sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1))) ' Provide worth indications of indications
Next
sKeyVal = Format$("&h" + sKeyVal) ' DWORD convert in character sequence
End Select

GetKeyValue = sKeyVal ' Return worth
rc = RegCloseKey(hKey) ' Registration keys close
Exit Function ' End

GetKeyError: ' Settle, after an error arose...
GetKeyValue = vbNullString ' Set Return value on empty character sequence
rc = RegCloseKey(hKey) ' Registration keys close
End Function
Private Function ErrorMsg(lErrorCode As Long) As String

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
GetErrorMsg = "Undefined Error Code: " & Str$(lErrorCode)
End Select

End Function

Private Function GetMainKeyHandle(MainKeyName As String) As Long

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

Private Sub ParseKey(KeyName As String, Keyhandle As Long)

rtn = InStr(KeyName, "\")

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then
MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName
Exit Sub
ElseIf rtn = 0 Then
Keyhandle = GetMainKeyHandle(KeyName)
KeyName = ""
Else
Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1))
KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub

Private Function DeleteKey(KeyName As String)

Call ParseKey(KeyName, MainKeyHandle)

If MainKeyHandle Then
rtn = RegDeleteKey(MainKeyHandle, KeyName)
End If

End Function

Private Function CreateKey(SubKey As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
If rtn = ERROR_SUCCESS Then 'if the key was created then
rtn = RegCloseKey(hKey) 'close the key
End If
End If

End Function

Private Function SetBinaryValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
If rtn = ERROR_SUCCESS Then
lDataSize = Len(Value)
ReDim ByteArray(lDataSize)
For i = 1 To lDataSize
ByteArray(i) = Asc(Mid$(Value, i, 1))
Next
rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize)
If Not rtn = ERROR_SUCCESS Then
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn)
End If
End If
rtn = RegCloseKey(hKey)
Else '
If DisplayErrorMsg = True Then
MsgBox ErrorMsg(rtn)
End If
End If
End If

End Function

'Here the function to provide an entry into the control panel.
Public Function CreateEntryToSystemPanel(GUID As String, Titel As String, ToolTipText As String, IconDatei As String, FileToOpen As String)

' Attitudes for the entry specify
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID, "", Titel
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID, "InfoTip", ToolTipText
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID & "\DefaultIcon", "", IconDatei
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID & "\InProcServer32", "", "shell32.dll"
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID & "\InProcServer32", "ThreadingModel", "Apartment"
UpdateKey HKEY_CLASSES_ROOT, "CLSID\" & GUID & "\Shell\Open\Command", "", FileToOpen

' Entry into the list "activate"
UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\" & GUID, "", ""
UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\" & GUID, "", ""
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder"
SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder", "Attributes", Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)

End Function

Public Function DeleteEntryFromSystemPanel(GUID As String)

DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID
DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\DefaultIcon"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\InProcServer32"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\Shell\Open\Command"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellEx\PropertySheetHandlers\" & GUID & ""
DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder"
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\ CurrentVersion\Explorer\Desktop\NameSpace\" & GUID
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\ CurrentVersion\Explorer\ControlPanel\NameSpace\" & GUID

End Function


