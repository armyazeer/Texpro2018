Attribute VB_Name = "Module2"
'Module to write to Registry
'required to write  ODBC Database Name
'Yazeer 2013/08/15
'Sourse Internet

Option Explicit

Public Enum REG_TOPLEVEL_KEYS
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_CURRENT_USER = &H80000001
 HKEY_DYN_DATA = &H80000006
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_USERS = &H80000003
End Enum


Private Declare Function RegCreateKey Lib _
   "advapi32.dll" Alias "RegCreateKeyA" _
   (ByVal Hkey As Long, ByVal lpSubKey As _
   String, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib _
   "advapi32.dll" (ByVal Hkey As Long) As Long

Private Declare Function RegSetValueEx Lib _
   "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal Hkey As Long, ByVal _
   lpValueName As String, ByVal _
   Reserved As Long, ByVal dwType _
   As Long, lpData As Any, ByVal _
   cbData As Long) As Long

Private Const REG_SZ = 1



Public Function WriteStringToRegistry(Hkey As _
  REG_TOPLEVEL_KEYS, strPath As String, strValue As String, _
  strdata As String) As Boolean
 
'WRITES A STRING VALUE TO REGISTRY:
'PARAMETERS:

'Hkey: Top Level Key as defined by
'REG_TOPLEVEL_KEYS Enum (See Declarations)

'strPath - 'Full Path of Subkey
'if path does not exist it will be created

'strValue ValueName

'strData - Value Data

'Returns: True if successful, false otherwise

'EXAMPLE:
'WriteStringToRegistry(HKEY_LOCAL_MACHINE, _
"Software\Microsoft", "CustomerName", "FreeVBCode.com")

Dim bAns As Boolean

On Error GoTo ErrorHandler
   Dim keyhand As Long
   Dim r As Long
   r = RegCreateKey(Hkey, strPath, keyhand)
   If r = 0 Then
        r = RegSetValueEx(keyhand, strValue, 0, _
           REG_SZ, ByVal strdata, Len(strdata))
        r = RegCloseKey(keyhand)
    End If
    
   WriteStringToRegistry = (r = 0)

Exit Function

ErrorHandler:
    WriteStringToRegistry = False
    Exit Function
    
End Function

