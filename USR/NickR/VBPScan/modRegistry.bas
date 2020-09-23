Attribute VB_Name = "modRegistry"
Option Explicit

'**************************************
'Windows API/Global Declarations for Registry
'**************************************

'Registry Constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
'Registry Specific Access Rights
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = &H3F
'Open/Create Options
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_OPTION_VOLATILE = &H1
'Key creation/open disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
'masks for the predefined standard access types
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
'Define severity codes
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_NO_MORE_ITEMS = 259
'Predefined Value Types
Public Const REG_NONE = (0) 'No value type
Public Const REG_SZ = (1) 'Unicode nul terminated string
Public Const REG_EXPAND_SZ = (2) 'Unicode nul terminated string w/enviornment var
Public Const REG_BINARY = (3) 'Free form binary
Public Const REG_DWORD = (4) '32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = (4) '32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = (5) '32-bit number
Public Const REG_LINK = (6) 'Symbolic Link (unicode)
Public Const REG_MULTI_SZ = (7) 'Multiple Unicode strings
Public Const REG_RESOURCE_LIST = (8) 'Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9) 'Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)

'Structures Needed For Registry Prototypes
Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Registry Function Prototypes
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

