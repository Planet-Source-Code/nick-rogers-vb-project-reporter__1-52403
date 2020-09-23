VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************
' Name: clsRegistry
' Description:A neat class module to give access to the registry.
'   This class allows:
'   get/set of registry entries,
'   checking if keys exist in the registry,
'   enumeration of entries within a section and
'   enumeration of sub sections within a section.
' By: Steve McMahon
'
' Returns:None
'
'Assumes:Save the source code into a file called
'   clsRegistry.cls, and name the class 'clsRegistry'.
'   Then follow the example code listed above.
'
'Side Effects:Currently, the class only
'     returns string values from the registry.
'     Because VB makes automatic ANSI to UNICODE
'     conversions, querying or writing other
'     types of values is not recommended,
'     particulary binary values in the registry.
'
'This code is copyrighted and has limited warranties.
'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.609/lngWId.1/qx/vb/scripts/ShowCode.htm
'for details.
'**************************************

Option Explicit
' =========================================================
' Description:
' A nice class wrapper around the registry functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Sample code: finds the location of the Common Files
' directory on the user's machine:
'
'Dim cR As New clsRegistry
'With cR
'.ClassKey = HKEY_LOCAL_MACHINE
'.SectionKey = "\SOFTWARE\Microsoft\Windows\CurrentVersion"
'.ValueKey = "CommonFilesDir"
' .Default = "?WHERE?"
'If (.Value <> .Default) Then
' MsgBox "Program files at: " & .Value,
'     vbInformation
'Else
'MsgBox "Foobar- Failed to find.", vbExclamation
'End If
'
' Class:clsRegistry
' Author:Steve McMahon (steve-mcmahon@pa
'     -consulting.com)
' Date :21 Feb 1997
'
' ======================================
'     ===================
' Store the current user settings:
Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_sValue As String
Private m_sSetValue As String
Private m_sDefault As String

Property Get ClassKey() As Long
    ' The Registry Class to search in, e.g. HKEY_CLASSES_ROOT, HKEY_CLASSES_LOCAL_MACHINE etc
    ClassKey = m_hClassKey
End Property
Property Let ClassKey(ByVal lKey As Long)
    ' The Registry Class to search in, e.g. HKEY_CLASSES_ROOT, HKEY_CLASSES_LOCAL_MACHINE etc
    m_hClassKey = lKey
End Property

Property Get SectionKey() As String
    ' The "directory" to search in, e.g. "\SOFTWARE\Microsoft\Windows\CurrentVersion"
    SectionKey = m_sSectionKey
End Property
Property Let SectionKey(ByVal sSectionKey As String)
    ' The "directory" to search in, e.g. "\SOFTWARE\Microsoft\Windows\CurrentVersion"
    m_sSectionKey = sSectionKey
End Property

Property Get ValueKey() As String
    ' The value to look at, e.g. "" for default, "CommonFilesDir" for the key named CommonFilesDir
        ValueKey = m_sValueKey
End Property
Property Let ValueKey(ByVal sValueKey As String)
    ' The value to look at, e.g. "" for default, "CommonFilesDir" for the key named CommonFilesDir
        m_sValueKey = sValueKey
End Property

Property Get KeyExists() As Boolean
    ' Returns whether the "directory" set up in
    ' SectionKey exists within the current ClassKey
    KeyExists = bCheckKeyExists(m_hClassKey, m_sSectionKey)
End Property

Property Get Default() As String
    ' Default to return if anything goes awry:
    Default = m_sDefault
End Property
Property Let Default(ByVal sDefault As String)
    ' Default to return if anything goes awry:
    m_sDefault = sDefault
End Property

Property Get Value() As String
    Dim sValue As String
    ' Gets the value associated with the current ClassKey and Section
    If (bGetRegValueSearchInKey(m_hClassKey, m_sSectionKey, m_sValueKey, sValue)) Then
        Value = sValue
    Else
        ' Return default
        Value = m_sDefault
    End If
End Property
Property Let Value(ByVal sValue As String)
    ' Sets the value associated with the current ClassKey and Section
    If (bSetRegValue(m_hClassKey, m_sSectionKey, m_sValueKey, sValue)) Then
        m_sValue = sValue
    Else
        Err.Raise 26001, App.EXEName & ".clsRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_sValue & "'"
    End If
End Property

Public Sub EnumerateValues(ByRef sKeys() As String, ByRef iKeyCount As Integer)
    ' Returns all the value names and values within a section into a string array.
    ' The string array dimensioned
    '(1,n) = Value Name
    '(2,n) = Value
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim szBuffer2 As String
    Dim lBuffSize2 As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    iKeyCount = 0
    Erase sKeys
    lIndex = 0
    lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, phkResult)

    Do While lResult = ERROR_SUCCESS
        'Set buffer space
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        'Get next value
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)


        If (lResult = ERROR_SUCCESS) Then
            iKeyCount = iKeyCount + 1
            ReDim Preserve sKeys(1 To 2, 1 To iKeyCount) As String
            sKeys(1, iKeyCount) = Left(szBuffer, lBuffSize)
            sKeys(2, iKeyCount) = Left$(szBuffer2, lBuffSize2)
        End If
        lIndex = lIndex + 1
    Loop
    RegCloseKey phkResult
End Sub

Public Sub EnumerateSections(ByRef sSect() As String, ByRef iSectCount As Integer)
    ' Returns the names of all the sub-sections (sub "directories") within the current section
    ' in a 1 dimensional array:
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    iSectCount = 0
    Erase sSect
    lIndex = 0
    lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, phkResult)


    Do While lResult = ERROR_SUCCESS
        'Set buffer space
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        'Get next value
        lResult = RegEnumKey(phkResult, lIndex, szBuffer, lBuffSize)

        If (lResult = ERROR_SUCCESS) Then
            iSectCount = iSectCount + 1
            ReDim Preserve sSect(1 To iSectCount) As String
            sSect(iSectCount) = Left(szBuffer, lBuffSize)
        End If
        lIndex = lIndex + 1
    Loop
    RegCloseKey phkResult
End Sub

Public Function CreateKey() As Boolean
    ' Create the current section
    CreateKey = bCreateNewKey(m_hClassKey, m_sSectionKey)
End Function

Private Function bSetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    ' Private function to set a registry value
    On Error GoTo ERROR_HANDLER
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    'Note: This function will create the key or value if it doesn't exist.
    'Open or Create the key
    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue, CLng(Len(sValue) + 1))
    'Close the key
    RegCloseKey phkResult
    'Return SetRegValue Result
    bSetRegValue = (lResult = ERROR_SUCCESS)
    If lResult = 87 Then bSetRegValue = True
    Exit Function

ERROR_HANDLER:
    MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) & "Please exit and try again."
    bSetRegValue = False
End Function

Private Function bGetRegValueSearchInKey(ByVal hKey As Long, ByVal sKey As String, ByVal sSubKey As String, ByRef sValue As String) As Boolean
    ' Private function servicing get value calls.
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim bFound As Integer
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim szBuffer2 As String
    Dim lBuffSize2 As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    lIndex = 0
    lResult = RegOpenKeyEx(hKey, sKey, 0, 1, phkResult)

    Do While lResult = ERROR_SUCCESS And Not (bFound)
        'Set buffer space
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        'Get next value
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)


        If (lResult = ERROR_SUCCESS) Then
            sCompKey = Left(szBuffer, lBuffSize)

            If (sCompKey = sSubKey) Then
                sValue = Left(szBuffer2, lBuffSize2 - 1)
                bFound = True
            End If
        End If
        lIndex = lIndex + 1
    Loop
    RegCloseKey phkResult
    bGetRegValueSearchInKey = bFound
End Function

Private Function bCheckKeyExists(ByVal hKey As Long, ByVal strKey As String) As Boolean
    ' Private function servicing CheckIfKeyExists call
    Dim phkResult As Long

    If RegOpenKeyEx(hKey, strKey, 0, 1, phkResult) = ERROR_SUCCESS Then
        bCheckKeyExists = True
        RegCloseKey phkResult
    Else
        bCheckKeyExists = False
    End If
End Function

Private Function bCreateNewKey(hKey As Long, strKey As String) As Boolean
    ' Private function to create a new subkey if not already present.
    Dim phkResult As Long
    Dim tSA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    'Create default SubKey if it does not exist
    If RegCreateKeyEx(hKey, strKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, phkResult, lCreate) <> ERROR_SUCCESS Then
        'Close default SubKey
        RegCloseKey phkResult
        bCreateNewKey = True
    Else
        bCreateNewKey = False
    End If
End Function
