VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================
' Class:    cRegistry
' Author:   Steve McMahon
' Date  :   21 Feb 1997
' =========================================================

'Registry Specific Access Rights
Private Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_INVALID_DATA = 13&

'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_eValueType As ERegistryValueTypes

Public Property Let Value( _
        ByVal vValue As Variant _
    )
Dim ordType As Long
Dim c As Long
Dim hKey As Long
Dim e As Long
Dim lCreate As Long
Dim tSA As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
'        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
            Case REG_SZ, REG_EXPAND_SZ
                Dim s As String, iPos As Long
                s = vValue
                ordType = REG_SZ
                ' Assume anything with two non-adjacent percents is expanded string
                iPos = InStr(s, "%")
                If iPos Then
                    If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
                End If
                c = Len(s) + 1
                e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
                
            ' User should convert to a compatible type before calling
            Case Else
                e = ERROR_INVALID_DATA
                
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
'            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
        
        'Close the key
        RegCloseKey hKey
    
    End If
    
End Property

Public Function DeleteValue() As Boolean
    Dim e           As Long
    Dim hKey        As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
    If e Then
'        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" _
'        & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
    Else
        e = RegDeleteValue(hKey, m_sValueKey)
        If e Then
'            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete" _
'            registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey _
'             &              "',Key: '" & m_sValueKey
        Else
            DeleteValue = (e = ERROR_SUCCESS)
        End If
    End If
End Function

Public Function DeleteKey() As Boolean
Dim e As Long
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
'        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
    Else
        DeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function

Public Property Get KeyExists() As Boolean
Dim hKey As Long
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
    Else
        KeyExists = False
    End If
    
End Property

Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property

Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
    m_hClassKey = eKey
End Property

Public Property Let SectionKey(ByVal sSectionKey As String)
    m_sSectionKey = sSectionKey
End Property

Public Property Let ValueKey(ByVal sValueKey As String)
    m_sValueKey = sValueKey
End Property

