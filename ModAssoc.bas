Attribute VB_Name = "ModAssoc"
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" _
(ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Function Associate(ByVal apPath As String, ByVal Ext As String) As Boolean
'Borrowed this association function from a submission by
' Insomniaque modified by Dj's Computer Labs
'Da rest is all Bobo Enterprises copyright
  Dim sKeyName As String
  Dim sKeyValue As String
  Dim ret&
  Dim lphKey&
  Dim apTitle As String
  apTitle = ParseName(apPath)
  If InStr(Ext, ".") = 0 Then Ext = "." & Ext
   sKeyName = Ext
  sKeyValue = apTitle
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
  If ret& <> 0 Then GoTo AssocFailed
   sKeyName = apTitle
  sKeyValue = apPath & " %1"
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
  If ret& <> 0 Then GoTo AssocFailed
    sKeyValue = apPath
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
  If ret& <> 0 Then GoTo AssocFailed
   Associate = True
  Exit Function
AssocFailed:
  Associate = False
End Function
Public Function ParseName(ByVal sPath As String) As String
  Dim strX As String
  Dim intX As Integer
  intX = InStrRev(sPath, "\")
  strX = Trim(Right(sPath, Len(sPath) - intX))
  If Right(strX, 1) = Chr(0) Then
    ParseName = Left(strX, Len(strX) - 1)
  Else
    ParseName = strX
  End If
End Function
Public Sub FileSave(Text As String, FilePath As String)
On Error Resume Next
Dim Directory As String
              Directory$ = FilePath
              Open Directory$ For Output As #1
           Print #1, Text
       Close #1
Exit Sub
End Sub
Function TrimVoid(Expre)
  On Error Resume Next
  Dim i As Integer
  Dim beg As String
  Dim expr As String
  For i = 1 To Len(Expre)
        beg = Mid(Expre, i, 1)
        If beg Like "[a-zA-Z0-9]" Then expr = expr & beg
    Next
    TrimVoid = expr
End Function
Public Function GetShortCut(cboindex As Integer) As String
Select Case cboindex
    Case 1
        GetShortCut = "^" + "A"
    Case 2
        GetShortCut = "^" + "B"
    Case 3
        GetShortCut = "^" + "C"
    Case 4
        GetShortCut = "^" + "D"
    Case 5
        GetShortCut = "^" + "E"
    Case 6
        GetShortCut = "^" + "F"
    Case 7
        GetShortCut = "^" + "G"
    Case 8
        GetShortCut = "^" + "H"
    Case 9
        GetShortCut = "^" + "I"
    Case 10
        GetShortCut = "^" + "J"
    Case 11
        GetShortCut = "^" + "K"
    Case 12
        GetShortCut = "^" + "L"
    Case 13
        GetShortCut = "^" + "M"
    Case 14
        GetShortCut = "^" + "N"
    Case 15
        GetShortCut = "^" + "O"
    Case 16
        GetShortCut = "^" + "P"
    Case 17
        GetShortCut = "^" + "Q"
    Case 18
        GetShortCut = "^" + "R"
    Case 19
        GetShortCut = "^" + "S"
    Case 20
        GetShortCut = "^" + "T"
    Case 21
        GetShortCut = "^" + "U"
    Case 22
        GetShortCut = "^" + "V"
    Case 23
        GetShortCut = "^" + "W"
    Case 24
        GetShortCut = "^" + "X"
    Case 25
        GetShortCut = "^" + "Y"
    Case 26
        GetShortCut = "^" + "Z"
    Case 27
        GetShortCut = "{F1}"
    Case 28
        GetShortCut = "{F2}"
    Case 29
        GetShortCut = "{F3}"
    Case 30
        GetShortCut = "{F4}"
    Case 31
        GetShortCut = "{F5}"
    Case 32
        GetShortCut = "{F6}"
    Case 33
        GetShortCut = "{F7}"
    Case 34
        GetShortCut = "{F8}"
    Case 35
        GetShortCut = "{F9}"
    Case 36
        GetShortCut = "{F10}"
    Case 37
        GetShortCut = "{F11}"
    Case 38
        GetShortCut = "{F12}"
    Case 39
        GetShortCut = "^{F1}"
    Case 40
        GetShortCut = "^{F2}"
    Case 41
        GetShortCut = "^{F3}"
    Case 42
        GetShortCut = "^{F4}"
    Case 43
        GetShortCut = "^{F5}"
    Case 44
        GetShortCut = "^{F6}"
    Case 45
        GetShortCut = "^{F7}"
    Case 46
        GetShortCut = "^{F8}"
    Case 47
        GetShortCut = "^{F9}"
    Case 48
        GetShortCut = "^{F10}"
    Case 49
        GetShortCut = "^{F11}"
    Case 50
        GetShortCut = "^{F12}"
    Case 51
        GetShortCut = "+{F1}"
    Case 52
        GetShortCut = "+{F2}"
    Case 53
        GetShortCut = "+{F3}"
    Case 54
        GetShortCut = "+{F4}"
    Case 55
        GetShortCut = "+{F5}"
    Case 56
        GetShortCut = "+{F6}"
    Case 57
        GetShortCut = "+{F7}"
    Case 58
        GetShortCut = "+{F8}"
    Case 59
        GetShortCut = "+{F9}"
    Case 60
        GetShortCut = "+{F10}"
    Case 61
        GetShortCut = "+{F11}"
    Case 62
        GetShortCut = "+{F12}"
    Case 63
        GetShortCut = "+^{F1}"
    Case 64
        GetShortCut = "+^{F2}"
    Case 65
        GetShortCut = "+^{F3}"
    Case 66
        GetShortCut = "+^{F4}"
    Case 67
        GetShortCut = "+^{F5}"
    Case 68
        GetShortCut = "+^{F6}"
    Case 69
        GetShortCut = "+^{F7}"
    Case 70
        GetShortCut = "+^{F8}"
    Case 71
        GetShortCut = "+^{F9}"
    Case 72
        GetShortCut = "+^{F10}"
    Case 73
        GetShortCut = "+^{F11}"
    Case 74
        GetShortCut = "+^{F12}"
    Case 75
        GetShortCut = "^{INSERT}"
    Case 76
        GetShortCut = "+{INSERT}"
    Case 77
        GetShortCut = "{DEL}"
    Case 78
        GetShortCut = "+{DEL}"
    Case 79
        GetShortCut = "%{BKSP}"
End Select
End Function

