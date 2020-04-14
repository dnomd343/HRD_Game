Attribute VB_Name = "Module"
Option Explicit
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hkey As Long, ByVal pszSubKey As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public debug_mode As Boolean, on_top As Boolean, playing As Boolean, solve_compete As Boolean
Public block_line_width As Integer, case_line_width As Integer
Public block_color, block_line_color, case_color, case_line_color
Public change_case As Boolean, change_case_title As String, change_case_code As String
Public Favourite_Cases_name() As String, Favourite_Cases_code() As String
Public favourite_add_name As String, favourite_add_code As String, favourite_add_confirm As Boolean
Public favourite_add_init_name As String, favourite_add_init_code As String, favourite_add_save As Boolean
Public wait_file_name As String, wait_cancel As Boolean, waiting As Boolean
Public start_code As String
Public Sub FindKeys(hkey As Long, SubKey As String)
  Dim phkRet As Long, lRet As Long, index As Long, lName As Long, lReserved As Long, lClass As Long
  Dim name As String, Class As String
  Dim LWT As FILETIME
  lReserved = 0
  index = 0
  lRet = RegOpenKey(hkey, SubKey, phkRet)
  If lRet = 0 Then
    Do
      name = String(255, Chr(0)): lName = Len(name)
      lRet = RegEnumKeyEx(phkRet, index, name, lName, lReserved, Class, lClass, LWT)
      If lRet = 0 Then
        ReDim Preserve Favourite_Cases_name(UBound(Favourite_Cases_name) + 1)
        Favourite_Cases_name(UBound(Favourite_Cases_name)) = name
      Else
        Exit Do
      End If
      index = index + 1
    Loop While lRet = 0
  End If
  Call RegCloseKey(phkRet)
End Sub
Public Sub Get_Favourite_Cases()
  Dim i As Long, w
  Dim temp As String
  Set w = CreateObject("WScript.Shell")
  ReDim Favourite_Cases_name(0)
  Call FindKeys(HKEY_CURRENT_USER, "Software\HRD_Game\Favourite")
  ReDim Favourite_Cases_code(UBound(Favourite_Cases_name))
  For i = 1 To UBound(Favourite_Cases_name)
    temp = Favourite_Cases_name(i)
    temp = Left(temp, InStr(1, temp, Chr(0)) - 1)
    Favourite_Cases_code(i) = w.RegRead("HKEY_CURRENT_USER\Software\HRD_Game\Favourite\" & temp & "\")
    temp = Right(temp, Len(temp) - InStr(1, temp, "."))
    Favourite_Cases_name(i) = temp
  Next i
End Sub
Public Sub Save_Favourite_Cases()
  Dim i As Long, length As Integer, w
  Dim temp As String
  Set w = CreateObject("WScript.Shell")
  Call SHDeleteKey(HKEY_CURRENT_USER, "Software\HRD_Game\Favourite")
  length = Len(Trim(UBound(Favourite_Cases_name)))
  For i = 1 To UBound(Favourite_Cases_name)
    temp = i
    temp = String(length - Len(temp), "0") & temp
    w.regWrite "HKEY_CURRENT_USER\Software\HRD_Game\Favourite\" & temp & "." & Favourite_Cases_name(i) & "\", Favourite_Cases_code(i), "REG_SZ"
  Next i
End Sub


