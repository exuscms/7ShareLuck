VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDefaultIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API Constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_READ = &H20019 'To allow us to READ the registry keys

'API Declarations
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub GetDefaultIcon(FileName As String, Picture_hDC As Long)
    'Parameters:
    'FileName: The extension of the filename, with the "." e.g .doc
    'Picture_hDC: The Handle to the device context of the Picture Box you want the icon
    'to be displayed on.
    'Example:
    'Call GetDefaultIcon(".doc",Picture1.hDC)
        
    Dim TempFileName As String
    Dim lngError As Long
    Dim lngRegKeyHandle As Long
    Dim strProgramName As String
    Dim strDefaultIcon As String
    Dim lngStringLength As Long
    Dim lngIconNumber As Long
    Dim lngIcon As Long
    Dim intN As Integer

    TempFileName = Right(FileName, Len(FileName) - InStrRev(FileName, ".") + 1)

    If LCase(TempFileName) = ".exe" Then
        strDefaultIcon = Space(260)
        lngStringLength = GetSystemDirectory(strDefaultIcon, 260)
        strDefaultIcon = Left(strDefaultIcon, lngStringLength) & "\SHELL32.DLL"
        lngIconNumber = 2
        GoTo Draw_Icon
    End If

    lngError = RegOpenKey(HKEY_CLASSES_ROOT, TempFileName, lngRegKeyHandle)
    If lngError Then GoTo No_Icon
    lngStringLength = 260
    strProgramName = Space$(260)
    lngError = RegQueryValueEx(lngRegKeyHandle, vbNullString, 0, 0, strProgramName, lngStringLength)
    If lngError Then
        lngError = RegCloseKey(lngRegKeyHandle)
        GoTo No_Icon
    End If
    lngError = RegCloseKey(lngRegKeyHandle)
    strProgramName = Left(strProgramName, lngStringLength - 1)
    lngError = RegOpenKey(HKEY_CLASSES_ROOT, strProgramName & "\DefaultIcon", lngRegKeyHandle)
    If lngError Then GoTo No_Icon
    lngStringLength = 260
    strDefaultIcon = Space$(260)
    lngError = RegQueryValueEx(lngRegKeyHandle, vbNullString, 0, 0, strDefaultIcon, lngStringLength)
    If lngError Then
        lngError = RegCloseKey(lngRegKeyHandle)
        GoTo No_Icon
    End If
    lngError = RegCloseKey(lngRegKeyHandle)
    strDefaultIcon = Trim$(Left(strDefaultIcon, lngStringLength - 1))

    intN = InStrRev(strDefaultIcon, ",")
    If intN < 1 Then GoTo No_Icon
    lngIconNumber = Trim$(Right(strDefaultIcon, Len(strDefaultIcon) - intN))
    strDefaultIcon = Trim$(Left(strDefaultIcon, intN - 1))

Draw_Icon:
    lngIcon = ExtractIcon(App.hInstance, strDefaultIcon, lngIconNumber)
    If lngIcon = 1 Or lngIcon = 0 Then GoTo No_Icon
    lngError = DrawIcon(Picture_hDC, 0, 0, lngIcon)
    lngError = DestroyIcon(lngIcon)
    Exit Sub
No_Icon:
    'No icon could be found so we use the normal windows icon
    'This icon is held in shell32.dll in the system directory, Icon 0
    strDefaultIcon = Space(260)
    lngStringLength = GetSystemDirectory(strDefaultIcon, 260)
    strDefaultIcon = Left(strDefaultIcon, lngStringLength) & "\SHELL32.DLL"
    lngIconNumber = 0
    GoTo Draw_Icon
End Sub
