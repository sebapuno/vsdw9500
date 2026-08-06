Attribute VB_Name = "GLOBALES"
Option Explicit

Global Version As String * 45

Rem     ==================================================
Rem               Declaraciones de Manejo archivos (.INI)
Rem     ==================================================

Declare Function WriteProfileString Lib "KERNEL32.DLL" Alias "WriteProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
Declare Function GetProfileString Lib "KERNEL32.DLL" Alias "GetProfileStringA" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Rem     ===========================================
Rem                 Declaraciones de AIS
Rem     ===========================================
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
' FI2026 SHG - waislog.dll 16-bit deshabilitado
'Declare Function LInsert Lib \"...waislog.dll\"


Rem     ===========================================
Rem           Declaraciones de Llamados Shell
Rem     ===========================================

Declare Function FindWindow% Lib "USER32.DLL" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpCaption As Any)

' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Declare Function GetModuleUsage% Lib \"KERNEL32.DLL\"

