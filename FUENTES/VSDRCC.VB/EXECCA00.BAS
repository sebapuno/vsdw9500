Attribute VB_Name = "ExeCCA00"
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long

Public Const DIRECTORIO_01 = "C:\Data\Visado\"

Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSrm.MDB"

Public Const RUTDEFECTO = "12345678"
Public Const SIMBOLOLARGO = "L"
Public Const SIMBOLOPOSICION = "P"

