Attribute VB_Name = "Vld2CB01"
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
Declare Function SetWindowPos Lib "USER32.DLL" (ByVal h&, ByVal hb%, ByVal x&, ByVal y&, ByVal cx%, ByVal cy%, ByVal f&) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Private Declare Function GetModuleUsage Lib "KERNEL32.DLL" (ByVal handler As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZENOMOVE = 3

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const EJECUTABLE = "VSDVL???.EXE"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"
Public Const PRODUCTOACCESADO = "(2CB) CrÃƒÆ’Ã‚Â©dito en Cuotas"

Public Const DIRECTORIO_02 = "C:\Bin\S\"
Public Const DIRECTORIO_03 = "C:\Data\V_CtaCte\"
Public Const ARCHIVO_INI03 = "V_CTACTE.INI"

Public Const ARCHIVO_PROFI = "V_CTACTE.EXE"

Global Rut As String
Global Modo As String
Global Producto As String
Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
Dim Retorno As Integer
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
  ' Call WaitProcess(Retorno)  ' requiere modRSrm32.bas

     ' DoEvents  ' FI2026 SHG - orphan from GetModuleUsage
  ' Wend  ' FI2026 SHG - orphan from GetModuleUsage
End Sub

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(800, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
