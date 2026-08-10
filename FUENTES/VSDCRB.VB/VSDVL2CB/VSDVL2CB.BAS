Attribute VB_Name = "Vld2CB01"
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
Declare Function SetWindowPos Lib "USER32.DLL" (ByVal h&, ByVal hb%, ByVal x&, ByVal y&, ByVal cx%, ByVal cy%, ByVal f&) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
' FI2026 - GetModuleUsage es API Win16 (kernel.dll) y no existe en Win32. Se
' usaba para esperar a que terminara el programa lanzado con Shell. El
' equivalente Win32 es OpenProcess(SYNCHRONIZE) + WaitForSingleObject, que
' se implementa mas abajo en WaitProcess. Este proyecto no incluye
' modRSrm32.bas, por eso las declaraciones van aca.
Private Declare Function OpenProcess Lib "KERNEL32.DLL" (ByVal dwAccess As Long, ByVal fInherit As Long, ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32.DLL" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32.DLL" (ByVal hObject As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZENOMOVE = 3

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const EJECUTABLE = "VSDVL???.EXE"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"
Public Const PRODUCTOACCESADO = "(2CB) Crédito en Cuotas"

Public Const DIRECTORIO_02 = "C:\Bin\S\"
Public Const DIRECTORIO_03 = "C:\Data\V_CtaCte\"
Public Const ARCHIVO_INI03 = "V_CTACTE.INI"

Public Const ARCHIVO_PROFI = "V_CTACTE.EXE"

Global Rut As String
Global Modo As String
Global Producto As String
Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
' FI2026 - Retorno era Integer: Shell() en VB6 32-bit devuelve el process ID
' (muy por encima de 32767), que desborda un Integer.
Dim Retorno As Long
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 - restaura la espera del original 16-bit, que era:
  '     While GetModuleUsage(Retorno) <> 0 : DoEvents : Wend
  ' EjecutarPrograma es SINCRONA: quien la llama asume que el programa
  ' externo ya termino al volver.
  Call WaitProcess(Retorno)
End Sub

' FI2026 - equivalente Win32 del patron Win16 While GetModuleUsage(pid) <> 0
Public Sub WaitProcess(ByVal pid As Long)
Dim hProc As Long
Const SYNCHRONIZE As Long = &H100000
  hProc = OpenProcess(SYNCHRONIZE, 0, pid)
  If hProc <> 0 Then
    Do While WaitForSingleObject(hProc, 100) <> 0
      DoEvents
    Loop
    CloseHandle hProc
  End If
End Sub

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(800, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
