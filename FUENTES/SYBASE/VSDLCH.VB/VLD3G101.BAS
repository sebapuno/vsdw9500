Attribute VB_Name = "Vld1GR01"
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
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"

Public Const DIRECTORIO_02 = "C:\Bin\S\"
Public Const DIRECTORIO_03 = "C:\Data\V_CtaCte\"
Public Const DIRECTORIO_SGT = "C:\Data\Admsgt\"

Public Const ARCHIVO_INI03 = "V_CtaCte.INI"
Public Const ARCHIVO_PROFI = "V_CtaCte.Exe"

Public Const SINNODO = "INHABILITADO"
Public Const EJECUTABLE = "VSDVL???.EXE"
Public Const RUTDEFECTO = "12345678"
Public Const NODOREGIONAL = "SRM_NS"
Public Const PRODUCTOACCESADO = "(LCH) Línea de Crédito Homologada"
Public Const AMBIENTEDESARROLLO = "DESA"


Global Rut As String
Global Modo As String
Global Producto As String
' FI2026 - la ruta del log de validacion estaba hardcodeada a "d:\VSD_VAL.TXT".
' Viene asi del original de 16 bits: en el PC del banco D: era una unidad de
' datos. En otra maquina D: suele ser el CD-ROM o no existir, y el Open For
' Append corta con error 75 (Path/File access error) -- compila igual, falla
' recien al ejecutar. Se pasa al INI, [VISADO] LogValidacion, con fallback al
' directorio de datos para no depender nunca mas de D:.
Public Function RutaLogValidacion() As String
Dim r As String
  r = Trim$(SacaValorINI("VISADO", "LogValidacion", DIRECTORIO_01 + ARCHIVO_INI01))
  If r = "" Then r = DIRECTORIO_01 + "VSD_VAL.TXT"
  RutaLogValidacion = r
End Function

Function registra_las_llamadas(xnodo As String, xaplica As String, xfuncion As String, xstatus As String, xdata As String)

Dim archivo_LECTURA As String
Dim Archivo_DESTINO As String
Dim XGRABAR As String
Dim MiRegistrox As String * 1000 ' Registro

   
Archivo_DESTINO = RutaLogValidacion()
       
Open Archivo_DESTINO For Append As #2
MiRegistrox = ""
MiRegistrox = Date & " " & Time & "VSDVL3G1.EXE" & " " & xnodo & " " & xaplica & " " & xfuncion & " " & xstatus & " " & xdata
Print #2, MiRegistrox
Close #2
       
End Function

Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
' FI2026 - Retorno era Integer: Shell() en VB6 32-bit devuelve el process ID
' (muy por encima de 32767), que desborda un Integer.
Dim Retorno As Long
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 - restaura la espera del original 16-bit:
  '     While GetModuleUsage(Retorno) <> 0 : DoEvents : Wend
  ' EjecutarPrograma es SINCRONA: quien la llama asume que el programa
  ' externo ya termino al volver.
  Call WaitProcess(Retorno)
End Sub

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function


