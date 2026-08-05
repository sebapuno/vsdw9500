Attribute VB_Name = "Vld1GR01"
Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer
Private Declare Function GetModuleUsage Lib "kernel.dll" (ByVal handler As Integer) As Integer

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

Function registra_las_llamadas(xnodo As String, xaplica As String, xfuncion As String, xstatus As String, xdata As String)

Dim archivo_LECTURA As String
Dim Archivo_DESTINO As String
Dim XGRABAR As String
Dim MiRegistrox As String * 1000 ' Registro

   
Archivo_DESTINO = "d:\VSD_VAL.TXT"
       
Open Archivo_DESTINO For Append As #2
MiRegistrox = ""
MiRegistrox = Date & " " & Time & "VSDVL3G1.EXE" & " " & xnodo & " " & xaplica & " " & xfuncion & " " & xstatus & " " & xdata
Print #2, MiRegistrox
Close #2
       
End Function

Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
Dim Retorno As Integer
  Retorno = Shell(Programa, Modo)
  DoEvents
  While GetModuleUsage(Retorno) <> 0
     DoEvents
  Wend
End Sub

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function


