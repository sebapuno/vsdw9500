Attribute VB_Name = "GesATC01"
Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZENOMOVE = 3

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const EJECUTABLE = "GES_???.EXE"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"
Public Const ARCHIVO_MDB03 = "Oficina.MDB"
Public Const PRODUCTOACCESADO = "(ATC) Activación Tarjeta Bancaria"

Global Rut As String
Global Modo As String
Global Producto As String
Public Function GlosaOficina(CodigoOficina As String) As String
Dim BaseSucur As Database
Dim PunteroSucursales As Recordset
  Set BaseSucur = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB03)
  Set PunteroSucursales = BaseSucur.OpenRecordset("Sucursales", dbOpenTable)
  PunteroSucursales.Index = PunteroSucursales.Indexes(0).Name
  PunteroSucursales.Seek "=", CodigoOficina
  If PunteroSucursales.NoMatch Then
    GlosaOficina = "Oficina " + CodigoOficina + " (No Determinada)"
  Else
    GlosaOficina = PunteroSucursales("Suc_Nombre")
  End If
  GlosaOficina = GlosaOficina + Space(30 - Len(GlosaOficina))
  PunteroSucursales.Close
  BaseSucur.Close
End Function

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
