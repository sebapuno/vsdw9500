Attribute VB_Name = "PTAFCT00"
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function SetWindowPos Lib "USER32.DLL" (ByVal h&, ByVal hb%, ByVal x&, ByVal y&, ByVal cx%, ByVal cy%, ByVal f&) As Long

Global BaseVisado As Database
Global PunteroVsdPrd As Recordset
Global DataPatron As String

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"

Public Const PRODUCTOACCESADO = "(MHB) MAC HistÃƒÂ³rico Empresas B"
Public Const EJECUTABLE = "PTAMHB00.EXE"

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZENOMOVE = 3
Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(1000, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
