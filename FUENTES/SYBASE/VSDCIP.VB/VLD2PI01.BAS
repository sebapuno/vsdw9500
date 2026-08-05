Attribute VB_Name = "VldCIP01"
Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZENOMOVE = 3

Public Const DIRECTORIO_01 = "C:\Data\Visado\"

Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"

Public Const SINNODO = "INACTIVO"
Public Const EJECUTABLE = "VSDVL???.EXE"
Public Const RUTDEFECTO = "12345678"
Public Const NODOREGIONAL = "SRM_NS"
Public Const PRODUCTOACCESADO = "(CIP) Cta. Cte. Con Intereses"
Public Const AMBIENTEDESARROLLO = "DESA"

Global Rut As String
Global Modo As String
Global Producto As String

Global IdUsuario As String * 8
Global CadenaDatos As String
Global BaseVisado As Database

Public Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  PunteroVistas.Index = PunteroVistas.Indexes(0).Name
  PunteroVistas.Seek "=", Vista
  If Not PunteroVistas.NoMatch Then
    If Environ$(AMBIENTEDESARROLLO) = "DESA" Then
      If PunteroVistas("Vta_NodoDesa") = "U???" Then
        DeterminaNodo = LCase$(Environ$(NODOREGIONAL))
      Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoDesa"))
      End If
    Else
      If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Environ$(NODOREGIONAL))
      Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoProd"))
      End If
    End If
    
    Aplicacion = PunteroVistas("Vta_Aplicacion")
    
  End If
  PunteroVistas.Close
  BaseEnvioSRM.Close
End Function

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
