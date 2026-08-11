Attribute VB_Name = "VldCPH01"
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
Declare Function SetWindowPos Lib "USER32.DLL" (ByVal h&, ByVal hb%, ByVal x&, ByVal y&, ByVal cx%, ByVal cy%, ByVal f&) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long

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
Public Const PRODUCTOACCESADO = "(CPH) Cuenta Corriente Homologada"
Public Const AMBIENTEDESARROLLO = "DESA"

Global Rut As String
Global Modo As String
Global Producto As String

Global IdUsuario As String * 8
Global CadenaDatos As String

Public Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - era: PunteroVistas.Index = PunteroVistas.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  PunteroVistas.Index = BaseEnvioSRM.TableDefs("Vistas").Indexes(0).Name
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
MiRegistrox = Date & " " & Time & "VSDVL2PH.EXE" & " " & xnodo & " " & xaplica & " " & xfuncion & " " & xstatus & " " & xdata
Print #2, MiRegistrox
Close #2
       
End Function
Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
