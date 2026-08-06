Attribute VB_Name = "PROCVLD"
Option Explicit

Global Ap_VsdPrd  As Recordset
   
Global NumVista  As String * 4
Global IdUsuario As String * 8
Global GetAis    As String * 70
Global Llave     As String
Global Ap_modo   As Integer
Global Modo      As String

Dim Datado As String * 1500
Dim BaseEnvioSRM As Database
Dim BaseVisado   As Database

Dim Usuario As String
Dim Sd_Llave As String
Dim DataUp As String
Dim Nodo_Dst As String
Dim Apl_Dst As String
Global DataPatron As String

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const ARCHIVO_INI01 = "Visado.INI"







Public Function GrabaParametro(Grupo As String, Llave As String, parametro As String, archivo As String) As Boolean
    Dim f As Integer, lpReturnedString As String
    
    lpReturnedString = String$(255, 0)
    parametro = parametro + Mid(lpReturnedString, Len(parametro) + 2)
    f = WritePrivateProfileString(ByVal Grupo, ByVal Llave, ByVal parametro, ByVal archivo)
    GrabaParametro = True

End Function

Public Function ObtieneParametro(Grupo As String, Llave As String, archivo As String) As String
    Dim f As Integer, lpReturnedString As String
    Dim nSize%
    Dim parametro
    
    lpReturnedString = String$(255, 0)
    nSize% = Len(lpReturnedString)
    f = GetPrivateProfileString(Grupo, ByVal Llave, "", lpReturnedString, nSize%, archivo)
    parametro = Left$(lpReturnedString$, f)

    ObtieneParametro = parametro

End Function


Function Ap_Visado()

Dim Ap_Rut  As String
Dim Ap_Producto As String
Dim Ap_Correlativo As String
Dim Continuar As Boolean
Dim OperCRS As String
Dim FechaOtor As String
Dim VarSiebel As String
Dim Ret As Integer
Dim p As Integer

' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado

'Ret = AisGetUsr(GetAis)
IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
 
' OBTIENE LOS PARAMETROS
Ap_Rut = ObtieneParametro("PARAMETROS", "PARAM1", "C:\DATA\VISADO\VISADO.INI")
Ap_Producto = ObtieneParametro("PARAMETROS", "PARAM2", "C:\DATA\VISADO\VISADO.INI")
Ap_Correlativo = ObtieneParametro("PARAMETROS", "PARAM3", "C:\DATA\VISADO\VISADO.INI")
Ap_modo = ObtieneParametro("PARAMETROS", "PARAM4", "C:\DATA\VISADO\VISADO.INI")
VarSiebel = ObtieneParametro("SYBASE", "SIEBEL", "C:\DATA\VISADO\VISADO.INI")

Set BaseVisado = OpenDatabase("C:\DATA\VISADO\VISADO.MDB")
Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
' FI2026 - Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
Ap_VsdPrd.Seek "=", Ap_Rut, Ap_Producto, Ap_Correlativo
   
If Not Ap_VsdPrd.NoMatch Then
   
  If VarSiebel <> "NORMAL" Then
    OperCRS = Mid$(Ap_VsdPrd!PRD_DATA, 40, 1)
    FechaOtor = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 225, 8)
    FechaOtor = Format$(FechaOtor, "&&/&&/&&&&")
 
    If Not ExisteUF(FechaOtor) Then
      Continuar = False
      MsgBox "No se encuentra el valor de la UF para la fecha de otorgamiento del crÃ©dito.", 16, "Error de ValidaciÃ³n"
    Else
      Continuar = True
    End If
  Else
    Continuar = True
  End If
       
  If Continuar Then
    p = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")
  End If
         
End If
   
Ap_VsdPrd.Close
BaseVisado.Close
End
   
End Function



Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal largo As String) As Boolean

Dim FiltroVsdPrd  As Recordset

DataPatron = ""
  
ProductoAsociadoEncontrado = False
  
Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + CStr(Len(PatronBusqueda)) + ")='" + PatronBusqueda + "';", dbOpenSnapshot)
  
If FiltroVsdPrd.RecordCount > 0 Then
  DataPatron = FiltroVsdPrd("Prd_Data")
  ProductoAsociadoEncontrado = True
End If
  
FiltroVsdPrd.Close
  
End Function




Public Function ExisteUF(FechaOtorgamiento As String) As Boolean

Dim LlaveUp As String
Dim DataUp As String
Dim NodoDestino As String
Dim Aplicacion As String
Dim Retorno As String

'  FechaOtorgamiento="05/10/2005"
ExisteUF = False

If IsDate(FechaOtorgamiento) Then

  NodoDestino = NodoSybase("plt_emp01")
  Aplicacion = Mid$(NodoDestino, 5, 4)
  NodoDestino = Mid$(NodoDestino, 1, 4)

  LlaveUp = "b12256781234W00000012plt_emp01 exec sp_busca_unidad_fomento '" & FechaOtorgamiento & "'"

  Retorno = Llamada(NodoDestino, Aplicacion, LlaveUp, "8K", "C")

  If Mid$(Retorno, 1, 6) = "000000" Then
    If Mid$(Retorno, 19, 2) = "SI" Then
       ExisteUF = True
    End If
  End If

End If

End Function
Function NodoSybase(C_SY As String) As String
  Dim PunteroSy  As Recordset
  
  Set BaseEnvioSRM = OpenDatabase("C:\DATA\VISADO\ENVIOSRM.MDB")
  
  Set PunteroSy = BaseEnvioSRM.OpenRecordset("SYBASE", dbOpenTable)
  ' FI2026 - PunteroSy.Index = PunteroSy.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroSy.Seek "=", C_SY
  If Not PunteroSy.NoMatch Then
     If Environ("DESA") = "DESA" Then
        NodoSybase = PunteroSy("Sy_NodoDesa")
     Else
        NodoSybase = PunteroSy("Sy_NodoProd")
     End If
     NodoSybase = NodoSybase + PunteroSy("Sy_Aplicacion")
  Else
    NodoSybase = "XXXXXXXX"
  End If
  PunteroSy.Close
  
  BaseEnvioSRM.Close
End Function
