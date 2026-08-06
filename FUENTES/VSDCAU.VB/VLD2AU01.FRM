VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   2820
   ClientTop       =   4245
   ClientWidth     =   4350
   Height          =   2220
   Left            =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   3900
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String * 8
Dim CadenaDatos As String

Private Function DeterminaNodo(Aplicacion As String, NumeroVista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset
Dim BaseEnvioSRM As Database
Dim BaseSGT      As Database
Dim PunteroORG   As Recordset
Dim Oficina      As String
Dim Region       As String
Dim Color        As String

If Trim$(NumeroVista) = "" Then
  '* oficina de la cuenta
  Oficina = SacaValorINI("DATOS", "OFICINA", DIRECTORIO_03 + ARCHIVO_INI03)
  Color = SacaValorINI("DATOS", "COLOR", DIRECTORIO_03 + ARCHIVO_INI03)
  NumeroVista = ""

  Set BaseSGT = OpenDatabase(DIRECTORIO_SGT + "TABSGT.MDB")
  Set PunteroORG = BaseSGT.OpenRecordset("Select * From TBLORG Where ORG_ORGCUI=" + Oficina, dbOpenDynaset)
  If PunteroORG.RecordCount > 0 Then
    DeterminaNodo = PunteroORG!Org_OrgScj
    Region = PunteroORG!Org_OrgReg
    If Color <> "A" Or Region = "13" Then
       NumeroVista = "3030"
    Else
       NumeroVista = "6391"
    End If
  Else
    BuscaNodoOfi = SINNODO
  End If
  PunteroORG.Close
  BaseSGT.Close
End If

If NumeroVista <> "" Then
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVista.Index = PunteroVista.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroVista.Seek "=", NumeroVista
  If PunteroVista.NoMatch Then
    DeterminaNodo = SINNODO
  Else
    If PunteroVista!Vta_NodoProd <> "U???" Then
       DeterminaNodo = PunteroVista!Vta_NodoProd
    End If
   Aplicacion = PunteroVista!Vta_Aplicacion
  End If
  PunteroVista.Close
  BaseEnvioSRM.Close
End If

End Function
Private Function ExisteChileCardVax(Nodo As String, Apli As String, Vista As String, Llave As String) As Boolean
'* CONSULTA CUENTA REGIONAL
Dim Usuario As String
  
Modo = "L"
ExisteChileCardVax = False

Usuario = IdUsuario + Vista + Modo
CadenaDatos = Llamada(Nodo, Apli, Usuario, Llave, "1K")
If Mid$(CadenaDatos, 1, 6) = "000096" Then
  MsgBox "No Existe la Cuenta Corriente Asociada" + Chr(10) + _
         "al Producto Otorgamiento ChileCard", 0 + 16, "Problemas de ValidaciÃ³n"
Else
  If Mid$(CadenaDatos, 1, 6) = "000000" Then
    ExisteChileCardVax = True
  End If
End If

End Function

Private Sub Form_Load()
Dim GetAis         As String * 70
Dim Correlativo    As String
Dim Respuesta      As Integer
Dim BaseVisado     As Database
Dim PunteroVsdPrd  As Recordset
Dim Existe         As String
Dim Valido         As String
Dim Var_Ofi        As String
Dim A              As Boolean
Dim CtaCorriente   As String
Dim NodoDestino  As String
Dim ApliDestino  As String * 4
Dim Vista        As String
Dim TipoCuenta    As String
Dim RetCon  As Boolean
Dim Continuar   As Boolean

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  ' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado
  'Respuesta = AisGetUsr(GetAis)
  IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
  PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
  
  CtaCorriente = "0" + Mid$(PunteroVsdPrd("Prd_Data"), 67, 10)
  '** tipo cuenta empresa   1=A , 2=B, 3=C, 4=D
  TipoCuenta = Mid$(PunteroVsdPrd("Prd_Data"), 146, 1)
  Continuar = True
  
  Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + CtaCorriente, 1)
  Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
  Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
     
  If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
    '* Var_Ofi = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
    '* A = PoneValorINI("PRODUCTO", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
    '* A = PoneValorINI("PRODUCTO", "Oficina", Var_Ofi, DIRECTORIO_01 + ARCHIVO_INI01)
      
     If Val(Mid$(Rut, 1, Len(Rut) - 1)) > 50000000 And TipoCuenta <> "3" Then '* 3=C
       If ExisteCtaCTR("0" + CtaCorriente, RetCon) Then
          MsgBox "El producto Chilecard asociado a una Cuenta Corriente Empresa" + Chr(10) + _
                 "con InterÃ©s sÃ³lo estÃ¡ habilitado para Cuentas tipo C.", 16, "Problemas de ValidaciÃ³n"
          Continuar = False
       ElseIf RetCon = False Then
          Continuar = False
       End If
     End If
       
     If Continuar Then
       
       Vista = ""
       NodoDestino = DeterminaNodo(ApliDestino, Vista)
    
       If Vista = "6391" Then '* REGIONAL *
         If ExisteChileCardVax(NodoDestino, ApliDestino, Vista, "0" + CtaCorriente) Then
           Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
         End If
       Else  '* METRO *
         If ExisteChileCardIBM(NodoDestino, ApliDestino, Vista, CtaCorriente) Then
           Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
         End If
       End If
     End If
     
  Else
    MsgBox "La Cuenta Corrriente asociada al producto Chilecard NO EXISTE.", 16, "Problemas de ValidaciÃ³n"
  End If
  
  PunteroVsdPrd.Close
  BaseVisado.Close
End If

End


End Sub


Public Function ExisteCtaCTR(eCuenta As String, RetCons As Boolean) As Boolean

Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  
ExisteCtaCTR = False
RetCons = False
NodoDestino = DeterminaNodo(AplicacionDestino, "CTR1")
If Trim$(NodoDestino) <> SINNODO Then
  Usuario = "02"
  Llave = Right(String(12, "0") + Trim$(eCuenta), 12)
  CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave, "8K")
  If Mid$(CadenaDatos, 1, 4) = "0000" Then
    If Mid$(CadenaDatos, 5, 2) = "00" And Mid$(CadenaDatos, 7, 12) = Mid$(Llave, 1, 12) Then
      RetCons = True
      ExisteCtaCTR = True
    ElseIf Mid$(CadenaDatos, 5, 2) = "12" Then
      RetCons = True
      ' CUENTA NO EXISTE EN MAESTRO DE ATRIBUTOS
    Else
      ' ERRORES DEL SERVIDOR CTR1
    End If
  End If
Else
  MsgBox "Vista NÂº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
End If

End Function

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function
Private Function ExisteChileCardIBM(Nodo As String, Apli As String, Vista As String, Llave As String) As Boolean
'* CONSULTA CUENTA METRO
Dim Usuario As String
Dim DatResp As String

Modo = "L"
ExisteChileCardIBM = False
  
Usuario = IdUsuario + Vista + Modo
CadenaDatos = Llamada(Nodo, Apli, Usuario, Llave, "1K")
If Mid$(CadenaDatos, 1, 6) = "000096" Then
  MsgBox "No Existe la Cuenta Corriente Asociada" + Chr(10) + _
         "al Producto Otorgamiento ChileCard", 0 + 16, "Problemas de ValidaciÃ³n"
Else
  If Mid$(CadenaDatos, 1, 6) = "000000" Then
    DatResp = Mid$(CadenaDatos, 7)
    If (Mid$(DatResp, 151, 1) <> " ") And (Mid$(DatResp, 151, 1) <> "0") Then
      MsgBox "La Cuenta Corriente Asociada al Producto" + Chr(10) + _
             "Otorgamiento de ChileCard EstÃ¡ Cerrada", 0 + 16, "Problemas de ValidaciÃ³n"
    Else
      If Mid$(DatResp, 154, 1) = "4" Then
       MsgBox "La Cuenta Corriente Asociada al Producto Otorgamiento" + Chr(10) + _
               "de ChileCard EstÃ¡ en Proceso de ser Cerrada", 0 + 16, "Problemas de ValidaciÃ³n"
      Else
        ExisteChileCardIBM = True
      End If
    End If
  End If
End If

End Function
