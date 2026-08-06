VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   2490
   ClientTop       =   2475
   ClientWidth     =   4350
   Height          =   2220
   Left            =   2430
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   2130
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario   As String * 8
Dim CadenaDatos As String
Dim BaseVisado  As Database
Dim Data_Vax    As String
Public Function AsignaNodoOficinaRegion() As String

Dim Oficina As String
Dim BaseSGT As Database
Dim p_Org As Recordset

Oficina = SacaValorINI("PRODUCTO", "Oficina", DIRECTORIO_01 + ARCHIVO_INI01)
  
Set BaseSGT = OpenDatabase(DIRECTORIO_SGT + "TABSGT.MDB")
Set p_Org = BaseSGT.OpenRecordset("Select * From TBLORG Where ORG_ORGCUI=" + Oficina, dbOpenDynaset)
If p_Org.RecordCount > 0 Then
     AsignaNodoOficinaRegion = p_Org!Org_OrgScj
Else
     AsignaNodoOficinaRegion = SINNODO
End If
p_Org.Close
BaseSGT.Close
    
End Function
Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim DataCampana    As String
Dim GestionLCD     As Boolean
Dim Operacion      As String
Dim Correlativo    As String
Dim CuentaCorriente As String
Dim LineaDeCredito As String * 11
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
Dim FiltroCliente  As Recordset
Dim ResultadoFinal As Boolean
Dim A              As Boolean
Dim Valido         As String
Dim Existe         As String
Dim Var_Ofi        As String
Dim PunteroAval    As Recordset

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabPrd.Index = "Tab_Prd"
  GestionLCD = False
  ResultadoFinal = False
  IdUsuario = ObtieneRutUsuarioLogueado()
  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
  PunteroTabPrd.Seek "=", Producto
  PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
  Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
  LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
  Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
  Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
  CuentaCorriente = Mid$(Dat, 394, 11)
  
  If Operacion <> "M" Then
    Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + CuentaCorriente, 1)
    Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
    Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
       
    If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
      Var_Ofi = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
      A = PoneValorINI("PRODUCTO", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
      A = PoneValorINI("PRODUCTO", "Oficina", Var_Ofi, DIRECTORIO_01 + ARCHIVO_INI01)
    
      If ExisteLineaDeCredito(LineaDeCredito, "6391", False) Then
        'Reapertura Operacion "R"
        PunteroVsdPrd.Edit
        PunteroVsdPrd("Prd_Data") = Mid$(PunteroVsdPrd("Prd_Data"), 1, 39) + "R" + Mid$(PunteroVsdPrd("Prd_Data"), 41)
        PunteroVsdPrd.Update
        ResultadoFinal = True
      Else
         If ValidarCtaConIntereses(Rut, CuentaCorriente) Then
           ResultadoFinal = True
         Else
            MsgBox "Este producto tiene asociado una Cuenta Corriente Empresa con InterÃƒÆ’Ã‚Â©s." + Chr(10) + _
                  "Por definiciÃƒÆ’Ã‚Â³n no se puede activar LÃƒÆ’Ã‚Â­nea de CrÃƒÆ’Ã‚Â©dito.", 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
         End If
      End If
      
      If ResultadoFinal = True Then
        If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
           MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
               "Error de Escritura"
        End If
      End If
      
      Set PunteroAval = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE Mid$(Prd_Data, 1, 9+3+12)='" + Mid$(PunteroVsdPrd("Prd_Data"), 1, 6) + "1EG" + "0" + LineaDeCredito + "'", dbOpenDynaset)
      If PunteroAval.RecordCount > 0 Then
        Modo = "L"
        If ObtieneInformacion("4400", Mid$(PunteroVsdPrd("Prd_Data"), 10, 12), False, False) Then
          PunteroAval.Edit
          PunteroAval("Prd_Data") = Mid$(PunteroAval("Prd_Data"), 1, 39) + "M" + Mid$(PunteroAval("Prd_Data"), 41)
          PunteroAval.Update
        End If
      End If
      PunteroAval.Close
      
    Else
      MsgBox "No existe la Cuenta Corriente Asociada.", 16, "Errores de ValidaciÃƒÆ’Ã‚Â³n"
    End If
   
  Else 'Modificacion de Cupo
    Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + LineaDeCredito, 1)
    Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
    Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
       
    If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
      Var_Ofi = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
      A = PoneValorINI("PRODUCTO", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
      A = PoneValorINI("PRODUCTO", "Oficina", Var_Ofi, DIRECTORIO_01 + ARCHIVO_INI01)
      
      If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
        Data_Visa = Actualiza(Mid(Dat, 13))
        PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
        PunteroVsdPrd.Edit
        PunteroVsdPrd("Prd_Data") = Reg + Mid(Dat, 1, 12) + Data_Visa
        PunteroVsdPrd.Update
        ResultadoFinal = True
        If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
          MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                "Error de Escritura"
        End If
      End If
            
    Else
      PunteroTabPrd.Close
      PunteroVsdPrd.Close
      BaseVisado.Close
      MsgBox "No se encontrÃƒÆ’Ã‚Â³ la LÃƒÆ’Ã‚Â­nea de CrÃƒÆ’Ã‚Â©dito " + Format$(LineaDeCredito, "&-&&&-&&&&&-&&"), 16, "VISADO - Consulta Producto"
      End
    End If
  End If
     
  ' Estudia parte Cliente
  If ResultadoFinal Then
    Data_Visa = Obtiene5391("0" + LineaDeCredito + Rut)
    Set FiltroCliente = BaseVisado.OpenRecordset("SELECT Prd_Correl FROM Vsd_Prd WHERE " + "Prd_CodProd='3TL'" + " AND " + "Mid$(Prd_Data, 11, 11)='" + LineaDeCredito + "';", dbOpenSnapshot)
    If FiltroCliente.RecordCount > 0 Then
      FiltroCliente.MoveFirst
      Correlativo = Format$(FiltroCliente("Prd_Correl"), "00")
      PunteroVsdPrd.Seek "=", Rut, "3TL", Val(Correlativo)
      PunteroVsdPrd.Edit
      If Data_Visa = "XXX" Then
         PunteroVsdPrd("Prd_Data") = Mid$(PunteroVsdPrd("Prd_Data"), 1, 39) + "O" + Mid$(PunteroVsdPrd("Prd_Data"), 41)
      Else
         PunteroVsdPrd("Prd_Data") = Mid$(PunteroVsdPrd("Prd_Data"), 1, 39) + "M" + Mid$(PunteroVsdPrd("Prd_Data"), 41)
      End If
      PunteroVsdPrd.Update
    Else
      If Data_Visa <> "XXX" Then
         PunteroVsdPrd.AddNew
         PunteroVsdPrd("Prd_RutCli") = Rut
         PunteroVsdPrd("Prd_CodProd") = "3TL"
         PunteroVsdPrd("Prd_Correl") = Val(Correlativo)
         PunteroVsdPrd("Prd_Data") = Mid$(Reg, 1, 6) + "CTL" + Mid$(Data_Visa, 1, 22) + Space(8) + "M" + "  " + Format(Now, "yyyymmdd") + Rut + Mid$(Data_Visa, 23)
         PunteroVsdPrd.Update
      End If
    End If
    FiltroCliente.Close
  End If
  ' Fin Estudio
  PunteroTabPrd.Close
  PunteroVsdPrd.Close
  BaseVisado.Close
End If
End
  
End Sub

Public Function ValidarCtaConIntereses(vRut As String, vCuenta As String) As Boolean
Dim RespCons As Boolean

ValidarCtaConIntereses = False
RespCons = False

If Val(Mid$(vRut, 1, Len(vRut) - 1)) > 50000000 Then
   If Not ExisteCtaCTR(vCuenta, RespCons) Then
     If RespCons = True Then
       ValidarCtaConIntereses = True
     End If
   End If
Else
   ValidarCtaConIntereses = True
End If

End Function

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
  CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
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
  MsgBox "Vista NÃƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
End If

End Function


Function Actualiza(ByVal Data_Visa As String)
   Dim Data_Ret As String
   
   Data_Ret = Mid$(Data_Vax, 1, 28)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 29, 2)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 31, 113)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 144, 46)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 190, 133)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 323)
   
   Actualiza = Data_Ret

End Function

Private Function Es_FCampana(ByVal Rut1 As String, ByVal Cuenta1 As String) As String
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4

  NodoDestino = DeterminaNodo(AplicacionDestino, "2147")
  Modo = "L"
  Llave = Rut1 + "0" + Cuenta1
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + "2147" + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
       Es_FCampana = Mid$(CadenaDatos, 7, 26)
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
         Es_FCampana = "XXX"
      Else
         Es_FCampana = "SRM"
      End If
    End If
  Else
    MsgBox "Vista NÃƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
    Es_FCampana = "SRM"
  End If
End Function

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
Dim VarTmp As String
  Modo = "V"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, "0" + Llave, False, Estado)
  
  If ExisteLineaDeCredito Then
     Data_Vax = Mid$(CadenaDatos, 28)
  End If

End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean, ByVal DebeExistirLinea As Boolean) As Boolean
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  ObtieneInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      If SeEstaBuscandoCuenta Then
        If Mid$(Llave, 3, 8) = Mid$(CadenaDatos, 18, 8) Then
          ObtieneInformacion = True
        Else
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
        End If
      Else
        If Llave = Mid$(CadenaDatos, 16, 12) Then
           ObtieneInformacion = True
        End If
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        If SeEstaBuscandoCuenta Then
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
        Else
          If DebeExistirLinea Then
            MsgBox "No Existe la LÃƒÆ’Ã‚Â­nea de CrÃƒÆ’Ã‚Â©dito", 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
          End If
        End If
      End If
    End If
  Else
    MsgBox "Vista NÃƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function Obtiene5391(ByVal Llave As String) As String
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  NodoDestino = DeterminaNodo(AplicacionDestino, "5391")
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + "5391" + "L"
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
       Obtiene5391 = Mid$(CadenaDatos, 7)
    Else
       Obtiene5391 = "XXX"
    End If
  Else
    MsgBox "Vista NÃƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function

Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "00" + Llave + "00", True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal NumeroVista As String) As String
Dim PunteroVista As Recordset
Dim BaseEnvioSRM As Database
  
     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     ' FI2026 - PunteroVista.Index = PunteroVista.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
     PunteroVista.Seek "=", NumeroVista
     If PunteroVista.NoMatch Then
       DeterminaNodo = SINNODO
     Else
       If Environ$("DESA") = "DESA" Then
         DeterminaNodo = PunteroVista("Vta_NodoDesa")
       Else
         If PunteroVista("Vta_NodoProd") = "U???" Then
           DeterminaNodo = AsignaNodoOficinaRegion ' LCase$(Environ$("SRM_NS"))
         Else
           DeterminaNodo = PunteroVista("Vta_NodoProd")
         End If
       End If
       Aplicacion = PunteroVista("Vta_Aplicacion")
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
