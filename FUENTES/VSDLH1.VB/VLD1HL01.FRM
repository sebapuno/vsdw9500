VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1035
   ClientTop       =   990
   ClientWidth     =   4350
   Height          =   2220
   Left            =   975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   645
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario   As String * 8
Dim CadenaDatos As String
Dim Data_Vax    As String
Dim BaseVisado  As Database

Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim DataCampana    As String
Dim Operacion      As String
Dim Correlativo    As String
Dim LineaDeCredito As String * 11
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
Dim Var_Ofi        As String

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroTabPrd.Index = "Tab_Prd"
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
    
    Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + LineaDeCredito, 1)
    
    Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
    Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
    If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
       Var_Ofi = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
       A = PoneValorINI("PRODUCTO", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
       A = PoneValorINI("PRODUCTO", "Oficina", Var_Ofi, DIRECTORIO_01 + ARCHIVO_INI01)
    End If
        
    If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
       Data_Visa = Actualiza(Dat)
       PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
       PunteroVsdPrd.Edit
       PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
       PunteroVsdPrd.Update
    
       If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
          MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                 "Error de Escritura"
       End If
    End If
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
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
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
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
  
  Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + Format$(Val(Mid$(Llave, 2)), "00000000000"), 1)

  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      If SeEstaBuscandoCuenta Then
        If Mid$(Llave, 3, 8) = Mid$(CadenaDatos, 18, 8) Then
          ObtieneInformacion = True
        Else
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de Validación"
        End If
      Else
        If DebeExistirLinea Then
          If Mid$(Llave, 1, 10) = Mid$(CadenaDatos, 16, 10) Then
            If Val(Mid$(CadenaDatos, 149, 6)) <> 0 Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              If Mid$(CadenaDatos, 291, 1) = "4" Then
                MsgBox "La Línea de Crédito Está" + Chr(10) + "en Proceso de ser Cerrada", 0 + 16, _
                              "Problemas de Validación"
              Else
                ObtieneInformacion = True
              End If
            End If
          Else
            MsgBox "No Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        Else
          If Llave = Mid$(CadenaDatos, 16, 12) Then
            ObtieneInformacion = True
            'MsgBox "Ya Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        End If
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        If SeEstaBuscandoCuenta Then
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de Validación"
        Else
          If DebeExistirLinea Then
            MsgBox "No Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        End If
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "00" + Llave + "00", True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset

     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     ' FI2026 - PunteroVista.Index = PunteroVista.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
     PunteroVista.Seek "=", Vista
     If PunteroVista.NoMatch Then
       DeterminaNodo = SINNODO
     Else
       If Environ$("DESA") = "DESA" Then
         DeterminaNodo = PunteroVista("Vta_NodoDesa")
       Else
         If PunteroVista("Vta_NodoProd") = "U???" Then
           '** busca oficina
           DeterminaNodo = AsignaNodoOficinaRegion 'LCase$(Environ$("SRM_NS"))
         Else
           DeterminaNodo = PunteroVista("Vta_NodoProd")
         End If
       End If
       
       Aplicacion = PunteroVista("Vta_Aplicacion")
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function
Public Function AsignaNodoOficinaRegion() As String

Dim Valido As String
Dim Existe As String
Dim Oficina As String
Dim BaseSGT As Database
Dim p_Org As Recordset

Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
      
If UCase(Valido) = "TRUE" And UCase(Existe) = "TRUE" Then
  Oficina = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
  
  Set BaseSGT = OpenDatabase(DIRECTORIO_SGT + "TABSGT.MDB")
  Set p_Org = BaseSGT.OpenRecordset("Select * From TBLORG Where ORG_ORGCUI=" + Oficina, dbOpenDynaset)
  If p_Org.RecordCount > 0 Then
     AsignaNodoOficinaRegion = p_Org!Org_OrgScj
  Else
     AsignaNodoOficinaRegion = SINNODO
  End If
  p_Org.Close
  BaseSGT.Close
    
Else
  AsignaNodoOficinaRegion = SINNODO
End If

End Function

Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
