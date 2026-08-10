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
Dim IdUsuario As String * 8
Dim CadenaDatos As String
Dim Data_IBM    As String

Dim BaseVisado   As Database
Dim BaseEnvioSRM As Database

Function Actualiza(ByVal Data_Visa As String, ByVal GestionLCD As Boolean)
   Dim Data_Ret As String
   Dim Fech_Tmp As String * 8
   
   Data_Ret = Mid$(Data_IBM, 1, 51)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 52, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 66, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 80, 14)

   'la idea es que el campo LCD_FEC_RENOVACION no sea modificado
   If Mid$(Data_IBM, 229, 8) = Mid$(Data_Visa, 229, 8) Then
      Data_Ret = Data_Ret + Mid$(Data_IBM, 94, 8)
   Else
      Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   End If

   Data_Ret = Data_Ret + Mid$(Data_IBM, 102, 97)

   'la idea es que el campo LCD_FEC_APERTURA no sea modificado
   Data_Ret = Data_Ret + Mid$(Data_IBM, 199, 8)

   Data_Ret = Data_Ret + Mid$(Data_IBM, 207, 13)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 220, 1)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 221, 3)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 224, 15)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 239, 1)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 240, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 244, 6)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 250, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 254, 8)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 262, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 266, 7)

   If Mid$(Data_IBM, 229, 8) = Mid$(Data_Visa, 229, 8) Or GestionLCD Then
      ' MODIFICACION
      Data_Ret = Data_Ret + Mid$(Data_IBM, 273, 8)
   Else
      ' RENOVACION
      Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   End If
   
   'Se guarda la fecha del campo LCD_FEC_ULT_ACTUALIZ.
   'la idea es guardar dicha fecha en el campo LCD_FEC_PENUL_ACTUAL
   Fech_Tmp = Mid$(Data_IBM, 281, 8)

   Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")

   'Se guarda la fecha proveniente del campo LCD_FEC_ULT_ACTUALIZ
   'en el campo LCD_FEC_PENUL_ACTUAL
   Data_Ret = Data_Ret + Fech_Tmp
   
   If Mid$(Data_IBM, 229, 8) <> Mid$(Data_Visa, 229, 8) And GestionLCD Then
      Data_Ret = Data_Ret + Format$(Date, "yyyy")
      Data_Ret = Data_Ret + Mid$(Data_IBM, 301, 17)
   Else
      Data_Ret = Data_Ret + Mid$(Data_IBM, 297, 21)
   End If

   Actualiza = Data_Ret

End Function


Private Sub Form_Load()
Dim Operacion      As String
Dim Correlativo    As String
Dim LineaDeCredito As String
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
Dim A              As Boolean
Dim Valido         As String
Dim Existe         As String
Dim Var_Ofi        As String
Dim Reg            As String
Dim Data_Visa       As String

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
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Data_Visa = Mid$(PunteroVsdPrd("Prd_Data"), 61)
    LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
    
    Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + LineaDeCredito, 1)
    Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
    Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
       
    If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
       Var_Ofi = SacaValorINI("Datos", "Oficina", DIRECTORIO_03 + ARCHIVO_INI03)
       A = PoneValorINI("PRODUCTO", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
       A = PoneValorINI("PRODUCTO", "Oficina", Var_Ofi, DIRECTORIO_01 + ARCHIVO_INI01)
       
       'Data_Visa = Actualiza_Fechas(Data_Visa)
               
       'PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
       'PunteroVsdPrd.Edit
       'PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
       'PunteroVsdPrd.Update
               
       If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
           MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                  "Error de Escritura"
       End If
    Else
       MsgBox "No se encontró la Línea de Crédito " + Format$(LineaDeCredito, "&-&&&-&&&&&-&&"), 0 + 16, _
                  "VISADO - Módulo Identificador de Producto"
    End If
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
  
End Sub
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
Dim DataTmp As String

  Modo = "V0000"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, Llave, False, Estado)
  If ExisteLineaDeCredito Then
     Data_IBM = Mid$(CadenaDatos, 27)
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
        ObtieneInformacion = True
      Else
        If DebeExistirLinea Then
          If Llave = Mid$(CadenaDatos, 15 + 1, 11) Then
            ' En año 2000 SD20 cambiar 170 a 173
            If Mid$(CadenaDatos, 15 + 176, 1) <> " " And Mid$(CadenaDatos, 15 + 176, 1) <> "0" Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              ' En año 2000 SD20 cambiar 173 a 176
              If Mid$(CadenaDatos, 15 + 180, 1) = "4" Then
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
          If Llave = Mid$(CadenaDatos, 16, 11) Then
            ObtieneInformacion = True
            MsgBox "Ya Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
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
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "0" + Mid$(Llave, 2, 8) + Format((Val(Mid$(Llave, 11, 2)) + 10) Mod 11, "00"), True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset

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
           '*** Obtener Nodo Según llave producto
            DeterminaNodo = AsignaNodoOficinaRegion
           '***
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
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
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


Public Function Actualiza_Fechas(Texto As String) As String

Dim Datos As String
Dim FchTemp  As String

Datos = Space(11) + Texto

FchTemp = Mid$(Datos, 290, 8)

'* Fecha de ultma actualiz
Datos = Mid$(Datos, 1, 289) + Format$(Date, "ddmmyyyy") + FchTemp + Mid$(Datos, 306)

'* Fecha de ult actualiz homolog
Datos = Mid$(Datos, 1, 360) + Format$(Date, "ddmmyyyy") + Mid$(Datos, 369)

Actualiza_Fechas = Mid$(Datos, 12)

End Function
