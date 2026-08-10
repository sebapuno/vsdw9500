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
Dim DatoPatron As String
Dim BaseVisado As Database
Dim CadenaDatos As String
Dim Data_IBM    As String
Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Seguir         As Boolean
Dim Operacion      As String
Dim Correlativo    As String
Dim Encabezado     As String
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
Dim NumeroLineaCredito As String

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
    Seguir = True
    Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
    NumeroLineaCredito = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
    If Operacion <> "M" Then 'OTORGAMIENTO
      If Not ProductoAsociadoEncontrado(Encabezado + "L01", Mid$(Rut, 1, 9) + "LCD" + NumeroLineaCredito, "10", False, False) Then
        Seguir = False
        MsgBox "Falta Registro de Producto" + Chr(10) + _
               "Deben Ser Enviados Nuevamente", 0 + 16, "Problemas de Validación"
      End If
      If Seguir Then
        If ProductoAsociadoEncontrado(Encabezado + "AL0", NumeroLineaCredito, "10", False, True) Then
          If ProductoAsociadoEncontrado(Encabezado + "LRC", NumeroLineaCredito, "10", True, True) Then
            If ProductoAsociadoEncontrado(Encabezado + "CC2", DatoPatron, "10", False, True) Then
              If ProductoAsociadoEncontrado(Encabezado + "D01", DatoPatron, "22", False, True) Then
                If ProductoAsociadoEncontrado(Encabezado + "D02", DatoPatron, "22", False, True) Then
                  PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
                  If Not ExisteLineaDeCredito(NumeroLineaCredito, PunteroTabPrd("Prd_Vista_Dst")) Then
                    If ExisteCuentaCorriente(Mid(NumeroLineaCredito, 2, 8), "0391") Then
                      If Not ExisteProductoCC2(DatoPatron, "0391") Then
                        If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
                          MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                                 "Error de Escritura"
                        End If
                      Else
                         MsgBox "Ya Existe la Línea de Crédito Activa", 0 + 16, "Problemas de Validación"
                      End If
                    Else
                      MsgBox "No Existe la Cuenta Corriente", 0 + 16, "Problemas de Validación"
                    End If
                  Else
                    MsgBox "Ya Existe la Línea de Crédito Activa", 0 + 16, "Problemas de Validación"
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Else ' MODIFICACION
      If ExisteLineaDeCredito(NumeroLineaCredito, PunteroTabPrd("Prd_Vista_Dst")) Then
         Data_IBM = Mid$(CadenaDatos, 18)
         Data_Visa = ActualizaMetro(Dat)
         PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
         PunteroVsdPrd.Edit
         PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
         PunteroVsdPrd.Update
         If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
          MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                 "Error de Escritura"
        End If
      Else
         MsgBox "No Existe la Línea de Crédito Activa", 0 + 16, "Problemas de Validación"
      End If
    End If
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V0000"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "0" + Llave + "00", True)
End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean) As Boolean
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  ObtieneInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      ObtieneInformacion = True
      If SeEstaBuscandoCuenta Then
        If Mid$(Llave, 1, 9) <> Mid$(CadenaDatos, 16, 9) Then
          ObtieneInformacion = False
        End If
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset
Dim BaseEnvioSRM As Database
  
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
           DeterminaNodo = LCase$(Environ$("SRM_NS"))
         Else
           DeterminaNodo = PunteroVista("Vta_NodoProd")
         End If
       End If
       
       ' Analiza si el producto ya es 2000
       'Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
       'Set PunteroPrd = BaseVisado.OpenRecordset("TAB_PRD", dbOpenTable)
       ' FI2026 - 'PunteroPrd.Index = 'PunteroPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
       'PunteroPrd.Seek "=", Producto
       'If PunteroPrd("Prd_2000") And PunteroVista("Vta_Aplicacion") = "SPLA" Then
       '   Aplicacion = "SPL2"
       'Else
          Aplicacion = PunteroVista("Vta_Aplicacion")
       'End If
       'PunteroPrd.Close
       'BaseVisado.Close
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function
Function ActualizaMetro(ByVal Data_Visa As String)
   Dim Data_Ret As String
   Dim Fech_Tmp As String * 8
   
   Data_Ret = Mid$(Data_IBM, 1, 49)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 50, 14)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 64, 14)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 78, 14)
   Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   Data_Ret = Data_Ret + Mid$(Data_IBM, 100, 91)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 191, 8)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 199, 13)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 212, 1)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 213, 3)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 216, 15)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 231, 1)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 232, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 236, 6)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 242, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 246, 8)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 254, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 258, 15)
   
   'Se guarda la fecha del campo LCD_FEC_ULT_ACTUALIZ.
   'la idea es guardar dicha fecha en el campo LCD_FEC_PENUL_ACTUAL
   Fech_Tmp = Mid$(Data_IBM, 273, 8)

   Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")

   'Se guarda la fecha proveniente del campo LCD_FEC_ULT_ACTUALIZ
   'en el campo LCD_FEC_PENUL_ACTUAL

   Data_Ret = Data_Ret + Fech_Tmp
   Data_Ret = Data_Ret + Mid$(Data_IBM, 289, 21)

   ActualizaMetro = Data_Ret

End Function

Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, Llave, False)
End Function
Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal CuentaBuscada As String, ByVal Posicion As String, ByVal CambiaDatoPatron As Boolean, ByVal VerMensaje As String) As Boolean
Dim Largo As String
Dim FiltroVsdPrd  As Recordset
  Largo = LTrim$(Str(Len(CuentaBuscada)))
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                                                                 "Mid$(Prd_Data, 1, 9)='" + PatronBusqueda + "' AND " + _
                                                                                 "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + CuentaBuscada + "';", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    If CambiaDatoPatron Then
      DatoPatron = Mid$(FiltroVsdPrd("Prd_Data"), 61, 11)
    End If
    ProductoAsociadoEncontrado = True
  Else
    If VerMensaje Then
      MsgBox "Errores en Integridad de los Datos" + Chr(10) + _
                    "Deben Ser Enviados Nuevamente", 0 + 16, "Problemas de Validación"
    End If
  End If
  FiltroVsdPrd.Close
End Function
Private Function ExisteProductoCC2(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteProductoCC2 = ObtieneInformacion(Vista, Llave, False)
End Function
