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
Dim BaseVisado As Database
Dim CadenaDatos As String
Private Sub Form_Load()
Dim Correlativo As String
Dim Encabezado As String
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
Dim NumeroCuentaCorriente As String
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    PunteroTabPrd.Index = "Tab_Prd"
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
    NumeroCuentaCorriente = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
    If ProductoAsociadoEncontrado(Encabezado + "D01", Mid(Rut, 1, 9) + "CTD" + NumeroCuentaCorriente, "10") Then
      If ProductoAsociadoEncontrado(Encabezado + "D02", Mid(Rut, 1, 9) + "CTD" + NumeroCuentaCorriente, "10") Then
        If ProductoAsociadoEncontrado(Encabezado + "RE1", NumeroCuentaCorriente, "10") Then
          If Not ExisteCuentaCorriente(NumeroCuentaCorriente, PunteroTabPrd("Prd_Vista_Dst")) Then
            If ExisteClasificacionDeRiesgo(Mid(Rut, 1, 9), "1348") Then
              If Not EsCuentaDOS(Mid(PunteroVsdPrd("Prd_Data"), 42, 1)) Then
                If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
                  MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                                "Error de Escritura"
                End If
              End If
            End If
          End If
        End If
      End If
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
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String)
  Modo = "L"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, Llave, True)
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
      If SeEstaBuscandoCuenta Then
        ObtieneInformacion = True
        MsgBox "Ya Existe la Cuenta Corriente", 0 + 16, "Problemas de Validación"
      Else
        If Mid$(CadenaDatos, 132, 2) = "SC" Then
          MsgBox "No Existe la Clasificación de Riesgo", 0 + 16, "Problemas de Validación"
        Else
          If Mid$(CadenaDatos, 184, 1) = "N" Then
            MsgBox "Falta Visar la Dirección Legal del Cliente", 0 + 16, "Problemas de Validación"
          Else
            ObtieneInformacion = True
          End If
        End If
      End If
    Else
      If SeEstaBuscandoCuenta Then
        If Mid$(CadenaDatos, 1, 6) <> "000096" Then
          ObtieneInformacion = True
        End If
      Else
          If Mid$(CadenaDatos, 1, 6) = "000096" Then
            MsgBox "El Cliente No Está Creado" + Chr(10) + "Consulte con el Ejecutivo", 0 + 16, "Problemas de Validación"
          End If
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal NumeroVista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset
Dim BaseEnvioSRM As Database
  
     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     PunteroVista.Index = PunteroVista.Indexes(0).Name
     PunteroVista.Seek "=", NumeroVista
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
       'Set PunteroPrd = BaseVisado.OpenRecordset("TAB_PRD", dbOpenTable)
       'PunteroPrd.Index = PunteroPrd.Indexes(0).Name
       'PunteroPrd.Seek "=", Producto
       'If PunteroPrd("Prd_2000") And PunteroVista("Vta_Aplicacion") = "SPLA" Then
       '   Aplicacion = "SPL2"
       'Else
          Aplicacion = PunteroVista("Vta_Aplicacion")
       'End If
       'PunteroPrd.Close
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function

Private Function ObtieneRutUsuarioLogueado() As String
Dim Rut As String * 70
Dim Retorno As Integer
  Retorno = AisGetUsr(Rut)
  If Retorno = 1 Then
    ObtieneRutUsuarioLogueado = Format$(Mid(Rut, 1, 8), "00000000")
  Else
    ObtieneRutUsuarioLogueado = RUTDEFECTO
  End If
End Function
Private Function ExisteClasificacionDeRiesgo(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteClasificacionDeRiesgo = ObtieneInformacion(Vista, Llave, False)
End Function

Private Function EsCuentaDOS(ByVal FlagDOS As String) As Boolean
  EsCuentaDOS = False
  If FlagDOS = "R" Then
    EsCuentaDOS = True
    MsgBox "La Cuenta es del Tipo DOS", 0 + 16, "Problemas de Validación"
  End If
End Function
Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal CuentaBuscada As String, ByVal Posicion As String) As Boolean
Dim Largo As String
Dim FiltroVsdPrd  As Snapshot
  Largo = LTrim$(Str(Len(CuentaBuscada)))
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                                                                 "Mid$(Prd_Data, 1, 9)='" + PatronBusqueda + "' AND " + _
                                                                                 "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                                                 CuentaBuscada + "';")
  If FiltroVsdPrd.RecordCount > 0 Then
    ProductoAsociadoEncontrado = True
  Else
    MsgBox "Errores en Integridad de los Datos" + Chr(10) + _
                  "Deben Ser Enviados Nuevamente", 0 + 16, "Problemas de Validación"
  End If
  FiltroVsdPrd.Close
End Function
