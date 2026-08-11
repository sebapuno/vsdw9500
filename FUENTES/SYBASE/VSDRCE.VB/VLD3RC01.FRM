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

Dim BaseVisado As Database

Private Sub Form_Load()
Dim Correlativo As String
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
Dim CuentaCorriente As String * 12
Dim FiltroCliente  As Recordset
Dim ResultadoFinal As Boolean
Dim Encabezado As String
Dim FiltroVsdPrd As Recordset

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
    ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
    ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
    PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
    PunteroTabPrd.Index = "Tab_Prd"
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    
    ResultadoFinal = False
    CuentaCorriente = Mid$(PunteroVsdPrd("Prd_Data"), 10, 12)
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
    If ProductoAsociadoEncontrado(Encabezado + "2FD", Mid(Rut, 1, 9) + "CTD" + Mid$(CuentaCorriente, 2, 11), "10") Then
      If ProductoAsociadoEncontrado(Encabezado + "2FP", Mid(Rut, 1, 9) + "CTD" + Mid$(CuentaCorriente, 2, 11), "10") Then
        If ProductoAsociadoEncontrado(Encabezado + "2FR", Mid$(CuentaCorriente, 2, 11), "10") Then
          If ProductoAsociadoEncontrado(Encabezado + "3FT", "0" + Mid$(CuentaCorriente, 2, 11) + Rut, "10") Then
             If Not ExisteCuentaCorriente(CuentaCorriente, PunteroTabPrd("Prd_Vista_Dst")) Then
                   '** Si Tiene Producto Asociado "ACR"
                   Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE prd_rutcli='" + Rut + "' AND " + _
                                      "Mid$(prd_data, 1, 31 )='" + Encabezado + "ACR" + Left(Mid(CuentaCorriente, 2) + Space(21), 21) + "' ;", dbOpenSnapshot)
                   If FiltroVsdPrd.RecordCount > 0 Then
                       Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE prd_rutcli='" + Rut + "' AND " + _
                                      "Mid$(prd_data, 1, 31 )='" + Encabezado + "DR3" + Left(Mid(CuentaCorriente, 2) + Space(21), 21) + "' ;", dbOpenSnapshot)
                       If FiltroVsdPrd.RecordCount > 0 Then
                           If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
                               MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                                      "Error de Escritura"
                           End If
                           ResultadoFinal = True
                       Else
                           MsgBox "No se encuentra registro relación cliente producto " + Chr(10) + "para Cuenta Corriente Remunerada.", 16, "Problemas de Validación"
                       End If
                   Else
                     MsgBox "No se encuentra registro de atributos " + Chr(10) + "para Cuenta Corriente Remunerada.", 16, "Problemas de Validación"
                   End If
                   FiltroVsdPrd.Close
                   '****
              End If
          End If
        End If
      End If
    End If
    
    ' Estudia parte Cliente
    If ResultadoFinal Then
      Data_Visa = Obtiene5391(CuentaCorriente + Rut)
      Set FiltroCliente = BaseVisado.OpenRecordset("SELECT Prd_Correl FROM Vsd_Prd WHERE " + "Prd_CodProd='3FT'" + " AND " + "Mid$(Prd_Data, 10, 12)='" + CuentaCorriente + "';", dbOpenSnapshot)
      If FiltroCliente.RecordCount > 0 Then
         FiltroCliente.MoveFirst
         Correlativo = Format$(FiltroCliente("Prd_Correl"), "00")
         PunteroVsdPrd.Seek "=", Rut, "3FT", Val(Correlativo)
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
            PunteroVsdPrd("Prd_CodProd") = "3FT"
            PunteroVsdPrd("Prd_Correl") = Val(Correlativo)
            PunteroVsdPrd("Prd_Data") = Mid$(Reg, 1, 6) + "3FT" + Mid$(Data_Visa, 1, 22) + Space(8) + "M" + "  " + Format(Now, "yyyymmdd") + Rut + Mid$(Data_Visa, 23)
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
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
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
          ObtieneInformacion = True
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
     ' FI2026 - era: PunteroVista.Index = PunteroVista.Indexes(0).Name
     ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
     ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
     ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
     PunteroVista.Index = BaseEnvioSRM.TableDefs("Vistas").Indexes(0).Name
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
       ' FI2026 - 'PunteroPrd.Index = 'PunteroPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
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

Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal CuentaBuscada As String, ByVal Posicion As String) As Boolean
Dim Largo As String
Dim FiltroVsdPrd  As Recordset
  Largo = LTrim$(Str(Len(CuentaBuscada)))
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                                                                 "Mid$(Prd_Data, 1, 9)='" + PatronBusqueda + "' AND " + _
                                                                                 "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                                                 CuentaBuscada + "';", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    ProductoAsociadoEncontrado = True
  Else
    MsgBox "Errores en Integridad de los Datos" + Chr(10) + _
                  "Deben Ser Enviados Nuevamente", 0 + 16, "Problemas de Validación"
  End If
  FiltroVsdPrd.Close
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
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function


Private Function ExisteClasificacionDeRiesgo(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteClasificacionDeRiesgo = ObtieneInformacion(Vista, Llave, False)
End Function
