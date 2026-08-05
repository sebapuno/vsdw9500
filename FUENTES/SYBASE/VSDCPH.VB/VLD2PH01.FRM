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

Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal CuentaBuscada As String, ByVal Posicion As String) As Boolean
Dim Largo         As String
Dim FiltroVsdPrd  As Snapshot
Dim SQL           As String

  Largo = LTrim$(Str(Len(CuentaBuscada)))
  ProductoAsociadoEncontrado = False
  Set BaseVisado = Workspaces(0).OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  SQL = "SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                               "Mid$(Prd_Data, 1, 9)='" + PatronBusqueda + "' AND " + _
                                               "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                               CuentaBuscada + "';"
                                               
  Set FiltroVsdPrd = BaseVisado.CreateSnapshot(SQL)
  If FiltroVsdPrd.RecordCount > 0 Then
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
  
End Function
Private Sub Form_Load()
Dim Correlativo As String
Dim BaseVisado As Database
Dim DatoProducto As String * 11
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
Dim FiltroVsdPrd As Snapshot

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
  DatoProducto = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
  
  CuentaCorriente = DatoProducto
  
  If Not ExisteCuentaCorriente(DatoProducto, PunteroTabPrd("Prd_Vista_Dst")) Then
      If ExisteClasificacionDeRiesgo(Mid(Rut, 1, 9), "1348") Then
         '** Si Tiene Producto Asociado "LGL"
         Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Prd WHERE prd_rutcli='" + Rut + "' AND " + _
                              "Mid$(prd_data, 1, 39 )='" + Mid$(PunteroVsdPrd("Prd_Data"), 1, 6) + "LGLCTD" + Mid$(PunteroVsdPrd("Prd_Data"), 10, 27) + "' ;")
         If FiltroVsdPrd.RecordCount > 0 Then
           If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
              MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                       "Error de Escritura"
           End If
         Else
           MsgBox "No se encuentra dirección legal para Cuenta Corriente.", 16, "Faltan Datos"
         End If
         FiltroVsdPrd.Close
      End If
  End If
  PunteroTabPrd.Close
  PunteroVsdPrd.Close
  BaseVisado.Close
End If
End

End Sub
Private Sub BuscarFirma(ByVal Llave As String, ByVal Vista As String)
  Modo = "V0000"
  If ObtieneInformacion(Vista, Llave, False) Then
    If Mid$(CadenaDatos, 16, 9) <> Mid$(Llave, 1, 9) Or Mid$(CadenaDatos, 36, 1) <> "S" Then
      MsgBox "Debe Digitalizar y/o Visar la Firma del Cliente", 0 + 64, "ADVERTENCIA"
    End If
  Else
    If Mid$(CadenaDatos, 1, 6) = "000096" Then
      MsgBox "Debe Digitalizar y/o Visar la Firma del Cliente", 0 + 64, "ADVERTENCIA"
    End If
  End If
End Sub
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
Private Function ExisteClasificacionDeRiesgo(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  
  If INI_REMOTO() = "SI" Then
     ExisteClasificacionDeRiesgo = ObtieneInformacion(Vista, Llave, False)
  Else
     ExisteClasificacionDeRiesgo = True
  End If
  
End Function
