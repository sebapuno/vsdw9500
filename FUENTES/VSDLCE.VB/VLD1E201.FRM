VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   2370
   ClientTop       =   3795
   ClientWidth     =   4350
   Height          =   2220
   Left            =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   3450
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario   As String * 8
Dim CadenaDatos As String
Dim Data_Vax    As String
Dim BaseVisado     As Database

Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim DataCampana    As String
Dim Operacion      As String
Dim Correlativo    As String
Dim Encabezado     As String
Dim LineaDeCredito As String * 11
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
Dim FiltroCliente  As Recordset
Dim ResultadoFinal As Boolean

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    PunteroTabPrd.Index = "Tab_Prd"
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
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
      
    If ProductoAsociadoEncontrado(Encabezado + "AE0", LineaDeCredito, "10") Then
       If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
          Data_Visa = Actualiza(Dat, False)
          PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
          PunteroVsdPrd.Edit
          PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
          PunteroVsdPrd.Update
      
          ResultadoFinal = True
          If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
             MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
          End If
       End If
    End If
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
Function Actualiza(ByVal Data_Visa As String, ByVal GestionLCD As Boolean)
   Dim Data_Ret As String
   
   Data_Ret = Mid$(Data_Vax, 1, 28)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 29, 2)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 31, 113)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 144, 46)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 190, 133)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 323, 19)
   
   Actualiza = Data_Ret

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
            If Val(Mid$(CadenaDatos, 275, 1)) = 1 Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              ObtieneInformacion = True
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

Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "00" + Llave + "00", True, False)
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
