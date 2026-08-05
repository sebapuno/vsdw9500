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
Dim BaseVisado  As Database
Dim CadenaDatos As String
Dim Data_IBM    As String

Function Actualiza(ByVal Data_Visa As String)
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

   If Mid$(Data_IBM, 229, 8) = Mid$(Data_Visa, 229, 8) Then
      ' MODIFICACION
      Data_Ret = Data_Ret + Mid$(Data_IBM, 273, 8)
   Else
      ' RENOVACION
      Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   End If
   
   'Se guarda la fecha del campo LCD_FEC_ULT_ACTUALIZ.
   'la idea es guardar dicha fecha en el campo LCD_FEC_PENUL_ACTUAL
   Fech_Tmp = Mid$(Data_IBM, 289, 8)

   Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")

   'Se guarda la fecha proveniente del campo LCD_FEC_ULT_ACTUALIZ
   'en el campo LCD_FEC_PENUL_ACTUAL
   Data_Ret = Data_Ret + Fech_Tmp
   
   Data_Ret = Data_Ret + Mid$(Data_IBM, 297, 21)

   Actualiza = Data_Ret
End Function

Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Operacion      As String
Dim Correlativo    As String
Dim Encabezado     As String
Dim Data_Visa      As String
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
Dim NumeroLineaCredito As String

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
    Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
    NumeroLineaCredito = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
     If ProductoAsociadoEncontrado(Encabezado + "AE0", NumeroLineaCredito, "10") Then
       If ExisteLineaDeCredito(NumeroLineaCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
          Data_Visa = Actualiza(Dat)
          PunteroVsdPrd.Edit
          PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
          PunteroVsdPrd.Update
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
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
  Modo = "V0000"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, Llave, False, Estado, False)
  If ExisteLineaDeCredito Then
     Data_IBM = Mid$(CadenaDatos, 15 + 12)
  End If
End Function
Private Function ClienteEstaActivado(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ClienteEstaActivado = ObtieneInformacion(Vista, Llave, False, False, True)
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
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean, ByVal DebeExistirLinea As Boolean, ByVal SeEstaBuscandoAlCliente As Boolean) As Boolean
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
        If Mid$(Llave, 1, 9) = Mid$(CadenaDatos, 16, 9) Then
          ObtieneInformacion = True
        Else
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de Validación"
        End If
      Else
        If DebeExistirLinea Then
          If Llave = Mid$(CadenaDatos, 15 + 1, 11) Then
            If Mid$(CadenaDatos, 15 + 176, 1) <> " " And Mid$(CadenaDatos, 15 + 176, 1) <> "0" Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              If Mid$(CadenaDatos, 15 + 179, 1) = "4" Then
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
          If SeEstaBuscandoAlCliente Then
            ObtieneInformacion = True
          Else
            If Llave = Mid$(CadenaDatos, 16, 11) Then
              ObtieneInformacion = True
              MsgBox "Ya Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
            End If
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
          Else
            MsgBox "El Cliente No Está Creado" + Chr(10) + "Consulte con el Ejecutivo", 0 + 16, "Problemas de Validación"
          End If
        End If
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V0000"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "0" + Llave + "00", True, False, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset
Dim BaseEnvioSRM As Database
  
     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     PunteroVista.Index = PunteroVista.Indexes(0).Name
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
