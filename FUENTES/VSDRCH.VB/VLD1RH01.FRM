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

Private Sub Form_Load()
Dim Cuenta As String * 11
Dim Correlativo As String
Dim BaseVisado As Database
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
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
    Cuenta = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
    If ExisteCuentaCorriente(Cuenta, PunteroTabPrd("Prd_Vista_Dst")) Then
      If Not ExisteLineaCredito("01" + Mid$(Cuenta, 3, 8) + Format((Val(Mid$(Cuenta, 11, 2)) + 1) Mod 11, "00"), "6391") Then
        If Not ExisteLineaCreditoActiva("3" + Mid$(Cuenta, 3, 8) + Format((Val(Mid$(Cuenta, 11, 2)) + 3) Mod 11, "00"), "3031") Then
           '** Si Tiene Producto Asociado "1CA"
           Dataini = Mid$(PunteroVsdPrd("Prd_Data"), 1, 40)
           Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE prd_rutcli='" + Rut + "' AND " + _
                             "Mid$(prd_data, 1, 40 )='" + Mid$(Dataini, 1, 6) + "CCA" + Mid$(Dataini, 10, 31) + "' ;", dbOpenSnapshot)
           If FiltroVsdPrd.RecordCount > 0 Then
              If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
              MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                          "Error de Escritura"
              End If
           Else
              MsgBox "No se encuentra registro de atributos " + Chr(10) + "para Cierre de Cuenta Corriente Remunerada.", 16, "Faltan Datos"
           End If
           FiltroVsdPrd.Close
           '****
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
Private Function ExisteLineaCreditoActiva(ByVal Llave As String, ByVal Vista As String)
  Modo = "L"
  ExisteLineaCreditoActiva = ObtieneInformacion(Vista, Llave, False, "Activa")
End Function
Private Function ExisteLineaCredito(ByVal Llave As String, ByVal Vista As String)
  Modo = "L"
  ExisteLineaCredito = ObtieneInformacion(Vista, Llave, False, "")
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String)
  Modo = "L"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, Llave, True, "")
End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean, ByVal TipoLinea As String) As Boolean
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  ObtieneInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      CadenaDatos = Mid$(CadenaDatos, 7)
      If SeEstaBuscandoCuenta Then
        ObtieneInformacion = True
      Else
        If (Mid$(CadenaDatos, 176, 1) = " " Or Mid$(CadenaDatos, 176, 1) = "0") And Mid$(CadenaDatos, 179, 1) <> "4" Then
          ObtieneInformacion = True
          MsgBox "La Línea de Crédito " + TipoLinea + " " + Format$(Llave, "0-000-00000-00") + Chr(13) + "No ha Sido Cerrada", 0 + 16, _
                 "Problemas de Validación"
        End If
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        If SeEstaBuscandoCuenta Then
          MsgBox "No Existe la Cuenta Corriente", 0 + 16, "Problemas de Validación"
        End If
      Else
        If Not SeEstaBuscandoCuenta Then
          ObtieneInformacion = True
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
Dim BaseVisado   As Database
  
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

Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
