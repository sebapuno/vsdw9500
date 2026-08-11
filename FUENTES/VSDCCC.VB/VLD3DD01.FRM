VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   2490
   ClientTop       =   2475
   ClientWidth     =   4350
   Height          =   2160
   Left            =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   2160
   Visible         =   0   'False
   Width           =   4410
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String * 8
Dim CadenaDatos As String

Private Sub Form_Load()
Dim Cuenta As String
Dim Correlativo As String
Dim Respuesta As Integer
Dim BaseVisado As Database
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
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
    Cuenta = Mid$(PunteroVsdPrd("Prd_Data"), 10, 12)
    If ExisteCuentaCorriente(Cuenta, PunteroTabPrd("Prd_Vista_Dst")) Then
      If Not ExisteLineaCredito("01" + Mid$(Cuenta, 3, 8) + Format((Val(Mid$(Cuenta, 11, 2)) + 1) Mod 11, "00"), "4391") Then
        If Not ExisteLineaCreditoActiva("3" + Mid$(Cuenta, 3, 8) + Format((Val(Mid$(Cuenta, 11, 2)) + 3) Mod 11, "00"), "0392") Then
          If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
             MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
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
Private Function ExisteLineaCreditoActiva(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteLineaCreditoActiva = ObtieneInformacion(Vista, Llave, False, 9)
End Function
Private Function ExisteLineaCredito(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteLineaCredito = ObtieneInformacion(Vista, Llave, False, 10)
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, Llave, True, 12)
End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean, ByVal FinCorte As Integer) As Boolean
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
        If Val(Mid(CadenaDatos, 149, 6)) = 0 Then
          If Mid(CadenaDatos, 291, 1) = 0 Then
            ObtieneInformacion = True
          Else
            MsgBox "La Cuenta Está en Proceso de Cierre", 0 + 16, "Problemas de Validación"
          End If
        Else
          MsgBox "La Cuenta Corriente Está Cerrada", 0 + 16, "Problemas de Validación"
        End If
      Else
        If Vista = "0392" Then
          If Mid$(CadenaDatos, 185, 1) = "0" Then
            ObtieneInformacion = True
            MsgBox "La Línea de Crédito Activa" + Chr(10) + "Aún No ha Sido Cerrada", 0 + 16, _
                   "Problemas de Validación"
          Else
            If Mid$(CadenaDatos, 188, 1) = "4" Then
              ObtieneInformacion = True
              MsgBox "La Línea de Crédito Activa Está" + Chr(10) + "en Proceso de ser Cerrada", 0 + 16, _
                     "Problemas de Validación"
            End If
          End If
        Else
          If Val(Mid$(CadenaDatos, 140, 6)) = 0 Or Mid$(CadenaDatos, 146, 1) = "0" Then
            ObtieneInformacion = True
            MsgBox "La Línea de Crédito Asociada" + Chr(10) + "Aún No ha Sido Cerrada", 0 + 16, _
                   "Problemas de Validación"
          End If
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
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - era: PunteroVistas.Index = PunteroVistas.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  PunteroVistas.Index = BaseEnvioSRM.TableDefs("Vistas").Indexes(0).Name
  PunteroVistas.Seek "=", Vista
  If Not PunteroVistas.NoMatch Then
    If Environ$(AMBIENTEDESARROLLO) = "DESA" Then
      DeterminaNodo = LCase$(PunteroVistas("Vta_NodoDesa"))
    Else
      If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Environ$(NODOREGIONAL))
      Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoProd"))
      End If
    End If
    Aplicacion = PunteroVistas("Vta_Aplicacion")
  End If
  PunteroVistas.Close
  BaseEnvioSRM.Close
End Function
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
