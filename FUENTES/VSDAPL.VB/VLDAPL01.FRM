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
Dim Correlativo As String
Dim BaseVisado As Database
Dim DatoProducto As String * 11
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
    
    DatoProducto = Mid$(PunteroVsdPrd("Prd_Data"), 21, 11)
    If ValCuentaCorriente(DatoProducto, PunteroTabPrd("Prd_Vista_Dst")) Then
       ' Cambios sobre la cuenta nueva
       CadenaDatos = Mid$(CadenaDatos, 1, 97) + "0" + Format(Day(Date), "00") + Format(Month(Date), "00") + Mid$(Format(Year(Date), "0000"), 3, 2) + _
                     Mid$(CadenaDatos, 105, 47) + "1" + Mid$(CadenaDatos, 153)
       CadenaDatos = Mid$(PunteroVsdPrd("Prd_Data"), 1, 119) + Mid$(CadenaDatos, 12)
       
       PunteroVsdPrd.Edit
       PunteroVsdPrd("Prd_Data") = CadenaDatos
       PunteroVsdPrd.Update
       ' Cambios sobre la cuenta antigua
       DatoProducto = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
       If ValCuentaCorriente(DatoProducto, PunteroTabPrd("Prd_Vista_Dst")) Then
          CadenaDatos = Mid$(CadenaDatos, 1, 149) + "2" + Mid$(CadenaDatos, 151, 1) + "85" + Mid$(CadenaDatos, 154)
          CadenaDatos = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6) + "BCA" + Mid$(PunteroVsdPrd("Prd_Data"), 10, 51) + Mid$(CadenaDatos, 12)
          PunteroVsdPrd.Seek "=", Rut, "BCA", Val(Correlativo)
          If PunteroVsdPrd.NoMatch Then
            PunteroVsdPrd.AddNew
            PunteroVsdPrd("Prd_CodProd") = "BCA"
            PunteroVsdPrd("Prd_RutCli") = Rut
            PunteroVsdPrd("Prd_Correl") = Val(Correlativo)
            PunteroVsdPrd("Prd_Data") = CadenaDatos
            PunteroVsdPrd.Update
          End If
          '** Si Tiene Producto Asociado "LGL"
          Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE prd_rutcli='" + Rut + "' AND " + _
                      "Mid$(prd_data, 1, 40 )='" + Encabezado + "LGL" + Mid$(dataini, 10, 31) + "' ;", dbOpenSnapshot)
          If FiltroVsdPrd.RecordCount > 0 Then
              If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
                   MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                                "Error de Escritura"
              End If
          Else
              MsgBox "No se encuentra dirección legal para Cuenta Corriente Remunerada.", 16, "Faltan Datos"
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
Private Function ValCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
Dim ExisteCuentaCorriente As Boolean

  Modo = "L"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, Llave, True)
  If ExisteCuentaCorriente Then
     CadenaDatos = Mid$(CadenaDatos, 7)
     ValCuentaCorriente = True
  Else
     ValCuentaCorriente = False
  End If
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
    Else
       'If Mid$(CadenaDatos, 1, 6) <> "000096" Then
          ObtieneInformacion = False
       'End If
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

