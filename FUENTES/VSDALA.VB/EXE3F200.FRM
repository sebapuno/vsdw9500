VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1800
   ClientTop       =   2265
   ClientWidth     =   3480
   Height          =   2160
   Left            =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3480
   Top             =   1950
   Visible         =   0   'False
   Width           =   3540
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String
Dim VistaOrigen As String
Dim BaseVisado As Database
Dim OperacionBorrado As String

Private Sub Form_Load()
Dim Rut As String * 10
Dim Producto As String
Dim Correlativo As String
Dim Encabezado As String
Dim PunteroVsdPrd As Recordset
Dim NumeroLineaCredito As String
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
    ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
    ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
    PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    IdUsuario = ObtieneRutUsuarioLogueado()
    Encabezado = Mid$(PunteroVsdPrd("Prd_Data"), 1, 6)
    NumeroLineaCredito = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
    Call LocalizaProductoAsociado(Rut, "2FP", NumeroLineaCredito, "22", False)
    Call LocalizaProductoAsociado(Rut, "2FD", NumeroLineaCredito, "22", False)
    Call LocalizaProductoAsociado(Rut, "3F2", NumeroLineaCredito, "11", False)
    Call LocalizaProductoAsociado(Rut, "3F3", NumeroLineaCredito, "11", False)
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
Private Sub ActualizarEnvioSRM(ByVal LlaveEnvio As String, ByVal DataEnvio As String)
Dim BaseSRM As Database
Dim PunteroEnvioSRM As Recordset
  Set BaseSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroEnvioSRM = BaseSRM.OpenRecordset("Mensaje")
  PunteroEnvioSRM.MoveLast
  PunteroEnvioSRM.Edit
  If IsNull(PunteroEnvioSRM("Data")) Then
    PunteroEnvioSRM("Data") = LlaveEnvio + DataEnvio + "##"
  Else
    PunteroEnvioSRM("Data") = PunteroEnvioSRM("Data") + LlaveEnvio + DataEnvio + "##"
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  BaseSRM.Close
End Sub
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
Private Function Interpreta(Cadena As String, DatoBuscado As String) As Integer
Dim PosicionComa As Integer
  PosicionComa = InStr(Cadena, ",")
  If PosicionComa <> 0 Then
    If DatoBuscado = SIMBOLOPOSICION Then
      Interpreta = Val(Mid(Cadena, 1, PosicionComa - 1))
    Else
      Interpreta = Val(Mid(Cadena, PosicionComa + 1, Len(Cadena)))
    End If
  Else
    Interpreta = -1
  End If
End Function
Private Sub ObtenerDatosDeCorte(ByVal Producto As String, IniLlave As Integer, LarLlave As Integer, IniData As Integer, LarData As Integer, VistaDestino As String)
Dim PunteroTabPrd As Recordset
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabPrd.Seek "=", Producto
  IniData = Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOPOSICION)
  IniLlave = Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOPOSICION)
  LarData = Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOLARGO)
  LarLlave = Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOLARGO)
  VistaOrigen = PunteroTabPrd("Prd_Vista_Ori")
  VistaDestino = PunteroTabPrd("Prd_Vista_Dst")
  PunteroTabPrd.Close
End Sub
Private Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
Private Function ActividadProducto(ByVal Producto As String, ByVal Data As String) As String
Dim PunteroTabPrd As Recordset
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabPrd.Seek "=", Producto
  If Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION) = -1 Then
    ActividadProducto = Mid$(PunteroTabPrd("Prd_Pos_Act"), 1, 1)
  Else
    ActividadProducto = Mid$(Data, _
                                              Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION), _
                                              Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOLARGO))
  End If
  PunteroTabPrd.Close
End Function
Private Sub LocalizaProductoAsociado(ByVal RutPatron As String, ByVal Producto As String, LineaBuscada As String, ByVal Posicion As String, ByVal CambiaNumeroLinea As Boolean)
Dim Largo As String
Dim FiltroVsdPrd  As Recordset
  Largo = LTrim$(Str(Len(LineaBuscada)))
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                                                                 "Prd_RutCli='" + RutPatron + "' AND Prd_CodProd='" + Producto + _
                                                                                 "' AND " + "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                                                 LineaBuscada + "';", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    If CambiaNumeroLinea Then
      LineaBuscada = Mid$(FiltroVsdPrd("Prd_Data"), 61, 11)
    Else
      Call GuardarEnEnvioSRM(Producto, FiltroVsdPrd("Prd_Data"))
    End If
  End If
  FiltroVsdPrd.Close
End Sub

Private Function ObtenerOperacionSRM(ByVal Producto As String, ByVal Actividad As String) As String
Dim PunteroTabOper As Recordset
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - era: PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  PunteroTabOper.Index = BaseVisado.TableDefs("Tab_Oper").Indexes(0).Name
  PunteroTabOper.Seek "=", Producto, Actividad
  OperacionBorrado = Trim$(PunteroTabOper("Oper_Borrado"))
  ObtenerOperacionSRM = Trim$(PunteroTabOper("Oper_Srm"))
  PunteroTabOper.Close
End Function
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function

Private Sub GuardarEnEnvioSRM(ByVal Producto As String, ByVal DataProducto As String)
Dim Data As String
Dim Llave As String
Dim InicioData As Integer
Dim InicioLlave As Integer
Dim LargoData As Integer
Dim Operacion As String
Dim LargoLlave As Integer
Dim VistaDestino As String
Dim ActProd As String * 1

  ActProd = ActividadProducto(Producto, DataProducto)
  Operacion = ObtenerOperacionSRM(Producto, ActProd)
  Call ObtenerDatosDeCorte(Producto, InicioLlave, LargoLlave, InicioData, LargoData, VistaDestino)
  Data = Mid$(DataProducto, InicioData, LargoData)
  Llave = Mid$(DataProducto, InicioLlave, LargoLlave)
  Call ActualizarEnvioSRM(IdUsuario + VistaDestino + Operacion + Llave, Data)
  Call BorrarEnVistaOrigen(Producto, DataProducto, ActProd)
  
End Sub
Private Function BorrarEnVistaOrigen(ByVal Producto As String, ByVal Data As String, ByVal Actividad As String) As Boolean
Dim Id As String
Dim DataBorrado    As String
Dim VistaOrigen    As String * 4
Dim PunteroTabPrd  As Recordset
Dim PunteroTabVsd  As Recordset
Dim PunteroTabOper As Recordset
  
  BorrarEnVistaOrigen = True
  
   Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
   PunteroTabPrd.Index = "Tab_Prd"
   
   Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
   ' FI2026 - era: PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
   ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
   ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
   ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
   PunteroTabOper.Index = BaseVisado.TableDefs("Tab_Oper").Indexes(0).Name
   
   Set PunteroTabVsd = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
   ' FI2026 - era: PunteroTabVsd.Index = PunteroTabVsd.Indexes(0).Name
   ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
   ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
   ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
   PunteroTabVsd.Index = BaseVisado.TableDefs("Tab_Vsd").Indexes(0).Name

  PunteroTabPrd.Seek "=", Producto
  VistaOrigen = PunteroTabPrd("PRD_VISTA_ORI")
  PunteroTabOper.Seek "=", Producto, Actividad
  If VistaOrigen <> "XXXX" Then
    PunteroTabVsd.Seek "=", VistaOrigen
    If Not PunteroTabVsd.NoMatch Then
      Id = IdUsuario + VistaOrigen + PunteroTabOper("OPER_BORRADO")
      DataBorrado = Mid$(Data, 1, PunteroTabVsd("Vsd_L_LlaveOri"))
      Call ActualizarEnvioSRM(Id, DataBorrado)
    Else
      BorrarEnVistaOrigen = False
    End If
  End If
  PunteroTabVsd.Close
  PunteroTabPrd.Close
  PunteroTabOper.Close
End Function


