VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1650
   ClientLeft      =   1665
   ClientTop       =   2460
   ClientWidth     =   3390
   Height          =   2055
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3390
   Top             =   2115
   Visible         =   0   'False
   Width           =   3510
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim BaseVisado As Database
Dim IdUsuario As String * 8

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
Private Function BorrarEnVistaOrigen(ByVal Producto As String, ByVal Data As String, ByVal Actividad As String, VistaBase As String) As Boolean
Dim Id As String
Dim DataBorrado    As String
Dim VistaOrigen    As String * 4
Dim PunteroTabPrd  As Recordset
Dim PunteroTabVsd  As Recordset
Dim PunteroTabOper As Recordset
Dim OperBorrado As String

  BorrarEnVistaOrigen = True
  
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabOper.Seek "=", Producto, Actividad
  OperBorrado = PunteroTabOper("OPER_BORRADO")
  PunteroTabOper.Close
 
  Set PunteroTabVsd = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
  ' FI2026 - PunteroTabVsd.Index = PunteroTabVsd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  
  If Trim$(VistaBase) <> "" Then
    VistaOrigen = VistaBase
  Else
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    PunteroTabPrd.Index = "Tab_Prd"
    PunteroTabPrd.Seek "=", Producto
    VistaOrigen = PunteroTabPrd("PRD_VISTA_ORI")
    PunteroTabPrd.Close
  End If
    
  '***
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSrm.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroVistas.Seek "=", VistaOrigen
  If PunteroVistas("Vta_OperBorr") <> "-" Then
    OperBorrado = PunteroVistas("Vta_OperBorr")
  End If
  PunteroVistas.Close
  BaseEnvioSrm.Close
  '***
 
  If VistaOrigen <> "XXXX" Then
    PunteroTabVsd.Seek "=", VistaOrigen
    If Not PunteroTabVsd.NoMatch Then
      Id = IdUsuario + VistaOrigen + OperBorrado
      DataBorrado = Mid$(Data, 1, PunteroTabVsd("Vsd_L_LlaveOri"))
      'Call Graba_SRM("", Id, DataBorrado, "", ModoIngreso)
      Call ActualizarEnvioSRM(Id, DataBorrado)

    Else
      BorrarEnVistaOrigen = False
    End If
  End If
  PunteroTabVsd.Close
  
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
  VistaDestino = PunteroTabPrd("Prd_Vista_Dst")
  PunteroTabPrd.Close
End Sub
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
Private Function ObtenerOperacionSRM(ByVal Producto As String, ByVal Actividad As String) As String
Dim PunteroTabOper As Recordset
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabOper.Seek "=", Producto, Actividad
  ObtenerOperacionSRM = LTrim$(PunteroTabOper("Oper_Srm"))
  PunteroTabOper.Close
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
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function SacaValorINI(ByVal Grupo As String, ByVal Titulo As String, ByVal NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
Private Sub Form_Load()
Dim Rut As String
Dim Data As String
Dim Vista As String
Dim Llave As String
Dim Producto As String
Dim InicioData As Integer
Dim InicioLlave As Integer
Dim Operacion As String
Dim LargoData As Integer
Dim Correlativo As String
Dim LargoLlave As Integer
Dim ActProd As String * 1
Dim PunteroVsdPrd As Recordset
  
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Call ObtenerDatosDeCorte(Producto, InicioLlave, LargoLlave, InicioData, LargoData, Vista)
    Data = Mid$(PunteroVsdPrd("Prd_Data"), InicioData, LargoData)
    Llave = Mid$(PunteroVsdPrd("Prd_Data"), InicioLlave, LargoLlave)
    
    Data = Llave + Data
    ' * fecha cierre
    Data = Mid$(Data, 1, 171) + Format$(Now, "ddmmyyyy") + Mid$(Data, 180)
    ' * fecha ult homolog
    Data = Mid$(Data, 1, 287) + Format$(Now, "ddmmyyyy") + Mid$(Data, 296)
        
    IdUsuario = ObtieneRutUsuarioLogueado()
    ActProd = ActividadProducto(Producto, PunteroVsdPrd("Prd_Data"))
    Operacion = ObtenerOperacionSRM(Producto, ActProd)
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, Data)
    Call BorrarEnVistaOrigen(Producto, PunteroVsdPrd("Prd_Data"), ActProd, PunteroVsdPrd("Prd_VtaOri"))
    If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
      MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
    End If
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
