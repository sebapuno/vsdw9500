VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1800
   ClientTop       =   2265
   ClientWidth     =   3480
   Height          =   2220
   Left            =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3480
   Top             =   1920
   Visible         =   0   'False
   Width           =   3600
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim BaseVisado As Database
Dim IdUsuario As String * 8

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
Dim PunteroVsdPrd As Recordset
Dim ActProd As String * 1
Dim StrEstado  As String
Dim Parametros As String

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    StrEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)

    If StrEstado = "1" Then
      'Aceptado
      Llave = "b12256781234W00000010plt_corp  exec sp_genera_maci_vigente "
    Else
       'Rechazado
      Llave = "b12256781234W00000010plt_corp  exec sp_rechaza_maci "
    End If
    
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Call ObtenerDatosDeCorte(Producto, InicioLlave, LargoLlave, InicioData, LargoData, Vista)
    Data = Mid$(PunteroVsdPrd("Prd_Data"), InicioData, LargoData)
    
    Nro_Rut = Mid$(PunteroVsdPrd("Prd_Data"), 61, 10)
    Nro_Correl = Mid$(PunteroVsdPrd("Prd_Data"), 71, 3)

    Parametros = Nro_Grupo + "," + Nro_Correl
    
    Call ActualizarEnvioSRM(Llave, Data)
    'Call BorrarEnVistaOrigen(Producto, PunteroVsdPrd("Prd_Data"), ActProd)

    If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
      MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
    End If
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
Private Function BorrarEnVistaOrigen(ByVal Producto As String, ByVal Data As String, ByVal Actividad As String) As Boolean
Dim Id As String
Dim DataBorrado As String
  BorrarEnVistaOrigen = True
  
  Id = "b12256781234W0000001Nplt_emp01 exec modifica_estado_empresa "
  DataBorrado = "'" + Mid$(Data, 1, 3) + "'," + "'" + Mid$(Data, 4, 3) + "'," _
                + "'" + Mid$(Data, 7, 3) + "'," + "'" + Mid$(Data, 10, 30) + "'," _
                + "'" + Mid$(Data, 40, 1) + "'," + "'" + Mid$(Data, 41, 2) + "'," _
                + "'" + Mid$(Data, 43, 8) + "'," + "'" + Mid$(Data, 51, 10) + "'"
  Call ActualizarEnvioSRM(Id, DataBorrado)
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
Private Function ObtenerOperacionSRM(ByVal Producto As String, ByVal Actividad As String) As String
Dim PunteroTabOper As Recordset
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
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
