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
Dim IdUsuario As String * 8
Dim DVUsuario As String * 1
Dim DataINI As String
Dim StrEstado As String * 1
Dim Vista As String * 4
Dim Operacion As String * 1
Dim Cantidad As Integer

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
    DVUsuario = Mid$(Rut, 9, 1)
  Else
    ObtieneRutUsuarioLogueado = RUTDEFECTO
    DVUsuario = "9"
  End If
End Function
Private Function ObtenerOperacionBorradoSRM(ByVal producto As String, ByVal Actividad As String) As String
Dim PunteroTabOper As Recordset

  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
  PunteroTabOper.Seek "=", producto, Actividad
  ObtenerOperacionBorradoSRM = LTrim$(PunteroTabOper("Oper_Borrado"))
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
Private Function SacaValorINI(ByVal Grupo As String, ByVal Titulo As String, ByVal NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(800, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub Form_Load()

Dim Rut As String
Dim Correlativo As String
Dim MiSql As String
Dim Ap_VsdPrd As Recordset
Dim Valor As Boolean
Dim Operacion As String
Dim VarSiebel As String

If Not App.PrevInstance Then
  
  IdUsuario = ObtieneRutUsuarioLogueado()
 
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  
  Ap_Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Ap_Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Ap_Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
  VarSiebel = SacaValorINI("SYBASE", "SIEBEL", DIRECTORIO_01 + ARCHIVO_INI01)
  
  Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name
  Ap_VsdPrd.Seek "=", Ap_Rut, Ap_Producto, Ap_Correlativo
  If Not Ap_VsdPrd.NoMatch Then
    DataINI = Trim$(Ap_VsdPrd!PRD_DATA)
  End If
  Ap_VsdPrd.Close
   
  Operacion = Mid$(DataINI, 40, 1)
  
  Cantidad = Val(Mid$(DataINI, 61 + 6, 2))
  
  StrEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
  
  If StrEstado = "1" Then
    StrEstado = "A"
    
    If Operacion = "R" Or VarSiebel = "NORMAL" Then
      Valor = Arma_Data_Aceptada
    Else
      Valor = True
    End If
    
  ElseIf StrEstado = "0" Then
    StrEstado = "R"
    Valor = Arma_Data_Borrado
  End If
  
  If Valor Then
    If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
     MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
    End If
  End If
  BaseEnvioSrm.Close
  BaseVisado.Close
End If

End 'FIN DE LA APLICACIÓN
 
End Sub



Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal Largo As String) As Boolean
Dim FiltroVsdPrd  As Snapshot
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + PatronBusqueda + "';")
  If FiltroVsdPrd.RecordCount > 0 Then
    DataPatron = FiltroVsdPrd("Prd_Data")
    VistaOrigen = FiltroVsdPrd("Prd_VtaOri")
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function



Public Function ObtenerVistaDestino(CodPrd As String) As String

Dim PunteroTabPrd As Recordset

Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
PunteroTabPrd.Index = PunteroTabPrd.Indexes(0).Name
PunteroTabPrd.Seek "=", CodPrd
ObtenerVistaDestino = LTrim$(PunteroTabPrd!Prd_Vista_Dst)
PunteroTabPrd.Close

End Function

Public Function ObtenerOperacionSRM(producto As String, Actividad As String) As String
Dim PunteroTabOper As Recordset

Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
PunteroTabOper.Seek "=", producto, Actividad
ObtenerOperacionSRM = LTrim$(PunteroTabOper!Oper_srm)
PunteroTabOper.Close

End Function






Public Function Arma_Data_Aceptada() As Boolean

Dim Datos As String
Dim i As Integer
Dim j As Integer
Dim VistaDestino As String

For i = 1 To Cantidad
  If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "SC" + Trim$(Str(i)) + Mid$(DataINI, 10, 30), "1", "39") Then
    DataPatron = Mid$(DataPatron, 1, 59 + 8 + 175) + "00" + Mid$(DataPatron, 60 + 8 + 175 + 2)
    If i = 1 Then
      VistaDestino = ObtenerVistaDestino("SC" + Trim$(Str(i)))
      Datos = IdUsuario + VistaDestino + DVUsuario + Mid$(DataPatron, 61, 643 + 8)
    Else
      Datos = Mid$(DataPatron, 61 + 8, 643)
    End If
    Call ActualizarEnvioSRM_2("", Datos) '* AGREGA REGISTROS SIN ## FINAL
  End If
Next i
Call ActualizarEnvioSRM_2("", "##")
  
 ' BORRADO
For j = 1 To 7
  If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "SC" + Trim$(Str(j)) + Mid$(DataINI, 10, 30), "1", "39") Then
    Datos = Mid$(DataPatron, 1, 60)
    'Operacion = ObtenerOperacionBorradoSRM(Ap_Producto, Mid$(DataPatron, 86, 1))
    Operacion = "E"
    Call ActualizarEnvioSRM(IdUsuario + VistaOrigen + Operacion, Datos)
  'Else
  '  Call ActualizarEnvioSRM_2("", "NULL##")
  End If
Next j

Arma_Data_Aceptada = True

End Function

Public Function Arma_Data_Borrado() As Boolean

Dim Datos As String

Arma_Data_Borrado = False

 ' BORRADO
For j = 1 To 7
  If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "SC" + Trim$(Str(j)) + Mid$(DataINI, 10, 30), "1", "39") Then
    Datos = Mid$(DataPatron, 1, 60)
    Operacion = "E"
    Call ActualizarEnvioSRM(IdUsuario + VistaOrigen + Operacion, Datos)
    Arma_Data_Borrado = True
  End If
Next j

End Function

Public Sub ActualizarEnvioSRM_2(ByVal LlaveEnvio As String, ByVal DataEnvio As String)
Dim BaseSRM As Database
Dim PunteroEnvioSRM As Recordset

  Set BaseSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroEnvioSRM = BaseSRM.OpenRecordset("Mensaje")
  PunteroEnvioSRM.MoveLast
  PunteroEnvioSRM.Edit
  If IsNull(PunteroEnvioSRM("Data")) Then
    PunteroEnvioSRM("Data") = LlaveEnvio + DataEnvio
  Else
    PunteroEnvioSRM("Data") = PunteroEnvioSRM("Data") + LlaveEnvio + DataEnvio
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  BaseSRM.Close
  
End Sub
