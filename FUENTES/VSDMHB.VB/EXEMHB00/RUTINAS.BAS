Attribute VB_Name = "RutinasVSDW"
Option Explicit

Public Function SacaValorINI(ByVal grupo As String, ByVal Titulo As String, ByVal NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
Dim nSize%

    lpReturnedString = String$(800, 0)
    nSize% = Len(lpReturnedString)
    ValorGet = GetPrivateProfileString(grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
    SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Public Function GrabaParametro(grupo As String, llave As String, parametro As String, archivo As String) As Boolean
    Dim f As Integer, lpReturnedString As String
    
    lpReturnedString = String$(255, 0)
    parametro = parametro + Mid(lpReturnedString, Len(parametro) + 2)
    f = WritePrivateProfileString(ByVal grupo, ByVal llave, ByVal parametro, ByVal archivo)
    GrabaParametro = True

End Function
Public Function ObtieneRutUsuarioLogueado() As String
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
End Function

Public Function ObtenerVistaDestino(CodPrd As String)
Dim PunteroTabprd As Recordset
  
  Set PunteroTabprd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  ' FI2026 - PunteroTabprd.Index = PunteroTabprd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabprd.Seek "=", CodPrd
  ObtenerVistaDestino = Trim$(PunteroTabprd!Prd_Vista_Dst)
  PunteroTabprd.Close
End Function

Public Function ObtenerOperacionSRM(Producto As String, Actividad As String) As String
Dim PunteroTabOper As Recordset
  
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabOper.Seek "=", Producto, Actividad
  ObtenerOperacionSRM = Trim$(PunteroTabOper!Oper_SRM)
  PunteroTabOper.Close
End Function

Public Sub ActualizarEnvioSRM(ByVal LlaveEnvio As String, ByVal DataEnvio As String)
Dim PunteroEnvioSRM As Recordset
  
  Set PunteroEnvioSRM = BaseEnvioSrm.OpenRecordset("Mensaje")
  PunteroEnvioSRM.MoveLast
  PunteroEnvioSRM.Edit
  If IsNull(PunteroEnvioSRM("Data")) Then
    PunteroEnvioSRM("Data") = LlaveEnvio + DataEnvio + "##"
  Else
    PunteroEnvioSRM("Data") = PunteroEnvioSRM("Data") + LlaveEnvio + DataEnvio + "##"
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  
End Sub

Public Function ObtenerOperacionBorrado(Producto As String, Actividad As String) As String
Dim PunteroTabOper As Recordset
  
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabOper.Seek "=", Producto, Actividad
  ObtenerOperacionBorrado = Trim$(PunteroTabOper!Oper_Borrado)
  PunteroTabOper.Close
  
End Function

Public Function ObtenerVistaOrigen(CodPrd As String)
Dim PunteroTabprd As Recordset
  
Set PunteroTabprd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  
' FI2026 - PunteroTabprd.Index = PunteroTabprd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
PunteroTabprd.Seek "=", CodPrd
ObtenerVistaOrigen = Trim$(PunteroTabprd!Prd_Vista_Ori)
PunteroTabprd.Close
  
End Function
