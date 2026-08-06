Attribute VB_Name = "RutinasVSDW"
Option Explicit

Public Function DatoClienteEncontrado(pRutcliente As String) As String

Dim FiltroVsdCli  As Recordset
Dim SQL           As String

Set BaseVisado = Workspaces(0).OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
SQL = "SELECT Cli_Nombre FROM Vsd_Cli WHERE CLI_RUT='" + pRutcliente + "';"
                                               
Set FiltroVsdCli = BaseVisado.OpenRecordset(SQL)
If FiltroVsdCli.RecordCount > 0 Then
  DatoClienteEncontrado = Left$(FiltroVsdCli!Cli_Nombre + Space(35), 35)
Else
  DatoClienteEncontrado = Space(35)
End If
FiltroVsdCli.Close

End Function
Public Function TraeRutEjecutivo(Cod_Ejc As String) As String

Dim BaseSGT As Database
Dim FiltroTabEjc  As Recordset
Dim SQL           As String
Dim CODOFI As String
Dim CODEJC As String

CODOFI = Mid(Cod_Ejc, 1, 3)
CODEJC = Mid(Cod_Ejc, 4, 3)

Set BaseSGT = Workspaces(0).OpenDatabase(DIRECTORIO_02 + ARCHIVO_MDB03)
SQL = "SELECT EJC_EJCRUT FROM TBLEJC WHERE EJC_EJCOFI=" + CODOFI + " AND EJC_EJCCOD=" + CODEJC + ";"
                                               
Set FiltroTabEjc = BaseSGT.OpenRecordset(SQL, dbOpenSnapshot)
If FiltroTabEjc.RecordCount > 0 Then
  TraeRutEjecutivo = Left$(FiltroTabEjc!EJC_EJCRUT + Space(10), 10)
Else
  TraeRutEjecutivo = Space(10)
End If
FiltroTabEjc.Close

BaseSGT.Close

End Function

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

Public Sub ActualizarEnvioSRM_Directo(ByVal LlaveEnvio As String, ByVal DataEnvio As String)
Dim PunteroEnvioSRM As Recordset
  
  Set PunteroEnvioSRM = BaseEnvioSrm.OpenRecordset("Mensaje")
  PunteroEnvioSRM.MoveLast
  PunteroEnvioSRM.Edit
  If IsNull(PunteroEnvioSRM("Directo")) Then
    PunteroEnvioSRM("Directo") = LlaveEnvio + DataEnvio + "##"
  Else
    PunteroEnvioSRM("Directo") = PunteroEnvioSRM("Directo") + LlaveEnvio + DataEnvio + "##"
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  
End Sub

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
ObtenerVistaOrigen = Trim$(PunteroTabprd!Prd_Vista_Ori, dbOpenSnapshot)
PunteroTabprd.Close
  
End Function
