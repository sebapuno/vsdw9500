Attribute VB_Name = "Modulo1"
'Private ' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer

Global BaseVisado As Database
Global BaseEnvioSRM As Database
Global BaseVldCta As Database

Global IdUsuario As String
Global CadenaDatos As String
Global GReintento As Boolean
Global Terminar As Boolean
Public Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function


Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function


Public Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
  
  DeterminaNodo = SINNODO
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroVistas.Seek "=", Vista
  If Not PunteroVistas.NoMatch Then
    If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Frm_vsd.tx_Nodo.Text)
    Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoProd"))
    End If
    Aplicacion = PunteroVistas("Vta_Aplicacion")
  End If
  PunteroVistas.Close
End Function
